import {XML,MIME,limits,xml,parse,children,first,all,bookReader,checkBounds,address,rowsOf,stylesOf,errorValue} from './xml.js';
import {makeZip} from './package.js';
const ns={};for(const [prefix,name] of Object.entries({office:'office',table:'table',style:'style',text:'text',number:'datastyle',fo:'xsl-fo-compatible',manifest:'manifest'}))ns[prefix]='urn:oasis:names:tc:opendocument:xmlns:'+name+':1.0';
ns.of='urn:oasis:names:tc:opendocument:xmlns:of:1.2';
const declarations=Object.entries(ns).map(([k,v])=>' xmlns:'+k+'="'+v+'"').join('');
const a=(el,p,attr)=>el?.getAttributeNS(ns[p],attr)||'';
const styleName=n=>a(n,'style','name');
const cellFormatLoss='Additional number formats, font settings, borders and other table styles are simplified.';
function repeated(n,key){const raw=a(n,'table',key);if(!raw)return 1;const v=Number(raw);if(!Number.isSafeInteger(v)||v<1||v>1048576)throw Error('Invalid or oversized ODS repeat count.');return v}
function paragraph(n,warn){
  let s='';
  function walk(el){
    for(const c of el.childNodes){
      if(c.nodeType===3||c.nodeType===4)s+=c.nodeValue;
      else if(c.localName==='s'){
        const n=Number(a(c,'text','c')||1);if(!Number.isSafeInteger(n)||n<1||s.length+n>32767)throw Error('A text cell exceeds the supported length.');
        s+=' '.repeat(n);
      }else if(c.localName==='line-break')s+='\n';else if(c.localName==='tab')s+='\t';else if(c.nodeType===1){warn('Inline text formatting and fields are reduced to text.');walk(c)}
      if(s.length>32767)throw Error('A text cell exceeds 32,767 characters.');
    }
  }walk(n);return s
}
function formulaIn(f){
  if(!/^(of|oooc):=/.test(f))return f;
  return f.replace(/^(of|oooc):=/,'').split(/("(?:[^"]|"")*")/g).map((part,i)=>i%2?part:part.replace(/\[\.(\$?[A-Z]+\$?\d+)(?::\.(\$?[A-Z]+\$?\d+))?\]/gi,(_,start,end)=>start+(end?':'+end:'')).replaceAll(';',',')).join('');
}
function formulaOut(f){
  return 'of:='+f.split(/("(?:[^"]|"")*")/g).map((part,i)=>i%2?part:part.replace(/(?<![\w.$])(\$?[A-Z]+\$?\d+)(?::(\$?[A-Z]+\$?\d+))?(?![\w.])/gi,(_,start,end)=>'[.'+start+(end?':.'+end:'')+']').replaceAll(',',';')).join('');
}
function collectStyles(docs,warn){
  const data=new Map(),styles=new Map(),live=new Set(),cache=new Map();
  for(const doc of docs){
    for(const s of all(doc,'style'))if(a(s,'style','family')==='table-cell')styles.set(styleName(s),s);
    for(const s of [...doc.getElementsByTagNameNS(ns.number,'*')]){
      if(!styleName(s))continue;
      const n=first(s,'number'),places=a(n,'number','decimal-places'),grouping=a(n,'number','grouping')==='true';
      let fmt;
      if(s.localName==='number-style'&&children(s).length===1&&!grouping)fmt=places==='2'?'number':!places?'general':null;
      if(s.localName==='percentage-style'&&places==='2'&&first(s,'text')?.textContent==='%')fmt='percent';
      if(s.localName==='currency-style'&&places==='2'&&first(s,'currency-symbol')?.textContent==='$')fmt='currency';
      data.set(styleName(s),fmt);
    }
  }
  function get(name){
    if(!name)return {};
    if(cache.has(name))return cache.get(name);
    if(live.has(name))throw Error('Cyclic ODS style inheritance.');
    if(live.size>64)throw Error('ODS style inheritance exceeds the supported depth.');
    const n=styles.get(name);if(!n){warn(cellFormatLoss);return {}}
    live.add(name);const s={...get(a(n,'style','parent-style-name'))},fmt=a(n,'style','data-style-name');
    if(fmt){const known=data.get(fmt);if(known==='general')delete s.fmt;else if(known)s.fmt=known;else warn(cellFormatLoss)}
    for(const p of children(n))for(const att of p.attributes){
      const key=att.localName,val=att.value;
      if(key==='font-weight')s.bold=val==='bold';
      else if(key==='font-style')s.italic=val==='italic';
      else if(key==='text-align'){const align={start:'left',end:'right',left:'left',right:'right',center:'center'}[val];if(align)s.align=align;else warn(cellFormatLoss)}
      else if(key==='color'||key==='background-color'){
        if(/^#[0-9a-f]{6}$/i.test(val))s[key==='color'?'color':'fill']=val.toLowerCase();
        else if(val!=='transparent'&&val!=='automatic')warn(cellFormatLoss);
      }else if(!['font-name','font-size','font-name-asian','font-name-complex','font-size-asian','font-size-complex','font-style-asian','font-style-complex','font-weight-asian','font-weight-complex','vertical-align'].includes(key))warn(cellFormatLoss);
    }
    live.delete(name);cache.set(name,s);return s
  }
  return get
}
export async function readOds(zip){
  const ctx=bookReader(),{warn}=ctx;
  const root=parse(await zip.text('content.xml'));
  const body=first(first(root,'body'),'spreadsheet');if(!body||body.namespaceURI!==ns.office)throw Error('The package is not an OpenDocument spreadsheet.');
  const docs=zip.names.includes('styles.xml')?[parse(await zip.text('styles.xml')),root]:[root];
  const getStyle=collectStyles(docs,warn);
  const ordinary=new Set(['mimetype','content.xml','styles.xml','meta.xml','settings.xml','META-INF/manifest.xml','Thumbnails/thumbnail.png','manifest.rdf']);
  for(const name of zip.names){
    if(ordinary.has(name)||name.endsWith('/'))continue;
    if(/script|basic|macro/i.test(name))warn('Macros and scripts are omitted.');
    else if(/signature/i.test(name))warn('Digital signatures are omitted from exported copies.');
    else warn('Embedded objects, images or additional package parts are omitted.');
  }
  if(zip.names.includes('META-INF/manifest.xml')){
    const manifest=parse(await zip.text('META-INF/manifest.xml'));
    if(all(manifest,'encryption-data').length)throw Error('Encrypted OpenDocument workbooks are unsupported.');
    for(const part of all(manifest,'file-entry')){
      const mime=a(part,'manifest','media-type');
      if(/script|basic|macro|executable/i.test(mime))warn('Macros and scripts are omitted.');
      if(/chart|image|oleobject/i.test(mime))warn('Embedded objects, images or additional package parts are omitted.');
    }
  }
  const omissions={
    'data-pilot-tables':'Pivot tables are omitted.', 'data-pilot-table':'Pivot tables are omitted.',
    'named-expressions':'Named ranges and expressions are omitted.', 'database-ranges':'Database ranges and filters are omitted.',
    'content-validations':'Data validation is omitted.', 'tracked-changes':'Tracked changes are omitted.',
    'shapes':'Drawings, charts and images are omitted.', 'frame':'Drawings, charts and images are omitted.',
    'annotation':'Annotations are omitted.', 'a':'Hyperlink targets are omitted.', 'table-source':'External table links are omitted.',
    'cell-range-source':'External ranges are omitted.', 'event-listeners':'Scripts and event listeners are omitted.',
    'conditional-formats':'Conditional formatting is omitted.', 'covered-table-cell':'Merged cells are unmerged.',
    'scenario':'Sheet scenarios are omitted.', 'detective':'Formula auditing information is omitted.',
    'object':'Drawings, charts and images are omitted.', 'object-ole':'Embedded objects are omitted.',
    'scripts':'Macros and scripts are omitted.', 'forms':'Forms and controls are omitted.',
    'data-pilot-source-cell-range':'Pivot tables are omitted.'
  };
  for(const e of root.getElementsByTagNameNS('*','*')){
    if(omissions[e.localName])warn(omissions[e.localName]);
    if(e.hasAttributeNS('http://www.w3.org/1999/xlink','href'))warn('Linked content and hyperlink targets are omitted.');
  }
  for(const doc of docs){
    for(const el of all(doc,'table-properties'))if(a(el,'table','display')==='false')warn('Hidden sheets become visible.');
  }
  if(children(body).some(el=>!['table','calculation-settings'].includes(el.localName)))
    warn('Additional spreadsheet structures or extension data are omitted.');
  const tables=children(body,'table');
  for(const t of tables){
    const sheet=ctx.addSheet(a(t,'table','name'));
    if(a(t,'table','protected')==='true')warn('Sheet protection is omitted.');
    if(a(t,'table','display')==='false')warn('Hidden sheets become visible.');
    let row=1;
    // The physical row/column structure can include enormous repeated empty tails.
    const visitRows=node=>{
      for(const el of children(node)){
        if(['table-header-rows','table-rows','table-row-group'].includes(el.localName)){warn('Row groups or repeated print headers are omitted.');visitRows(el);continue}
        if(el.localName!=='table-row'){
          if(!['table-column','table-column-group','table-columns','table-header-columns'].includes(el.localName))warn('Additional sheet structures, objects or extension data are omitted.');
          continue;
        }
        if(a(el,'table','visibility')&&a(el,'table','visibility')!=='visible')warn('Hidden rows become visible.');
        const nr=repeated(el,'number-rows-repeated');let col=1;
        for(const c of children(el)){
          if(!['table-cell','covered-table-cell'].includes(c.localName)){warn('Additional row data or extension records are omitted.');continue}
          if(children(c).some(n=>n.localName!=='p')||Array.from(c.attributes).some(v=>v.namespaceURI===ns.table&&!['formula','style-name','number-columns-repeated'].includes(v.localName)))warn('Additional cell markup, attributes or embedded content are omitted.');
          const nc=repeated(c,'number-columns-repeated'), f=a(c,'table','formula'),type=a(c,'office','value-type');
          if(a(c,'table','number-matrix-columns-spanned')||a(c,'table','number-matrix-rows-spanned'))throw Error('Matrix formulas are unsupported. The original ODS is unchanged.');
          if(a(c,'table','number-columns-spanned')||a(c,'table','number-rows-spanned'))warn('Merged cells are unmerged.');
          if(f&&(nr>1||nc>1))throw Error('Repeated formula cells require per-cell translation and are unsupported.');
          let v='';
          if(f)v=''; // Recalculate in the editor. Never import a cached formula result as data.
          else if(['float','percentage','currency'].includes(type)){
            const val=a(c,'office','value');v=Number(val);if(!val||!Number.isFinite(v))throw Error('Invalid numeric value in ODS.');
          }else if(type==='boolean'){const val=a(c,'office','boolean-value');if(!['true','false','1','0'].includes(val))throw Error('Invalid Boolean value in ODS.');v=['1','true'].includes(val)}
          else {
            const paras=children(c,'p');v=a(c,'office','string-value')||paras.map(p=>paragraph(p,warn)).join('\n');
            if(type&&!['string','void'].includes(type)){warn('Date, time or other unsupported cell types become text.');v=v||a(c,'office','date-value')||a(c,'office','time-value')}
          }
          const style=getStyle(a(c,'table','style-name')),specific=Object.entries(style).some(([key,val])=>!(key==='color'&&val==='#000000')&&val);
          if(f||type||v!==''||specific){
            checkBounds(row+nr-1,col+nc-1);
            if(nr*nc>limits.cells)throw Error('Repeated cells exceed the workbook limit.');
            for(let dr=0;dr<nr;dr++)for(let dc=0;dc<nc;dc++){
              const cell={v};if(f)cell.f=formulaIn(f);if(Object.keys(style).length)cell.style={...style};ctx.put(sheet,address(row+dr,col+dc),cell)
            }
          }
          col+=nc;
        }
        row+=nr;
      }
    };
    visitRows(t);
  }
  warn('Document metadata, view settings, print layouts, row and column styles are rebuilt. Only documented cell formats are preserved.');
  return ctx.result();
}
function numberStyles(){
  const num='<number:number number:decimal-places="2" number:min-decimal-places="2" number:min-integer-digits="1"/>';
  return '<number:number-style style:name="number">'+num+'</number:number-style><number:percentage-style style:name="percent">'+num+'<number:text>%</number:text></number:percentage-style><number:currency-style style:name="currency"><number:currency-symbol number:language="en" number:country="US">$</number:currency-symbol>'+num.replace('/>',' number:grouping="true"/>')+'</number:currency-style>'
}
function para(value){
  return '<text:p>'+String(value).split(/( +|\r\n|\n|\r|\t)/).map(s=>s[0]===' '?'<text:s text:c="'+s.length+'"/>':s==='\t'?'<text:tab/>':/^[\r\n]/.test(s)?'<text:line-break/>':xml(s)).join('')+'</text:p>';
}
export function writeOds(book,getValue){
  const {styles,id}=stylesOf(book);
  const cellStyles=styles.map((s,i)=>{
    let props='';
    if(s.fill)props+='<style:table-cell-properties fo:background-color="'+xml(s.fill)+'"/>';
    if(s.align)props+='<style:paragraph-properties fo:text-align="'+xml(s.align)+'"/>';
    if(s.color||s.bold||s.italic)props+='<style:text-properties'+(s.color?' fo:color="'+xml(s.color)+'"':'')+(s.bold?' fo:font-weight="bold"':'')+(s.italic?' fo:font-style="italic"':'')+'/>';
    return '<style:style style:name="c'+i+'" style:family="table-cell"'+(s.fmt?' style:data-style-name="'+xml(s.fmt)+'"':'')+'>'+props+'</style:style>';
  });
  const tables=book.sheets.map((sheet,index)=>{
    const rows=[];let lastRow=0;
    for(const [r,items] of rowsOf(sheet)){
      if(r>lastRow+1)rows.push('<table:table-row table:number-rows-repeated="'+(r-lastRow-1)+'"><table:table-cell/></table:table-row>');lastRow=r;
      const cells=[];let lastCol=0;
      for(const {c,addr,cell} of items){
        if(c>lastCol+1)cells.push('<table:table-cell table:number-columns-repeated="'+(c-lastCol-1)+'"/>');lastCol=c;
        const value=cell.f?getValue(index,addr):cell.v;
        let att=' table:style-name="c'+id(cell.style)+'"';
        if(cell.f)att+=' table:formula="'+xml(formulaOut(cell.f))+'"';
        if(typeof value==='boolean')att+=' office:value-type="boolean" office:boolean-value="'+value+'"';
        else if(typeof value==='number'){
          const type=({percent:'percentage',currency:'currency'})[cell.style?.fmt]||'float';
          att+=' office:value-type="'+type+'" office:value="'+xml(value)+'"'+(type==='currency'?' office:currency="USD"':'');
        }else att+=' office:value-type="string"'+(value===''?' office:string-value=""':'');
        const display=typeof value==='boolean'?String(value).toUpperCase():value??'';
        cells.push('<table:table-cell'+att+'>'+para(display)+'</table:table-cell>');
      }
      rows.push('<table:table-row>'+cells.join('')+'</table:table-row>');
    }
    return '<table:table table:name="'+xml(sheet.name)+'"><table:table-column/>'+(rows.join('')||'<table:table-row><table:table-cell/></table:table-row>')+'</table:table>';
  });
  const content=XML+'<office:document-content'+declarations+' office:version="1.3"><office:automatic-styles>'+numberStyles()+cellStyles.join('')+'</office:automatic-styles><office:body><office:spreadsheet>'+tables.join('')+'</office:spreadsheet></office:body></office:document-content>';
  const manifest=XML+'<manifest:manifest xmlns:manifest="'+ns.manifest+'" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="'+MIME.ods+'"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>';
  return makeZip([['mimetype',MIME.ods],['content.xml',content],['META-INF/manifest.xml',manifest]],MIME.ods);
}
