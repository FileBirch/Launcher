import {XML, MS, REL, PKG, MIME, parse, xml, children, first, all, attr, address, coords, rowsOf, stylesOf, isSheetNameExcel, errorValue, bookReader} from './xml.js';
import {makeZip} from './package.js';
const GLOSS = 'Package metadata, custom themes, view and print settings are rebuilt on export. Only the documented cell formats are preserved.';
const STYLELOSS = 'Some fonts, borders, number formats or other advanced cell styles are simplified.';
// Inspect package declarations as well as filenames. Producers may use arbitrary part names.
function partLoss(type){
  if(/macroenabled|vba|macrosheet|activex|oleobject|embedded(package|object)|ctrlprop|customui/i.test(type))
    return 'Macros, active content and embedded objects are omitted.';
  if(/pivot|slicer|timeline/i.test(type))return 'Pivot tables, caches, slicers and timelines are omitted.';
  if(/externallink|connection|querytable|hyperlink|webextension/i.test(type))return 'External links, connections and queries are omitted.';
  if(/drawing|chart|image|diagram|vml/i.test(type))return 'Drawings, charts and images are omitted.';
  if(/comment|person/i.test(type))return 'Comments and annotations are omitted.';
  if(/\.table\+xml|\/table$/i.test(type))return 'Structured tables are omitted.';
  if(/signature/i.test(type))return 'Digital signatures are omitted from exported copies.';
  return '';
}
function inspectChildren(node,allowed,warn,label){
  if(children(node).some(c=>!allowed.includes(c.localName)))warn(label);
}
function partName(base,target){
  if(!target || /^[a-z][\w+.-]*:/i.test(target)||target.startsWith('//')||/[?#\\]/.test(target))throw Error('Unsupported package relationship target.');
  return decodeURIComponent(new URL(target,'https://package.invalid/'+base).pathname).slice(1);
}
function stringValue(n){return (n?.textContent||'').replace(/_x([0-9a-f]{4})_/gi,(_,hex)=>String.fromCharCode(parseInt(hex,16)))}
function richText(n,warn){
  if(!n)return '';
  if(children(n).some(c=>c.localName!=='t'))warn('Rich text, phonetic annotations or inline markup become plain text.');
  return children(n).map(c=>c.localName==='t'?stringValue(c):c.localName==='r'?stringValue(first(c,'t')):'').join('');
}
function readStyles(root,warn){
  const fonts=children(first(root,'fonts'),'font'), fills=children(first(root,'fills'),'fill'),nums=new Map(children(first(root,'numFmts'),'numFmt').map(n=>[attr(n,'numFmtId'),attr(n,'formatCode')]));
  const codes={'0.00':'number','0.00%':'percent','$#,##0.00':'currency','"$"#,##0.00':'currency'};
  function color(n){
    const rgb=attr(n,'rgb');
    if(/^(FF)?[A-F\d]{6}$/i.test(rgb))return '#'+rgb.slice(-6).toLowerCase();
    if(n&&!(attr(n,'theme')==='1'&&!attr(n,'tint'))&&attr(n,'auto')!=='1')warn(STYLELOSS);
    return '';
  }
  return children(first(root,'cellXfs'),'xf').map(x=>{
    const s={},font=fonts[Number(attr(x,'fontId'))],fill=first(fills[Number(attr(x,'fillId'))],'patternFill'),a=first(x,'alignment');
    for(const [tag,k] of [['b','bold'],['i','italic']]){const el=first(font,tag);if(el&&!['0','false'].includes(attr(el,'val')))s[k]=true}
    if(children(font).some(n=>!['b','i','color','sz','name','family','scheme','charset'].includes(n.localName))||Number(attr(x,'borderId'))||Number(attr(x,'xfId'))||first(x,'protection'))warn(STYLELOSS);
    if(fonts[0]&&font&&['sz','name','family','scheme','charset'].some(k=>attr(first(font,k),'val')!==attr(first(fonts[0],k),'val')))warn(STYLELOSS);
    const fg=color(first(font,'color'));if(fg)s.color=fg;
    if(attr(fill,'patternType')==='solid'){const bg=color(first(fill,'fgColor'));if(bg)s.fill=bg;else warn(STYLELOSS)}
    else if(fill&&!['none','gray125',''].includes(attr(fill,'patternType'))||Number(attr(x,'fillId'))>1&&!fill)warn(STYLELOSS);
    if(a){const horizontal=attr(a,'horizontal');if(['left','center','right'].includes(horizontal))s.align=horizontal;
      if(Array.from(a.attributes).some(v=>v.localName!=='horizontal')||(horizontal&&!s.align))warn(STYLELOSS)}
    const fmt=attr(x,'numFmtId')||'0', known=({'2':'number','10':'percent'})[fmt]||codes[nums.get(fmt)];
    if(known)s.fmt=known;else if(fmt!=='0')warn(STYLELOSS);
    return s;
  });
}
export async function readXlsx(zip){
  const ctx=bookReader(),{warn}=ctx;
  const rels=new Map();
  const read=async name=>parse(await zip.text(name));
  const types=await read('[Content_Types].xml');
  for(const node of children(types)){
    const category=partLoss(attr(node,'ContentType'));
    if(category)warn(category);
  }
  for(const name of zip.names){
    if(name.endsWith('/'))continue;
    if(name.endsWith('.rels')){
      const map=new Map();
      for(const r of children(await read(name),'Relationship')){
        const external=attr(r,'TargetMode').toLowerCase()==='external';
        if(external)warn('External links and hyperlink targets are omitted.');
        const relation=attr(r,'Type'),category=partLoss(relation);
        if(category)warn(category);
        else if(!['officeDocument','worksheet','styles','sharedStrings','theme','calcChain','core-properties','extended-properties'].includes(relation.split('/').pop()))
          warn('Other package relationships, custom properties or extension data are omitted.');
        map.set(attr(r,'Id'),{type:attr(r,'Type').split('/').pop(),target:attr(r,'Target'),external});
      }
      rels.set(name,map);
    }
  }
  const office=[...(rels.get('_rels/.rels')?.values()||[])].find(r=>r.type==='officeDocument'&&!r.external);
  if(!office)throw Error('No workbook relationship found in XLSX.');
  const wbName=partName('',office.target),doc=await read(wbName);
  if(doc.localName!=='workbook')throw Error('The package does not contain a spreadsheet workbook.');
  inspectChildren(doc,['fileVersion','workbookPr','bookViews','sheets','calcPr','definedNames','workbookProtection','externalReferences'],warn,'Workbook settings, extension records or additional workbook structures are omitted.');
  const dir=wbName.slice(0,wbName.lastIndexOf('/')+1), wbRels=rels.get(dir+'_rels/'+wbName.slice(dir.length)+'.rels');
  if(!wbRels)throw Error('Missing internal worksheet relationships.');
  const known=new Set(['[Content_Types].xml',wbName,...rels.keys()]);
  const lookup=type=>[...wbRels.values()].find(r=>r.type===type&&!r.external);
  let styles=[],strings=[];
  for(const type of ['styles','sharedStrings']){
    const rel=lookup(type);if(!rel)continue;
    const name=partName(wbName,rel.target);known.add(name);const root=await read(name);
    if(type==='styles')styles=readStyles(root,warn);
    else strings=children(root,'si').map(n=>richText(n,warn));
  }
  if(all(doc,'definedName').length)warn('Defined names, including print ranges and named formulas, are omitted.');
  if(first(doc,'workbookProtection'))warn('Workbook protection is omitted.');
  if(first(doc,'externalReferences'))warn('External workbook references are omitted.');
  for(const node of children(first(doc,'sheets'),'sheet')){
    const name=attr(node,'name'),id=node.getAttributeNS(REL,'id')||node.getAttribute('r:id'),rel=wbRels.get(id);
    if(!rel||rel.external)throw Error('Invalid or external worksheet relationship.');
    if(rel.type!=='worksheet'){warn('Chart sheets and other non-worksheet sheets are omitted.');continue}
    const sheet=ctx.addSheet(name),file=partName(wbName,rel.target);known.add(file);
    if(attr(node,'state')&&attr(node,'state')!=='visible')warn('Hidden sheets and rows or columns become visible.');
    const ws=await read(file);
    if(ws.localName!=='worksheet')throw Error('Invalid worksheet XML.');
    const labels={mergeCells:'Merged cells are unmerged.',drawing:'Drawings, charts and images are omitted.',legacyDrawing:'Legacy drawings and comments are omitted.',hyperlinks:'Hyperlink targets are omitted.',conditionalFormatting:'Conditional formatting is omitted.',dataValidations:'Data validation is omitted.',tableParts:'Structured tables are omitted.',autoFilter:'Filters are omitted.',sheetProtection:'Sheet protection is omitted.',cols:'Custom column widths, visibility and column styles are omitted.',extLst:'Worksheet extension data is omitted.'};
    for(const k of children(ws)){
      if(labels[k.localName])warn(labels[k.localName]);else if(!['dimension','sheetViews','sheetFormatPr','sheetData','pageMargins','sheetPr','printOptions','pageSetup','headerFooter'].includes(k.localName))warn('Other worksheet features ('+k.localName+') are omitted.');
    }
    if(all(ws,'pane').length)warn('Frozen and split panes are omitted.');
    let rowIndex=0;
    for(const row of children(first(ws,'sheetData'),'row')){
      inspectChildren(row,['c'],warn,'Additional row extension data is omitted.');
      rowIndex=Number(attr(row,'r')||rowIndex+1);let col=0;
      if(['s','ht','hidden','customFormat','outlineLevel','collapsed'].some(a=>row.hasAttribute(a)))warn('Row formatting, height, visibility and outlines are omitted.');
      for(const c of children(row,'c')){
        inspectChildren(c,['f','v','is'],warn,'Additional cell metadata or extension data is omitted.');
        if(Array.from(c.attributes).some(v=>!['r','s','t'].includes(v.localName)))
          warn('Additional cell metadata or extension data is omitted.');
        const addr=attr(c,'r')||address(rowIndex,col+1); col=coords(addr)[1];
        const f=first(c,'f'),v=first(c,'v'), type=attr(c,'t')||'n', style=styles[Number(attr(c,'s')||0)];
        const cell={v:''};
        if(f){
          if(attr(f,'t') && attr(f,'t')!=='normal')throw Error('Shared, array, or data-table formulas are unsupported. The original workbook was left intact.');
          if(!f.textContent)throw Error('A formula is missing its expression.');
          cell.f=f.textContent;
        }else if(type==='s'){
          const index=Number(v?.textContent);if(!v||!Number.isInteger(index)||index<0||index>=strings.length)throw Error('Invalid shared string index.');
          cell.v=strings[index];
        }else if(type==='inlineStr')cell.v=richText(first(c,'is'),warn);
        else if(type==='b'){if(!v||!['1','0','true','false'].includes(v.textContent))throw Error('Invalid Boolean cell.');cell.v=['1','true'].includes(v.textContent)}
        else if(type==='str')cell.v=stringValue(v);
        else if(type==='n'&&v){cell.v=Number(v.textContent);if(!Number.isFinite(cell.v))throw Error('Invalid numeric value in workbook.')}
        else if(type!=='n'){cell.v=v?.textContent||'';warn('Error, date or unrecognized cell types are imported as text.')}
        if(style&&Object.keys(style).length)cell.style={...style};
        if(attr(c,'s')&&!style)warn(STYLELOSS);
        if(cell.v!==''||cell.f||cell.style||type==='inlineStr'||type==='s')ctx.put(sheet,addr,cell);
      }
    }
  }
  for(const file of zip.names){
    if(known.has(file)||file.endsWith('/'))continue;
    if(/vba|macro|activeX|embeddings/i.test(file))warn('Macros, active content and embedded objects are omitted.');
    else if(/pivot|slicer|timeline/i.test(file))warn('Pivot tables, caches, slicers and timelines are omitted.');
    else if(/externalLink|connections|queryTable/i.test(file))warn('External links, connections and queries are omitted.');
    else if(/chart|drawing|media/i.test(file))warn('Drawings, charts and images are omitted.');
    else if(/comment|person/i.test(file))warn('Comments and annotations are omitted.');
    else if(/\/tables\//i.test(file))warn('Structured tables are omitted.');
    else if(!/^(docProps\/|xl\/theme\/|xl\/calcChain\.xml$)/.test(file))warn('Additional package parts or custom data are omitted.');
  }
  warn(GLOSS);return ctx.result();
}
function xtext(value){return xml(String(value).replace(/_x[\da-f]{4}_/gi,s=>'_x005F_'+s.slice(1)))}
function styleXML(styles){
  const rgb=c=>'FF'+c.slice(1).toUpperCase(),fonts=[],fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'],xfs=[];
  for(const s of styles){
    const font='<font>'+(s.bold?'<b/>':'')+(s.italic?'<i/>':'')+'<sz val="11"/>'+(s.color?'<color rgb="'+rgb(s.color)+'"/>':'<color theme="1"/>')+'<name val="Calibri"/></font>';
    let fontId=fonts.indexOf(font);if(fontId<0){fontId=fonts.length;fonts.push(font)}
    let fillId=0;if(s.fill){const fill='<fill><patternFill patternType="solid"><fgColor rgb="'+rgb(s.fill)+'"/><bgColor indexed="64"/></patternFill></fill>';fillId=fills.indexOf(fill);if(fillId<0){fillId=fills.length;fills.push(fill)}}
    const numFmtId=({number:2,percent:10,currency:164})[s.fmt]||0;
    xfs.push('<xf numFmtId="'+numFmtId+'" fontId="'+fontId+'" fillId="'+fillId+'" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"'+(s.align?' applyAlignment="1"><alignment horizontal="'+xml(s.align)+'"/></xf>':'/>'));
  }
  return XML+'<styleSheet xmlns="'+MS+'"><numFmts count="1"><numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0.00"/></numFmts><fonts count="'+fonts.length+'">'+fonts.join('')+'</fonts><fills count="'+fills.length+'">'+fills.join('')+'</fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="'+xfs.length+'">'+xfs.join('')+'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
}
export function writeXlsx(book,getValue){
  const {styles,id}=stylesOf(book),parts=[], sheets=[],relationships=[],overrides=[];
  const relation=(id,type,target)=>'<Relationship Id="rId'+id+'" Type="'+REL+'/'+type+'" Target="'+target+'"/>';
  for(const [index,sheet] of book.sheets.entries()){
    if(!isSheetNameExcel(sheet.name))throw Error('Rename the sheet “'+sheet.name+'” to an Excel-compatible name before exporting XLSX.');
    const n=index+1,file='worksheets/sheet'+n+'.xml';
    sheets.push('<sheet name="'+xml(sheet.name)+'" sheetId="'+n+'" r:id="rId'+n+'"/>');
    relationships.push(relation(n,'worksheet',file));
    overrides.push('<Override PartName="/xl/'+file+'" ContentType="'+MIME.xlsx+'.worksheet+xml"/>');
    const rows=[];
    for(const [r,items] of rowsOf(sheet)){
      const cells=items.map(({addr,cell})=>{
        const value=cell.f?getValue(index,addr):cell.v,style=id(cell.style);
        const at=' r="'+addr+'"'+(style?' s="'+style+'"':'');
        if(typeof value==='string'&&!cell.f)return '<c'+at+' t="inlineStr"><is><t xml:space="preserve">'+xtext(value)+'</t></is></c>';
        const type=typeof value==='boolean'?'b':typeof value==='string'?(errorValue(value)?'e':'str'):'n';
        return '<c'+at+' t="'+type+'">'+(cell.f?'<f>'+xml(cell.f)+'</f>':'')+'<v>'+xtext(type==='b'?Number(value):value??'')+'</v></c>';
      });rows.push('<row r="'+r+'">'+cells.join('')+'</row>');
    }
    parts.push(['xl/'+file,XML+'<worksheet xmlns="'+MS+'"><sheetData>'+rows.join('')+'</sheetData></worksheet>']);
  }
  const types=XML+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="'+MIME.xlsx+'.main+xml"/><Override PartName="/xl/styles.xml" ContentType="'+MIME.xlsx+'.styles+xml"/>'+overrides.join('')+'</Types>';
  relationships.push(relation(book.sheets.length+1,'styles','styles.xml'));
  const relDoc=content=>XML+'<Relationships xmlns="'+PKG+'">'+content+'</Relationships>';
  parts.push(['[Content_Types].xml',types],['_rels/.rels',relDoc(relation(1,'officeDocument','xl/workbook.xml'))],['xl/_rels/workbook.xml.rels',relDoc(relationships.join(''))],['xl/workbook.xml',XML+'<workbook xmlns="'+MS+'" xmlns:r="'+REL+'"><sheets>'+sheets.join('')+'</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>'],['xl/styles.xml',styleXML(styles)]);
  // Spreadsheet part types use the spreadsheetml namespace, without the package suffix.
  parts.find(([p])=>p==='[Content_Types].xml')[1]=types.replaceAll('.sheet.worksheet+xml','.worksheet+xml').replace('.sheet.styles+xml','.styles+xml');
  return makeZip(parts,MIME.xlsx);
}
