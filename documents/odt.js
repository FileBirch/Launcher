import {is,attr,ch,one,node,int,measure,odtRun,odtParagraph,mergeStyles} from './format.js';
import {internalPath,utf8} from './package.js';
export async function odt(zip,ctx){
  if(utf8.decode(await zip.bytes('mimetype',160))!=='application/vnd.oasis.opendocument.text')throw Error('The archive is not a supported OpenDocument Text document.');
  const root=await zip.xml('content.xml'), styles=await zip.xml('styles.xml',true), manifest=await zip.xml('META-INF/manifest.xml',true);
  if(manifest?.getElementsByTagNameNS('urn:oasis:names:tc:opendocument:xmlns:manifest:1.0','encryption-data').length)throw Error('Password-protected OpenDocument files are not supported.');
  const body=one(one(root,'office','body'),'office','text');
  if(!is(root,'office','document-content')||!body)throw Error('The archive has no supported OpenDocument text body.');
  const pStyles=new Map(), rStyles=new Map(), tStyles=new Map(), listStyles=new Map(), listCounts=new Map();
  for(const base of [styles,root]){
    for(const area of [...ch(base,'office','styles'),...ch(base,'office','automatic-styles')]){
      for(const e of [...ch(area,'style','default-style'),...ch(area,'style','style')]){
        const name=attr(e,'style','name')||'#default', family=attr(e,'style','family');
        const p=one(e,'style','paragraph-properties'),r=one(e,'style','text-properties'),c=one(e,'style','table-cell-properties');
        const item={parent:attr(e,'style','parent-style-name'),css:{},before:attr(p,'fo','break-before'),after:attr(p,'fo','break-after')};
        if(family==='paragraph'){item.css={...odtParagraph(p),...odtRun(r)};pStyles.set(name,item);}
        if(family==='text'){item.css=odtRun(r);rStyles.set(name,item);}
        if(family==='table-cell'){const fill=attr(c,'fo','background-color');if(/^#[0-9a-f]{6}$/i.test(fill))item.css.backgroundColor=fill;tStyles.set(name,item);}
      }
      for(const e of ch(area,'text','list-style'))listStyles.set(attr(e,'style','name'),Array.from(e.children));
    }
  }
  const pp=one(one(one(styles,'office','automatic-styles'),'style','page-layout'),'style','page-layout-properties');
  const max=measure(attr(pp,'fo','page-width'),320,1500);if(max)ctx.root.style.maxWidth=max;
  if(styles?.getElementsByTagNameNS('*','header').length||styles?.getElementsByTagNameNS('*','footer').length)ctx.warn('Page headers and footers are not rendered.');
  const tagmark=text=>{const el=node('span',text);el.className='papers-break';return el;};
  function frame(x,dst){
    const img=one(x,'draw','image'),path=internalPath('',attr(img,'xlink','href'));
    const desc=one(x,'svg','desc')?.textContent||one(x,'svg','title')?.textContent||'';
    const width=Number.parseFloat(measure(attr(x,'svg','width'),1,1500))||0;
    if(img)ctx.image(dst,/^Pictures\//.test(path)?path:null,desc,width);
    else ctx.omit(dst,'Drawing or object');
    if(attr(x,'text','anchor-type')&&attr(x,'text','anchor-type')!=='as-char')ctx.warn('Floating pictures are shown inline.');
  }
  function inline(src,dst){
    for(const x of src.childNodes){
      if(x.nodeType===3||x.nodeType===4){ctx.text(dst,x.textContent);continue;}
      if(is(x,'text','span')){
        const span=ctx.el('span');Object.assign(span.style,mergeStyles(rStyles,attr(x,'text','style-name')));inline(x,span);dst.append(span);
      }else if(is(x,'text','s'))ctx.text(dst,' '.repeat(int(attr(x,'text','c'),1,1,512)));
      else if(is(x,'text','tab'))dst.append('\t');
      else if(is(x,'text','line-break'))dst.append(node('br'));
      else if(is(x,'text','soft-page-break'))dst.append(tagmark('Page break'));
      else if(is(x,'text','a')){const span=ctx.el('span');span.className='papers-linktext';span.title='Link disabled in read-only preview';inline(x,span);dst.append(span);}
      else if(is(x,'draw','frame'))frame(x,dst);
      else if(is(x,'text','note')||is(x,'office','annotation'))ctx.omit(dst,'Note');
      else if(['date','time','page-number','page-count','sequence','reference-ref','bookmark-ref','conditional-text','user-field-get'].some(k=>is(x,'text',k))){
        ctx.text(dst,x.textContent);ctx.warn('Field results are shown as stored; fields are not evaluated.');
      }
    }
  }
  function pageBreakFor(key,side){
    for(let n=0;n<16&&pStyles.has(key);n++){const s=pStyles.get(key);if(s[side])return s[side]==='page';key=s.parent;}
    return false;
  }
  function paragraph(x,dst){
    const key=attr(x,'text','style-name'),p=ctx.el(is(x,'text','h')?'h'+int(attr(x,'text','outline-level'),1,1,6):'p');
    Object.assign(p.style,mergeStyles(pStyles,'#default'),mergeStyles(pStyles,key));
    if(is(x,'text','h')&&pStyles.has(key))p.classList.add('papers-styled-heading');
    if(pageBreakFor(key,'before'))dst.append(tagmark('Page break'));
    inline(x,p);if(!p.hasChildNodes())p.append(node('br'));dst.append(p);
    if(pageBreakFor(key,'after'))dst.append(tagmark('Page break'));
  }
  function list(src,dst,level,parentKey=''){
    const key=attr(src,'text','style-name')||parentKey;
    const def=listStyles.get(key)?.find(e=>attr(e,'text','level')===String(level));
    if(is(def,'text','list-level-style-image'))ctx.warn('Picture bullets use standard bullets.');
    const fmt=attr(def,'style','num-format'),ordered=is(def,'text','list-level-style-number')&&fmt!=='';
    const el=ctx.el(ordered?'ol':'ul');dst.append(el);
    const countKey=key+':'+level,id=src.getAttributeNS('http://www.w3.org/XML/1998/namespace','id');
    let sequence=int(attr(def,'text','start-value'),1,0,100000);
    const reference=attr(src,'text','continue-list');
    if(ordered){el.type=['a','A','i','I','1'].includes(fmt)?fmt:'1';if(el.type!==fmt)ctx.warn('Some custom list numbering uses simple numbering.');
      if(reference)sequence=listCounts.get('#'+reference)??sequence;
      else if(attr(src,'text','continue-numbering')==='true')sequence=listCounts.get(countKey)??sequence;
      el.start=sequence;
    }
    for(const entry of src.children){
      const item=is(entry,'text','list-item'),head=is(entry,'text','list-header');
      if(!item&&!head)continue;
      const li=ctx.el('li');el.append(li);if(head)li.style.listStyleType='none';
      else if(ordered){const restart=attr(entry,'text','start-value');if(restart)li.value=sequence=int(restart,sequence,0,100000);sequence++;}
      blocks(entry,li,level,key);
    }
    listCounts.set(countKey,sequence);if(id)listCounts.set('#'+id,sequence);
  }
  function table(src,dst){
    const box=ctx.el('div');box.className='papers-table';const t=ctx.el('table');box.append(t);dst.append(box);
    const repeats=(raw,max)=>{
      if(raw&&(!/^\d{1,8}$/.test(raw)||Number(raw)>max))throw Error('A repeated table structure exceeds the safe preview limit.');
      return int(raw,1,1,max);
    };
    function rows(parent,target,header){
      for(const row of parent.children){
        if(is(row,'table','table-header-rows')){const thead=ctx.el('thead');t.append(thead);rows(row,thead,true);continue;}
        if(is(row,'table','table-row-group')||is(row,'table','table-rows')){rows(row,target,header);continue;}
        if(!is(row,'table','table-row'))continue;
        const copies=repeats(attr(row,'table','number-rows-repeated'),512);
        for(let k=0;k<copies;k++){
          const tr=ctx.el('tr');target.append(tr);
          for(const cell of row.children){
            if(!is(cell,'table','table-cell'))continue;
            const times=repeats(attr(cell,'table','number-columns-repeated'),128);
            for(let j=0;j<times;j++){
              const td=ctx.el(header?'th':'td');if(header)td.scope='col';td.colSpan=int(attr(cell,'table','number-columns-spanned'),1,1,128);td.rowSpan=int(attr(cell,'table','number-rows-spanned'),1,1,512);tr.append(td);
              Object.assign(td.style,mergeStyles(tStyles,attr(cell,'table','style-name')));blocks(cell,td,0,'');
            }
          }
        }
      }
    }
    rows(src,t,false);
  }
  function blocks(src,dst,level=0,key=''){
    for(const x of src.children){
      if(is(x,'text','p')||is(x,'text','h'))paragraph(x,dst);
      else if(is(x,'text','list'))list(x,dst,level+1,key);
      else if(is(x,'table','table'))table(x,dst);
      else if(is(x,'text','section')){
        const section=ctx.el('section');dst.append(section);if(dst.childNodes.length>1)section.append(tagmark('Section break'));blocks(x,section,level,key);
      }else if(is(x,'draw','frame'))frame(x,dst);
      else if(is(x,'text','table-of-content')||is(x,'text','alphabetical-index')){
        const index=one(x,'text','index-body');if(index)blocks(index,dst,level,key);ctx.warn('Generated indexes are shown as stored and do not update.');
      }
    }
  }
  blocks(body,ctx.root);
  if([...zip.names].some(n=>/^(Scripts|Basic)\//.test(n)))ctx.warn('Macros and scripts are ignored; they never run.');
}
