import {W,is,attr,ch,one,node,int,wordRun,wordParagraph,mergeStyles,hex} from './format.js';
import {internalPath} from './package.js';
const REL=['http://schemas.openxmlformats.org/officeDocument/2006/relationships','http://purl.oclc.org/ooxml/officeDocument/relationships'];
const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
const DRAW=['http://schemas.openxmlformats.org/drawingml/2006/main','http://purl.oclc.org/ooxml/drawingml/main'];
const val=(n,k)=>attr(one(n,'w',k),'w','val');
export async function word(zip,ctx){
  const root=await zip.xml('word/document.xml');
  const body=one(root,'w','body');
  if(!is(root,'w','document')||!body)throw Error('The archive is not a supported Word document.');
  const [styles,numbering,rels]=await Promise.all(['word/styles.xml','word/numbering.xml','word/_rels/document.xml.rels'].map(n=>zip.xml(n,true)));
  const styleMap=new Map(), lists=new Map(), abstract=new Map(), images=new Map(), counters=new Map();
  const settings=one(styles,'w','docDefaults');
  const defaultP=one(one(settings,'w','pPrDefault'),'w','pPr');
  const defaultR=one(one(settings,'w','rPrDefault'),'w','rPr');
  for(const s of ch(styles,'w','style')){
    const key=attr(s,'w','styleId'),type=attr(s,'w','type'),p=one(s,'w','pPr');
    if(!key)continue;
    styleMap.set(key,{parent:val(s,'basedOn'),name:val(s,'name'),level:val(p,'outlineLvl'),num:one(p,'w','numPr'),css:type==='paragraph'?{...wordParagraph(p),...wordRun(one(s,'w','rPr'))}:wordRun(one(s,'w','rPr'))});
  }
  for(const s of ch(numbering,'w','abstractNum')){
    const levels=new Map();for(const x of ch(s,'w','lvl'))levels.set(attr(x,'w','ilvl'),x);
    abstract.set(attr(s,'w','abstractNumId'),levels);
  }
  for(const x of ch(numbering,'w','num')){
    const overrides=new Map();for(const ov of ch(x,'w','lvlOverride'))overrides.set(attr(ov,'w','ilvl'),ov);
    lists.set(attr(x,'w','numId'),{levels:abstract.get(val(x,'abstractNumId')),overrides});
  }
  for(const x of rels?.children??[]){
    if(x.namespaceURI!==PKG||x.localName!=='Relationship'||x.getAttribute('TargetMode')==='External')continue;
    if(!REL.some(n=>x.getAttribute('Type')===n+'/image'))continue;
    const id=x.getAttribute('Id');
    const path=internalPath('word',x.getAttribute('Target'));
    if(/^word\/media\//.test(path))images.set(id,path);
  }
  const explicitBreak=(kind='page')=>{const mark=node('span',kind==='section'?'Section break':'Page break');mark.className='papers-break';return mark;};
  function sectionGeometry(dst,sect){
    const sz=one(sect,'w','pgSz'), n=attr(sz,'w','w');
    if(/^\d{3,6}$/.test(n))dst.style.maxWidth=Math.min(1500,Math.max(320,Number(n)/15))+'px';
    if(ch(sect,'w','headerReference').length||ch(sect,'w','footerReference').length)ctx.warn('Page headers and footers are not rendered.');
    if(int(attr(one(sect,'w','cols'),'w','num'),1,1,100)>1)ctx.warn('Multi-column sections are rendered in a single reading column.');
  }
  function styleMeta(key,k){for(let n=0;n<16&&styleMap.has(key);n++){const s=styleMap.get(key);if(s[k]!==undefined&&s[k]!=='')return s[k];key=s.parent;}return '';}
  function drawing(n,dst){
    let found=false;
    for(const uri of DRAW)for(const x of n.getElementsByTagNameNS(uri,'blip')){
      const id=REL.map(s=>x.getAttributeNS(s,'embed')).find(Boolean);
      const props=Array.from(n.getElementsByTagNameNS('*','docPr')).find(el=>el.namespaceURI?.includes('wordprocessingDrawing'));
      const ext=Array.from(n.getElementsByTagNameNS('*','extent')).find(el=>el.namespaceURI?.includes('wordprocessingDrawing'));
      const width=Number(ext?.getAttribute('cx'))/9525;
      ctx.image(dst,images.get(id),props?.getAttribute('descr')||props?.getAttribute('title')||'',Number.isFinite(width)&&width>0?Math.min(1500,width):0);found=true;
    }
    if(!found)ctx.omit(dst,'Drawing');
    else if(Array.from(n.getElementsByTagNameNS('*','anchor')).some(el=>el.namespaceURI?.includes('wordprocessingDrawing')))ctx.warn('Floating pictures are shown inline.');
  }
  function inline(src,dst,depth=0){
    if(depth>40)return;
    for(const x of src.children){
      if(!W.includes(x.namespaceURI))continue;
      switch(x.localName){
        case 'r':{const r=ctx.el('span');const p=one(x,'w','rPr');Object.assign(r.style,mergeStyles(styleMap,val(p,'rStyle')),wordRun(p));inline(x,r,depth+1);dst.append(r);break;}
        case 't':ctx.text(dst,x.textContent);break;
        case 'tab':ctx.text(dst,'\t');break;
        case 'cr':dst.append(node('br'));break;
        case 'br':dst.append(attr(x,'w','type')==='page'?explicitBreak():node('br'));break;
        case 'noBreakHyphen':dst.append('\u2011');break;
        case 'softHyphen':dst.append('\u00ad');break;
        case 'drawing':drawing(x,dst);break;
        case 'hyperlink':{const span=node('span');span.className='papers-linktext';span.title='Link disabled in read-only preview';inline(x,span,depth+1);dst.append(span);break;}
        case 'sdt':{const contents=one(x,'w','sdtContent');if(contents)inline(contents,dst,depth+1);break;}
        case 'ins':inline(x,dst,depth+1);ctx.warn('Tracked changes are shown as the current text only.');break;
        case 'del':ctx.warn('Tracked changes are shown as the current text only.');break;
        case 'fldSimple':inline(x,dst,depth+1);ctx.warn('Field results are shown as stored; fields are not evaluated.');break;
        case 'instrText':case 'fldChar':ctx.warn('Field results are shown as stored; fields are not evaluated.');break;
        case 'footnoteReference':case 'endnoteReference':case 'commentReference':ctx.omit(dst,'Note');break;
        case 'object':case 'pict':ctx.omit(dst,'Object');break;
      }
    }
  }
  function paragraph(src){
    const p=one(src,'w','pPr'), key=val(p,'pStyle'), n=val(p,'outlineLvl')||styleMeta(key,'level');
    const heading=/^[0-8]$/.test(n)?Math.min(Number(n)+1,6):Number(/heading\s*([1-6])(?:\b|$)/i.exec(styleMeta(key,'name')||key)?.[1]);
    const el=ctx.el(heading?'h'+heading:'p');
    Object.assign(el.style,wordParagraph(defaultP),wordRun(defaultR),mergeStyles(styleMap,key),wordParagraph(p),wordRun(one(p,'w','rPr')));
    if(heading&&styleMap.has(key))el.classList.add('papers-styled-heading');
    inline(src,el);
    if(!el.hasChildNodes())el.append(node('br'));
    return {el,p,num:one(p,'w','numPr')||styleMeta(key,'num')};
  }
  function numInfo(info){
    const id=val(info,'numId');if(!id||id==='0')return null;
    const lev=String(int(val(info,'ilvl'),0,0,8)), list=lists.get(id), ov=list?.overrides.get(lev);
    const level=one(ov,'w','lvl')||list?.levels?.get(lev);
    const fmt=val(level,'numFmt')||'bullet';
    const types={decimal:'1',lowerLetter:'a',upperLetter:'A',lowerRoman:'i',upperRoman:'I'};
    if(fmt!=='bullet'&&!types[fmt])ctx.warn('Some custom list numbering uses simple numbering.');
    if(val(level,'lvlText')&&!/^(%[1-9][.)]?)$/.test(val(level,'lvlText'))&&fmt!=='bullet')ctx.warn('Some custom list labels are simplified.');
    if(fmt==='bullet'&&val(level,'lvlText')&&!['•','◦','▪'].includes(val(level,'lvlText')))ctx.warn('Some custom bullets use standard bullets.');
    return {id,lev:Number(lev),ordered:fmt!=='bullet',type:types[fmt]||'1',start:int(val(ov,'startOverride')||val(level,'start'),1,0,100000)};
  }
  function table(src){
    const frame=ctx.el('div');frame.className='papers-table';const t=ctx.el('table');frame.append(t);
    let previous=new Map();
    for(const row of ch(src,'w','tr')){
      const tr=ctx.el('tr');t.append(tr);const next=new Map();
      const header=one(one(row,'w','trPr'),'w','tblHeader');
      let column=int(val(one(row,'w','trPr'),'gridBefore'),0,0,128);
      for(const c of ch(row,'w','tc')){
        const p=one(c,'w','tcPr'), span=int(val(p,'gridSpan'),1,1,128), vm=one(p,'w','vMerge');
        if(vm&&attr(vm,'w','val')!=='restart'&&previous.has(column)&&previous.get(column).span===span){
          const prev=previous.get(column);prev.el.rowSpan++;next.set(column,prev);column+=span;continue;
        }
        const td=ctx.el(header?'th':'td');td.colSpan=span;if(header)td.scope='col';tr.append(td);
        const fill=hex(attr(one(p,'w','shd'),'w','fill'));if(fill)td.style.backgroundColor=fill;
        blocks(c,td);
        if(vm)next.set(column,{el:td,span});column+=span;
      }
      previous=next;
    }
    return frame;
  }
  function blocks(src,dst){
    let stack=[];
    for(const x of src.children){
      if(!W.includes(x.namespaceURI)){if(x.localName==='AlternateContent')ctx.omit(dst,'Alternate content');continue;}
      if(x.localName==='p'){
        const {el,p,num}=paragraph(x), cfg=numInfo(num);
        if(one(p,'w','pageBreakBefore')&&val(p,'pageBreakBefore')!=='0'){stack=[];dst.append(explicitBreak());}
        if(!cfg){stack=[];dst.append(el);}
        else{
          const depth=Math.min(cfg.lev,stack.length);
          stack.length=Math.min(stack.length,depth+1);
          let curr=stack[depth];
          if(!curr||curr.id!==cfg.id||curr.ordered!==cfg.ordered){
            const list=ctx.el(cfg.ordered?'ol':'ul');if(cfg.ordered){list.type=cfg.type;list.start=counters.get(cfg.id+':'+cfg.lev)??cfg.start;}
            ((depth&&stack[depth-1]?.last)||dst).append(list);curr={...cfg,list};stack[depth]=curr;
          }
          const li=ctx.el('li');li.append(el);curr.list.append(li);curr.last=li;
          if(cfg.ordered){const key=cfg.id+':'+cfg.lev;counters.set(key,(counters.get(key)??cfg.start)+1);}
        }
        const sect=one(p,'w','sectPr');
        if(sect){sectionGeometry(dst,sect);stack=[];dst.append(explicitBreak('section'));}
      }else if(x.localName==='tbl'){stack=[];dst.append(table(x));}
      else if(x.localName==='sectPr')sectionGeometry(dst,x);
      else if(x.localName==='altChunk')ctx.omit(dst,'Embedded document');
      else if(x.localName==='sdt'){stack=[];const child=one(x,'w','sdtContent');if(child)blocks(child,dst);}
    }
  }
  blocks(body,ctx.root);
  if(zip.names.has('word/vbaProject.bin'))ctx.warn('Macros are ignored; they never run.');
}
