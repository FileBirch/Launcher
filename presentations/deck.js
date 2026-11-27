import {openZip,MAX_INPUT} from '../spreadsheets/package.js';

const OFFICEREL=['http://schemas.openxmlformats.org/officeDocument/2006/relationships','http://purl.oclc.org/ooxml/officeDocument/relationships'];
const MAX_SLIDE_RUNS=12000;
const SAFE_COLORS=/^#[a-f0-9]{6}$/i;
const CRC=Uint32Array.from({length:256},(_,n)=>{for(let i=0;i<8;i++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
const children=(x,n)=>[...x?.children||[]].filter(e=>!n||e.localName===n);
const child=(x,n)=>children(x,n)[0];
const descendants=(x,n)=>[...x?.getElementsByTagNameNS('*',n)||[]];
const attr=(x,n)=>x?.getAttribute(n)||[...x?.attributes||[]].find(a=>a.localName===n)?.value||'';
const relAttr=(x,name)=>OFFICEREL.map(uri=>x?.getAttributeNS(uri,name)).find(Boolean)||'';
const getText=x=>String(x?.textContent||'').slice(0,20000);
const finite=(value)=>Number.isFinite(value)?value:0;
const bound=(n,min,max)=>Math.min(Math.max(finite(n),min),max);
const points=v=>Number(v)/100;
const emu=v=>Number(v)/9525;
const color=v=>SAFE_COLORS.test(v||'')?v:null;
function length(v){
  const m=/^(-?[0-9]*\.?[0-9]+)(cm|mm|in|pt|pc|px)$/.exec(v||'');
  return m?Number(m[1])*({cm:96/2.54,mm:96/25.4,in:96,pt:96/72,pc:16,px:1})[m[2]]:0;
}
function dim(w,h){
  if(w<24||h<24||w>10000||h>10000)throw Error('The slide size is outside the viewer’s supported range.');
  return {width:w,height:h};
}
function resolvePart(base,value){
  if(typeof value!=='string'||!value||value.length>1024||/[\\?#\u0000-\u001f]|(^|\/)\/|^[a-z][a-z0-9+.-]*:/i.test(value))return null;
  let path;
  try{path=decodeURIComponent(value);}catch{return null;}
  if(/[\\?#\u0000-\u001f]|(^|\/)\/|^[a-z][a-z0-9+.-]*:/i.test(path))return null;
  const parts=path.startsWith('/')?[]:base.split('/').slice(0,-1);
  for(const p of path.split('/')){if(!p||p==='.')continue;if(p==='..'){if(!parts.length)return null;parts.pop();}else parts.push(p);}
  return parts.join('/');
}
function archive(buffer){
  if(!(buffer instanceof ArrayBuffer)||buffer.byteLength<22||buffer.byteLength>MAX_INPUT)throw Error('Use an unencrypted presentation under 20 MiB.');
  let zip;
  try{zip=openZip(buffer);}catch{throw Error('This presentation package is malformed, encrypted, or exceeds the ZIP safety limits.');}
  const names=new Set(zip.names),xmlCache=new Map(),artCache=new Map();
  let xmlTotal=0,pictureTotal=0;
  async function xml(name){
    if(xmlCache.has(name))return xmlCache.get(name);
    if(!names.has(name))return null;
    const data=zip.packed(name);
    if(data.size>2*1024*1024||(xmlTotal+=data.size)>20*1024*1024)throw Error('Presentation XML exceeds the safe processing limit.');
    const src=await zip.text(name);
    if(/<!\s*(DOCTYPE|ENTITY)/i.test(src))throw Error('Document types and XML entities are not supported.');
    const doc=new DOMParser().parseFromString(src,'application/xml');
    if(doc.getElementsByTagNameNS('*','parsererror').length)throw Error('A presentation part has invalid XML.');
    xmlCache.set(name,doc.documentElement);return doc.documentElement;
  }
  async function picture(name){
    if(!name||!names.has(name))return null;
    if(artCache.has(name))return artCache.get(name);
    const unsafe=()=>{artCache.set(name,null);return null;};
    const p=zip.packed(name);
    if(p.size<12||p.size>3*1024*1024||pictureTotal+p.size>12*1024*1024)return unsafe();
    try{
      let b=p.bytes;
      if(p.method===8){
        b=new Uint8Array(p.size);let used=0;
        const r=new Blob([p.bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
        try{while(true){const {done,value}=await r.read();if(done)break;if(used+value.length>p.size)return unsafe();b.set(value,used);used+=value.length;}if(used!==p.size)return unsafe();}
        finally{await r.cancel().catch(()=>{});r.releaseLock();}
      }
      if(b.length!==p.size)return unsafe();
      let crc=0xffffffff;for(const v of b)crc=CRC[(crc^v)&255]^(crc>>>8);
      if(((crc^0xffffffff)>>>0)!==p.crc)return unsafe();
      // Validate declared raster dimensions before decoding; SVG, animation and other media stay inert.
      const view=new DataView(b.buffer,b.byteOffset,b.byteLength),ascii=(off,len)=>String.fromCharCode(...b.subarray(off,off+len));
      let type,w,h;
      if(view.getUint32(0)===0x89504e47&&view.getUint32(4)===0x0d0a1a0a&&ascii(12,4)==='IHDR'&&b.length>=24){
        type='image/png';w=view.getUint32(16);h=view.getUint32(20);
        for(let i=8;i+12<=b.length;){const n=view.getUint32(i);if(n>b.length-i-12||ascii(i+4,4)==='acTL')return unsafe();i+=n+12;}
      }else if(view.getUint16(0)===0xffd8){
        for(let i=2;i+8<b.length;){if(b[i++]!==255)break;while(b[i]===255)i++;const code=b[i++];if(code===0xda||code===0xd9||i+2>b.length)break;const size=view.getUint16(i);
          if(size<2||i+size>b.length)break;
          if([0xc0,0xc1,0xc2].includes(code)&&size>=8){type='image/jpeg';h=view.getUint16(i+3);w=view.getUint16(i+5);break;}i+=size;
        }
      }
      if(!type||w<1||h<1||w*h>16000000||w>8000||h>8000)return unsafe();
      const url=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=reject;reader.readAsDataURL(new Blob([b],{type}));});
      pictureTotal+=b.length;artCache.set(name,url);return url;
    }catch{return unsafe();}
  }
  return {names,xml,picture,async text(name){if(!names.has(name)||zip.packed(name).size>512)return '';return zip.text(name);}};
}
async function relationships(pkg,part){
  const p=part.split('/'),n=p.pop();
  const xml=await pkg.xml([...p,'_rels',n+'.rels'].join('/')),out=new Map();
  for(const rel of children(xml,'Relationship')){
    const id=attr(rel,'Id');if(!id||out.has(id))throw Error('The deck contains invalid relationships.');
    const external=attr(rel,'TargetMode').toLowerCase()==='external';
    out.set(id,{type:attr(rel,'Type').split('/').at(-1),target:external?null:resolvePart(part,attr(rel,'Target')),external});
  }
  return out;
}
function PPTColor(x){const rgb=attr(child(x,'srgbClr'),'val');return color(rgb?'#'+rgb:null);}
function transform(x){
  return {x:emu(attr(child(x,'off'),'x')),y:emu(attr(child(x,'off'),'y')),
    width:emu(attr(child(x,'ext'),'cx')),height:emu(attr(child(x,'ext'),'cy')),rotation:bound(Number(attr(x,'rot'))/60000,-360,360)};
}
function trimData(items,notes){
  let left=MAX_SLIDE_RUNS,chars=120000;
  for(const item of items){
    if(!item.paragraphs)continue;
    for(const p of item.paragraphs){
      if(p.runs.length>left||p.runs.reduce((sum,r)=>sum+(r.text?.length||0),0)>chars)notes.add('Some text exceeded the slide preview limit.');
      p.runs=p.runs.slice(0,left).flatMap(r=>{
        if(!chars)return [];
        if(r.text){r.text=r.text.slice(0,chars);chars-=r.text.length;}
        return [r];
      });
      left-=p.runs.length;
    }
  }
}
function pptText(x){
  return children(x,'p').slice(0,128).map(p=>{
    const pr=child(p,'pPr');
    return {align:{l:'left',ctr:'center',r:'right',just:'justify'}[attr(pr,'algn')],runs:children(p).slice(0,512).flatMap(r=>{
      if(r.localName==='br')return [{break:true}];
      if(!['r','fld'].includes(r.localName))return [];
      const rp=child(r,'rPr'),def=child(pr,'defRPr');
      return [{text:getText(child(r,'t')),size:points(attr(rp,'sz')||attr(def,'sz'))*96/72,bold:attr(rp,'b')==='1',italic:attr(rp,'i')==='1',color:PPTColor(child(rp,'solidFill'))}];
    })};
  });
}
async function pptx(pkg){
  const root=await pkg.xml('ppt/presentation.xml');
  if(root?.localName!=='presentation')throw Error('Open a valid PPTX presentation.');
  const size=child(root,'sldSz'),geometry=size?dim(emu(attr(size,'cx')),emu(attr(size,'cy'))):{width:960,height:540};
  const rels=await relationships(pkg,'ppt/presentation.xml');
  const ids=children(child(root,'sldIdLst'),'sldId');
  if(!ids.length||ids.length>180)throw Error('Present supports decks with 1–180 slides.');
  const slides=[];
  for(const id of ids){
    const r=rels.get(relAttr(id,'id'));
    if(r?.type!=='slide'||!r.target?.startsWith('ppt/')||!r.target.endsWith('.xml'))throw Error('The deck has a missing or unsupported slide reference.');
    const slide=await pkg.xml(r.target);
    if(slide?.localName!=='sld')throw Error('A slide part is missing or invalid.');
    const s=child(slide,'cSld'),tree=child(s,'spTree'), sr=await relationships(pkg,r.target),notes=new Set(),items=[];
    const pptbg=child(child(s,'bg'),'bgPr'),background=PPTColor(child(pptbg,'solidFill'));
    if(pptbg&&!background)notes.add('Theme, gradient, or picture background approximated.');
    if(child(slide,'timing')||child(slide,'transition'))notes.add('Animation and slide transitions are not played.');
    if([...sr.values()].some(x=>['media','audio','video','oleObject','chart','diagramData'].includes(x.type)))notes.add('Media, charts, or interactive objects are not played or opened.');
    if([...sr.values()].some(x=>x.external))notes.add('External links or resources were not loaded.');
    let fallback=0;
    for(const item of children(tree).slice(0,500)){
      if(['nvGrpSpPr','grpSpPr','extLst'].includes(item.localName))continue;
      const sp=child(item,'spPr'),nv=children(item).find(c=>c.localName.startsWith('nv')),name=attr(child(nv,'cNvPr'),'descr')||attr(child(nv,'cNvPr'),'name');
      let pos=transform(child(sp,'xfrm'));
      if(!pos.width||!pos.height){
        pos={x:geometry.width*.07,y:geometry.height*(.10+Math.min(fallback++,6)*.12),width:geometry.width*.86,height:geometry.height*.16,rotation:0};
        notes.add('Some placeholder layout was approximated.');
      }
      const pic=item.localName==='pic'?child(child(item,'blipFill'),'blip'):null;
      if(pic){
        const p=sr.get(relAttr(pic,'embed')),image=p?.type==='image'?await pkg.picture(p.target):null;
        if(!image)notes.add('Linked, animated, or unsupported pictures are shown as placeholders.');
        items.push({...pos,image,placeholder:!image?'Image not shown':null,name});
      }else if(item.localName==='sp'){
        const shape=attr(child(sp,'prstGeom'),'prst')||'rect',paragraphs=pptText(child(item,'txBody'));
        const special=!['rect','roundRect','ellipse','line'].includes(shape);
        if(special)notes.add('Some special shapes are approximated.');
        if(special&&!paragraphs.length)items.push({...pos,placeholder:'Unsupported shape'});
        else items.push({...pos,shape,paragraphs,fill:PPTColor(child(sp,'solidFill')),line:PPTColor(child(child(sp,'ln'),'solidFill')),name});
      }else{
        const graphic=child(item,'xfrm'),sizeNode=child(graphic,'ext');
        if(sizeNode)pos=transform(graphic);
        items.push({...pos,placeholder:({graphicFrame:'Chart, table, or diagram not rendered',grpSp:'Grouped objects not rendered',cxnSp:'Connector not rendered'})[item.localName]||'Unsupported slide object'});
        notes.add('Some grouped objects, charts, diagrams, or media are shown as placeholders.');
      }
    }
    if(children(tree).length>500)notes.add('Only the first 500 elements on this slide are shown.');
    trimData(items,notes);
    slides.push({...geometry,background,items,notes:[...notes],title:attr(s,'name')||items.flatMap(a=>a.paragraphs||[]).find(p=>p.runs.some(r=>r.text?.trim()))?.runs.map(r=>r.text||'').join('').slice(0,100)||'Slide '+(slides.length+1)});
  }
  return {slides,format:'PPTX'};
}
function odpStyles(roots){
  const styles=new Map(),masters=new Map(),layouts=new Map();
  for(const root of roots){
    for(const n of descendants(root,'page-layout'))layouts.set(attr(n,'name'),child(n,'page-layout-properties'));
    for(const n of descendants(root,'master-page'))masters.set(attr(n,'name'),attr(n,'page-layout-name'));
    for(const n of descendants(root,'style'))styles.set(attr(n,'name'),n);
  }
  function properties(name,kind,field){
    const seen=new Set();
    while(name&&!seen.has(name)&&seen.size<12){seen.add(name);const s=styles.get(name),v=attr(child(s,kind),field);if(v)return v;name=attr(s,'parent-style-name');}
    return '';
  }
  return {properties,masters,layouts};
}
function odpText(flow,styles){
  function runs(parent,inherited='',depth=0){
    if(!parent||depth>20)return [];
    const result=[],style=attr(parent,'style-name')||inherited;
    const role=(field)=>styles.properties(style,'text-properties',field);
    for(const n of [...parent.childNodes].slice(0,500)){
      if(n.nodeType===3||n.nodeType===4)result.push({text:n.nodeValue.slice(0,20000),size:length(role('font-size')),bold:['bold','700'].includes(role('font-weight')),italic:role('font-style')==='italic',color:color(role('color'))});
      else if(n.nodeType===1){
        if(n.localName==='line-break')result.push({break:true});
        else if(n.localName==='s')result.push({text:' '.repeat(bound(Number(attr(n,'c')||1),1,1000))});
        else if(n.localName==='tab')result.push({text:'    '});
        else if(['span','a','bookmark','bookmark-start','bookmark-end'].includes(n.localName))result.push(...runs(n,style,depth+1));
      }
    }
    return result;
  }
  return [...descendants(flow,'p'),...descendants(flow,'h')].sort((a,b)=>a===b?0:a.compareDocumentPosition(b)&2?1:-1).slice(0,128).map(p=>({align:({start:'left',end:'right',left:'left',right:'right',center:'center',justify:'justify'})[styles.properties(attr(p,'style-name'),'paragraph-properties','text-align')],runs:runs(p)}));
}
async function odp(pkg){
  const mime=await pkg.text('mimetype');
  if(mime&&mime.trim()!=='application/vnd.oasis.opendocument.presentation')throw Error('Open a valid ODP presentation.');
  const root=await pkg.xml('content.xml'),pres=child(child(root,'body'),'presentation');
  if(root?.localName!=='document-content'||!pres)throw Error('Open a valid ODP presentation.');
  const styleXML=await pkg.xml('styles.xml'),styles=odpStyles([styleXML,root]);
  const pages=children(pres,'page');
  if(!pages.length||pages.length>180)throw Error('Present supports decks with 1–180 slides.');
  const slides=[];
  for(const page of pages){
    const layout=styles.layouts.get(styles.masters.get(attr(page,'master-page-name')))||styles.layouts.values().next().value;
    const geometry=layout?dim(length(attr(layout,'page-width')),length(attr(layout,'page-height'))):{width:960,height:540};
    const base=attr(page,'style-name'),background=color(styles.properties(base,'drawing-page-properties','fill-color'));
    const notes=new Set(),items=[];
    if(!layout)notes.add('Slide size used a standard fallback.');
    if(styles.properties(base,'drawing-page-properties','transition-style')||descendants(page,'animate').length||descendants(page,'par').length)notes.add('Animation and slide transitions are not played.');
    for(const item of children(page).slice(0,500)){
      if(['notes','forms','event-listeners','animations','par','seq'].includes(item.localName)){if(item.localName!=='notes')notes.add('Event and animation playback is disabled.');continue;}
      let pos={x:length(attr(item,'x')),y:length(attr(item,'y')),width:length(attr(item,'width')),height:length(attr(item,'height'))};
      if(attr(item,'transform'))notes.add('Some transformed items were approximated.');
      if(!pos.width||!pos.height){pos={x:geometry.width*.05,y:geometry.height*(.10+Math.min(items.length,6)*.12),width:geometry.width*.9,height:geometry.height*.16};notes.add('Some inherited object layout was approximated.');}
      const style=attr(item,'style-name'),fill=color(styles.properties(style,'graphic-properties','fill-color')),line=color(styles.properties(style,'graphic-properties','stroke-color'));
      const image=child(item,'image'),box=child(item,'text-box');
      if(image){
        const raster=await pkg.picture(resolvePart('content.xml',attr(image,'href')));
        if(!raster)notes.add('Linked, animated, or unsupported pictures are shown as placeholders.');
        items.push({...pos,image:raster,placeholder:!raster?'Image not shown':null,name:getText(child(item,'title'))});
      }else if(box||['rect','ellipse','circle','custom-shape','caption'].includes(item.localName)){
        items.push({...pos,shape:['ellipse','circle'].includes(item.localName)?'ellipse':'rect',fill,line,paragraphs:odpText(box||item,styles)});
        if(item.localName==='custom-shape')notes.add('Some special shapes are approximated.');
      }else{
        items.push({...pos,placeholder:'Unsupported slide object'});
        notes.add('Some grouped objects, charts, diagrams, or media are shown as placeholders.');
      }
      if(descendants(item,'a').length)notes.add('Links are shown as plain text only.');
    }
    if(children(page).length>500)notes.add('Only the first 500 elements on this slide are shown.');
    trimData(items,notes);
    slides.push({...geometry,background,items,notes:[...notes],title:attr(page,'name')||'Slide '+(slides.length+1)});
  }
  return {slides,format:'ODP'};
}
export async function readDeck(file){
  if(!file||file.size>MAX_INPUT)throw Error('Use an unencrypted presentation under 20 MiB.');
  const extension=/\.(odp|pptx)$/i.exec(file.name||'')?.[1].toLowerCase();
  if(!extension)throw Error('Choose a .pptx or .odp file.');
  const pkg=archive(await file.arrayBuffer());
  return extension==='pptx'?pptx(pkg):odp(pkg);
}
