import {packageReader,MAX_INPUT,utf8,raster,dataUri} from './package.js';
import {node} from './format.js';
import {word} from './word.js';
import {odt} from './odt.js';
const MAX_TXT=4*1024*1024, MAX_IMAGES=32, MAX_IMG_BYTES=12*1024*1024, MAX_PIXELS=24_000_000;
export async function renderDocument(file){
  const type=/\.(docx|odt|txt)$/i.exec(file.name)?.[1].toLowerCase();
  if(!type)throw Error('Choose a .docx, .odt or UTF-8 .txt file. Legacy .doc and macro-enabled .docm files are not supported.');
  if(file.size>(type==='txt'?MAX_TXT:MAX_INPUT))throw Error('The file exceeds the preview limit ('+(type==='txt'?'4 MiB for text':'20 MiB for packaged documents')+').');
  const buf=await file.arrayBuffer(), root=node('article'), warnings=new Set();
  root.className='papers-page';
  const warn=message=>warnings.add(message);
  if(type==='txt'){
    let content;
    try{content=utf8.decode(buf);}catch{throw Error('This file does not look like UTF-8 plain text.');}
    if(content.includes('\0'))throw Error('This file does not look like UTF-8 plain text.');
    root.append(node('pre',content));
    return {type,root,warnings};
  }
  const zip=packageReader(buf), jobs=[];
  let elements=0, characters=0;
  function el(tag){if(++elements>24000)throw Error('The document has too many elements to render safely.');return node(tag);}
  function text(dst,value){
    characters+=value.length;
    if(characters>6*1024*1024)throw Error('The document has too much text to render safely.');
    if(++elements>24000)throw Error('The document has too many elements to render safely.');
    dst.append(document.createTextNode(value));
  }
  function omit(dst,label){
    warn('Some embedded objects, notes or images are not supported.');
    const span=el('span');span.className='papers-omitted';span.textContent='['+label+' omitted]';dst.append(span);
  }
  function image(dst,path,alt,width){
    const span=el('span');dst.append(span);
    if(!path||!zip.names.has(path)||jobs.length>=MAX_IMAGES){omit(span,'Image');return;}
    jobs.push({span,path,alt:alt.slice(0,512),width});
  }
  await (type==='docx'?word:odt)(zip,{el,text,image,omit,warn,root});
  let shownBytes=0,pixels=0;
  const cached=new Map();
  for(const {span,path,alt,width} of jobs){
    if(!cached.has(path)){
      try{
        const data=await zip.bytes(path,3*1024*1024), details=raster(data);
        if(!details||shownBytes+data.length>MAX_IMG_BYTES)throw Error();
        shownBytes+=data.length;cached.set(path,{...details,src:dataUri(details.type,data)});
      }catch{cached.set(path,null);}
    }
    const item=cached.get(path);
    if(!item||item.w*item.h+pixels>MAX_PIXELS){omit(span,'Image');continue;}
    pixels+=item.w*item.h;
    const img=el('img');img.alt=alt||'Embedded picture';img.decoding='async';img.loading='lazy';
    img.width=Math.round(width)||item.w;
    img.onerror=()=>{img.replaceWith(node('span','[Image omitted]'));};
    img.src=item.src;span.append(img);
  }
  if(!root.hasChildNodes())warn('No supported document content was found.');
  return {type,root,warnings};
}
export function mount(host,{brand,base}){
  const header=node('div');header.className='papers-intro';
  const title=node('h1',brand?.apps?.documents?.name||'Papers');
  const caption=node('p','Read-only document viewer. Files are opened in this browser, not uploaded or edited.');
  const toolbar=node('div');toolbar.className='papers-toolbar';
  const pick=node('input');pick.type='file';pick.id='papers-file';pick.accept='.docx,.odt,.txt';
  const label=node('label','Open local document');label.htmlFor=pick.id;
  const clear=node('button','Close document');clear.type='button';clear.disabled=true;
  const link=node('a','Back to the launcher');link.href=base;
  toolbar.append(label,pick,clear,link);
  const status=node('p','Choose a .docx, .odt or UTF-8 .txt file.');status.role='status';status.className='papers-status';
  const info=node('details');info.className='papers-details';
  info.append(node('summary','Preview limits'),node('p','This is a reading view, not an exact print layout. Basic headings, styling, lists, tables, section/page break hints and supported embedded PNG or JPEG images are shown. Headers, footers, notes, exact pagination, complex layout, historical markup, interactive fields, fonts/themes and some custom numbering may look different or be missing. Links, remote resources, scripts and macros never run. The viewer does not save your document.'));
  const warnings=node('ul');warnings.className='papers-warnings';warnings.hidden=true;
  const view=node('div');view.className='papers-view';view.hidden=true;
  header.append(title,caption,toolbar,status,info,warnings);
  host.replaceChildren(header,view);
  let request=0;
  function reset(){request++;view.replaceChildren();view.hidden=true;warnings.replaceChildren();warnings.hidden=true;clear.disabled=true;status.classList.remove('papers-error');}
  clear.addEventListener('click',()=>{reset();pick.value='';status.textContent='Document closed.';pick.focus();});
  pick.addEventListener('change',async()=>{
    const file=pick.files?.[0];if(!file)return;
    reset();const ticket=request;pick.value='';clear.disabled=false;status.textContent='Opening '+file.name+'…';
    try{
      const result=await renderDocument(file);if(ticket!==request)return;
      view.replaceChildren(result.root);view.hidden=false;
      for(const warning of result.warnings)warnings.append(node('li',warning));
      warnings.hidden=!result.warnings.size;status.textContent=file.name+' · '+result.type.toUpperCase()+' preview';
    }catch(error){if(ticket!==request)return;status.classList.add('papers-error');status.textContent='Could not open '+file.name+'. '+error.message;}
  });
}
