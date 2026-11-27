import {readDeck} from './deck.js';
const element=(name,text,className)=>{
  const node=document.createElement(name);if(text!=null)node.textContent=text;if(className)node.className=className;return node;
};
function button(text,action){const node=element('button',text);node.type='button';node.addEventListener('click',action);return node;}
const length=(n,min=0,max=10000)=>Math.min(Math.max(Number.isFinite(n)?n:0,min),max)+'px';
function draw(slide){
  const canvas=element('div',null,'fbpr-slide');
  canvas.setAttribute('aria-label',slide.title);
  Object.assign(canvas.style,{width:length(slide.width),height:length(slide.height),backgroundColor:slide.background||'#ffffff'});
  for(const data of slide.items){
    const item=element('div',null,'fbpr-item');
    Object.assign(item.style,{left:length(data.x,-10000),top:length(data.y,-10000),width:length(data.width),height:length(data.height),
      backgroundColor:data.fill||'',borderColor:data.line||'',borderWidth:data.line?'1px':'',borderStyle:data.line?'solid':'',
      borderRadius:data.shape==='ellipse'?'50%':data.shape==='roundRect'?'10px':'',transform:data.rotation?'rotate('+data.rotation+'deg)':''});
    if(data.placeholder){item.classList.add('fbpr-placeholder');item.textContent=data.placeholder;}
    else if(data.image){
      const image=element('img');image.alt=data.name||'Embedded slide image';image.src=data.image;
      image.addEventListener('error',()=>{item.classList.add('fbpr-placeholder');item.textContent='Image not shown';},{once:true});item.append(image);
    }else{
      for(const p of data.paragraphs||[]){
        const paragraph=element('p');
        if(p.align)paragraph.style.textAlign=p.align;
        for(const r of p.runs){
          if(r.break){paragraph.append(element('br'));continue;}
          const run=element('span',r.text);
          if(r.size)run.style.fontSize=length(r.size,6,240);
          if(r.bold)run.style.fontWeight='bold';
          if(r.italic)run.style.fontStyle='italic';
          if(r.color)run.style.color=r.color;
          paragraph.append(run);
        }
        if(!paragraph.childNodes.length)paragraph.append(element('br'));
        item.append(paragraph);
      }
    }
    canvas.append(item);
  }
  return canvas;
}
export function mountViewer(host){
  let deck=null,index=0,generation=0,original='';
  const wrapper=element('section',null,'fbpr');
  host.replaceChildren(wrapper);
  const head=element('div',null,'fbpr-header');
  const title=element('h1','Present'); const badge=element('span','Read-only','fbpr-badge');
  const label=element('label','Open a PowerPoint or OpenDocument presentation');
  const input=element('input');input.type='file';input.accept='.pptx,.odp';label.append(input);
  const shut=button('Close document',()=>{generation++;deck=null;index=0;original='';input.value='';message.textContent='Document closed.';refresh();});shut.disabled=true;
  head.append(title,badge,label,shut);
  const explain=element('p','Files stay on this device. The preview approximates fonts, theme and master layouts. Animations, embedded media, charts, interactive links, and external content do not run. No editing or export is available.');
  const message=element('p','Choose a .pptx or .odp file.','fbpr-message');message.setAttribute('role','status');
  const tools=element('nav',null,'fbpr-nav');tools.setAttribute('aria-label','Slide controls');
  const prev=button('Previous slide',()=>go(index-1)),next=button('Next slide',()=>go(index+1));
  const picker=element('select');picker.setAttribute('aria-label','Go to slide');picker.addEventListener('change',()=>go(Number(picker.value)));
  tools.append(prev,picker,next);
  const stage=element('div',null,'fbpr-stage');stage.tabIndex=0;stage.setAttribute('aria-label','Slide preview');
  const warnings=element('ul',null,'fbpr-warnings');
  wrapper.append(head,explain,message,tools,stage,warnings);
  stage.addEventListener('keydown',event=>{
    const delta=event.key==='ArrowLeft'||event.key==='PageUp'?-1:event.key==='ArrowRight'||event.key==='PageDown'?1:0;
    if(delta){event.preventDefault();go(index+delta);}
    else if(deck&&['Home','End'].includes(event.key)){event.preventDefault();go(event.key==='Home'?0:deck.slides.length-1);}
  });
  function fit(){
    const s=deck?.slides[index],current=stage.firstElementChild;if(!s||!current||!stage.clientWidth)return;
    const scale=Math.min(stage.clientWidth/s.width, Math.min(0.73*innerHeight,760)/s.height,1.5);
    stage.style.height=length(s.height*scale);current.style.transform='scale('+scale+')';
  }
  new ResizeObserver(fit).observe(stage);
  function go(i){if(!deck||i<0||i>=deck.slides.length||i===index)return;index=i;refresh();}
  function refresh(){
    tools.hidden=stage.hidden=!deck;shut.disabled=!deck;warnings.replaceChildren();stage.replaceChildren();
    if(!deck)return;
    const s=deck.slides[index];
    message.textContent=original+' · '+deck.format+' · Slide '+(index+1)+' of '+deck.slides.length;
    picker.value=String(index);prev.disabled=index===0;next.disabled=index===deck.slides.length-1;
    for(const note of s.notes)warnings.append(element('li',note));
    stage.append(draw(s));fit();
  }
  input.addEventListener('change',async()=>{
    const file=input.files?.[0];if(!file)return;
    const request=++generation;deck=null;index=0;original='';refresh();shut.disabled=false;message.textContent='Reading presentation locally…';
    try{
      const loaded=await readDeck(file); if(generation!==request)return;
      deck=loaded;original=file.name;picker.replaceChildren(...loaded.slides.map((s,i)=>{const o=element('option',(i+1)+'. '+s.title);o.value=String(i);return o;}));refresh();
    }catch(err){if(generation===request){shut.disabled=true;message.textContent=err?.message||'This presentation could not be opened.';}}
    finally{if(generation===request)input.value='';}
  });
  refresh();
}
