import {address,column,limits,display,inputText,tsvEscape,parseTSV} from './model.js';
const rowHeight=27,colWidth=116,headWidth=47,headHeight=27;
const clamp=(x,top)=>Math.max(0,Math.min(top-1,x));
function tile(cls,text,x,y,w,h,role){
  const el=document.createElement('div');el.className=cls;el.textContent=text;
  Object.assign(el.style,{left:x+'px',top:y+'px',width:w+'px',height:h+'px'});
  if(role)el.setAttribute('role',role);return el
}
export class Grid {
  constructor(view,callbacks){
    this.view=view;this.cb=callbacks;this.plane=view.firstElementChild;this.layer=this.plane.firstElementChild;
    this.r=this.c=this.er=this.ec=0;this.edit=null;this.frame=0;
    this.plane.style.width=(headWidth+limits.cols*colWidth)+'px';this.plane.style.height=(headHeight+limits.rows*rowHeight)+'px';
    view.addEventListener('scroll',()=>{if(!this.frame)this.frame=requestAnimationFrame(()=>{this.frame=0;this.draw()})});
    new ResizeObserver(()=>this.draw()).observe(view);
    view.addEventListener('pointerdown',e=>{const pos=e.target.closest('[data-address]');if(!pos||e.button!==0)return;
      e.preventDefault();this.finish(true);view.focus({preventScroll:true});
      const ref=pos.dataset.address,now=performance.now();
      const double=this.lastPress?.ref===ref&&now-this.lastPress.time<420&&!e.shiftKey;
      this.lastPress=double?null:{ref,time:now};
      this.select(+pos.dataset.row,+pos.dataset.col,e.shiftKey,false);
      if(double)this.begin();else if(e.pointerType==='mouse')this.drag=true
    });
    view.addEventListener('pointermove',e=>{if(this.drag&&e.buttons===1){const pos=e.target.closest('[data-address]');if(pos&&(this.er!==+pos.dataset.row||this.ec!==+pos.dataset.col))this.select(+pos.dataset.row,+pos.dataset.col,true,false)}});
    window.addEventListener('pointerup',()=>this.drag=false);
    document.addEventListener('pointerdown',e=>{if(!view.contains(e.target))this.lastPress=null},true);
    view.addEventListener('keydown',e=>this.key(e));
    view.addEventListener('copy',e=>{if(this.edit)return;try{const b=this.bounds(),lines=[];this.checkSize();for(let r=b.top;r<=b.bottom;r++){const row=[];for(let c=b.left;c<=b.right;c++)row.push(tsvEscape(this.cb.value(address(r,c))));lines.push(row.join('\t'))}e.clipboardData.setData('text/plain',lines.join('\r\n'));e.preventDefault()}catch(err){e.preventDefault();this.cb.message(err.message)}});
    view.addEventListener('paste',e=>{if(this.edit)return;e.preventDefault();const s=e.clipboardData.getData('text/plain');if(s.length>8000000){this.cb.message('The clipboard selection is too large.');return}this.cb.paste(parseTSV(s))});
  }
  bounds(){return {top:Math.min(this.r,this.er),bottom:Math.max(this.r,this.er),left:Math.min(this.c,this.ec),right:Math.max(this.c,this.ec)}}
  checkSize(){if((Math.abs(this.r-this.er)+1)*(Math.abs(this.c-this.ec)+1)>limits.operation)throw Error('Select up to 50,000 cells at a time.')}
  select(r,c,extend=false,reveal=true){
    r=clamp(r,limits.rows);c=clamp(c,limits.cols);
    if(!extend){this.r=r;this.c=c}this.er=r;this.ec=c;
    if(reveal)this.reveal(r,c);this.draw();this.cb.selected()
  }
  reveal(r,c){const v=this.view,x=c*colWidth,y=r*rowHeight; if(x<v.scrollLeft)v.scrollLeft=x;else if(x+colWidth>v.scrollLeft+v.clientWidth-headWidth)v.scrollLeft=x+colWidth+headWidth-v.clientWidth;
    if(y<v.scrollTop)v.scrollTop=y;else if(y+rowHeight>v.scrollTop+v.clientHeight-headHeight)v.scrollTop=y+rowHeight+headHeight-v.clientHeight
  }
  draw(){
    const v=this.view, x=v.scrollLeft,y=v.scrollTop,b=this.bounds(),fr=Math.floor(y/rowHeight),fc=Math.floor(x/colWidth);
    const lr=Math.min(limits.rows,fr+Math.ceil(v.clientHeight/rowHeight)+1),lc=Math.min(limits.cols,fc+Math.ceil(v.clientWidth/colWidth)+1),frag=document.createDocumentFragment();
    const hrow=document.createElement('div');hrow.setAttribute('role','row');hrow.append(tile('corner','',x,y,headWidth,headHeight,'presentation'));
    for(let c=fc;c<lc;c++)hrow.append(tile('colhead'+(c>=b.left&&c<=b.right?' headsel':''),column(c),headWidth+c*colWidth,y,colWidth,headHeight,'columnheader'));
    frag.append(hrow);
    for(let r=fr;r<lr;r++){
      const row=document.createElement('div');row.setAttribute('role','row');row.setAttribute('aria-rowindex',r+2);
      row.append(tile('rowhead'+(r>=b.top&&r<=b.bottom?' headsel':''),r+1,x,headHeight+r*rowHeight,headWidth,rowHeight,'rowheader'));
      for(let c=fc;c<lc;c++){
        const ref=address(r,c),cell=this.cb.cell(ref),value=this.cb.value(ref),style=cell?.style||{},inRange=r>=b.top&&r<=b.bottom&&c>=b.left&&c<=b.right;
        const el=tile('gridcell'+(typeof value==='number'?' num':'')+(inRange?' sel':'')+(r===this.r&&c===this.c?' active':'')+(cell?.f!==undefined&&String(value).startsWith('#')?' error':''),display(value,style.fmt),headWidth+c*colWidth,headHeight+r*rowHeight,colWidth,rowHeight,'gridcell');
        el.dataset.address=ref;el.dataset.row=r;el.dataset.col=c;el.id='cell-'+ref;el.setAttribute('aria-selected',String(inRange));el.setAttribute('aria-colindex',c+2);el.setAttribute('aria-label',ref+' '+el.textContent);
        if(style.bold)el.style.fontWeight='700';if(style.italic)el.style.fontStyle='italic';if(style.align)el.style.textAlign=style.align;if(style.fill)el.style.backgroundColor=style.fill;if(style.color)el.style.color=style.color;
        row.append(el)
      }
      frag.append(row)
    }
    this.layer.replaceChildren(frag);const active='cell-'+address(this.r,this.c);if(document.getElementById(active))v.setAttribute('aria-activedescendant',active);else v.removeAttribute('aria-activedescendant');
  }
  begin(initial){
    if(this.edit)return;this.select(this.r,this.c);const ref=address(this.r,this.c),initialText=inputText(this.cb.cell(ref)),field=document.createElement('input');
    field.className='cellinput';field.setAttribute('aria-label','Edit '+ref);field.autocomplete='off';field.spellcheck=false;field.value=initial??initialText;
    Object.assign(field.style,{top:(headHeight+this.r*rowHeight)+'px',left:(headWidth+this.c*colWidth)+'px',width:colWidth+'px',height:rowHeight+'px'});
    this.edit={field,ref,initialText};this.plane.append(field);field.focus();field.setSelectionRange(field.value.length,field.value.length);this.cb.drafting(field.value,field.value!==initialText);
    field.oninput=()=>this.cb.drafting(field.value,field.value!==initialText);field.onblur=()=>this.finish(true);
  }
  finish(commit){const e=this.edit;if(!e)return;this.edit=null;e.field.onblur=null;const val=e.field.value;e.field.remove();if(commit&&e.initialText!==val)this.cb.commit(e.ref,val);this.cb.selected()}
  key(e){
    this.lastPress=null;
    if(e.isComposing)return;
    if(this.edit){if(['Enter','Tab','Escape'].includes(e.key)){e.preventDefault();this.finish(e.key!=='Escape');this.view.focus();if(e.key!=='Escape')this.select(this.r+(e.key==='Enter'?(e.shiftKey?-1:1):0),this.c+(e.key==='Tab'?(e.shiftKey?-1:1):0))}return}
    const mod=e.ctrlKey||e.metaKey,k=e.key;
    if(mod){if(k.toLowerCase()==='z'||k.toLowerCase()==='y'){e.preventDefault();this.cb.history(k.toLowerCase()==='y'||e.shiftKey)}else if(k==='Home'){e.preventDefault();this.select(0,0,e.shiftKey)}return}
    const r=e.shiftKey?this.er:this.r,c=e.shiftKey?this.ec:this.c,page=Math.max(1,Math.floor(this.view.clientHeight/rowHeight)-1);
    const moves={ArrowDown:[1,0],ArrowUp:[-1,0],ArrowRight:[0,1],ArrowLeft:[0,-1],PageDown:[page,0],PageUp:[-page,0]};
    if(moves[k]){e.preventDefault();this.select(r+moves[k][0],c+moves[k][1],e.shiftKey);return}
    if(k==='Home'){e.preventDefault();this.select(r,0,e.shiftKey)}
    else if(k==='Enter'||k==='Tab'){e.preventDefault();this.select(this.r+(k==='Enter'?(e.shiftKey?-1:1):0),this.c+(k==='Tab'?(e.shiftKey?-1:1):0))}
    else if(k==='F2'){e.preventDefault();this.begin()}
    else if(k==='Delete'||k==='Backspace'){e.preventDefault();this.cb.clear()}
    else if(!e.altKey&&k.length===1){e.preventDefault();this.begin(k)}
  }
}
