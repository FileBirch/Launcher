import {readFile} from './io.js';
import {address} from './xml.js';

const PAGE_ROWS=40,PAGE_COLS=16,MS_DAY=86400000;
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const dateString=d=>d.toISOString().split('T')[0];
const two=n=>String(n).padStart(2,'0');
function calendarLabel(value,type,date1904){
  const millis=Math.round(value*MS_DAY);
  if(!Number.isSafeInteger(millis))return String(value);
  const ts=Math.abs(millis),day=Math.floor(millis/MS_DAY),ms=((millis%MS_DAY)+MS_DAY)%MS_DAY;
  const clock=n=>two(Math.floor(n/3600000))+':'+two(Math.floor(n/60000)%60)+':'+two(Math.floor(n/1000)%60)+(n%1000?'.'+String(n%1000).padStart(3,'0'):'');
  if(type==='duration')return (millis<0?'−':'')+clock(ts);
  if(type==='time')return value<0?String(value)+' (time serial)':clock(ms);
  if(day<0)return String(value)+' (date serial)';
  if(!date1904&&day===60)return '1900-02-29 (Excel serial 60)'+(type==='datetime'?' '+clock(ms):'');
  const date=new Date((date1904?Date.UTC(1904,0,1):Date.UTC(1899,11,31))+MS_DAY*(day-(!date1904&&day>60?1:0)));
  if(!Number.isFinite(+date)||date.getUTCFullYear()>9999)return String(value)+' (date serial)';
  return dateString(date)+(type==='datetime'?' '+clock(ms):'');
}
export function visibleCell(cell,date1904){
  if(!cell)return '';
  if((cell.f||cell.group||cell.unsupportedFormula)&&!cell.saved)return '[No saved formula result]';
  if(cell.unsupportedType)return '[Unsupported cell type: '+cell.unsupportedType+'] '+String(cell.v??'');
  if(typeof cell.v==='boolean')return cell.v?'TRUE':'FALSE';
  if(typeof cell.v!=='number')return String(cell.v??'');
  const s=cell.style;
  if(s?.dateType)return calendarLabel(cell.v,s.dateType,date1904);
  if(s?.fmt==='percent')return (cell.v*100).toFixed(2)+'%';
  if(s?.fmt==='currency')return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(cell.v);
  if(s?.fmt==='number')return cell.v.toFixed(2);
  return String(cell.v);
}
function sheetTable(sheet,rowStart,colStart){
  const meta=sheet.preview,rowEnd=Math.min(rowStart+PAGE_ROWS-1,meta.rows),colEnd=Math.min(colStart+PAGE_COLS-1,meta.cols);
  if(!meta.rows||!meta.cols)return '<p>The sheet has no visible cell data in the supported grid.</p>';
  const spans=new Map(),covered=new Set();
  for(const m of meta.merges){
    const r=Math.max(m.r,rowStart), c=Math.max(m.c,colStart), bottom=Math.min(m.bottom,rowEnd), right=Math.min(m.right,colEnd);
    if(r>bottom||c>right)continue;
    spans.set(address(r,c),{m,bottom,right});
    for(let y=r;y<=bottom;y++)for(let x=c;x<=right;x++)if(y!==r||x!==c)covered.add(address(y,x));
  }
  let heading='<th scope="col" class="tables-corner" aria-label="Row number"></th>';
  for(let c=colStart;c<=colEnd;c++)heading+='<th scope="col">'+address(1,c).slice(0,-1)+'</th>';
  let grid='';
  for(let r=rowStart;r<=rowEnd;r++){
    let cells='<th scope="row">'+r+'</th>';
    for(let c=colStart;c<=colEnd;c++){
      const here=address(r,c);if(covered.has(here))continue;
      const span=spans.get(here),addr=span?.m.anchor||here,cell=sheet.cells[addr];
      const text=visibleCell(cell,meta.date1904),shown=text.length>480?text.slice(0,480)+'…':text;
      const formula=!!(cell?.f||cell?.group||cell?.unsupportedFormula);
      const style=cell?.style||{};
      const css=(style.bold?'font-weight:700;':'')+(style.italic?'font-style:italic;':'')+(['left','right','center'].includes(style.align)?'text-align:'+style.align+';':'')+
        (/^#[\da-f]{6}$/i.test(style.color||'')?'color:'+style.color+';':'')+(/^#[\da-f]{6}$/i.test(style.fill||'')?'background:'+style.fill+';':'');
      const clipped=span&&(span.m.r!==r||span.m.c!==c||span.m.bottom!==span.bottom||span.m.right!==span.right);
      const range=span?span.m.anchor+':'+address(span.m.bottom,span.m.right):addr;
      const detail=cell||span;
      const value=(formula?'<span aria-hidden="true" class="tables-formula">ƒ</span> ':'')+escape(shown)+(clipped?' <span class="tables-continuation">(merged '+range+')</span>':'');
      cells+='<td'+(span?' rowspan="'+(span.bottom-r+1)+'" colspan="'+(span.right-c+1)+'"':'')+
        (css?' style="'+css+'"':'')+' data-cell="'+here+'">'+(detail?'<button type="button" data-cell-detail="'+addr+'" title="'+range+'">'+(value||'<span aria-label="Empty formatted cell">·</span>')+'</button>':'')+'</td>';
    }
    grid+='<tr>'+cells+'</tr>';
  }
  return '<div class="tables-scroll" tabindex="0" aria-label="Spreadsheet grid"><table aria-label="'+escape(sheet.name)+'"><thead><tr>'+heading+'</tr></thead><tbody>'+grid+'</tbody></table></div>';
}
export async function mount(container){
  let book=null,name='',sheetIndex=0,row=1,col=1,generation=0;
  container.classList.add('tables-viewer');
  container.innerHTML='<header><h2>Open a spreadsheet</h2><p>Read-only · files stay on this device. Open an unencrypted .xlsx, .ods or UTF-8 .csv (up to 20 MiB).</p>'+
    '<label for="tables-input">Browse for a spreadsheet</label> <input id="tables-input" type="file" accept=".xlsx,.ods,.csv">'+
    ' <button type="button" data-clear hidden>Close file</button><p data-result role="status"></p></header><div data-book></div>';
  const input=container.querySelector('input'),result=container.querySelector('[data-result]'),area=container.querySelector('[data-book]'),clear=container.querySelector('[data-clear]');
  const status=text=>{result.textContent=text;};
  function render(){
    if(!book){area.replaceChildren();return;}
    const s=book.sheets[sheetIndex],m=s.preview,rowEnd=Math.min(row+PAGE_ROWS-1,m.rows),colEnd=Math.min(col+PAGE_COLS-1,m.cols);
    area.innerHTML='<h3>'+escape(name)+'</h3><label for="tables-sheets">Worksheet</label> <select id="tables-sheets">'+
      book.sheets.map((s,i)=>'<option value="'+i+'"'+(i===sheetIndex?' selected':'')+'>'+escape(s.name)+(s.preview.hidden?' (hidden in original)':'')+'</option>').join('')+
      '</select><p class="tables-note">Formulas are not recalculated. Values may be stale or missing. Select a cell to inspect its saved value and formula text.</p>'+
      (book.warnings.length?'<details><summary>Reading limits for this file ('+book.warnings.length+')</summary><ul>'+book.warnings.map(w=>'<li>'+escape(w.replaceAll('from exported copies','in this view'))+'</li>').join('')+'</ul></details>':'')+
      (m.rows&&m.cols?'<nav class="tables-page" aria-label="Grid position">'+
        '<button type="button" data-page="up"'+(row===1?' disabled':'')+'>Previous rows</button><button type="button" data-page="down"'+(rowEnd>=m.rows?' disabled':'')+'>Next rows</button>'+
        '<button type="button" data-page="left"'+(col===1?' disabled':'')+'>Previous columns</button><button type="button" data-page="right"'+(colEnd>=m.cols?' disabled':'')+'>Next columns</button>'+
        '<form data-jump><label for="tables-address">Go to cell</label> <input id="tables-address" type="text" inputmode="text" placeholder="A1" maxlength="8" required> <button>Go</button></form>'+
        '<p>Showing rows '+row+'–'+rowEnd+' of '+m.rows+'; columns '+col+'–'+colEnd+' of '+m.cols+'.</p></nav>':'')+
      sheetTable(s,row,col)+'<div data-cell-info aria-live="polite"></div>';
  }
  function inspect(addr){
    const sheet=book.sheets[sheetIndex],c=sheet.cells[addr],info=area.querySelector('[data-cell-info]');
    if(!info)return;
    info.innerHTML='<h4>'+escape(addr)+'</h4><pre>'+escape(visibleCell(c,sheet.preview.date1904))+'</pre>'+
      (c?.f?'<p>Formula text (not run)</p><pre>'+escape(c.f.startsWith('=')||/^(?:of|oooc):=/.test(c.f)?c.f:'='+c.f)+'</pre>':'')+
      (c?.group?'<p>Part of a grouped formula.</p>':'')+(c?.unsupportedFormula?'<p>Unsupported formula type: '+escape(c.unsupportedFormula)+'</p>':'');
  }
  input.addEventListener('change',async()=>{
    const file=input.files[0];if(!file)return;
    const gen=++generation;book=null;name=file.name;row=col=1;sheetIndex=0;clear.hidden=false;render();status('Opening '+file.name+'…');
    try{const next=await readFile(file,{viewOnly:true});if(gen!==generation)return;book=next;status('Opened '+file.name+' · read-only.');render();}
    catch(e){if(gen!==generation)return;book=null;render();status('Could not open '+file.name+': '+(e.message||'Unsupported file.'));}
    finally{if(gen===generation)input.value='';}
  });
  container.addEventListener('change',ev=>{
    if(!book||ev.target.id!=='tables-sheets')return;
    sheetIndex=Number(ev.target.value);row=col=1;render();area.querySelector('#tables-sheets').focus();
  });
  container.addEventListener('click',ev=>{
    if(ev.target.closest('[data-clear]')){generation++;book=null;render();clear.hidden=true;input.value='';status('File closed.');input.focus();return;}
    if(!book)return;
    const detail=ev.target.closest('[data-cell-detail]');if(detail){inspect(detail.dataset.cellDetail);return;}
    const page=ev.target.closest('[data-page]')?.dataset.page;
    if(!page)return;
    if(page==='up')row=Math.max(1,row-PAGE_ROWS);
    if(page==='down')row+=PAGE_ROWS;
    if(page==='left')col=Math.max(1,col-PAGE_COLS);
    if(page==='right')col+=PAGE_COLS;
    render();const same=area.querySelector('[data-page="'+page+'"]');(same&&!same.disabled?same:area.querySelector('[data-jump] input'))?.focus();
  });
  area.addEventListener('submit',ev=>{
    if(!book||!ev.target.matches('[data-jump]'))return;
    ev.preventDefault();
    const addr=ev.target.querySelector('input').value.trim().toUpperCase(),m=/^([A-Z]{1,3})([1-9]\d{0,4})$/.exec(addr);
    const c=m?[...m[1]].reduce((sum,x)=>sum*26+x.charCodeAt(0)-64,0):0,r=m?Number(m[2]):0;
    const meta=book.sheets[sheetIndex].preview;
    if(!r||!c||r>meta.rows||c>meta.cols){status('Cell is outside the displayed worksheet bounds.');return;}
    row=Math.floor((r-1)/PAGE_ROWS)*PAGE_ROWS+1;col=Math.floor((c-1)/PAGE_COLS)*PAGE_COLS+1;render();inspect(addr);
    status('Showing the page containing '+addr+'.');area.querySelector('[data-cell-detail="'+addr+'"]')?.focus();
  });
}
