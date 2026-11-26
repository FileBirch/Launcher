import {readFile,writeFile} from './io.js';
import {makeCalculator,isSupported} from './calc.js';
import {Grid} from './grid.js';
import {renameQualifier} from './formula-groups.js';
import {limits,blank,address,position,inputText,parseInput} from './model.js';

const styleLink=document.createElement('link');
styleLink.rel='stylesheet';styleLink.href=new URL('./editor.css',import.meta.url);document.head.append(styleLink);
const svg=d=>'<svg aria-hidden="true" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"><path d="'+d+'"/></svg>';
const icons={sheet:'M3 4h18v16H3zM3 9h18M3 14h18M9 9v11M15 9v11',open:'M2 19V5h8l2 3h8v3M2 19l4-8h16l-4 8z',save:'M12 3v11m-4-4 4 4 4-4M4 15v5h16v-5',undo:'M8 4 3 9l5 5M3 9h10a7 7 0 0 1 0 14',redo:'m16 4 5 5-5 5m5-5H11a7 7 0 0 0 0 14',expand:'m7 10 5 5 5-5'};
const button=(id,title,content)=>'<button class="toolbtn" id="'+id+'" title="'+title+'" aria-label="'+title+'">'+content+'</button>';
const help='<h3>Editing a workbook</h3><p>Choose a cell and type, double-click it, or edit the formula bar. Enter commits and moves down. Tab moves right. F2 edits the selected cell. Escape cancels the current entry.</p><p>Use the arrow keys to move and Shift + arrows or drag to select a rectangle. The name box jumps to an address. Copy transfers calculated values as tab-separated text. Paste accepts rows, columns, and formulas. Prefix an entry with an apostrophe to keep it as text.</p><h3>Calculations</h3><p>Use A1 references on this sheet or another worksheet, arithmetic, comparisons, SUM, AVERAGE, MIN, MAX, and IF. For example, =SUM(B2:B6) or =IF(B2&gt;0,"Yes","No"). Imported XLSX calculation ranges appear in the compatibility notes. Saved results are labeled in the grid.</p><h3>Files and formatting</h3><p>Open .ods, .xlsx, or UTF-8 CSV. Download .ods or .xlsx for multiple sheets, text, numbers, booleans, supported formulas, bold, italic, alignment, simple colors, and basic two-decimal number, percent, and USD currency formats. CSV contains values from the active sheet only.</p><p>Open and export use local browser memory. Your source file stays on disk. The editor lists detected features that a new copy may lose. Reopen the downloaded copy in the application where you intend to use it.</p><h3>Workbook limits</h3><p>Up to 32 sheets, 10,000 rows and 256 columns per sheet, and 100,000 stored cells per workbook. Each clipboard or formatting operation can cover 50,000 cells. Undo retains up to 50 recent actions, within a 200,000-cell history budget.</p>';

export function mount(host,{brand,base}){
  const app=brand.apps.spreadsheets.name;
  host.className='editor-shell';
  host.innerHTML='<nav class="editorcrumb" aria-label="Breadcrumb"><a id="back">Launcher</a> / <span id="appname"></span></nav><div class="editorbar"><div class="titlegroup"><span class="work-icon appicon">'+svg(icons.sheet)+'</span><div><h1 id="booktitle"></h1><div class="filename" id="filetype"></div></div></div><div class="fileactions"><button id="new" class="secondary">New</button><button id="open" class="secondary">'+svg(icons.open)+' Open file</button><input type="file" id="file" accept=".ods,.xlsx,.csv" hidden><button id="save" class="primary">'+svg(icons.save)+' Download .ods</button><div class="menuwrap"><button id="export" class="secondary" aria-expanded="false" aria-controls="formats">Export '+svg(icons.expand)+'</button><div class="exportmenu" id="formats" hidden><button data-out="xlsx">Excel workbook (.xlsx)</button><button data-out="csv">Active sheet (.csv)</button></div></div></div></div><section class="surface" aria-label="Workbook"><div class="toolrow" role="group" aria-label="Cell formatting and history"><div class="toolgroup">'+button('undo','Undo (Ctrl+Z)',svg(icons.undo))+button('redo','Redo (Ctrl+Y)',svg(icons.redo))+'</div><div class="toolgroup">'+button('bold','Bold','<b>B</b>')+button('italic','Italic','<i>I</i>')+'</div><div class="toolgroup"><select id="numberformat" aria-label="Number format"><option value="general">General</option><option value="number">Number · 0.00</option><option value="percent">Percent · 0.00%</option><option value="currency">Currency · USD</option></select></div><div class="toolgroup"><select id="align" aria-label="Cell alignment"><option value="">Auto align</option><option value="left">Align left</option><option value="center">Center</option><option value="right">Align right</option></select></div><div class="toolgroup"><label class="colorlabel">Text<input id="color" type="color" aria-label="Text color"></label><label class="colorlabel">Fill<input id="fill" type="color" aria-label="Cell fill color"></label>'+button('resetstyle','Clear selected formatting','Clear format')+'</div></div><div class="formulabar"><input id="addr" aria-label="Cell address" spellcheck="false" autocomplete="off"><span class="fx" aria-hidden="true">ƒx</span><input id="formula" aria-label="Cell value or formula" spellcheck="false" autocomplete="off"></div><div class="sheetviewport" role="grid" tabindex="0" aria-label="Worksheet" aria-rowcount="10001" aria-colcount="257"><div class="sheetplane" role="presentation"><div role="presentation"></div></div></div><div class="sheetbar"><div class="tablist" role="tablist" aria-label="Worksheets"></div><button id="addsheet" class="sheetbtn glyph" aria-label="Add worksheet" title="Add worksheet">+</button><button id="renamesheet" class="sheetbtn" title="Rename active worksheet">Rename</button></div></section><div class="statusrow"><div class="statusleft"><span class="saveflag" id="saved"></span><span id="status" role="status"></span><button class="warnlink" id="review" hidden>Review import notes</button></div><span class="cellcount" id="selection"></span></div><div class="editorbottom"><span><strong>Session only.</strong> Download a copy before leaving to keep your changes.</span><button class="plainlink" id="help">Editor guide</button></div><p class="filename" id="devwarning"></p><dialog id="editor-dialog" aria-labelledby="dlgtitle"><h2 id="dlgtitle"></h2><div id="dlgbody" class="helpbody"></div><div class="actions"><button class="secondary" id="cancel">Cancel</button><button class="primary" id="confirm">Continue</button></div></dialog>';
  const $=s=>host.querySelector(s), field=$('#formula'),view=$('.sheetviewport'),dlg=$('#editor-dialog');
  $('#back').href=base.href;$('#appname').textContent=app;$('#devwarning').textContent=brand.warning;
  let book=blank(),sheet=0,title='Untitled workbook',source='',calc=makeCalculator(book),history=[],future=[],rev=0,serial=0,saved=0,pending=false,busy=false,formulaRef='',formulaOriginal='',count=0;
  let formulaNotes=[],groupNotes=[],formatNotes=[],grid;
  const downloadURLs=new Set();
  const sheetNow=()=>book.sheets[sheet],active=()=>address(grid.r,grid.c),message=s=>$('#status').textContent=s;
  const setBusy=v=>{busy=v;host.inert=v;host.setAttribute('aria-busy',String(v))};
  function ask(heading,details,ok='Continue',cancel=true,editName){
    return new Promise(resolve=>{
      $('#dlgtitle').textContent=heading;const body=$('#dlgbody');body.replaceChildren();
      if(typeof details==='string'){const p=document.createElement('p');p.textContent=details;body.append(p)}else if(details)body.append(details);
      let input;if(editName!==undefined){input=document.createElement('input');input.className='dialoginput';input.setAttribute('aria-label','Worksheet name');input.value=editName;input.maxLength=31;body.append(input);input.onkeydown=e=>{if(e.key==='Enter')$('#confirm').click()}}
      $('#confirm').textContent=ok;$('#cancel').hidden=!cancel;let result=null;
      $('#confirm').onclick=()=>{result=input?input.value:true;dlg.close()};$('#cancel').onclick=()=>dlg.close();
      dlg.addEventListener('close',()=>resolve(result),{once:true});dlg.showModal();
      if(input){input.focus();input.select()}else (cancel?$('#cancel'):$('#confirm')).focus()
    })
  }
  function list(items){const ul=document.createElement('ul');ul.className='dialoglist';for(const s of items){const l=document.createElement('li');l.textContent=s;ul.append(l)}return ul}
  function updateNotes(){
    groupNotes=[];formatNotes=[];
    for(const [i,s] of book.sheets.entries()){
      const ranges=new Set(), unavailable=new Set();
      for(const [ref,c] of Object.entries(s.cells)){
        if(!c.group && !(book.xlsx && c.f))continue;
        const label=s.name+'!'+(c.group?.ref||ref);
        if(c.group && !ranges.has(c.group)){
          ranges.add(c.group);
          formatNotes.push(label+' ('+c.group.label+')');
        }
        if(calc.status(i,ref))unavailable.add(label+' · '+(c.group?.label||'formula')+' cannot recalculate. Cells show the last saved value or “No saved result”.');
      }
      groupNotes.push(...unavailable);
    }
    let n=0;const examples=[];for(const s of book.sheets)for(const [ref,cell] of Object.entries(s.cells))if(!book.xlsx&&cell.f!==undefined&&!isSupported(cell.f)){n++;if(examples.length<4)examples.push(s.name+'!'+ref)}
    formulaNotes=n?[n+' unsupported formula'+(n===1?'':'s')+' ('+examples.join(', ')+(n>4?', …':'')+'). Workbook downloads convert these expressions to literal text, which can change dependent formulas. CSV exports the expressions with a safety prefix.']:[];$('#review').hidden=!book.warnings.length&&!n&&!groupNotes.length
  }
  function saveFlag(){const flag=$('#saved'),dirty=rev!==saved||pending;flag.classList.toggle('changed',dirty);flag.textContent=dirty?'Unsaved changes':source||saved!==0?'No pending changes':'Blank workbook';$('#undo').disabled=!history.length;$('#redo').disabled=!future.length}
  function selected(){
    pending=false;const ref=active(),cell=sheetNow().cells[ref],sty=cell?.style||{};
    $('#addr').value=ref;field.readOnly=!!cell?.group;
    field.value=cell?.group ? groupText(cell.group) : inputText(cell);formulaRef=ref;formulaOriginal=field.value;
    for(const k of ['bold','italic'])$('#'+k).setAttribute('aria-pressed',String(!!sty[k]));
    $('#numberformat').value=sty.fmt||'general';$('#align').value=sty.align||'';$('#color').value=sty.color||'#292735';$('#fill').value=sty.fill||'#ffffff';
    const end=address(grid.er,grid.ec),state=calc.status(sheet,ref);$('#selection').textContent=sheetNow().name+' · '+ref+(ref===end?'':'–'+end)+(state?' · '+state:'');saveFlag()
  }
  function groupText(g){
    if(g.type==='dataTable')return 'TABLE('+[g.input1,g.input2].filter(Boolean).join(', ')+') · '+g.ref;
    return g.expression?'='+g.expression:g.label+' · '+g.ref;
  }
  function tabs(){
    const bar=$('.tablist');bar.replaceChildren();book.sheets.forEach((s,i)=>{const b=document.createElement('button');b.className='sheetbtn';b.setAttribute('role','tab');b.setAttribute('aria-selected',String(i===sheet));b.tabIndex=i===sheet?0:-1;b.textContent=s.name;
      const switchSheet=j=>{sheet=j;tabs();grid.select(0,0)};b.onclick=()=>switchSheet(i);b.onkeydown=e=>{if(['ArrowLeft','ArrowRight','Home','End'].includes(e.key)){e.preventDefault();const n=book.sheets.length;switchSheet(e.key==='Home'?0:e.key==='End'?n-1:(i+(e.key==='ArrowLeft'?-1:1)+n)%n);bar.children[sheet].focus()}};bar.append(b)
    });view.setAttribute('aria-label',sheetNow().name+' worksheet');$('#addsheet').disabled=book.sheets.length>=limits.sheets
  }
  function refresh(){calc=makeCalculator(book);updateNotes();grid.draw();selected()}
  function apply(action,forward){
    const s=book.sheets[action.sheet];
    if(action.type==='cells')for(const change of action.changes){
      const next=forward?change.after:change.before,had=Object.hasOwn(s.cells,change.ref);
      if(next){s.cells[change.ref]=next;if(!had)count++}else{delete s.cells[change.ref];if(had)count--}
    }
    else if(action.type==='add'){if(forward)book.sheets.push(action.added);else book.sheets.pop()}
    else if(action.type==='rename'){
      if(!action.formulas){
        action.formulas=[];
        const groups=new Map();
        for(const owner of book.sheets)for(const [ref,cell] of Object.entries(owner.cells)){
          const f=cell.f && renameQualifier(cell.f,action.before,action.after);
          const g=cell.group;
          if(g && !groups.has(g)){
            const expression=renameQualifier(g.expression,action.before,action.after);
            const source=g.source && {...g.source,qualifier:renameQualifier(g.source.qualifier,action.before,action.after)};
            groups.set(g,expression===g.expression?g:{...g,expression,source});
          }
          if(f!==cell.f || (g && groups.get(g)!==g)){
            const after={...cell};
            if(f!==undefined)after.f=f;
            if(g)after.group=groups.get(g);
            action.formulas.push({owner,ref,before:cell,after});
          }
        }
      }
      for(const entry of action.formulas)entry.owner.cells[entry.ref]=forward?entry.after:entry.before;
      s.name=forward?action.after:action.before;
    }
    sheet=forward&&action.type==='add'?book.sheets.length-1:action.sheet;rev=forward?action.next:action.prev;
    if(action.type!=='cells')tabs()
  }
  function record(act){act.sheet=sheet;act.prev=rev;act.next=++serial;apply(act,true);history.push(act);future=[];
    let size=history.reduce((n,a)=>n+(a.changes?.length||1),0);while(history.length>50||size>200000){size-=history.shift().changes?.length||1}
    tabs();refresh();message('')
  }
  function undo(redo){const from=redo?future:history,to=redo?history:future,a=from.pop();if(!a)return;apply(a,redo);to.push(a);tabs();refresh();message(redo?'Action redone.':'Action undone.')}
  function changeCells(entries,styleOnly=false){
    const changes=[];let delta=0;const cells=sheetNow().cells;
    for(const [ref,after] of entries){const before=cells[ref];if(before?.group && !styleOnly){message('Protected '+before.group.label+' at '+sheetNow().name+'!'+before.group.ref+'.');return false}if(JSON.stringify(before)===JSON.stringify(after))continue;changes.push({ref,before,after});delta+=(after?1:0)-(before?1:0)}
    if(count+delta>limits.cells){message('This edit exceeds the 100,000-cell workbook limit.');return false}
    if(changes.length)record({type:'cells',changes});return true
  }
  function commit(ref,text){if(text.length>32767){message('A cell can contain up to 32,767 characters.');return}changeCells([[ref,parseInput(text,sheetNow().cells[ref]?.style)]])}
  function rangeMap(fn,styleOnly=false){
    try{grid.checkSize();const b=grid.bounds(),entries=[];for(let r=b.top;r<=b.bottom;r++)for(let c=b.left;c<=b.right;c++){const ref=address(r,c);entries.push([ref,fn(sheetNow().cells[ref])])}changeCells(entries,styleOnly)}catch(e){message(e.message)}
  }
  function format(k,v){rangeMap(cell=>{const sty={...cell?.style};if(v===undefined||v===''||v===false||v==='general')delete sty[k];else sty[k]=v;
    const baseCell={...cell,v:cell?.v??''};delete baseCell.style;if(Object.keys(sty).length)baseCell.style=sty;
    return baseCell.v===''&&baseCell.f===undefined&&!baseCell.group&&!baseCell.style?undefined:baseCell},true);view.focus({preventScroll:true})
  }
  function paste(rows){
    const w=Math.max(0,...rows.slice(0,limits.rows+1).map(x=>x.length)),total=rows.reduce((n,r)=>n+r.length,0);
    if(total>limits.operation||grid.r+rows.length>limits.rows||grid.c+w>limits.cols){message('Paste is too large for this selection or worksheet.');return}
    const entries=[];for(let r=0;r<rows.length;r++)for(let c=0;c<rows[r].length;c++){if(rows[r][c].length>32767){message('A pasted cell exceeds 32,767 characters.');return}const ref=address(grid.r+r,grid.c+c);entries.push([ref,parseInput(rows[r][c],sheetNow().cells[ref]?.style)])}
    if(changeCells(entries)){grid.select(grid.r+rows.length-1,grid.c+Math.max(w-1,0),true,false);message('Pasted '+total+' cell'+(total===1?'':'s')+'.')}
  }
  grid=new Grid(view,{cell:ref=>sheetNow().cells[ref],status:ref=>calc.status(sheet,ref),value:ref=>calc(sheet,ref),selected,commit,history:undo,clear:()=>rangeMap(c=>parseInput('',c?.style)),message,paste,drafting:(s,dirty)=>{field.value=s;pending=dirty;saveFlag()}});
  field.onfocus=()=>{formulaRef=active();formulaOriginal=field.value};
  field.oninput=()=>{pending=field.value!==formulaOriginal;saveFlag()};
  function settleField(){if(!field.readOnly && field.value!==formulaOriginal)commit(formulaRef,field.value);selected()}
  field.onblur=settleField;field.onkeydown=e=>{if(e.isComposing)return;if(e.key==='Enter'||e.key==='Escape'){e.preventDefault();if(e.key==='Escape')field.value=formulaOriginal;view.focus();if(e.key==='Enter')grid.select(grid.r+(e.shiftKey?-1:1),grid.c)}};
  $('#addr').onkeydown=e=>{if(e.key==='Enter'){e.preventDefault();const pos=position(e.target.value);if(pos){grid.select(pos.r,pos.c);view.focus()}else{message('Enter a cell address between A1 and IV10000.');e.target.select()}}else if(e.key==='Escape'){selected();view.focus()}};
  $('#addr').onfocus=e=>e.target.select();$('#addr').onblur=()=>$('#addr').value=active();
  for(const k of ['bold','italic'])$('#'+k).onclick=()=>format(k,!sheetNow().cells[active()]?.style?.[k]);
  for(const [id,k] of [['numberformat','fmt'],['align','align'],['color','color'],['fill','fill']])$('#'+id).onchange=e=>format(k,e.target.value);
  $('#resetstyle').onclick=()=>rangeMap(c=>{if(!c)return undefined;const clean={...c};delete clean.style;return !clean.group&&clean.f===undefined&&clean.v===''?undefined:clean},true);
  $('#undo').onclick=()=>undo(false);$('#redo').onclick=()=>undo(true);
  $('#addsheet').onclick=()=>{let n=book.sheets.length+1;const used=new Set(book.sheets.map(s=>s.name.toLowerCase()));while(used.has('sheet'+n))n++;record({type:'add',added:{name:'Sheet'+n,cells:{}}});grid.select(0,0);view.focus()};
  $('#renamesheet').onclick=async()=>{const name=await ask('Rename worksheet','Use a distinct name of up to 31 characters.','Rename',true,sheetNow().name);if(name===null||name===sheetNow().name)return;
    if(!name.trim()||/[\\/:?*\[\]\x00-\x1f]/.test(name)||/^'|'$/.test(name)||book.sheets.some((s,i)=>i!==sheet&&s.name.toLowerCase()===name.toLowerCase())){message('Choose a unique worksheet name without reserved characters.');return}record({type:'rename',before:sheetNow().name,after:name})};
  $('#help').onclick=()=>{const div=document.createElement('div');div.innerHTML=help;ask('Editor guide',div,'Close',false)};
  $('#review').onclick=()=>ask('Workbook compatibility',list(importNotes()),'Close',false);
  function importNotes(){
    return book.xlsx ? [...groupNotes,'XLSX export retains the original calculation records and package parts, including embedded content. Other formats have a separate export review.'] : [...book.warnings,...formulaNotes];
  }
  function reset(newBook,name,ext){
    book=newBook;title=name;source=ext;sheet=0;history=[];future=[];serial=0;saved=0;rev=0;pending=false;calc=makeCalculator(book);count=book.sheets.reduce((n,s)=>n+Object.keys(s.cells).length,0);
    $('#booktitle').textContent=title;$('#filetype').textContent=ext?ext.toUpperCase()+' · Local working copy':'Local workbook';tabs();grid.select(0,0);updateNotes();message('')
  }
  const allowReplace=()=>rev===saved&&!pending||ask('Discard unsaved changes?','Replacing this workbook discards edits kept in this session. Download a copy first if you want to keep them.','Discard edits');
  $('#new').onclick=async()=>{if(await allowReplace()){reset(blank(),'Untitled workbook','');view.focus()}};
  $('#open').onclick=()=>$('#file').click();$('#file').onchange=async e=>{
    const file=e.target.files[0];e.target.value='';if(!file||busy)return;
    // Retain the current workbook if reading or validating the replacement fails.
    try{setBusy(true);message('Reading local file…');const result=await readFile(file);setBusy(false);
      if(!await allowReplace()){message('Open cancelled.');return}reset(result,file.name.replace(/\.[^.]+$/,'')||'Workbook',file.name.split('.').pop());message('Opened a local copy.');if(book.warnings.length||formulaNotes.length)await ask('Import notes',list(importNotes()),'Continue',false)
    }catch(err){setBusy(false);message('File could not be opened.');await ask('Open failed',err.message,'Close',false)}
  };
  const menu=$('#formats'),exportButton=$('#export');
  const closeMenu=()=>{menu.hidden=true;exportButton.setAttribute('aria-expanded','false')};
  exportButton.onclick=()=>{if(menu.hidden){menu.hidden=false;exportButton.setAttribute('aria-expanded','true');menu.firstElementChild.focus()}else closeMenu()};
  document.addEventListener('pointerdown',e=>{if(!e.target.closest('.menuwrap'))closeMenu()});
  menu.onkeydown=e=>{if(e.key==='Escape'){closeMenu();exportButton.focus()}};
  menu.onfocusout=e=>{if(!e.relatedTarget?.closest('.menuwrap'))closeMenu()};
  async function download(format){
    if(busy)return;closeMenu();grid.finish(true);
    const reasons=format==='xlsx' && book.xlsx ? [] : [...book.warnings,...formulaNotes];
    if(format!=='xlsx' && book.xlsx){
      for(const item of formatNotes)reasons.unshift(item+' will contain displayed values only. XLSX preserves the calculation.');
      for(const note of groupNotes)reasons.push(note);
      if(groupNotes.length)reasons.push('This copy stores saved results as values and leaves missing results blank. XLSX retains the original formulas and cache records.');
    }
    if(format==='csv')reasons.unshift('Only the active worksheet, “'+sheetNow().name+'”, will be included. CSV contains calculated values and drops formatting and formula definitions. Formula-like text is prefixed with an apostrophe for safe opening.');
    if(reasons.length){
      const details=list(reasons);
      const alternate=format==='ods' && book.xlsx && (formatNotes.length || groupNotes.length);
      let useXlsx=false;
      if(alternate){
        const choice=document.createElement('button');choice.type='button';choice.className='secondary';
        choice.textContent='Download XLSX with formulas';
        choice.onclick=()=>{useXlsx=true;$('#confirm').click()};
        const body=document.createElement('div');body.append(details,choice);
        if(!await ask('Review export changes',body,'Download ODS copy')){message('Download cancelled.');return}
        if(useXlsx)format='xlsx';
      }else if(!await ask('Review export changes',details,'Download copy')){message('Download cancelled.');return}
    }
    try{
      setBusy(true);message('Preparing download…');let out=book,value=calc;
      if(book.xlsx && format!=='xlsx'){
        out={...book,sheets:book.sheets.map((s,i)=>({...s,cells:Object.fromEntries(Object.entries(s.cells).map(([ref,cell])=>{
          if(!cell.group && !(cell.f && calc.status(i,ref)))return [ref,cell];
          return [ref,{v:calc(i,ref),style:cell.style}];
        }))}))};
      }
      if(formulaNotes.length){out={...book,sheets:book.sheets.map(s=>({...s,cells:Object.fromEntries(Object.entries(s.cells).map(([ref,cell])=>{
        if(cell.f===undefined||isSupported(cell.f))return [ref,cell];const copy={...cell,v:'='+cell.f};delete copy.f;return [ref,copy]
      }))}))};value=format==='csv'?calc:makeCalculator(out)}
      const blob=await writeFile(out,format,sheet,value);const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;
      const filename=title+(format==='csv'?' - '+sheetNow().name:'')+' - copy.'+format;a.download=filename.replace(/[<>:"/\\|?*\x00-\x1f]/g,'_');document.body.append(a);downloadURLs.add(url);a.click();a.remove();setTimeout(()=>{URL.revokeObjectURL(url);downloadURLs.delete(url)},30000);
      if(format==='xlsx' && book.xlsx || format!=='csv'&&!formulaNotes.length&&!formatNotes.length&&!groupNotes.length)saved=rev;
      message('Download requested. Check your Downloads folder.');saveFlag()
    }catch(err){setBusy(false);message('Download failed.');await ask('Export failed',err.message,'Close',false)}finally{setBusy(false)}
  }
  $('#save').onclick=()=>download('ods');menu.querySelectorAll('button').forEach(b=>b.onclick=()=>download(b.dataset.out));
  addEventListener('beforeunload',e=>{if(rev!==saved||pending){e.preventDefault();e.returnValue=''}});
  addEventListener('pagehide',()=>{
    // A browser may keep a departing page in its back/forward cache.
    // Release workbook state and remove cell text from its DOM snapshot.
    for(const url of downloadURLs)URL.revokeObjectURL(url);
    downloadURLs.clear();grid.finish(false);book=blank();sheet=0;history=[];future=[];
    calc=makeCalculator(book);formulaNotes=[];formulaRef='';formulaOriginal='';title='';source='';
    saved=rev;pending=false;field.value='';grid.layer.replaceChildren();dlg.replaceChildren();host.replaceChildren();
  },{once:true});
  document.addEventListener('keydown',e=>{if(dlg.open||document.querySelector('#termsdialog')?.open||busy)return;if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==='s'){e.preventDefault();if(document.activeElement===field)field.blur();download('ods')}});
  reset(book,title,source)
}
