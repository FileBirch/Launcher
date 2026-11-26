import {address, coords, MS, REL, MIME, first, children, parse, isSheetNameExcel} from './xml.js';
import {makeZip,packText} from './package.js';
import {renameQualifier} from './formula-groups.js';

function element(parent,name,attributes={}) {
  const el=parent.ownerDocument.createElementNS(parent.namespaceURI,name);
  for(const [key,value] of Object.entries(attributes)) el.setAttribute(key,value);
  return el;
}
function append(parent,name,attributes) {
  const node=element(parent,name,attributes);
  parent.append(node);
  return node;
}
function uniqueName(base,exists) {
  for(let n=1;;n++){const name=base+n;if(!exists(name))return name;}
}
const encodedString=s=>String(s).replace(/_x[\da-f]{4}_/gi,m=>'_x005F_'+m.slice(1))
  .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g,ch=>'_x'+ch.charCodeAt(0).toString(16).padStart(4,'0')+'_');

function patchValue(node,value,formula,error) {
  for(const old of children(node)) if(['v','is'].includes(old.localName)) old.remove();
  const type=typeof value==='boolean'?'b':typeof value==='number'?'n':error?'e':formula?'str':'inlineStr';
  node.setAttribute('t',type);
  const entry=element(node,type==='inlineStr'?'is':'v');
  node.insertBefore(entry,first(node,'extLst')||null);
  if(type==='inlineStr'){
    const t=append(entry,'t');
    t.setAttributeNS('http://www.w3.org/XML/1998/namespace','xml:space','preserve');
    t.textContent=encodedString(value);
  }else entry.textContent=type==='b'?String(Number(value)):type==='str'?encodedString(value):String(value);
}

// Append only new styles. Existing XF, font, fill and custom format indexes stay valid.
function styleAppender(root,styleXML) {
  const ids=new Map();
  function collection(name,before) {
    let list=first(root,name);
    if(!list){list=element(root,name);root.insertBefore(list,children(root).find(n=>before.includes(n.localName))||null);}
    return list;
  }
  const order=['numFmts','fonts','fills','borders','cellStyleXfs','cellXfs','cellStyles','dxfs','tableStyles','colors','extLst'];
  const defaults=parse(styleXML([{}]));
  for(const name of ['fonts','fills','borders','cellStyleXfs','cellXfs']){
    if(first(root,name))continue;
    const list=collection(name,order.slice(order.indexOf(name)+1));
    for(const node of children(first(defaults,name)))list.append(node.cloneNode(true));
    list.setAttribute('count',list.children.length);
  }
  function add(list,node){
    const id=list.children.length;
    list.append(node.cloneNode(true));list.setAttribute('count',id+1);
    return id;
  }
  return style=>{
    const key=JSON.stringify(style||{});
    if(ids.has(key))return ids.get(key);
    const donor=parse(styleXML([{},style||{}]));
    const xf=children(first(donor,'cellXfs'),'xf')[1];
    for(const [list,attr] of [['fonts','fontId'],['fills','fillId']]){
      const at=Number(xf.getAttribute(attr));
      if(list==='fills' && !at)continue;
      xf.setAttribute(attr,add(first(root,list),first(donor,list).children[at]));
    }
    if(style?.fmt==='currency'){
      const list=collection('numFmts',['fonts','fills','borders','cellStyleXfs','cellXfs']);
      const num=first(first(donor,'numFmts'),'numFmt'),entries=children(list,'numFmt');
      let id=entries.find(x=>x.getAttribute('formatCode')===num.getAttribute('formatCode'))?.getAttribute('numFmtId');
      if(!id){id=Math.max(163,...entries.map(x=>Number(x.getAttribute('numFmtId'))))+1;num.setAttribute('numFmtId',id);add(list,num);}
      xf.setAttribute('numFmtId',id);
    }
    const id=add(first(root,'cellXfs'),xf);ids.set(key,id);return id;
  };
}

function worksheetIndex(ws) {
  const data=first(ws,'sheetData'),rows=new Map(),cells=new Map();
  let row=0;
  for(const node of children(data,'row')){
    row=Number(node.getAttribute('r')||row+1); rows.set(row,node);
    let col=0;
    for(const c of children(node,'c')){
      const ref=c.getAttribute('r')||address(row,col+1);
      col=coords(ref)[1]; cells.set(ref,c);
      // Make implicit positions explicit before inserting/removing adjacent cells.
      if(!c.hasAttribute('r'))c.setAttribute('r',ref);
    }
    if(!node.hasAttribute('r'))node.setAttribute('r',row);
  }
  function ensure(ref){
    if(cells.has(ref))return cells.get(ref);
    const [r,col]=coords(ref);
    let row=rows.get(r);
    if(!row){
      row=element(data,'row',{r});const last=data.lastElementChild;
      data.insertBefore(row,last && Number(last.getAttribute('r'))>r ? children(data,'row').find(n=>Number(n.getAttribute('r'))>r) : null);
      rows.set(r,row);
    }
    const c=element(row,'c',{r:ref});
    row.insertBefore(c,children(row).find(n=>n.localName!=='c'||coords(n.getAttribute('r'))[1]>col)||null);
    cells.set(ref,c);return c;
  }
  return {cells,ensure};
}

export async function patchXlsx(book,calculate,styleXML){
  const {zip,originals,wbName,relsName}=book.xlsx;
  const parts=new Map(zip.names.map(name=>[name,zip.packed(name)]));
  const load=async name=>parse(await zip.text(name));
  const write=async(name,root)=>parts.set(name,await packText(new XMLSerializer().serializeToString(root)));
  const wb=await load(wbName),rels=await load(relsName),types=await load('[Content_Types].xml');
  const sheets=first(wb,'sheets'),dir=wbName.slice(0,wbName.lastIndexOf('/')+1);
  let styleFile=book.xlsx.styleFile, styleRoot,styleId,packageChanged=false;
  function addPart(stem,kind){
    const path=uniqueName(stem,n=>parts.has(n+'.xml'))+'.xml';
    const id=uniqueName('rIdFB',n=>children(rels).some(rel=>rel.getAttribute('Id')===n));
    append(rels,'Relationship',{Id:id,Type:REL+'/'+kind,Target:'/'+path});
    append(types,'Override',{PartName:'/'+path,ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.'+kind+'+xml'});
    packageChanged=true;
    return {id,path};
  }
  for(const [i,sheet] of book.sheets.entries()){
    const original=originals[i];
    if(!isSheetNameExcel(sheet.name))throw Error('Rename '+sheet.name+' before XLSX export.');
    let ws,file;
    if(original){
      file=original.file;ws=await load(file);
      if(sheet.name!==original.name)children(sheets,'sheet').find(s=>(s.getAttributeNS(REL,'id')||s.getAttribute('r:id'))===original.id).setAttribute('name',sheet.name);
    }else{
      const {id,path}=addPart(dir+'worksheets/filebirch','worksheet');file=path;
      const sheetId=Math.max(0,...children(sheets,'sheet').map(s=>Number(s.getAttribute('sheetId'))))+1;
      append(sheets,'sheet',{name:sheet.name,sheetId}).setAttributeNS(REL,'r:id',id);
      ws=parse('<worksheet xmlns="'+MS+'"><sheetData/></worksheet>');
    }
    const {cells,ensure}=worksheetIndex(ws),before=original?.cells||{};
    // Preserve an opaque shared group intact when its formulas are unchanged.
    // If a member was edited, give any remaining opaque members independent groups.
    const opaque=new Set(),modified=new Set();
    let nextShared=0;
    for(const [ref,old] of Object.entries(before)){
      if(old.shared===undefined)continue;
      nextShared=Math.max(nextShared,Number(old.shared)+1);
      if(sheet.cells[ref]?.f!==old.f)modified.add(old.shared);
      else if(calculate.status(i,ref))opaque.add(old.shared);
    }
    let changed=!original;
    for(const ref of new Set([...Object.keys(before),...Object.keys(sheet.cells)])){
      const cell=sheet.cells[ref],old=before[ref];
      if(!cell){cells.get(ref)?.remove();changed=true;continue;}
      const edited=old!==cell,formula=!!(cell.f||cell.group);
      const status=formula ? calculate.status(i,ref) : '';
      const value=formula ? calculate(i,ref) : cell.v;
      const fresh=formula && !status && (!old?.saved || old.v!==value);
      const expand=cell.shared!==undefined && cell.f && !(opaque.has(cell.shared) && !modified.has(cell.shared));
      if(!edited && !fresh && !expand)continue;
      const node=ensure(ref);
      if(edited && JSON.stringify(old?.style)!==JSON.stringify(cell.style)){
        if(cell.style){
          if(!styleRoot){
            if(styleFile)styleRoot=await load(styleFile);
            else{styleFile=addPart(dir+'filebirchStyles','styles').path;styleRoot=parse(styleXML([{}]));}
            styleId=styleAppender(styleRoot,styleXML);
          }
          node.setAttribute('s',styleId(cell.style));
        }else node.removeAttribute('s');
      }
      const contentEdit=edited && (old?.f!==cell.f || old?.v!==cell.v || !old);
      if(expand || (contentEdit && !cell.group)){
        let f=first(node,'f');
        if(!cell.f)f?.remove();
        else{
          if(!f){f=element(node,'f');node.insertBefore(f,node.firstChild);}
          for(const key of ['t','ref','si'])f.removeAttribute(key);
          if(expand && status){
            f.setAttribute('t','shared');f.setAttribute('ref',ref);f.setAttribute('si',nextShared++);
          }
          f.textContent=cell.f;
        }
      }
      if(cell.group?.type==='array' && cell.group.anchor===ref && old?.f!==cell.f){
        first(node,'f').textContent=cell.f;
      }
      if(fresh || (contentEdit && !status)){
        patchValue(node,value,formula,formula ? calculate.isError(i,ref) : cell.error);
      }else if(contentEdit && status && !cell.group && !cell.saved){
        // A newly typed unsupported formula has no result from a spreadsheet engine.
        first(node,'v')?.remove();first(node,'is')?.remove();node.removeAttribute('t');
      }
      changed=true;
    }
    if(changed){
      // dimension is optional. Let readers infer it after inserting or deleting cells.
      first(ws,'dimension')?.remove();
      await write(file,ws);
    }
  }
  // Defined names used by preserved ranges follow renamed worksheets too.
  for(const name of children(first(wb,'definedNames'),'definedName')){
    for(const [i,original] of originals.entries()){
      const current=book.sheets[i].name;
      if(current!==original.name)name.textContent=renameQualifier(name.textContent,original.name,current);
    }
  }
  if(styleRoot)await write(styleFile,styleRoot);
  // Cached results for preserved formulas may predate the edits. Ask Excel to calculate on open.
  let calc=first(wb,'calcPr');
  if(!calc){
    calc=element(wb,'calcPr');
    wb.insertBefore(calc,children(wb).find(n=>['oleSize','customWorkbookViews','pivotCaches','smartTagPr','smartTagTypes','webPublishing','fileRecoveryPr','webPublishObjects','extLst'].includes(n.localName))||null);
  }
  calc.setAttribute('fullCalcOnLoad','1');calc.setAttribute('calcCompleted','0');
  const chain=children(rels).find(n=>n.getAttribute('Type')?.endsWith('/calcChain'));
  if(chain){
    const path=decodeURIComponent(new URL(chain.getAttribute('Target'),'https://package.invalid/'+wbName).pathname).slice(1);
    parts.delete(path);children(types).find(n=>n.getAttribute('PartName')==='/'+path)?.remove();
    chain.remove();packageChanged=true;
  }
  if(packageChanged){await write(relsName,rels);await write('[Content_Types].xml',types);}
  await write(wbName,wb);
  return makeZip([...parts],MIME.xlsx);
}

