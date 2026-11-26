export const limits={rows:10000,cols:256,cells:100000,sheets:32};
export const XML='<?xml version="1.0" encoding="UTF-8"?>';
export const MS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
export const REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const PKG='http://schemas.openxmlformats.org/package/2006/relationships';
export const MIME={xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',ods:'application/vnd.oasis.opendocument.spreadsheet',csv:'text/csv;charset=utf-8'};
export function parse(text){
  if(/<!DOCTYPE|<!ENTITY/i.test(text)) throw Error('XML with document type or entity declarations is unsupported.');
  const doc=new DOMParser().parseFromString(text,'application/xml');
  if(doc.getElementsByTagNameNS('*','parsererror').length)throw Error('A workbook XML part is malformed.');
  return doc.documentElement;
}
export function xml(value){
  const s=String(value);
  if(/[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/u.test(s))throw Error('A cell contains a character that cannot be represented in workbook XML.');
  return s.replace(/[&<>"'\r\n\t]/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;','\r':'&#13;','\n':'&#10;','\t':'&#9;'}[ch]));
}
export function children(n,name){return Array.from(n?.children||[]).filter(c=>!name||c.localName===name)}
export function first(n,name){return children(n,name)[0]}
export function all(n,name){return Array.from(n?.getElementsByTagNameNS('*',name)||[])}
export function attr(n,k){return n?.getAttribute(k)??''}
export function coords(addr){
  const m=/^([A-Z]{1,3})([1-9]\d{0,6})$/.exec(addr);if(!m)throw Error('Invalid cell address.');
  let col=0;for(const c of m[1])col=col*26+c.charCodeAt(0)-64;return [Number(m[2]),col]
}
export function address(row,col){
  let out='';while(col>0){col--;out=String.fromCharCode(65+col%26)+out;col=Math.floor(col/26)}return out+row
}
export function checkBounds(row,col){
  if(row>limits.rows||col>limits.cols||row<1||col<1)throw Error('Content exceeds the 10,000-row or 256-column limit.');
}
export function bookReader(){
  const warns=new Set(),sheets=[];let total=0;
  function addSheet(name){
    if(sheets.length>=limits.sheets)throw Error('The workbook exceeds 32 sheets.');
    if(!name||sheets.some(s=>s.name.toLowerCase()===name.toLowerCase()))throw Error('Empty or duplicate sheet name.');
    const sheet={name,cells:Object.create(null)};sheets.push(sheet);return sheet;
  }
  function put(sheet,addr,cell){
    checkBounds(...coords(addr));if(Object.hasOwn(sheet.cells,addr))throw Error('Overlapping or duplicate cells.');
    if(++total>limits.cells)throw Error('The workbook exceeds 100,000 populated or styled cells.');
    if(String(cell.v).length>32767 || (cell.f?.length||0)>8192)throw Error('A cell exceeds the supported text or formula length.');
    sheet.cells[addr]=cell;
  }
  function result(){if(!sheets.length)throw Error('The package has no supported worksheets.');return {sheets,warnings:[...warns]}}
  return {warn:s=>warns.add(s),addSheet,put,result}
}
export function rowsOf(sheet) {
  const rows=new Map();
  for(const [addr,cell] of Object.entries(sheet.cells)){
    const [r,c]=coords(addr);checkBounds(r,c);
    if(!rows.has(r))rows.set(r,[]); rows.get(r).push({c,addr,cell});
  }
  return [...rows.entries()].sort((a,b)=>a[0]-b[0]).map(([r,items])=>[r,items.sort((a,b)=>a.c-b.c)]);
}
export function stylesOf(book){
  const styles=[{}],ids=new Map([['{}',0]]);
  function id(style){
    const flat={};for(const k of ['bold','italic','align','fmt','color','fill'])if(style?.[k] && !(k==='fmt'&&style[k]==='general'))flat[k]=style[k];
    const key=JSON.stringify(flat);if(!ids.has(key)){ids.set(key,styles.length);styles.push(flat)}return ids.get(key);
  }
  for(const sheet of book.sheets)for(const cell of Object.values(sheet.cells))id(cell.style);
  return {styles,id}
}
export function isSheetNameExcel(name){return name.length>0&&name.length<=31&&!/[\\/:*?\[\]]/.test(name)&&!/^'|'$/.test(name)&&name.toLowerCase()!=='history'}
export const errorValue=v=>typeof v==='string'&&/^#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|NULL!)$/.test(v);

