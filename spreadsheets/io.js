import {initPreview,finishPreview} from './preview-metadata.js';
import {openZip,MAX_INPUT,utf8} from './package.js';
import {bookReader,checkBounds,address,rowsOf,MIME} from './xml.js';
import {readXlsx,writeXlsx} from './xlsx.js';
import {readOds,writeOds} from './ods.js';
const numeric=/^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/;
function csvRead(text,viewOnly){
  const ctx=bookReader(),sheet=ctx.addSheet('Sheet1');if(viewOnly)initPreview(sheet);let r=1,c=1,field='',quoted=false,closed=false;
  const put=()=>{
    checkBounds(r,c);
    if(viewOnly){sheet.preview.rows=Math.max(sheet.preview.rows,r);sheet.preview.cols=Math.max(sheet.preview.cols,c);}
    if(field!==''){
      let v=field;
      if(!viewOnly&&numeric.test(field)&&field.replace(/[^0-9]/g,'').length<=15&&Number.isFinite(Number(field)))v=Number(field);
      ctx.put(sheet,address(r,c),{v});
    }
    field='';closed=false;c++;
  };
  for(let i=0;i<text.length;i++){
    const ch=text[i];
    if(ch==='"'){
      if(quoted&&text[i+1]==='"'){field+='"';i++}
      else if(quoted){quoted=false;closed=true}
      else if(!field.length&&!closed)quoted=true;
      else throw Error('Unexpected quote in CSV on row '+r+'.');
    }else if(quoted)field+=ch;
    else if(ch===',')put();
    else if(ch==='\r'||ch==='\n'){if(ch==='\r'&&text[i+1]==='\n')i++;put();r++;c=1}
    else if(closed)throw Error('Unexpected text after a closing CSV quote on row '+r+'.');
    else field+=ch;
    if(field.length>32767)throw Error('A CSV field exceeds 32,767 characters.');
  }
  if(quoted)throw Error('The CSV contains an unclosed quoted field.');
  if(field||closed||c>1)put();
  if(viewOnly)finishPreview(sheet);
  return ctx.result();
}
function csvWrite(sheet,index,getValue){
  const rows=rowsOf(sheet);
  const lastCol=rows.reduce((n,[,items])=>Math.max(n,items.at(-1).c),1);
  const lines=[];let next=1;
  const blank=','.repeat(lastCol-1);
  for(const [r,items] of rows){
    while(next<r){lines.push(blank);next++}
    const fields=Array(lastCol).fill('');
    for(const {c,cell,addr} of items){
      const v=cell.f?getValue(index,addr):cell.v;
      let s=String(v??'');
      if(typeof v==='string'&&/^(?:\s*[=+@-]|[\t\r\n])/.test(s))s="'"+s;
      if(/[",\r\n]/.test(s))s='"'+s.replaceAll('"','""')+'"';
      fields[c-1]=s;
    }
    lines.push(fields.join(','));next++;
  }
  const blob=new Blob(['\ufeff',lines.join('\r\n'),'\r\n'],{type:MIME.csv});
  if(blob.size>MAX_INPUT)throw Error('The CSV would exceed 20 MiB. Reduce its contents first.');return blob
}
export async function readFile(file,{viewOnly=false}={}){
  if(!file||!file.size)throw Error('Choose a nonempty local file.');
  if(file.size>MAX_INPUT)throw Error('The file exceeds the 20 MiB input limit.');
  const ext=file.name.split('.').at(-1).toLowerCase();
  if(!['xlsx','ods','csv'].includes(ext))throw Error('Supported inputs are .ods, .xlsx and UTF-8 .csv. Macro-enabled and legacy workbooks are unsupported.');
  const bytes=await file.arrayBuffer();
  if(ext==='csv'){
    let source;try{source=utf8.decode(bytes)}catch{throw Error('CSV must use UTF-8 text encoding.')}
    if(source.includes('\0'))throw Error('CSV contains null characters or a non-UTF-8 text encoding.');
    return csvRead(source,viewOnly);
  }
  const zip=openZip(bytes);
  if(ext==='ods'){
    if(!zip.names.includes('mimetype')||(await zip.text('mimetype')).trim()!==MIME.ods)throw Error('Invalid OpenDocument spreadsheet.');
    return readOds(zip,{viewOnly})
  }
  return readXlsx(zip,{viewOnly});
}
export async function writeFile(book,format,activeSheetIndex,getValue){
  switch(format.toLowerCase()){
    case 'ods':return writeOds(book,getValue);
    case 'xlsx':return writeXlsx(book,getValue);
    case 'csv':return csvWrite(book.sheets[activeSheetIndex],activeSheetIndex,getValue);
    default:throw Error('Unsupported download format.');
  }
}
