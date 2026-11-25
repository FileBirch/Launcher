export const limits={rows:10000,cols:256,sheets:32,cells:100000,operation:50000};
export function column(n){let out='';for(n++;n>0;n=Math.floor((n-1)/26))out=String.fromCharCode(65+(n-1)%26)+out;return out}
export function address(r,c){return column(c)+(r+1)}
export function position(ref){const m=/^([A-Z]{1,3})([1-9][0-9]{0,4})$/i.exec(ref.trim());if(!m)return null;let c=0;for(const ch of m[1].toUpperCase())c=c*26+ch.charCodeAt(0)-64;const r=Number(m[2]);return c<=limits.cols&&r<=limits.rows?{r:r-1,c:c-1}:null}
export const blank=()=>({sheets:[{name:'Sheet1',cells:{}}],warnings:[]});
export function inputText(cell){return cell?.f!==undefined?'='+cell.f:String(cell?.v??'')}
export function parseInput(text,style){
  if(text==='')return style&&Object.keys(style).length?{v:'',style}:undefined;
  let v=text,f;
  if(text.startsWith("'"))v=text.slice(1);
  else if(text.startsWith('=')){f=text.slice(1);v=''}
  else if(/^(?:true|false)$/i.test(text))v=text.toLowerCase()==='true';
  else if(/^[+-]?(?:0(?:\.\d*)?|[1-9]\d*(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(text)&&Number.isFinite(Number(text)))v=Number(text);
  return {v,...(f===undefined?{}:{f}),...(style?{style}:{})}
}
const numbers={number:new Intl.NumberFormat('en-US',{useGrouping:false,minimumFractionDigits:2,maximumFractionDigits:2}),percent:new Intl.NumberFormat('en-US',{style:'percent',minimumFractionDigits:2,maximumFractionDigits:2}),currency:new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'})};
export function display(value,format){if(typeof value==='number'&&numbers[format])return numbers[format].format(value);return typeof value==='boolean'?String(value).toUpperCase():String(value??'')}
export const tsvEscape=value=>{const s=String(value??'');return /["\r\n\t]/.test(s)?'"'+s.replaceAll('"','""')+'"':s};
export function parseTSV(text){
  const rows=[];let row=[],s='',quoted=false;
  for(let i=0;i<text.length;i++){const ch=text[i];
    if(ch==='"'&&quoted&&text[i+1]==='"'){s+='"';i++}
    else if(ch==='"'&&(quoted||s===''))quoted=!quoted;
    else if(!quoted&&ch==='\t'){row.push(s);s=''}
    else if(!quoted&&(ch==='\r'||ch==='\n')){if(ch==='\r'&&text[i+1]==='\n')i++;row.push(s);rows.push(row);row=[];s=''}
    else s+=ch
  }
  if(row.length||s||!rows.length||!/[\r\n]$/.test(text)){row.push(s);rows.push(row)}
  return rows
}