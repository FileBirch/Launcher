// Metadata used only by the read-only viewer. Editing import/export behavior is unchanged.
import {address,checkBounds,coords,limits} from './xml.js';
export function addMerge(sheet,first,last){
  const [r,c]=coords(first),[bottom,right]=coords(last);
  checkBounds(r,c);checkBounds(bottom,right);
  if(bottom<r||right<c)throw Error('Invalid merged cell range.');
  if(bottom===r&&right===c)return;
  const meta=sheet.preview;
  if(meta.merges.length>=10000)throw Error('The sheet exceeds 10,000 merged ranges.');
  // Check only intersected cells and bound memory to the supported worksheet dimensions.
  meta.mergeUse??=new Map();
  for(let y=r;y<=bottom;y++){
    if(!meta.mergeUse.has(y))meta.mergeUse.set(y,new Uint8Array(limits.cols+1));
    const row=meta.mergeUse.get(y);
    for(let x=c;x<=right;x++){
      if(row[x])throw Error('The workbook contains overlapping merged cells.');
      row[x]=1;
    }
  }
  meta.merges.push({r,c,bottom,right,anchor:address(r,c)});
}
export function initPreview(sheet,hidden=false){
  sheet.preview={hidden,merges:[],rows:0,cols:0};
  return sheet.preview;
}
export function finishPreview(sheet){
  const m=sheet.preview;
  delete m.mergeUse;
  for(const cell of Object.keys(sheet.cells)){const [r,c]=coords(cell);m.rows=Math.max(m.rows,r);m.cols=Math.max(m.cols,c);}
  for(const {bottom,right} of m.merges){m.rows=Math.max(m.rows,bottom);m.cols=Math.max(m.cols,right);}
}
// Detect ordinary Office date/time formats without interpreting arbitrary formulas or
// claiming pixel-accurate rendering of custom, conditional or localized format codes.
export function excelDateFormat(id,code){
  if([14,15,16,17,27,28,29,30,31,34,35,36,50,51,52,53,54,55,56,57,58].includes(id))return 'date';
  if([18,19,20,21,32,33,45,47].includes(id))return 'time';
  if(id===22)return 'datetime';
  if(id===46)return 'duration';
  if(!code)return '';
  const clean=code.replace(/"(?:""|[^"])*"|\\.|_.|\*./g,'').split(';')[0].toLowerCase();
  const duration=/\[[hms]+\]/i.test(clean);
  const mask=clean.replace(/\[[^\]]*\]/g,'');
  if(duration)return 'duration';
  const clock=/[hs]|am\/pm|a\/p/.test(mask);
  const calendar=/[yd]/.test(mask)||(!clock&&/m/.test(mask));
  if(/[abcefgijklnopqrtuvwxz]/.test(mask.replace(/am\/pm|a\/p/g,'')))return '';
  return clock?(calendar?'datetime':'time'):(calendar?'date':'');
}
