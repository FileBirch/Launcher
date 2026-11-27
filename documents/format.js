export const W=['http://schemas.openxmlformats.org/wordprocessingml/2006/main','http://purl.oclc.org/ooxml/wordprocessingml/main'];
export const O=id=>'urn:oasis:names:tc:opendocument:xmlns:'+id+':1.0';
export const namespaces={w:W,office:[O('office')],style:[O('style')],fo:['urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0'],text:[O('text')],table:[O('table')],draw:[O('drawing')],svg:['urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0'],xlink:['http://www.w3.org/1999/xlink']};
export function is(n,ns,k){return n?.localName===k&&namespaces[ns].includes(n.namespaceURI);}
export function attr(n,ns,k){for(const uri of namespaces[ns]){const x=n?.getAttributeNS(uri,k);if(x!=null)return x;}return '';}
export const ch=(n,ns,k)=>Array.from(n?.children??[]).filter(x=>is(x,ns,k));
export const one=(n,ns,k)=>ch(n,ns,k)[0];
export function node(tag,text){const n=document.createElement(tag);if(text!==undefined)n.textContent=text;return n;}
export const int=(s,d,lo,hi)=>/^\d{1,8}$/.test(s)?Math.min(hi,Math.max(lo,Number(s))):d;
export const hex=s=>/^[a-f0-9]{6}$/i.test(s)?'#'+s:'';
const csshex=s=>/^#[a-f0-9]{6}$/i.test(s)?s:'';
const align=v=>({'left':'left','center':'center','right':'right','start':'start','end':'end','justify':'justify','both':'justify'})[v]||'';
const capped=(v,min,max)=>Math.max(min,Math.min(max,v));
export function measure(v,min=0,max=1000){
  const m=/^(-?\d{1,5}(?:\.\d{1,5})?)(cm|mm|in|pt|px|pc)$/.exec(v);
  return m?capped(Number(m[1])*({cm:96/2.54,mm:96/25.4,in:96,pt:4/3,px:1,pc:16})[m[2]],min,max)+'px':'';
}
const wv=(n,k)=>attr(one(n,'w',k),'w','val');
const twip=(v,min=-180,max=1500)=>/^-?\d{1,7}$/.test(v)?capped(Number(v)/15,min,max)+'px':'';
const toggle=(n,k)=>{const x=one(n,'w',k);return x?!/^(0|false|off)$/i.test(attr(x,'w','val')):null};
export function wordRun(n){
  const s={};
  for(const [tag,key,yes,no] of [['b','fontWeight','bold','normal'],['i','fontStyle','italic','normal']]){
    const v=toggle(n,tag);if(v!==null)s[key]=v?yes:no;
  }
  const u=one(n,'w','u'),st=toggle(n,'strike');
  if(u||st!==null)s.textDecorationLine=[u&&wv(n,'u')!=='none'?'underline':'',st?'line-through':''].filter(Boolean).join(' ')||'none';
  const size=wv(n,'sz');if(size)s.fontSize=int(size,24,12,192)/2+'pt';
  const color=hex(wv(n,'color'));if(color)s.color=color;
  const bg=hex(attr(one(n,'w','shd'),'w','fill'));if(bg)s.backgroundColor=bg;
  const high=wv(n,'highlight');if(['yellow','green','cyan','magenta','blue','red','gray','black','white'].includes(high))s.backgroundColor=high;
  const va=wv(n,'vertAlign');if(va==='superscript'||va==='subscript'){s.verticalAlign=va==='superscript'?'super':'sub';s.fontSize='0.8em';}
  return s;
}
export function wordParagraph(n){
  const s={},j=align(wv(n,'jc'));if(j)s.textAlign=j;
  const spacing=one(n,'w','spacing'), indent=one(n,'w','ind');
  for(const [src,tag,key] of [[spacing,'before','marginTop'],[spacing,'after','marginBottom'],[indent,'left','marginInlineStart'],[indent,'right','marginInlineEnd'],[indent,'firstLine','textIndent']]){
    const p=twip(attr(src,'w',tag));if(p)s[key]=p;
  }
  const hanging=attr(indent,'w','hanging');if(/^\d{1,6}$/.test(hanging))s.textIndent=twip('-'+hanging);
  const line=attr(spacing,'w','line');if(/^\d{1,6}$/.test(line))s.lineHeight=attr(spacing,'w','lineRule')==='auto'||!attr(spacing,'w','lineRule')?String(capped(Number(line)/240,0.7,5)):twip(line,7,140);
  return s;
}
export function odtRun(n){
  const s={};
  for(const [key,source,allowed] of [['fontWeight','font-weight',['bold','normal']],['fontStyle','font-style',['italic','normal','oblique']]]){
    const v=attr(n,'fo',source);if(allowed.includes(v))s[key]=v;
  }
  const c=csshex(attr(n,'fo','color'));if(c)s.color=c;
  const bg=csshex(attr(n,'fo','background-color'));if(bg)s.backgroundColor=bg;
  const size=measure(attr(n,'fo','font-size'),8,130);if(size)s.fontSize=size;
  const u=attr(n,'style','text-underline-style'),strike=attr(n,'style','text-line-through-style');
  if(u||strike)s.textDecorationLine=[u&&u!=='none'?'underline':'',strike&&strike!=='none'?'line-through':''].filter(Boolean).join(' ')||'none';
  const pos=attr(n,'style','text-position');if(/^(super|sub)( |$)/.test(pos)){s.verticalAlign=pos.split(' ')[0];s.fontSize='0.8em';}
  return s;
}
export function odtParagraph(n){
  const s={},j=align(attr(n,'fo','text-align'));if(j)s.textAlign=j;
  for(const [tag,key] of [['margin-top','marginTop'],['margin-bottom','marginBottom'],['margin-left','marginInlineStart'],['margin-right','marginInlineEnd'],['text-indent','textIndent']]){
    const p=measure(attr(n,'fo',tag),-180);if(p)s[key]=p;
  }
  const line=attr(n,'fo','line-height');if(/^\d{1,3}%$/.test(line))s.lineHeight=String(capped(Number.parseInt(line,10)/100,0.7,5));
  else {const px=measure(line,7,140);if(px)s.lineHeight=px;}
  return s;
}
export function mergeStyles(map,key){
  const sources=[],seen=new Set();
  while(map.has(key)&&sources.length<16&&!seen.has(key)){seen.add(key);const entry=map.get(key);sources.unshift(entry);key=entry.parent;}
  return Object.assign({},...sources.map(s=>s.css));
}
