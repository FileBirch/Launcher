import {openZip, utf8, MAX_INPUT} from '../spreadsheets/package.js';
import {parse} from '../spreadsheets/xml.js';
export {utf8, MAX_INPUT};
const crc = Uint32Array.from({length:256}, (_,n)=>{for(let i=0;i<8;i++)n=(n>>>1)^((n&1)?0xedb88320:0);return n>>>0});
export function packageReader(buffer){
  let zip;
  try{zip=openZip(buffer);}
  catch(error){throw Error(error.message.replace('workbook','document'));}
  const names=new Set(zip.names);
  async function bytes(name, max){
    if(!names.has(name)) throw Error('A required document part is missing.');
    const e=zip.packed(name);
    if(e.size>max) throw Error('A document part exceeds the safe preview size.');
    let raw=e.bytes;
    if(e.method===8){
      const read=new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
      raw=new Uint8Array(e.size);let used=0;
      try{for(;;){const {done,value}=await read.read();if(done)break;
        if(used+value.length>raw.length)throw Error('A compressed part exceeded its declared size.');
        raw.set(value,used);used+=value.length;
      }if(used!==raw.length)throw Error('A compressed part is incomplete.');}
      finally{await read.cancel().catch(()=>{});read.releaseLock();}
    }
    let sum=0xffffffff;for(const b of raw)sum=crc[(sum^b)&255]^(sum>>>8);
    if(raw.length!==e.size||((sum^0xffffffff)>>>0)!==e.crc)throw Error('A document part failed its integrity check.');
    return raw;
  }
  async function xml(name, optional=false){
    if(optional&&!names.has(name))return null;
    const source=utf8.decode(await bytes(name,6*1024*1024));
    let root;
    try{root=parse(source);}
    catch(error){throw Error(error.message.replace('workbook','document'));}
    // Bound traversal, stack depth, and the quantity of attributes in retained XML.
    const pending=[[root,0]];let count=0;
    while(pending.length){const [n,depth]=pending.pop();
      if(++count>90000||depth>48||n.attributes.length>64)throw Error('The document is too structurally complex for this viewer.');
      for(const child of n.children)pending.push([child,depth+1]);
    }
    return root;
  }
  return {names, bytes, xml};
}
export function internalPath(folder, target){
  if(!target||target.length>384||/[:\\?#%\u0000-\u001f]/.test(target))return '';
  const out=target.startsWith('/')?[]:folder.split('/').filter(Boolean);
  for(const p of target.split('/')){if(!p||p==='.')continue;
    if(p==='..'){if(!out.length)return '';out.pop();}else out.push(p);
  }
  return out.join('/');
}
export function raster(bytes){
  const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);
  let type,w,h;
  if(bytes.length>=33&&v.getUint32(0)===0x89504e47&&v.getUint32(4)===0x0d0a1a0a&&v.getUint32(12)===0x49484452){
    type='image/png';w=v.getUint32(16);h=v.getUint32(20);
    // Animated PNG frames can carry much more decoded data than the first image.
    for(let i=8;i+12<=bytes.length;){const size=v.getUint32(i);if(size>bytes.length-i-12)return null;
      const tag=v.getUint32(i+4);if([0x6163544c,0x6663544c,0x66644154].includes(tag))return null;i+=size+12;}
  }else if(bytes.length>12&&bytes[0]===255&&bytes[1]===216){
    for(let i=2;i+4<bytes.length;){
      if(bytes[i++]!==255)break;
      while(bytes[i]===255)i++;
      const tag=bytes[i++];if(tag===218||tag===217||i+2>bytes.length)break;
      const len=v.getUint16(i);if(len<2||i+len>bytes.length)break;
      if([192,193,194].includes(tag)&&len>=8){type='image/jpeg';h=v.getUint16(i+3);w=v.getUint16(i+5);break;}i+=len;
    }
  }
  return type&&w&&h&&w<=4096&&h<=4096&&w*h<=8000000?{type,w,h}:null;
}
export function dataUri(type,bytes){
  let binary='';for(let i=0;i<bytes.length;i+=16384)binary+=String.fromCharCode(...bytes.subarray(i,i+16384));
  return 'data:'+type+';base64,'+btoa(binary);
}
