// A bounded ZIP32 reader and an uncompressed ZIP32 writer for browser-local workbooks.
const encoder = new TextEncoder();
export const utf8 = new TextDecoder('utf-8', {fatal:true});
export const MAX_INPUT = 20 * 1024 * 1024;
const MAX_EXPANDED = 64 * 1024 * 1024;
const crcTable = Uint32Array.from({length:256}, (_, i) => {
  for (let j=0; j<8; j++) i = i & 1 ? 0xedb88320 ^ (i >>> 1) : i >>> 1;
  return i >>> 0;
});
function checksum(bytes) {
  let crc = 0xffffffff;
  for (const b of bytes) crc = crcTable[(crc ^ b) & 255] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}
function requireZip(condition) {
  if (!condition) throw Error('Invalid, encrypted, or unsupported ZIP package. Open a simple unencrypted workbook.');
}
export function openZip(buffer) {
  const b = new Uint8Array(buffer), v = new DataView(buffer);
  const u16 = i=>v.getUint16(i,true), u32 = i=>v.getUint32(i,true);
  let end = b.length - 22;
  while(end>=Math.max(0,b.length-65557) && !(u32(end)===0x06054b50 && end+22+u16(end+20)===b.length)) end--;
  requireZip(end>=0 && u32(end)===0x06054b50 && !u16(end+4) && !u16(end+6));
  const count=u16(end+10), start=u32(end+16), entries=new Map();
  requireZip(count===u16(end+8) && count<=4096 && start+u32(end+12)===end);
  let pos=start, total=0;
  for(let i=0; i<count; i++){
    requireZip(pos+46<=end && u32(pos)===0x02014b50 && u16(pos+34)===0);
    const flags=u16(pos+8), method=u16(pos+10), crc=u32(pos+16), packed=u32(pos+20), size=u32(pos+24);
    const len=u16(pos+28), local=u32(pos+42);
    const next=pos+46+len+u16(pos+30)+u16(pos+32);
    requireZip(next<=end && len && !(flags & 0x2041) && [0,8].includes(method));
    const name=utf8.decode(b.subarray(pos+46,pos+46+len));
    requireZip(!entries.has(name) && !name.startsWith('/') && !name.includes('\\') && !name.split('/').includes('..') && !name.includes('\0'));
    requireZip(local+30<=start && u32(local)===0x04034b50 && u16(local+6)===flags && u16(local+8)===method);
    const data=local+30+u16(local+26)+u16(local+28);
    requireZip(data+packed<=start && utf8.decode(b.subarray(local+30,local+30+u16(local+26)))===name);
    total+=size;
    if(total>MAX_EXPANDED) throw Error('The expanded workbook exceeds 64 MiB.');
    entries.set(name,{data,packed,size,crc,method});pos=next;
  }
  requireZip(pos===end);
  return {
    names:[...entries.keys()],
    async text(name) {
      const e=entries.get(name); if(!e) throw Error('A required workbook part is missing.');
      const compressed=b.subarray(e.data,e.data+e.packed);
      let bytes=compressed;
      if(e.method===8){
        const reader=new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
        bytes=new Uint8Array(e.size);let used=0;
        try{while(true){
          const {done,value}=await reader.read();if(done)break;
          requireZip(used+value.length<=e.size);
          bytes.set(value,used);used+=value.length;
        }requireZip(used===e.size)}finally{await reader.cancel().catch(()=>{});reader.releaseLock()}
      }
      requireZip(bytes.length===e.size && checksum(bytes)===e.crc);
      return utf8.decode(bytes);
    }
  };
}
export function makeZip(parts, type) {
  const local=[], central=[];let position=0, centralSize=0;
  for(const [name,content] of parts){
    const filename=encoder.encode(name), bytes=encoder.encode(content), crc=checksum(bytes);
    const a=new Uint8Array(30+filename.length), c=new Uint8Array(46+filename.length);
    for(const [rec,isCentral] of [[a,false],[c,true]]){
      const v=new DataView(rec.buffer), off=isCentral?2:0;
      v.setUint32(0,isCentral?0x02014b50:0x04034b50,true);
      if(isCentral)v.setUint16(4,20,true);
      v.setUint16(4+off,20,true);v.setUint16(6+off,2048,true);
      v.setUint16(12+off,33,true); // DOS date 1980-01-01
      v.setUint32(14+off,crc,true);v.setUint32(18+off,bytes.length,true);
      v.setUint32(22+off,bytes.length,true);v.setUint16(26+off,filename.length,true);
      rec.set(filename,isCentral?46:30);
    }
    new DataView(c.buffer).setUint32(42,position,true);local.push(a,bytes);central.push(c);
    position+=a.length+bytes.length;centralSize+=c.length;
  }
  const end=new Uint8Array(22), v=new DataView(end.buffer);
  v.setUint32(0,0x06054b50,true);v.setUint16(8,parts.length,true);v.setUint16(10,parts.length,true);
  v.setUint32(12,centralSize,true);v.setUint32(16,position,true);
  if(position+centralSize+22>MAX_INPUT) throw Error('This export would exceed the 20 MiB workbook limit. Reduce its contents first.');
  return new Blob([...local,...central,end],{type});
}
