import {createServer} from 'node:http';
import {readFile} from 'node:fs/promises';
import {fileURLToPath,pathToFileURL} from 'node:url';
import {extname,resolve} from 'node:path';
const root=fileURLToPath(new URL('.',import.meta.url));
const MIME={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.svg':'image/svg+xml','.md':'text/plain; charset=utf-8'};
const rootFiles=new Set(['index.html','brand.js','release.js','shell.js','shell.css','quill-art.js','quill.svg','tos.md']);
const appFile=/^(?:spreadsheets|documents|presentations)\/(?:index\.html|[a-z][a-z0-9-]*\.(?:js|css))$/;
export function createFileBirchServer(){return createServer(async(req,res)=>{
  const send=(status,body,mime='text/plain; charset=utf-8')=>{
    res.writeHead(status,{'Content-Type':mime,'X-Content-Type-Options':'nosniff','Cache-Control':'no-store','Referrer-Policy':'no-referrer'});
    res.end(req.method==='HEAD'?undefined:body);
  };
  if(!['GET','HEAD'].includes(req.method)){send(405,'This viewer does not accept uploads.');return;}
  try{
    const pathname=decodeURIComponent(new URL(req.url,'http://localhost').pathname).slice(1);
    const path=!pathname||pathname.endsWith('/')?pathname+'index.html':pathname;
    if(!rootFiles.has(path)&&!appFile.test(path)){send(404,'This file is not available.');return;}
    try{const data=await readFile(resolve(root,path));send(200,data,MIME[extname(path)]);}catch(e){send(e.code==='ENOENT'?404:500,'The viewer file could not be read.');}
  }catch{send(400,'Invalid viewer request.');}
});}
if(process.argv[1]&&pathToFileURL(resolve(process.argv[1])).href===import.meta.url){
  const port=Number(process.env.FILEBIRCH_PORT??4286);
  if(!Number.isSafeInteger(port)||port<1||port>65535)throw Error('Choose FILEBIRCH_PORT from 1 to 65535.');
  const server=createFileBirchServer();
  server.on('error',error=>{console.error(error.code==='EADDRINUSE'?'This port is in use. Choose another FILEBIRCH_PORT.':error.message);process.exitCode=1;});
  server.listen(port,'127.0.0.1',()=>console.log('FileBirch: http://127.0.0.1:'+port+'/  (keep this window open)'));
}
