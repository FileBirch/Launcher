const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const {createServer}=require('node:http');
const {join,resolve,extname}=require('node:path');
const {tmpdir}=require('node:os');
const {spawnSync}=require('node:child_process');
const executable=process.env.FILEBIRCH_BROWSER_EXE||process.env.FIRMQUILL_BROWSER_EXE;
const root=resolve(__dirname,'..');
test('Tables opens real XLSX, ODS and CSV files without editing, evaluation or external requests',{skip:!executable},async t=>{
 const temp=await fs.mkdtemp(join(tmpdir(),'filebirch-tables-'));t.after(()=>fs.rm(temp,{recursive:true,force:true}));
 const created=spawnSync('python',[join(__dirname,'tables-viewer-fixtures.py'),temp],{encoding:'utf8'});
 assert.equal(created.status,0,created.stderr);
 const server=createServer(async(req,res)=>{
   try{
     const url=new URL(req.url,'http://localhost');
     if(url.pathname==='/_view-test'){res.writeHead(200,{'Content-Type':'text/html'});res.end('<!doctype html><meta charset=utf-8><link rel=stylesheet href="/spreadsheets/viewer.css"><main id=root></main><script type=module src="/_view-test.js"></script>');return;}
     if(url.pathname==='/_view-test.js'){res.writeHead(200,{'Content-Type':'text/javascript'});res.end("import {mount} from '/spreadsheets/viewer.js'; await mount(document.querySelector('#root'));");return;}
     if(!/^\/spreadsheets\/[a-z-]+\.(js|css)$/.test(url.pathname)){res.writeHead(404);res.end();return;}
     const bytes=await fs.readFile(join(root,url.pathname));res.writeHead(200,{'Content-Type':extname(url.pathname)==='.css'?'text/css':'text/javascript'});res.end(bytes);
   }catch{res.writeHead(404);res.end();}
 });
 await new Promise((r,j)=>{server.once('error',j);server.listen(0,'127.0.0.1',r);});t.after(()=>new Promise(r=>server.close(r)));
 const url='http://127.0.0.1:'+server.address().port;
 const {chromium}=require(process.env.FILEBIRCH_PLAYWRIGHT||process.env.FIRMQUILL_PLAYWRIGHT||'playwright');
 const browser=await chromium.launch({executablePath:executable,headless:true});t.after(()=>browser.close());
 const page=await browser.newPage(),outside=[],errors=[];
 page.on('request',req=>{if(!req.url().startsWith(url))outside.push(req.url());});page.on('pageerror',e=>errors.push(e.message));
 await page.goto(url+'/_view-test');await page.getByRole('heading',{name:'Open a spreadsheet'}).waitFor();
 const file=page.getByLabel('Browse for a spreadsheet'),status=page.getByRole('status');
 const read=async(name)=>{await file.setInputFiles(join(temp,name));await status.getByText('Opened '+name+' · read-only.',{exact:true}).waitFor();};
 const cell=(addr)=>page.locator('[data-cell="'+addr+'"]');
 const cellText=async(addr,value)=>assert.ok((await cell(addr).innerText()).includes(value),addr);
 await read('workbook.xlsx');
 assert.deepEqual(await page.locator('#tables-sheets option').allTextContents(),['Transactions','Archive (hidden in original)']);
 await cellText('A1','Merged XLSX heading');assert.equal(await cell('A1').getAttribute('colspan'),'2');
 await cellText('A2','2024-01-01');await cellText('B2','08:30:00');await cellText('C2','2024-01-01 06:00:00');
 await cellText('D2','36:30:00');await cellText('E2','2024-02-29T12:00:00Z');
 await cellText('A3','42');await cellText('B3','No saved formula result');await cellText('C3','#REF!');await cellText('D3','FALSE');
 await cell('A3').getByRole('button').click();assert.match(await page.locator('[data-cell-info]').innerText(),/WEBSERVICE\("https:\/\/outside\.invalid/);
 await page.getByLabel('Go to cell').fill('Q41');await page.getByRole('button',{name:'Go',exact:true}).click();await cellText('Q41','Far value');
 await page.getByLabel('Worksheet').selectOption('1');await cellText('A1','Archived data');
 await read('calendar1904.xlsx');await cellText('A1','1904-01-02');
 await read('workbook.ods');await cellText('A1','Merged ODS heading');assert.equal(await cell('A1').getAttribute('colspan'),'2');
 await cellText('A2','2025-02-08');await cellText('B2','PT13H30M');await cellText('C2','1000');await cellText('D2','TRUE');await cellText('E2','No saved formula result');
 await cell('C2').getByRole('button').click();assert.match(await page.locator('[data-cell-info]').innerText(),/of:=SUM/);
 await page.getByRole('button',{name:'Close file'}).click();assert.equal(await page.locator('table').count(),0);
 const csv=Buffer.from('\ufeffID,Quoted,Literal\r\n00123,"quoted, value","=1+1"\r\n,,"<img src=https://outside.invalid/b.png>"\r\n');
 await file.setInputFiles({name:'data.csv',mimeType:'text/csv',buffer:csv});await status.getByText('Opened data.csv · read-only.').waitFor();
 await cellText('A2','00123');await cellText('B2','quoted, value');await cellText('C2','=1+1');await cellText('C3','<img src=https://outside.invalid/b.png>');
 for(const [name,buffer,want] of [
   ['bad.csv',Buffer.from('bad,"quoted'),/unclosed/],
   ['binary.csv',Buffer.from([255,254,253]),/UTF-8/],
   ['bad.xlsx',Buffer.from('not zip'),/unsupported ZIP/],
   ['macro.xlsm',Buffer.from('no macros'),/Macro-enabled/],
   ['long.csv',Buffer.alloc(20*1024*1024+1,65),/20 MiB/]
 ]){await file.setInputFiles({name,mimeType:'application/octet-stream',buffer});await status.getByText(want).waitFor();assert.equal(await page.locator('table').count(),0);}
 for(const [name,want] of [['broken-xml.xlsx',/malformed/],['merge-overlap.xlsx',/overlapping merged cells/],['bad-bounds.ods',/10,000-row or 256-column limit/],['expanded.xlsx',/64 MiB/]]){
   await file.setInputFiles(join(temp,name));await status.getByText(want).waitFor();assert.equal(await page.locator('table').count(),0);
 }
 assert.deepEqual(errors,[]);assert.deepEqual(outside,[]);
});
