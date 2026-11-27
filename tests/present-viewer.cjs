const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const {createServer}=require('node:http');
const {chromium}=require('playwright');
const root=path.resolve(__dirname,'..');
const p='http://schemas.openxmlformats.org/presentationml/2006/main',a='http://schemas.openxmlformats.org/drawingml/2006/main',r='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels=rows=>'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rows.map(([id,kind,target,ext])=>`<Relationship Id="${id}" Type="${r}/${kind}" Target="${target}"${ext?' TargetMode="External"':''}/>`).join('')+'</Relationships>';
const png=new Uint8Array(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=','base64'));
const xfm='<a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3810000" cy="914400"/></a:xfrm>';
const slide=(name,copy,extra='')=>`<p:sld xmlns:p="${p}" xmlns:a="${a}" xmlns:r="${r}"><p:cSld name="${name}"><p:spTree><p:sp><p:spPr>${xfm}</p:spPr><p:txBody><a:p><a:r><a:rPr sz="2500" b="1"/><a:t>${copy}</a:t></a:r></a:p></p:txBody></p:sp>${extra}</p:spTree></p:cSld><p:transition/></p:sld>`;
async function fixture(kind){
 const {makeZip, packText}=await import('../spreadsheets/package.js');
 const pptx=[
   ['ppt/presentation.xml',`<p:presentation xmlns:p="${p}" xmlns:r="${r}"><p:sldIdLst><p:sldId id="511" r:id="b"/><p:sldId id="510" r:id="a"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`],
   ['ppt/_rels/presentation.xml.rels',rels([['a','slide','slides/slide1.xml'],['b','slide','slides/slide2.xml']])],
   ['ppt/slides/slide1.xml',slide('Second','&lt;img src="not-an-image"&gt;')],
   ['ppt/slides/slide2.xml',slide('First','Designed first','<p:pic><p:blipFill><a:blip r:embed="img"/></p:blipFill><p:spPr>'+xfm.replace('914400" y="914400','5594400" y="914400')+'</p:spPr></p:pic><p:pic><p:blipFill><a:blip r:link="ext"/></p:blipFill><p:spPr>'+xfm.replace('y="914400','y="1914400')+'</p:spPr></p:pic><p:graphicFrame/>')],
   ['ppt/slides/_rels/slide2.xml.rels',rels([['img','image','../media/p.png'],['ext','image','https://blocked.invalid/should-not-fetch.png',true]])],['ppt/media/p.png',{bytes:png}]
 ];
 const odp=[
   ['mimetype','application/vnd.oasis.opendocument.presentation'],
   ['content.xml',`<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:presentation><draw:page draw:name="ODP first" draw:master-page-name="Main"><draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="1in"><draw:text-box><text:p>Before <text:a xlink:href="https://blocked.invalid/no-outbound">link</text:a></text:p></draw:text-box></draw:frame><draw:frame svg:x="6in" svg:y="1in" svg:width="1in" svg:height="1in"><draw:image xlink:href="Pictures/safe.png"/></draw:frame></draw:page><draw:page draw:name="ODP second"><draw:frame svg:x="1in" svg:y="1in" svg:width="4in" svg:height="1in"><draw:text-box><text:p>Second ODP slide</text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>`],
   ['styles.xml','<office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"><office:automatic-styles><style:page-layout style:name="wide"><style:page-layout-properties fo:page-width="10in" fo:page-height="6in"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name="Main" style:page-layout-name="wide"/></office:master-styles></office:document-styles>'],['Pictures/safe.png',{bytes:png}]
 ];
 const parts=kind==='pptx'?pptx:odp; if(kind==='unsafe')parts[1][1]='<!DOCTYPE a [<!ENTITY remote SYSTEM "https://blocked.invalid/entity">]>'+parts[1][1];
 return {name:'synthetic.'+(kind==='unsafe'?'odp':kind),mimeType:'application/octet-stream',buffer:Buffer.from(await makeZip(await Promise.all(parts.map(async ([n,v])=>[n,typeof v==='string'?await packText(v):v]))).arrayBuffer())};
}
(async()=>{
 const target=process.env.FILEBIRCH_REAL_ENTRY==1?'/presentations/':'/test-present';
 const server=createServer(async(req,res)=>{
   const route=new URL(req.url,'http://127.0.0.1').pathname;
   res.setHeader('Content-Security-Policy',"default-src 'none'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; object-src 'none';");
   try{
     if(route==='/test-present'){res.setHeader('Content-Type','text/html');res.end('<!doctype html><html><link rel="stylesheet" href="/presentations/viewer.css"><main id="mount"></main><script type="module" src="/test-present-entry.js"></script></html>');return;}
     if(route==='/test-present-entry.js'){res.setHeader('Content-Type','text/javascript');res.end("import {mountViewer} from '/presentations/viewer.js';mountViewer(document.querySelector('#mount'));");return;}
     const filename=path.resolve(root,'.'+route+(route.endsWith('/')?'index.html':''));
     if(!filename.startsWith(root+path.sep))throw Error();
     res.setHeader('Content-Type',({'.js':'text/javascript','.css':'text/css','.html':'text/html'})[path.extname(filename)]||'application/octet-stream');res.end(await fs.readFile(filename));
   }catch{res.writeHead(404).end();}
 });
 await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
 const base='http://127.0.0.1:'+server.address().port;
 let browser;
 try{
   browser=await chromium.launch(process.env.FILEBIRCH_BROWSER_EXE?{executablePath:process.env.FILEBIRCH_BROWSER_EXE,headless:true}:{channel:process.env.FILEBIRCH_BROWSER||'msedge',headless:true});
   const page=await browser.newPage({viewport:{width:390,height:844}}),errors=[],outbound=[];
   page.on('pageerror',e=>errors.push(e.message));page.on('request',r=>{if(!r.url().startsWith(base))outbound.push(r.url());});
   await page.goto(base+target);
   if(target==='/presentations/')await page.getByRole('button',{name:'I agree to the Terms of Service'}).click();
   const upload=page.getByLabel('Open a PowerPoint or OpenDocument presentation');
   await upload.setInputFiles(await fixture('pptx'));
   await page.getByText('Designed first',{exact:true}).waitFor();
   assert.equal(await page.locator('.fbpr-slide').evaluate(el=>el.style.height),'540px');
   assert.equal(await page.locator('.fbpr-item').first().evaluate(el=>el.style.left),'96px');
   assert.ok(await page.locator('.fbpr-slide img').evaluate(img=>img.src.startsWith('data:image/png;')&&img.naturalWidth===1));
   await page.getByText('Chart, table, or diagram not rendered').waitFor();
   await page.getByText('External links or resources were not loaded.').waitFor();
   await page.getByRole('button',{name:'Next slide'}).click();
   await page.getByText('<img src="not-an-image">',{exact:true}).waitFor();
   assert.equal(await page.locator('.fbpr-slide img').count(),0);
   await page.getByRole('button',{name:'Previous slide'}).click();
   await page.getByText('Designed first',{exact:true}).waitFor();
   await upload.setInputFiles(await fixture('odp'));
   await page.getByText('Before',{exact:true}).waitFor();
   assert.equal(await page.locator('.fbpr-slide a').count(),0);
   assert.equal(await page.locator('.fbpr-slide').evaluate(el=>el.style.height),'576px');
   const stage=page.getByLabel('Slide preview',{exact:true});await stage.focus();await stage.press('ArrowRight');
   await page.getByText('Second ODP slide',{exact:true}).waitFor();
   await page.getByRole('button',{name:'Close document'}).click();
   assert.equal(await page.locator('.fbpr-slide').count(),0);
   await upload.setInputFiles(await fixture('unsafe'));
   await page.getByText('Document types and XML entities are not supported.').waitFor();
   assert.deepEqual(errors,[]);assert.deepEqual(outbound,[]);
   console.log('Present Chromium verification passed: PPTX/ODP order, sizing, text, PNG, placeholders, keyboard, escaping, closed file and blocked external content.');
 }finally{if(browser)await browser.close();server.close();}
})().catch(error=>{console.error(error);process.exitCode=1;});
