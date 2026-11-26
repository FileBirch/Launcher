// Run with "node tests/formula-groups.cjs" and Playwright on the module path.
// Uses a loopback static server and a temporary download, just as the editor does.
const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const os = require('node:os');
const {createServer} = require('node:http');
const {chromium} = require('playwright');
const root=path.resolve(__dirname,'..'), fixture=path.join(__dirname,'formula-groups.xlsx');

async function run(){
  const temp=await fs.mkdtemp(path.join(os.tmpdir(),'filebirch-formulas-'));
  const copy=path.join(temp,'export.xlsx'),original=await fs.readFile(fixture);
  const server=createServer(async(req,res)=>{
    try{
      const route=new URL(req.url,'http://localhost').pathname;
      const target=path.resolve(root,'.'+route+(route.endsWith('/')?'index.html':''));
      if(!target.startsWith(root+path.sep))throw Error();
      const data=await fs.readFile(target);
      res.setHeader('Content-Type',({'.js':'text/javascript','.css':'text/css','.html':'text/html','.svg':'image/svg+xml'})[path.extname(target)]||'application/octet-stream');
      res.end(data);
    }catch{res.writeHead(404).end();}
  });
  await new Promise(resolve=>server.listen(0,'127.0.0.1',resolve));
  let browser;
  try {
    browser=await chromium.launch({channel:process.env.FILEBIRCH_BROWSER||'msedge',headless:true});
    const page=await browser.newPage({viewport:{width:1320,height:850}});
    page.setDefaultTimeout(6000);
    const errors=[];page.on('pageerror',err=>errors.push(err.message));
    await page.goto('http://127.0.0.1:'+server.address().port+'/spreadsheets/');
    await page.getByRole('button',{name:'I agree to the Terms of Service'}).click();
    const button=id=>page.locator('#'+id);
    const cell=ref=>page.locator('[role=gridcell][data-address="'+ref+'"]');
    const tab=name=>page.getByRole('tab',{name,exact:true}).click();
    const edit=async(ref,value)=>{
      await cell(ref).click();await button('formula').fill(value);await button('formula').press('Enter');
    };
    const value=async(ref,wanted)=>assert.equal(await cell(ref).textContent(),String(wanted),ref);
    const open=async file=>{
      await button('file').setInputFiles(file);
      await page.getByRole('dialog',{name:'Import notes'}).waitFor();
      const notes=await button('dlgbody').innerText();
      for(const label of ['Calculation!F6:F7','Scenarios!B9:B10'])assert.ok(notes.includes(label),notes);
      await button('confirm').click();
    };
    await open(fixture);
    await value('D3',38);
    // Excel itself uses ordinary formulas for cross-sheet expressions. Other
    // producers can share them. Exercise both encodings through the file picker.
    const qualified=await page.evaluate(async bytes=>{
      const {openZip,makeZip,packText}=await import('./package.js');
      const zip=openZip(Uint8Array.from(bytes).buffer),parts=[];
      for(const n of zip.names){
        const part=n==='xl/worksheets/sheet1.xml'
          ? await packText((await zip.text(n)).replace('+$A$1+J2',"+$A$1+'Rates 1'!B2"))
          : zip.packed(n);
        parts.push([n,part]);
      }
      return [...new Uint8Array(await makeZip(parts).arrayBuffer())];
    },[...original]);
    await open({name:'qualified-shared.xlsx',mimeType:'application/octet-stream',buffer:Buffer.from(qualified)});
    await value('C2',16);await value('D3',38);
    await cell('D3').click();
    assert.equal(await button('formula').inputValue(),"=$A3+C$1+$A$1+'Rates 1'!C3");
    await value('F2',4);await value('F3',6);await value('G2',10);
    await value('F6','41 · saved');await value('F7','No saved result');await value('G6','42 · saved');
    await cell('F7').click();assert.ok(await button('formula').evaluate(el=>el.readOnly));
    assert.match(await button('formula').inputValue(),/UNIQUE/);
    await page.keyboard.press('Delete');
    await cell('F6').click();await page.keyboard.press('F2');
    assert.equal(await page.locator('.cellinput').count(),0);
    await cell('E7').click();
    await page.locator('[role=grid]').evaluate(el=>{
      const data=new DataTransfer();data.setData('text/plain','changed\tchanged');
      el.dispatchEvent(new ClipboardEvent('paste',{clipboardData:data,bubbles:true,cancelable:true}));
    });
    await value('E7','');await value('F7','No saved result');
    await edit('A2','9');await value('C2',21);await value('F2',9);await value('G2',15);
    await value('F6','41 · saved');await value('F7','No saved result');
    await tab('Scenarios');
    for(const [key,n] of Object.entries({B5:73,B6:83,E5:47,F5:67,E6:48,F6:68,E9:43,F9:63}))await value(key,n);
    await edit('B1','5');await value('B5',75);await value('E9',45);await value('F6',68);
    await value('B9','77 · saved');await value('B10','No saved result');
    await edit('A9','11');await value('B9','77 · saved');
    await tab('Rates 1');await edit('A1','unrelated edit');await cell('A1').click();await button('bold').click();
    await button('renamesheet').click();
    await page.getByRole('textbox',{name:'Worksheet name'}).fill("Rate's input");
    await button('confirm').click();await tab('Calculation');await value('D3',38);
    await cell('D3').click();assert.match(await button('formula').inputValue(),/'Rate''s input'!C3/);
    await button('undo').click();await tab('Calculation');await value('D3',38);
    await button('save').click();
    assert.match(await button('dlgbody').innerText(),/Calculation!F6:F7/);
    assert.match(await button('dlgbody').innerText(),/Scenarios!E5:F6/);
    const download=page.waitForEvent('download');
    await page.getByRole('button',{name:'Download XLSX with formulas'}).click();
    const item=await download;assert.match(item.suggestedFilename(),/ - copy\.xlsx$/);
    await item.saveAs(copy);
    const result=await fs.readFile(copy);
    await open(copy);
    await value('C2',21);await value('F2',9);await value('D3',38);
    await value('F6','41 · saved');await value('F7','No saved result');
    await tab('Rates 1');await value('A1','unrelated edit');assert.equal(await cell('A1').evaluate(el=>el.style.fontWeight),'700');
    await tab('Scenarios');await value('E5',47);await value('B5',75);await value('B10','No saved result');
    const checks=await page.evaluate(async ({before,after})=>{
      const {openZip}=await import('./package.js');
      const {translateShared}=await import('./formula-groups.js');
      const source=openZip(Uint8Array.from(before).buffer),dest=openZip(Uint8Array.from(after).buffer);
      const xml=async name=>new DOMParser().parseFromString(await dest.text(name),'application/xml');
      const sheet=await xml('xl/worksheets/sheet1.xml'),tables=await xml('xl/worksheets/sheet3.xml');
      const c=(ws,addr)=>ws.querySelector('c[r="'+addr+'"]');
      const f=(ws,addr)=>c(ws,addr).querySelector('f');
      const stable=source.names.filter(n=>!['xl/styles.xml','xl/workbook.xml'].includes(n)&&!/^xl\/worksheets\/sheet\d+\.xml$/.test(n));
      return {
        sameParts:stable.every(n=>dest.names.includes(n)&&source.packed(n).bytes.every((v,i)=>dest.packed(n).bytes[i]===v)),
        metadata:c(sheet,'F6').getAttribute('cm'), spill:[f(sheet,'F6').getAttribute('ref'),f(sheet,'F6').textContent],
        saved:c(sheet,'F6').querySelector('v').textContent, missing:c(sheet,'F7')?.querySelector('v')?.textContent??null,
        shared:[f(sheet,'D3').textContent,c(sheet,'D3').querySelector('v').textContent],
        table:[...f(tables,'E5').attributes].map(a=>[a.name,a.value]),
        stale:c(tables,'B9').querySelector('v').textContent,
        translated:translateShared('IF(A1="B3",LOG10($C3),SUM(A:B,$2:4,Sheet1!A$1,Table1[A1]))','C2','D3')
      };
    },{before:[...original],after:[...result]});
    assert.equal(checks.sameParts,true);assert.equal(checks.metadata,'1');
    assert.deepEqual(checks.spill,['F6:F7','_xlfn.UNIQUE(A2:A3)']);
    assert.equal(checks.saved,'41');assert.equal(checks.missing,null);
    assert.deepEqual(checks.shared,["$A3+C$1+$A$1+'Rates 1'!C3",'38']);
    assert.deepEqual(Object.fromEntries(checks.table),{t:'dataTable',ref:'E5:F6',dt2D:'1',dtr:'1',r1:'A1',r2:'B1'});
    assert.equal(checks.stale,'77');
    assert.equal(checks.translated,'IF(B2="B3",LOG10($C4),SUM(B:C,$2:5,Sheet1!B$1,Table1[A1]))');
    // Explicitly choosing ODS writes the reviewed values. Ordinary translated
    // formulas, including cross-sheet references, still recalculate after reopen.
    await button('save').click();
    const odsDownload=page.waitForEvent('download');
    await page.getByRole('button',{name:'Download ODS copy'}).click();
    const ods=await odsDownload, odsPath=path.join(temp,'values.ods');
    assert.match(ods.suggestedFilename(),/\.ods$/);
    await ods.saveAs(odsPath);
    await button('file').setInputFiles(odsPath);
    await page.getByRole('dialog',{name:'Import notes'}).waitFor();
    await button('confirm').click();await value('C2',21);await value('D3',38);await value('F6',41);
    assert.deepEqual(await fs.readFile(fixture),original);
    const opaque=await page.evaluate(async()=>{
      const {openZip,makeZip}=await import('./package.js'),{readFile,writeFile}=await import('./io.js');
      const {makeCalculator}=await import('./calc.js');
      const simple={sheets:[{name:'Shared',cells:{A1:{v:1},B1:{f:'A1+1'},B2:{f:'A2+1'}}}]};
      const zip=openZip(await (await writeFile(simple,'xlsx',0,makeCalculator(simple))).arrayBuffer());
      const parts=[];
      for(const name of zip.names){
        let part=zip.packed(name);
        if(name.endsWith('sheet1.xml')){
          part=(await zip.text(name)).replace('<f>A1+1</f><v>2</v>','<f t="shared" ref="B1:B2" si="0">PRODUCT(A1,2)</f><v>13</v>')
            .replace('<f>A2+1</f><v>1</v>','<f t="shared" si="0"/>');
        }
        parts.push([name,part]);
      }
      const before=await readFile(new File([makeZip(parts)],'shared.xlsx'));
      before.sheets[0].cells.C1={v:'unrelated'};
      const exported=await writeFile(before,'xlsx',0,makeCalculator(before));
      const after=await readFile(new File([exported],'copy.xlsx')),calc=makeCalculator(after);
      return [calc(0,'B1'),calc.status(0,'B1'),calc.status(0,'B2'),after.sheets[0].cells.B2.f,calc(0,'C1')];
    });
    assert.deepEqual(opaque,[13,'Last saved value','No saved result','PRODUCT(A2,2)','unrelated']);
    assert.deepEqual(errors,[]);
    if(process.env.FILEBIRCH_TEST_EXPORT)await fs.copyFile(copy,process.env.FILEBIRCH_TEST_EXPORT);
    console.log('PASS: shared references, arrays, spills, what-if scenarios, read-only outputs, stale/missing results, ODS choice, XLSX download, XML fidelity and browser reopen.');
  }finally{
    await browser?.close();await new Promise(resolve=>server.close(resolve));
    await fs.rm(temp,{recursive:true,force:true});
  }
}
run().catch(e=>{console.error(e);process.exitCode=1;});

