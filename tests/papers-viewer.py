#!/usr/bin/env python3
"""Browser-level Papers test: python tests/papers-viewer.py. Requires Playwright + Chromium."""
import base64
import io
import os
from pathlib import Path
from functools import partial
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from threading import Thread
from zipfile import ZipFile, ZIP_DEFLATED
from playwright.sync_api import sync_playwright, expect

PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+nm2sAAAAASUVORK5CYII=")
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
X = "urn:oasis:names:tc:opendocument:xmlns:"
O = "application/vnd.oasis.opendocument.text"
ODT_NS = ' '.join(f'xmlns:{p}="{X}{u}:1.0"' for p,u in {
    'office':'office','text':'text','table':'table','draw':'drawing','style':'style',
    'fo':'xsl-fo-compatible','svg':'svg-compatible'}.items())+' xmlns:xlink="http://www.w3.org/1999/xlink"'

def archive(parts):
    out = io.BytesIO()
    with ZipFile(out, "w", ZIP_DEFLATED) as z:
        for path, content in parts.items(): z.writestr(path, content)
    return out.getvalue()

def drawing(id):
    return f'<w:r><w:drawing><a:blip r:embed="{id}"/></w:drawing></w:r>'

def docx(forbidden=''):
    return archive({
        'word/document.xml': f'{forbidden}<w:document xmlns:w="{W}" xmlns:a="{A}" xmlns:r="{R}"><w:body>'
         '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Viewing a DOCX</w:t></w:r></w:p>'
         '<w:p><w:hyperlink r:id="remote"><w:r><w:rPr><w:b/><w:i/><w:color w:val="334455"/></w:rPr><w:t>&lt;script&gt;danger&lt;/script&gt;</w:t></w:r></w:hyperlink>'
         + drawing('picture') + drawing('remote') + drawing('fakePicture') + '</w:p>'
         '<w:p><w:pPr><w:numPr><w:numId w:val="5"/></w:numPr></w:pPr><w:r><w:t>First item</w:t></w:r></w:p>'
         '<w:tbl><w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Table heading</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
         '<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240"/></w:sectPr></w:pPr><w:r><w:br w:type="page"/></w:r></w:p>'
         '</w:body></w:document>',
        'word/styles.xml': f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:sz w:val="44"/></w:rPr></w:style></w:styles>',
        'word/numbering.xml': f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="3"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum><w:num w:numId="5"><w:abstractNumId w:val="1"/></w:num></w:numbering>',
        'word/_rels/document.xml.rels':f'<Relationships xmlns="{P}"><Relationship Type="{R}/image" Id="picture" Target="media/real.png"/><Relationship Type="{R}/image" Id="fakePicture" Target="media/fake.png"/><Relationship Type="{R}/image" Id="remote" TargetMode="External" Target="https://preview.invalid/remote.png"/></Relationships>',
        'word/media/real.png':PNG, 'word/media/fake.png':'<svg onload="throw new Error(99)"/>', 'word/vbaProject.bin':b'ignored',
    })

def odt(repeat='1'):
    return archive({
        'mimetype':O,
        'content.xml':f'<office:document-content {ODT_NS}><office:automatic-styles><style:style style:name="strong" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style><text:list-style style:name="ordered"><text:list-level-style-number text:level="1" style:num-format="a" text:start-value="2"/></text:list-style></office:automatic-styles><office:body><office:text>'
         '<text:h text:outline-level="2">Viewing an ODT</text:h>'
         '<text:p><text:span text:style-name="strong">Run</text:span><text:s text:c="3"/>text<text:a xlink:href="https://preview.invalid/bad">link</text:a>'
         '<draw:frame text:anchor-type="as-char" svg:width="32px"><draw:image xlink:href="Pictures/test.png"/></draw:frame>'
         '<draw:frame text:anchor-type="as-char"><draw:image xlink:href="https://preview.invalid/bad.png"/></draw:frame></text:p>'
         '<text:list text:style-name="ordered"><text:list-item><text:p>Letter</text:p></text:list-item></text:list>'
         '<table:table><table:table-row table:number-rows-repeated="'+repeat+'"><table:table-cell table:number-columns-spanned="2"><text:p>Cell</text:p></table:table-cell></table:table-row></table:table>'
         '<text:section text:name="B"><text:p>Section<strong xmlns="https://preview.invalid/XML">ignored</strong></text:p></text:section>'
         '</office:text></office:body></office:document-content>',
        'Pictures/test.png':PNG,
    })

def upload(page,name,buffer):
    page.locator('#papers-file').set_input_files(dict(name=name,mimeType='application/octet-stream',buffer=buffer))

def main():
    top = Path(os.environ.get('FILEBIRCH_PAPERS_REPO', Path(__file__).absolute().parent.parent))
    httpd = ThreadingHTTPServer(('127.0.0.1', 0), partial(SimpleHTTPRequestHandler, directory=top))
    Thread(target=httpd.serve_forever, daemon=True).start()
    origin = 'http://127.0.0.1:'+str(httpd.server_port)
    try:
        with sync_playwright() as pw:
            browser = pw.chromium.launch(executable_path=os.environ.get('CHROMIUM_PATH','/usr/bin/chromium'))
            page = browser.new_page(viewport={"width":390,"height":844})
            browser_errors, outbound = [], []
            page.on('pageerror',lambda e:browser_errors.append(str(e)))
            page.on('request',lambda r:outbound.append(r.url) if not r.url.startswith((origin,'data:')) else None)
            page.goto(origin+'/documents/')
            expect(page.locator('#papers-file')).to_have_count(0)
            page.get_by_role('button',name='I agree to the Terms of Service').click()
            upload(page,'Papers.docx',docx())
            expect(page.get_by_text('Papers.docx · DOCX preview')).to_be_visible()
            expect(page.get_by_role('heading',name='Viewing a DOCX',level=1)).to_be_visible()
            expect(page.get_by_text('<script>danger</script>',exact=True)).to_have_css('font-style','italic')
            expect(page.locator('article.papers-page script,article.papers-page a')).to_have_count(0)
            expect(page.locator('.papers-page ol')).to_have_attribute('start','3')
            expect(page.get_by_role('columnheader',name='Table heading')).to_have_attribute('colspan','2')
            expect(page.get_by_text('Section break',exact=True)).to_be_visible()
            expect(page.locator('.papers-page img')).to_have_count(1)
            expect(page.locator('.papers-page img')).to_have_js_property('naturalWidth',1)
            expect(page.get_by_text('Macros are ignored; they never run.')).to_be_visible()
            page.get_by_role('button',name='Close document').click()
            expect(page.locator('article.papers-page')).to_have_count(0)
            page.reload()
            expect(page.locator('#papers-file')).to_be_visible()
            upload(page,'Papers.odt',odt())
            expect(page.get_by_role('heading',name='Viewing an ODT',level=2)).to_be_visible()
            expect(page.get_by_text('Run',exact=True)).to_have_css('font-weight','700')
            expect(page.locator('.papers-page ol')).to_have_attribute('type','a')
            expect(page.locator('.papers-page ol')).to_have_attribute('start','2')
            expect(page.get_by_role('cell',name='Cell')).to_have_attribute('colspan','2')
            expect(page.locator('.papers-page img')).to_have_js_property('naturalWidth',1)
            expect(page.get_by_text('[Image omitted]',exact=True)).to_be_visible()
            expect(page.get_by_text('ignored',exact=True)).to_have_count(0)
            upload(page,'blocked.docx',docx('<!DOCTYPE r [<!ENTITY x "bad">]>'))
            expect(page.get_by_role('status')).to_contain_text('entity declarations')
            expect(page.locator('article.papers-page')).to_have_count(0)
            upload(page,'repeat.odt',odt('900000'))
            expect(page.get_by_role('status')).to_contain_text('safe preview limit')
            upload(page,'wrong.docx',archive({'content.xml':'<document/>'}))
            expect(page.get_by_role('status')).to_contain_text('required document part')
            upload(page,'text.txt',b'<b>Plain</b>\nhello')
            expect(page.locator('.papers-page pre')).to_have_text('<b>Plain</b>\nhello',use_inner_text=True)
            expect(page.locator('.papers-page b')).to_have_count(0)
            upload(page,'not UTF-8.txt',b'\xff')
            expect(page.get_by_role('status')).to_contain_text('Could not open')
            assert not outbound, outbound
            assert not browser_errors, browser_errors
            print('Papers browser checks passed: terms, DOCX, ODT, TXT, images, styles, section/list/table, safety limits, and no outbound requests.')
            browser.close()
    finally:
        httpd.shutdown()

if __name__=='__main__': main()
