"""Generate fictional workbooks for the Tables browser tests; no Office install required."""
from pathlib import Path
from sys import argv
from zipfile import ZipFile, ZIP_DEFLATED

dest = Path(argv[1])
MS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
T = "application/vnd.openxmlformats-officedocument.spreadsheetml"
ODS = "application/vnd.oasis.opendocument.spreadsheet"

def zip_file(file, parts):
    with ZipFile(dest / file, "w", ZIP_DEFLATED) as z:
        for key, content in parts.items():
            z.writestr(key, content)

def xlsx(name, body, *, calendar1904=False, extras=None):
    def rel(i, part, filename):
        return f'<Relationship Id="{i}" Type="{R}/{part}" Target="{filename}"/>'
    wb = f'<workbook xmlns="{MS}" xmlns:r="{R}"><workbookPr date1904="{int(calendar1904)}"/><sheets><sheet name="Transactions" sheetId="1" r:id="s1"/><sheet name="Archive" state="veryHidden" sheetId="8" r:id="s8"/></sheets></workbook>'
    styles = f'<styleSheet xmlns="{MS}"><numFmts count="1"><numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/></numFmts><cellXfs count="5"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="20"/><xf numFmtId="165"/><xf numFmtId="46"/></cellXfs></styleSheet>'
    parts = {
        "[Content_Types].xml": f'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="{T}.sheet.main+xml"/></Types>',
        "_rels/.rels": f'<Relationships xmlns="{RELS}">{rel("doc","officeDocument","xl/workbook.xml")}</Relationships>',
        "xl/workbook.xml": wb,
        "xl/_rels/workbook.xml.rels": f'<Relationships xmlns="{RELS}">{rel("s1","worksheet","worksheets/sheet20.xml")}{rel("s8","worksheet","worksheets/sheet1.xml")}{rel("styles","styles","styles.xml")}</Relationships>',
        "xl/styles.xml": styles,
        "xl/worksheets/sheet1.xml": f'<worksheet xmlns="{MS}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Archived data</t></is></c></row></sheetData></worksheet>',
        "xl/worksheets/sheet20.xml": f'<worksheet xmlns="{MS}">{body}</worksheet>',
    }
    parts.update(extras or {})
    zip_file(name, parts)

body = '<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Merged XLSX heading</t></is></c></row>' \
    '<row r="2"><c r="A2" s="1"><v>45292</v></c><c r="B2" s="2"><v>0.35416666666666667</v></c><c r="C2" s="3"><v>45292.25</v></c><c r="D2" s="4"><v>1.520833333333333333</v></c><c r="E2" t="d"><v>2024-02-29T12:00:00Z</v></c></row>' \
    '<row r="3"><c r="A3"><f>WEBSERVICE("https://outside.invalid/data")</f><v>42</v></c><c r="B3"><f>SUM(1,2)</f><v/></c><c r="C3" t="e"><v>#REF!</v></c><c r="D3" t="b"><v>0</v></c></row>' \
    '<row r="41"><c r="Q41" t="inlineStr"><is><t>Far value</t></is></c></row></sheetData>'
xlsx("workbook.xlsx", body + '<mergeCells><mergeCell ref="A1:B1"/></mergeCells>')
xlsx("calendar1904.xlsx", '<sheetData><row><c r="A1" s="1"><v>1</v></c></row></sheetData>', calendar1904=True)
xlsx("merge-overlap.xlsx",body+'<mergeCells><mergeCell ref="A1:B1"/><mergeCell ref="B1:C1"/></mergeCells>')
xlsx("broken-xml.xlsx","<sheetData>")
xlsx("expanded.xlsx",body,extras={"metadata-large":b"a"*(65*1024*1024)})
namespaces = ' '.join(f'xmlns:{p}="urn:oasis:names:tc:opendocument:xmlns:{n}:1.0"' for p,n in [('office','office'),('table','table'),('text','text')])

def ods(name, rows):
    zip_file(name,{"mimetype":ODS,"content.xml":f'<office:document-content {namespaces}><office:body><office:spreadsheet><table:table table:name="ODS sheet">{rows}</table:table></office:spreadsheet></office:body></office:document-content>'})
ods("workbook.ods",'<table:table-row><table:table-cell table:number-columns-spanned="2" office:value-type="string"><text:p>Merged ODS heading</text:p></table:table-cell><table:covered-table-cell/></table:table-row>' \
    '<table:table-row><table:table-cell office:value-type="date" office:date-value="2025-02-08"/><table:table-cell office:value-type="time" office:time-value="PT13H30M"/>' \
    '<table:table-cell office:value-type="float" office:value="1000" table:formula="of:=SUM(950;50)"/><table:table-cell office:value-type="boolean" office:boolean-value="true" table:formula="of:=TRUE()"/>' \
    '<table:table-cell office:value-type="float" table:formula="of:=SUM(1;2)"/></table:table-row>')
ods("bad-bounds.ods",'<table:table-row table:number-rows-repeated="10001"><table:table-cell office:value-type="float" office:value="1"/></table:table-row>')
