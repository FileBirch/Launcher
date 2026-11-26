"""Generate the small regression workbook. Requires XlsxWriter only to rebuild the fixture."""
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import xml.etree.ElementTree as ET
import xlsxwriter

source = BytesIO()
with xlsxwriter.Workbook(source, {'in_memory': True}) as book:
    calc = book.add_worksheet('Calculation')
    rates = book.add_worksheet('Rates 1')
    tables = book.add_worksheet('Scenarios')
    calc.write_row('A1', [2, 3, 11])
    calc.write_column('A2', [4, 6])
    rates.write_row('B2', [7, 13])
    rates.write_row('B3', [17, 19])
    rates.add_table('A5:B7', {'data': [[1, 2], [3, 4]]})
    for target, source_ref in [('J2','B2'),('K2','C2'),('J3','B3'),('K3','C3')]:
        calc.write_formula(target, "='Rates 1'!" + source_ref)
    for row in range(1, 3):
        for col in range(2, 4):
            calc.write_formula(row, col, "=$A2+B$1+$A$1+J2", None, -99)
    calc.write_array_formula('F2:F3', '=A2:A3', None, -99)
    calc.write_array_formula('G2:G2', '=SUM(A2:A3)', None, -99)
    calc.write_dynamic_array_formula('F6:F7', '=UNIQUE(A2:A3)', None, 41)
    calc.write_formula('G6', '=F6+1', None, 42)
    tables.write_row('A1', [2, 3])
    formula = '=$A$1*10+$B$1'
    tables.write_formula('B4', formula)
    tables.write_column('A5', [7, 8])
    tables.write_formula('D4', formula)
    tables.write_row('E4', [4, 6])
    tables.write_column('D5', [7, 8])
    tables.write_formula('B8', '=PRODUCT(A1,B1)')
    tables.write_column('A9', [4, 6])
    tables.write_formula('D9', formula)
    tables.write_row('E8', [4, 6])
    for ref in ['B5', 'B6', 'E5', 'F5', 'E6', 'F6', 'E9', 'F9', 'B9']:
        tables.write_number(ref, 77)

ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
ET.register_namespace('', ns)
q = lambda tag: '{' + ns + '}' + tag
def cell(root, ref):
    return root.find(".//" + q('c') + "[@r='" + ref + "']")

result = Path(__file__).with_name('formula-groups.xlsx')
with ZipFile(source) as src, ZipFile(result, 'w', ZIP_DEFLATED) as out:
    for info in src.infolist():
        data = src.read(info)
        if info.filename == 'xl/worksheets/sheet1.xml':
            root = ET.fromstring(data)
            for ref in ['C2', 'D2', 'C3', 'D3']:
                f = cell(root, ref).find(q('f'))
                f.attrib.update(t='shared', si='0')
                if ref == 'C2':
                    f.set('ref', 'C2:D3')
                else:
                    f.text = None
            c = cell(root, 'F7')
            if c is not None:
                c.remove(c.find(q('v')))
            data = ET.tostring(root, encoding='utf-8')
        if info.filename == 'xl/worksheets/sheet3.xml':
            root = ET.fromstring(data)
            for anchor, ref, inputs in [
                ('B5', 'B5:B6', {'r1': 'A1'}),
                ('E5', 'E5:F6', {'dt2D': '1', 'dtr': '1', 'r1': 'A1', 'r2': 'B1'}),
                ('E9', 'E9:F9', {'dtr': '1', 'r1': 'A1'}),
                ('B9', 'B9:B10', {'r1': 'A1'}),
            ]:
                f = ET.Element(q('f'), {'t': 'dataTable', 'ref': ref, **inputs})
                cell(root, anchor).insert(0, f)
            data = ET.tostring(root, encoding='utf-8')
        out.writestr(info, data)

