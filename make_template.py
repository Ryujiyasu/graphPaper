"""Writes template.xlsx: two empty A4 sheets of square cells, portrait and
landscape, that the app copies for every sheet it saves.

Every cell is 20 x 20 pixels in Excel: the workbook's normal font is
Yu Gothic 11 (8-pixel digit width), so a column 2.5 wide is 20 pixels, and a
row 15 points high is 20 pixels. 34 columns x 50 rows fill an A4 page inside
15 mm margins (180 x 267 mm); 50 x 34 the same page turned. Printing fits
the width to the page and lets a taller grid run on to more pages.
"""
import sys, zipfile

import sys, zipfile

def col(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

MARGIN = 15 / 25.4
SHEETS = [('縦', 34, 50, 'portrait'), ('横', 50, 34, 'landscape')]

def sheet(cols, rows, orientation):
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr><dimension ref="A1:{col(cols)}{rows}"/><sheetViews><sheetView view="pageLayout" workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15" customHeight="1"/><cols><col min="1" max="{cols}" width="2.5" customWidth="1"/></cols><sheetData>{''.join(f'<row r="{r}" ht="15" customHeight="1"/>' for r in range(1, rows + 1))}</sheetData><pageMargins left="{MARGIN:.4f}" right="{MARGIN:.4f}" top="{MARGIN:.4f}" bottom="{MARGIN:.4f}" header="0" footer="0"/><pageSetup paperSize="9" orientation="{orientation}" fitToWidth="1" fitToHeight="0"/></worksheet>'''

files = {
'[Content_Types].xml': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>''',
'_rels/.rels': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>''',
'docProps/app.xml': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>''',
'docProps/core.xml': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>graphPaper</dc:creator></cp:coreProperties>''',
'xl/_rels/workbook.xml.rels': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/></Relationships>''',
'xl/workbook.xml': f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="縦" sheetId="1" r:id="rId1"/><sheet name="横" sheetId="2" r:id="rId3"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">縦!$A$1:$AH$50</definedName><definedName name="_xlnm.Print_Area" localSheetId="1">横!$A$1:$AX$34</definedName></definedNames></workbook>''',
'xl/styles.xml': '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><color theme="1"/><name val="游ゴシック"/><family val="3"/><charset val="128"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"><alignment vertical="center"/></xf></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment vertical="center"/></xf></cellXfs><cellStyles count="1"><cellStyle name="標準" xfId="0" builtinId="0"/></cellStyles></styleSheet>''',
'xl/worksheets/sheet1.xml': sheet(34, 50, 'portrait'),
'xl/worksheets/sheet2.xml': sheet(50, 34, 'landscape'),
}
with zipfile.ZipFile(sys.argv[1] if len(sys.argv) > 1 else 'template.xlsx', 'w', zipfile.ZIP_DEFLATED) as z:
    for name, text in files.items():
        z.writestr(name, text)
