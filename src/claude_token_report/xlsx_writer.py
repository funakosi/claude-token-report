"""最小限の .xlsx ライター (標準ライブラリのみ)。数式とキャッシュ値の両方を書く。

セル指定: Cell(value, style="", formula=None)
  - formula があれば value は計算済みのキャッシュ値として <v> に書く
"""
import zipfile
from dataclasses import dataclass
from datetime import date
from xml.sax.saxutils import escape

# style 名 -> (cellXfs の index)。順序は STYLES_XML と対応させる。
STYLE_IDS = {"": 0, "date": 1, "int": 2, "price": 3, "bold": 4, "usd": 5, "jpy": 6, "boldint": 7,
             "boldusd": 8, "boldjpy": 9, "rate": 10}

STYLES_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="5">
<numFmt numFmtId="164" formatCode="yyyy/mm/dd"/>
<numFmt numFmtId="165" formatCode="#,##0.0000"/>
<numFmt numFmtId="166" formatCode="&quot;$&quot;#,##0.00"/>
<numFmt numFmtId="167" formatCode="&quot;¥&quot;#,##0"/>
<numFmt numFmtId="168" formatCode="#,##0.00"/>
</numFmts>
<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>
<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="11">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
<xf numFmtId="3" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="166" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="167" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1"/>
<xf numFmtId="168" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
</styleSheet>"""

_EPOCH = date(1899, 12, 30)


@dataclass
class Cell:
    value: object = None
    style: str = ""
    formula: str = None


def col_letter(i):  # 0-based
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def _cell_xml(ref, c):
    if not isinstance(c, Cell):
        c = Cell(c)
    s = STYLE_IDS[c.style]
    v = c.value
    f = f"<f>{escape(c.formula)}</f>" if c.formula else ""
    if v is None and not f:
        return f'<c r="{ref}" s="{s}"/>' if s else ""
    if isinstance(v, date):
        v = (v - _EPOCH).days
    if isinstance(v, bool) or not isinstance(v, (int, float, str, type(None))):
        raise TypeError(type(v))
    if isinstance(v, str):
        if f:
            return f'<c r="{ref}" s="{s}" t="str">{f}<v>{escape(v)}</v></c>'
        return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{escape(v)}</t></is></c>'
    vx = "" if v is None else f"<v>{v!r}</v>"
    return f'<c r="{ref}" s="{s}">{f}{vx}</c>'


def _sheet_xml(rows, widths):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">']
    if widths:
        out.append("<cols>" + "".join(
            f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i, w in enumerate(widths)) + "</cols>")
    out.append("<sheetData>")
    for r, row in enumerate(rows, 1):
        cells = "".join(_cell_xml(f"{col_letter(c)}{r}", v) for c, v in enumerate(row))
        out.append(f'<row r="{r}">{cells}</row>')
    out.append("</sheetData></worksheet>")
    return "".join(out)


def write_xlsx(path, sheets):
    """sheets: [(シート名, 行のリスト, 列幅リスト or None)]"""
    n = len(sheets)
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                    for i in range(1, n + 1)) + "</Types>")
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>")
    wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          + "".join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>'
                    for i, (name, _, _) in enumerate(sheets, 1))
          + '</sheets><calcPr fullCalcOnLoad="1"/></workbook>')
    wb_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               + "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
                         for i in range(1, n + 1))
               + f'<Relationship Id="rId{n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
               "</Relationships>")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wb_rels)
        z.writestr("xl/styles.xml", STYLES_XML)
        for i, (_, rows, widths) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", _sheet_xml(rows, widths))
