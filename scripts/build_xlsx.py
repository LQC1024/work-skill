#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""排餐数据 JSON → xlsx（纯标准库，零依赖，不需要 LibreOffice / openpyxl）。

xlsx 本质是一个 zip 包，这里直接用 zipfile + 手写 OOXML 生成，
颜色按"fgColor = bgColor + applyFill=1"写，Excel / WPS / 网页预览都能正常显示。

用法:
    python3 build_xlsx.py <数据.json> <输出.xlsx>

数据结构见 examples/schedule-data.json。
"""
import json
import sys
import zipfile
from xml.sax.saxutils import escape

DEFAULT_COLORS = {"A": "FFFF00", "B": "FF0000", "C": "92D050", "D": "00B0F0", "E": "D9D9D9"}
DEFAULT_SLOTS = ["1菜品", "2菜品", "3菜品", "4主食"]


def is_dark(rgb):
    r, g, b = int(rgb[0:2], 16), int(rgb[2:4], 16), int(rgb[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 120


class Styles:
    def __init__(self, colors):
        self.colors = colors
        self.white_font = {}
        self.fonts = [
            '<font><sz val="11"/><name val="宋体"/></font>',                       # 0 常规
            '<font><b/><sz val="18"/><name val="宋体"/></font>',                   # 1 标题
            '<font><b/><sz val="11"/><name val="宋体"/></font>',                   # 2 粗体
            '<font><sz val="9"/><name val="宋体"/></font>',                        # 3 小字
        ]
        self.fills = [
            '<fill><patternFill patternType="none"/></fill>',                      # 0
            '<fill><patternFill patternType="gray125"/></fill>',                   # 1
            '<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/>'
            '<bgColor rgb="FFFFFFFF"/></patternFill></fill>',                      # 2 白
        ]
        self.fill_index = {"white": 2}
        for v, rgb in colors.items():
            self.fills.append(
                f'<fill><patternFill patternType="solid">'
                f'<fgColor rgb="FF{rgb}"/><bgColor rgb="FF{rgb}"/></patternFill></fill>')
            self.fill_index[v] = len(self.fills) - 1
            if is_dark(rgb):
                self.fonts.append(
                    f'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="宋体"/></font>')
                self.white_font.setdefault(v, len(self.fonts) - 1)

        self.borders = [
            '<border><left/><right/><top/><bottom/><diagonal/></border>',          # 0 无
            '<border><left style="thin"><color rgb="FF000000"/></left>'
            '<right style="thin"><color rgb="FF000000"/></right>'
            '<top style="thin"><color rgb="FF000000"/></top>'
            '<bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>',  # 1 全框
        ]
        self.xfs = []
        self._add(0, 0, 0, None)                        # 0 默认
        self.title = self._add(1, 0, 0, ("center", False))
        self.note = self._add(0, 0, 0, ("left", False))
        self.note_bold = self._add(2, 0, 0, ("left", False))
        self.week = self._add(2, self.fill_index["white"], 1, ("center", True))
        self.hdr = {}
        self.leg = {}
        for v in colors:
            fi = self.white_font.get(v, 2)
            self.hdr[v] = self._add(fi, self.fill_index[v], 1, ("center", True))
            self.leg[v] = self._add(fi, self.fill_index[v], 1, ("center", False))
        self.slot = self._add(2, 0, 1, ("center", False))
        self.dish = self._add(0, 0, 1, ("center", True))
        self.people = self._add(3, 0, 1, ("center", True))

    def _add(self, font, fill, border, align):
        self.xfs.append((font, fill, border, align))
        return len(self.xfs) - 1

    def xml(self):
        def xf(font, fill, border, align):
            a = '<alignment'
            if align:
                a += f' horizontal="{align[0]}" vertical="center"'
                if align[1]:
                    a += ' wrapText="1"'
            else:
                a += ' horizontal="center" vertical="center"'
            a += "/>"
            return (f'<xf numFmtId="0" fontId="{font}" fillId="{fill}" borderId="{border}" '
                    f'xfId="0" applyFont="1" applyFill="1" applyBorder="1" '
                    f'applyAlignment="1">{a}</xf>')
        fonts = "".join(self.fonts)
        fills = "".join(self.fills)
        borders = "".join(self.borders)
        xfs = "".join(xf(*x) for x in self.xfs)
        return (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<fonts count="{len(self.fonts)}">{fonts}</fonts>'
            f'<fills count="{len(self.fills)}">{fills}</fills>'
            f'<borders count="{len(self.borders)}">{borders}</borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(self.xfs)}">{xfs}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>'
        )


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def build_sheet(data, st):
    colors = data.get("colors", DEFAULT_COLORS)
    slots = data.get("slots", DEFAULT_SLOTS)
    notes = data.get("notes", [])
    people = data.get("people_note", "")
    ncol = 9

    rows = []          # (cells, height)
    merges = []

    def row(cells, height=None):
        rows.append((cells, height))

    # 标题
    row([(0, data.get("title", "排餐表"), st.title)], 34)
    merges.append(f"A1:{col_letter(ncol)}1")
    r = 2
    for i, note in enumerate(notes):
        row([(0, note, st.note_bold if i == 0 else st.note)], 18)
        merges.append(f"A{r}:{col_letter(ncol)}{r}")
        r += 1

    for week in data["weeks"]:
        days = week["days"]
        cells = [(0, week["label"], st.week)]
        for j, d in enumerate(days):
            cells.append((1 + j, d["head"], st.hdr[d["value"]]))
        cells.append((8, people, st.people))
        row(cells, 28)
        merges.append(f"I{r}:I{r + len(slots) - 1}")
        r += 1
        for s in range(len(slots)):
            cells = [(0, slots[s], st.slot)]
            for j, d in enumerate(days):
                cells.append((1 + j, d["dishes"][s], st.dish))
            row(cells, 44)
            r += 1

    row([], 10)
    r += 1
    row([(0, "备注：", st.note_bold)], 18)
    r += 1
    for v in colors:
        row([(0, f"{v}值", st.leg[v])], 22)
        r += 1

    # 列宽
    widths = [12, 16, 16, 16, 16, 16, 16, 16, 24]
    cols = ('<cols>' + "".join(
        f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(widths)) + '</cols>')

    out = []
    for ri, (cells, height) in enumerate(rows, start=1):
        h = f' ht="{height}" customHeight="1"' if height else ""
        if not cells:
            out.append(f'<row r="{ri}"{h}/>')
            continue
        cxml = []
        for ci, text, style in cells:
            ref = f"{col_letter(ci + 1)}{ri}"
            cxml.append(f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">'
                        f'{escape(str(text))}</t></is></c>')
        out.append(f'<row r="{ri}"{h}>' + "".join(cxml) + "</row>")

    merge_xml = ('<mergeCells count="%d">' % len(merges)
                 + "".join(f'<mergeCell ref="{m}"/>' for m in merges) + '</mergeCells>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>'
        f'<dimension ref="A1:{col_letter(ncol)}{len(rows)}"/>'
        '<sheetViews><sheetView workbookViewId="0"/></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        f'{cols}<sheetData>{"".join(out)}</sheetData>{merge_xml}'
        '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
        '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="1"/>'
        '</worksheet>'
    )


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)
ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)
WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)


def workbook(sheet_name):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets><sheet name="{escape(sheet_name)}" sheetId="1" r:id="rId1"/></sheets>'
        '</workbook>'
    )


def main():
    if len(sys.argv) < 3:
        sys.exit("用法: python3 build_xlsx.py <数据.json> <输出.xlsx>")
    src, dst = sys.argv[1], sys.argv[2]
    with open(src, encoding="utf-8") as f:
        data = json.load(f)

    st = Styles(data.get("colors", DEFAULT_COLORS))
    sheet = build_sheet(data, st)

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", CONTENT_TYPES)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook(data.get("sheet_name", "排餐")))
        z.writestr("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        z.writestr("xl/styles.xml", st.xml())
        z.writestr("xl/worksheets/sheet1.xml", sheet)

    ndays = sum(len(w["days"]) for w in data["weeks"])
    print(f"已生成: {dst}（{len(data['weeks'])} 周，{ndays} 天）")


if __name__ == "__main__":
    main()
