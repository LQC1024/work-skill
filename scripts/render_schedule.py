#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""排餐数据 JSON → Excel 2003 XML（SpreadsheetML）。

纯标准库实现，不依赖 openpyxl。

用法:
    python3 render_schedule.py <数据.json> <输出.xml>

数据结构见 examples/schedule-data.json。
生成的 .xml 用 Excel / LibreOffice 可直接打开，再另存为 .xlsx；
或用 scripts/xml_to_xlsx.sh 无头转换。
"""
import json
import sys
import html

DEFAULT_COLORS = {
    "A": "FFFF00",  # 黄
    "B": "FF0000",  # 红
    "C": "92D050",  # 绿
    "D": "00B0F0",  # 蓝
    "E": "D9D9D9",  # 灰
}
DEFAULT_SLOTS = ["1菜品", "2菜品", "3菜品", "4主食"]

BORDER = (
    '<Borders>'
    '<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1"/>'
    '<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1"/>'
    '<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1"/>'
    '<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1"/>'
    '</Borders>'
)


def esc(s):
    return html.escape(str(s), quote=True)


def is_dark(rgb):
    """是否需要白色字体（深底白字）。"""
    r = int(rgb[0:2], 16)
    g = int(rgb[2:4], 16)
    b = int(rgb[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) < 120


def style(sid, *, bold=False, size=11, color=None, fill=None,
          horiz="Center", wrap=False, border=True):
    parts = [f'<Style ss:ID="{sid}">']
    align = f'<Alignment ss:Horizontal="{horiz}" ss:Vertical="Center"'
    if wrap:
        align += ' ss:WrapText="1"'
    align += "/>"
    parts.append(align)
    font = f'<Font ss:FontName="宋体" ss:Size="{size}"'
    if bold:
        font += ' ss:Bold="1"'
    if color:
        font += f' ss:Color="#{color}"'
    font += "/>"
    parts.append(font)
    if fill:
        parts.append(f'<Interior ss:Color="#{fill}" ss:Pattern="Solid"/>')
    if border:
        parts.append(BORDER)
    parts.append("</Style>")
    return "".join(parts)


def build_styles(colors):
    out = [
        '<Style ss:ID="Default" ss:Name="Normal">'
        '<Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>'
        '<Font ss:FontName="宋体" ss:Size="11"/></Style>',
    ]
    out.append(style("sTitle", bold=True, size=18, wrap=False, border=False))
    out.append(style("sNote", horiz="Left", border=False))
    out.append(style("sNoteBold", bold=True, horiz="Left", border=False))
    out.append(style("sWeek", bold=True, fill="FFFFFF"))
    for val, rgb in colors.items():
        fc = "FFFFFF" if is_dark(rgb) else None
        out.append(style(f"sHdr{val}", bold=True, fill=rgb, color=fc, wrap=True))
        out.append(style(f"sLeg{val}", bold=True, fill=rgb, color=fc))
    out.append(style("sSlot", bold=True))
    out.append(style("sDish", wrap=True))
    out.append(style("sPeople", size=9, wrap=True))
    return "\n  ".join(out)


def cell(text, sid, merge=None):
    attrs = f' ss:StyleID="{sid}"'
    if merge:
        attrs += f' ss:{merge[0]}="{merge[1]}"'
    return f'<Cell{attrs}><Data ss:Type="String">{esc(text)}</Data></Cell>'


def build(data):
    colors = data.get("colors", DEFAULT_COLORS)
    slots = data.get("slots", DEFAULT_SLOTS)
    notes = data.get("notes", [])
    people = data.get("people_note", "")
    ncol = 1 + 7 + 1  # 周标签 + 7 天 + 备注人数

    rows = []
    # 标题
    rows.append(('<Row ss:Height="32">'
                 + cell(data.get("title", "排餐表"), "sTitle", ("MergeAcross", ncol - 1))
                 + "</Row>"))
    # 备注
    for i, note in enumerate(notes):
        sid = "sNoteBold" if i == 0 else "sNote"
        rows.append(f'<Row ss:Height="18">{cell(note, sid, ("MergeAcross", ncol - 1))}</Row>')

    for week in data["weeks"]:
        days = week["days"]
        # 表头行
        hdr = [cell(week["label"], "sWeek")]
        for d in days:
            hdr.append(cell(d["head"], f'sHdr{d["value"]}'))
        hdr.append(cell(people, "sPeople", ("MergeDown", len(slots) - 1)))
        rows.append('<Row ss:Height="28">' + "".join(hdr) + "</Row>")
        # 菜品行
        for s in range(len(slots)):
            line = [cell(slots[s], "sSlot")]
            for d in days:
                line.append(cell(d["dishes"][s], "sDish"))
            rows.append('<Row ss:Height="32">' + "".join(line) + "</Row>")

    # 图例
    rows.append('<Row ss:Height="10"/>')
    rows.append(f'<Row ss:Height="18">{cell("备注：", "sNoteBold")}</Row>')
    for val in colors:
        rows.append(f'<Row ss:Height="22">{cell(f"{val}值", f"sLeg{val}")}</Row>')

    widths = [62] + [98] * 7 + [150]
    colxml = "".join(f'<Column ss:Width="{w}"/>' for w in widths)

    return f'''<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  {build_styles(colors)}
 </Styles>
 <Worksheet ss:Name="排餐">
  <Table ss:DefaultRowHeight="15">
   {colxml}
   {"".join(rows)}
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Layout x:Orientation="Landscape"/>
    <PageMargins x:Bottom="0.5" x:Left="0.4" x:Right="0.4" x:Top="0.5"/>
   </PageSetup>
   <FitToPage/>
   <Print><FitWidth>1</FitWidth><FitHeight>1</FitHeight><ValidPrinterInfo/></Print>
  </WorksheetOptions>
 </Worksheet>
</Workbook>
'''


def main():
    if len(sys.argv) < 3:
        sys.exit("用法: python3 render_schedule.py <数据.json> <输出.xml>")
    src, dst = sys.argv[1], sys.argv[2]
    with open(src, encoding="utf-8") as f:
        data = json.load(f)
    xml = build(data)
    with open(dst, "w", encoding="utf-8") as f:
        f.write(xml)
    ndays = sum(len(w["days"]) for w in data["weeks"])
    print(f"已生成: {dst}（{len(data['weeks'])} 周，{ndays} 天）")


if __name__ == "__main__":
    main()
