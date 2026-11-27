#!/usr/bin/env bash
# Excel 2003 XML（SpreadsheetML）→ xlsx，调用 LibreOffice 无头转换。
#
# 用法:
#   bash scripts/xml_to_xlsx.sh out.xml .          # 输出到当前目录
#   bash scripts/xml_to_xlsx.sh out.xml /tmp/out
#
# 依赖: soffice / libreoffice
set -euo pipefail

if [ $# -lt 1 ]; then
  echo "用法: bash xml_to_xlsx.sh <文件.xml> [输出目录(默认: xml 所在目录)]"
  exit 1
fi

XML="$1"
OUTDIR="${2:-$(dirname "$XML")}"

if ! command -v soffice >/dev/null 2>&1 && ! command -v libreoffice >/dev/null 2>&1; then
  echo "未找到 LibreOffice（soffice）。请先安装，或改用 Excel/WPS 手动打开 .xml 另存为 .xlsx。" >&2
  exit 1
fi

BIN=$(command -v soffice || command -v libreoffice)
"$BIN" --headless --convert-to xlsx --outdir "$OUTDIR" "$XML" 2>&1 | grep -v javaldx || true

BASE=$(basename "$XML")
BASE="${BASE%.*}"
echo "已输出: $OUTDIR/$BASE.xlsx"
echo "下一步（重要）: python3 scripts/fix_excel_style.py \"$OUTDIR/$BASE.xlsx\""
