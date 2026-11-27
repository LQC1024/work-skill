# Excel 颜色兼容性（LibreOffice / WPS / 网页预览）

## 症状

- 在 LibreOffice 里看颜色正常，传到**手机 WPS** 或**网页预览**就觉得"没颜色 / 颜色不对"。
- 典型表现：边框、文字都在，**只有底色不显示**（白板）。

## 根因 1：`bgColor` 与 `fgColor` 不一致

LibreOffice 另存 xlsx 时，纯色块写成：

```xml
<patternFill patternType="solid">
  <fgColor rgb="FFFF0000"/>   <!-- 真实填充色：红 -->
  <bgColor rgb="FF993300"/>   <!-- 被换成内部调色板色：棕 -->
</patternFill>
```

- Excel / LibreOffice 读 `fgColor` → 正常
- WPS、部分网页预览读 `bgColor` → 显示棕/灰/青

**修复**：把 `bgColor` 改成与 `fgColor` 相同。

## 根因 2：`cellXfs` 缺 `applyFill="true"`

LibreOffice 导出的样式只带 `applyFont / applyBorder / applyAlignment / applyProtection`：

```xml
<xf numFmtId="164" fontId="7" fillId="3" borderId="1" xfId="0"
    applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="true"/>
```

缺 `applyFill` 时，Excel 仍会填充，但 **WPS 手机版严格按标记渲染 → 不画填充**。

**修复**：给每个 `<xf>` 补 `applyFill="true"`。

## 一键修复

```bash
python3 scripts/fix_excel_style.py 文件.xlsx
```

同时做三件事：
1. `bgColor` 归一到 `fgColor`
2. 所有 `<xf>` 补 `applyFill="true"`
3. 页面设为横向 + 缩放到一页宽

**只要在 LibreOffice 里编辑并另存为 xlsx，这两个问题就会重新出现，必须再跑一次。**

## 自查方法（不打开文件）

```bash
python3 - <<'PY'
import zipfile, re
z = zipfile.ZipFile("文件.xlsx")
st = z.read("xl/styles.xml").decode()
print("纯色块:")
for m in re.findall(r"<patternFill patternType=\"solid\">.*?</patternFill>", st, re.S):
    print(" ", m)
print("缺 applyFill 的 xf 数:",
      len([x for x in re.findall(r"<xf [^>]*>", st) if "applyFill" not in x]))
PY
```

- 每个纯色块的 `fgColor` 和 `bgColor` 应一致
- "缺 applyFill 的 xf 数" 应为 0

## 终极验证

把 xlsx 渲染成图，亲眼确认：

```bash
soffice --headless --convert-to pdf --outdir /tmp 文件.xlsx
pdftoppm -png -r 80 -f 1 -l 1 /tmp/文件.pdf /tmp/preview
```

## 补充：为什么不直接写 xlsx

- 有 `openpyxl` 时：直接写单元格 + `PatternFill(start_color=..., fill_type="solid")`，**写完仍建议跑一次 `fix_excel_style.py`**。
- 没有 `openpyxl` 时：本仓库走 **JSON → SpreadsheetML(.xml) → LibreOffice 转 xlsx** 的纯标准库路线，生成的 xlsx 同样需要跑 `fix_excel_style.py`。
