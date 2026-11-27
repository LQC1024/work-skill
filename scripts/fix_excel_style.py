#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""修复 LibreOffice 保存后 xlsx 的颜色兼容问题（Excel / WPS 手机版 / 网页预览）。

LibreOffice 导出 xlsx 有两个坑：
  1) 纯色块 patternType="solid" 的 bgColor 被写成它自己的内部调色板色
     （红→棕 #993300、绿→灰 #C0C0C0、蓝→青 #33CCCC），
     WPS / 网页预览按 bgColor 取色 → 显示不出原色。
  2) cellXfs 里的 <xf> 缺少 applyFill="true"，
     WPS 手机版严格按标记渲染 → 直接不画填充（白板）。

做法：
  - 纯色块 bgColor 统一改成与 fgColor 一致
  - 所有 <xf> 补上 applyFill="true"
  - 页面设为 横向 + 缩放到一页宽

用法:
    python3 fix_excel_style.py 文件.xlsx
    python3 fix_excel_style.py 文件.xlsx --no-page    # 不改打印设置

提示：在 LibreOffice 里编辑并另存为 xlsx 后，每次都要重新跑一次。
"""
import argparse
import os
import re
import shutil
import sys
import zipfile

FILL_PAT = re.compile(
    r'<patternFill patternType="solid">\s*'
    r'<fgColor rgb="([0-9A-Fa-f]{8})"/>\s*'
    r'(?:<bgColor rgb="[0-9A-Fa-f]{8}"/>)?\s*'
    r'</patternFill>'
)


def fix_fill(m):
    fg = m.group(1)
    return (f'<patternFill patternType="solid">'
            f'<fgColor rgb="{fg}"/><bgColor rgb="{fg}"/>'
            f'</patternFill>')


def add_applyfill(m):
    xf = m.group(0)
    if "applyFill=" in xf:
        return xf
    return re.sub(r'(fillId="\d+")', r'\1 applyFill="true"', xf, count=1)


def fix(path, fix_page=True, backup=True):
    path = os.path.abspath(path)
    if not os.path.exists(path):
        sys.exit(f"找不到文件: {path}")
    if backup:
        shutil.copy(path, path + ".bak")

    zin = zipfile.ZipFile(path)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()

    n_fill = n_apply = 0
    out = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    for info, data in items:
        if info.filename == "xl/styles.xml":
            s = data.decode("utf-8")
            s, n_fill = FILL_PAT.subn(fix_fill, s)
            m = re.search(r"<cellXfs.*?</cellXfs>", s, re.S)
            if m:
                block = m.group(0)
                newblock = re.sub(r"<xf [^>]*>", add_applyfill, block)
                n_apply = newblock.count('applyFill="true"')
                s = s.replace(block, newblock)
            data = s.encode("utf-8")
        elif fix_page and info.filename.startswith("xl/worksheets/sheet"):
            s = data.decode("utf-8")
            s = s.replace('<pageSetUpPr fitToPage="false"/>',
                          '<pageSetUpPr fitToPage="true"/>')
            s = s.replace('orientation="portrait"', 'orientation="landscape"')
            # 去掉 LibreOffice 自动加的"页眉=工作表名 / 页脚=页N"
            s = re.sub(r'<headerFooter.*?</headerFooter>|<headerFooter[^>]*/>', '', s, flags=re.S)
            data = s.encode("utf-8")
        out.writestr(info, data)
    out.close()
    return n_fill, n_apply


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--no-page", action="store_true", help="不修改打印设置")
    ap.add_argument("--no-backup", action="store_true", help="不生成 .bak 备份")
    args = ap.parse_args()

    n_fill, n_apply = fix(args.xlsx, fix_page=not args.no_page,
                          backup=not args.no_backup)
    print(f"已修正纯色块 {n_fill} 个，补 applyFill {n_apply} 处：{os.path.abspath(args.xlsx)}")
    if not args.no_backup:
        print(f"修复前备份：{os.path.abspath(args.xlsx)}.bak")


if __name__ == "__main__":
    main()
