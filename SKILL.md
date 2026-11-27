---
name: work-skill
description: 制作/续排"宵夜排餐表"等运行排班菜单 Excel。当用户提供上一周期的排班表（值别、班次、菜品槽位）与各值的新菜谱，需要按 10 天循环规律外推下一周期、生成带颜色分块的原版式 xlsx，或需要修复 LibreOffice / WPS 下 Excel 颜色不显示的问题时使用。
---

# work-skill — 排餐表制作

把"上一周期排班表 + 新菜谱"变成"下一周期的排餐表 xlsx"，并保证颜色在各种软件（Excel / WPS 手机版 / 网页预览）都正常显示。

## 什么时候用

- 用户提到：**排餐、宵夜菜单、排班表、值别轮值、菜谱续排、中班夜宵**
- 用户给了一张排班 Excel，要求"按这个规律排下个月/下四周"
- 用户抱怨"表格颜色在手机上不见了 / 打开没颜色"

## 核心规律（详见 references/roster-cycle.md）

```
作息 10 天一个周期（位置 1–10）：
中 · 中 · 白 · 休 · 休 · 白 · 夜 · 夜 · 休 · 休
```

- 每天**恰好一个值上中班**，且连上 **2 天（中中）**，之后交给下一个值
- 值交接顺序固定，每 5 对（10 天）一轮（例：C → D → E → A → B）
- **只有中班当天的那个值点夜宵**，该日的菜品即当日菜单
- 用"锚点"校验：用户会给出某值某天处于周期第几位，据此反推并校验各值相位

## 标准工作流

### 1. 读原表 → 还原规律
从上一周期表格还原：`值 × 日期 → 班次`，确认交接顺序、周期长度、锚点。

### 2. 把新菜谱整理成 JSON
按 `examples/schedule-data.json` 的结构：每周 7 天，每天含 `head`（星期+值+日期）、`value`（值别）、`dishes`（4 个槽位：1菜品/2菜品/3菜品/4主食）。

拆分规则：**前 3 项 → 1/2/3 菜品，其余项用 `/` 连接 → 4 主食**。若某套菜谱项数异常（3 个素菜、主食混在菜品位等），先向用户确认，不要自行脑补。

### 3. 生成 xlsx

**默认路线（零依赖，任何机器都能跑，不需要 LibreOffice / openpyxl）**
```bash
# 用户只给"月份 + 菜谱"时，两条命令搞定：
python3 scripts/plan_schedule.py --month 2027-01 --weeks 4 --menus menus.json --out data.json
python3 scripts/build_xlsx.py data.json 排餐表.xlsx
```

- `plan_schedule.py`：内置 10 天循环与锚点，自动算每天哪个值 + 按顺序套菜谱
- `build_xlsx.py`：纯标准库（zipfile + OOXML）直接写 xlsx，颜色按 `fgColor=bgColor + applyFill=1` 生成，Excel/WPS/网页都能显示

**备选路线（本机有 LibreOffice 时）**
```bash
python3 scripts/render_schedule.py data.json out.xml   # JSON → Excel 2003 XML
bash scripts/xml_to_xlsx.sh out.xml .                  # XML → xlsx
```

**已有第三方 xlsx 需要修颜色时**
```bash
python3 scripts/fix_excel_style.py 文件.xlsx
```

### 4. 交付前自检（必做）
- 颜色块：四个表头行逐格核对 `值 → 颜色` 对应
- 合并单元格：标题 1 + 备注 2 + 每周"人数备注"纵向合并
- 打印：横向 + 缩放到一页宽
- 把 xlsx 重新渲染成 PDF/PNG 亲自看一眼（`scripts/xml_to_xlsx.sh` 同目录有说明）

## 颜色标准（与原表一致）

| 值 | 颜色 | 色号 |
|---|---|---|
| A | 黄 | `FFFF00` |
| B | 红（白字） | `FF0000` |
| C | 绿 | `92D050` |
| D | 蓝 | `00B0F0` |
| E | 灰 | `D9D9D9` |

## 踩过的坑（详见 references/xlsx-compat.md）

1. **LibreOffice 存 xlsx 会把纯色块的 `bgColor` 写成它自己的调色板色**（红→棕、绿→灰、蓝→青）。
   WPS / 网页预览按 `bgColor` 取色 → 显示不出原色。
2. **`cellXfs` 里缺 `applyFill="true"`**，WPS 手机版直接不画填充 → 白板。

两种情况都跑 `python3 scripts/fix_excel_style.py 文件.xlsx` 即可修复。
**每次在 LibreOffice 里编辑并另存为 xlsx 后都要重新跑一次。**

## 脚本清单

| 脚本 | 作用 | 依赖 |
|---|---|---|
| `scripts/plan_schedule.py` | 月份 + 菜谱 → schedule-data.json（自动排班、套菜谱） | 无 |
| `scripts/build_xlsx.py` | JSON → xlsx（手写 OOXML） | 无 |
| `scripts/fix_excel_style.py` | 修 bgColor / applyFill / 打印设置 | 无 |
| `scripts/render_schedule.py` | JSON → Excel 2003 XML | 无 |
| `scripts/xml_to_xlsx.sh` | XML → xlsx | LibreOffice |

## 用户只需提供

1. **月份**（如"2027 年 1 月"或"1 月 1 日起四周"）
2. **各值菜谱**（A/B/C/D/E 各若干套，逗号/顿号分隔）

其余（每天哪个值、菜谱怎么轮、表格怎么排）全部自动完成。
若用户没给锚点/上月排班表，直接用内置锚点（2026-10-05 = C，周期第 1 位）。
