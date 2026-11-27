#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""按 10 天循环规律，自动算出某月"每天哪个值上中班"，并把各值菜谱按顺序排进去，
输出 render_schedule.py 需要的 schedule-data.json。

用户只需要提供：月份 + 各值菜谱。排班规律和锚点内置。

用法:
    python3 plan_schedule.py --month 2027-01 --menus menus.json --out data.json
    python3 plan_schedule.py --month 2027-01 --weeks 4 --menus menus.json --out data.json
    python3 plan_schedule.py --start 2027-01-01 --days 28 --menus menus.json --out data.json

menus.json 格式（每个值是"菜谱套"的列表，每套用字符串，顿号/逗号分隔）:
{
  "A": ["菜1、菜2、菜3、主食，汤", "..."],
  "B": ["...", "..."],
  "C": ["..."],
  "D": ["..."],
  "E": ["..."]
}

拆分规则：每套取前 3 项 → 1/2/3 菜品，其余项用 "/" 连接 → 4 主食。
菜谱套数少于该值中班天数时自动循环取用。
"""
import argparse
import datetime as dt
import json

# —— 排班规律常量（改这里即可适配别的站点/轮值）——
CYCLE_VALUES = ["C", "D", "E", "A", "B"]   # 交接顺序，每 2 天换一个值，5 值 10 天一循环
ANCHOR_DATE = dt.date(2026, 10, 5)          # 锚点日期
ANCHOR_VALUE = "C"                          # 该日期对应的值
ANCHOR_POS = 1                              # 该日期处于 10 天周期的第几位（1 = 中中第一天）

WEEKDAY_CN = ["一", "二", "三", "四", "五", "六", "日"]
DEFAULT_SLOTS = ["1菜品", "2菜品", "3菜品", "4主食"]
SEPS = "、，,；;"


def value_for(date):
    n = (date - ANCHOR_DATE).days
    base = CYCLE_VALUES.index(ANCHOR_VALUE)
    return CYCLE_VALUES[(base + n // 2) % len(CYCLE_VALUES)]


def pos_for(date):
    n = (date - ANCHOR_DATE).days
    return (ANCHOR_POS - 1 + n) % 10 + 1


def split_set(s, n_slots=4):
    items = [x.strip() for x in s.replace("，", "、").replace(",", "、").split("、") if x.strip()]
    dishes = items[: n_slots - 1]
    rest = items[n_slots - 1:]
    dishes.append("/".join(rest) if rest else "")
    while len(dishes) < n_slots:
        dishes.append("")
    return dishes


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", help="YYYY-MM，整月")
    ap.add_argument("--start", help="YYYY-MM-DD，起算日（默认月首日）")
    ap.add_argument("--days", type=int, help="共排几天（默认整月）")
    ap.add_argument("--weeks", type=int, help="共排几周（每周 7 天，优先于 --days）")
    ap.add_argument("--menus", required=True, help="菜谱 JSON")
    ap.add_argument("--out", required=True, help="输出 schedule-data.json")
    ap.add_argument("--title", default="", help="标题，默认自动")
    ap.add_argument("--people-note", default="20人以内，按照餐标每人预算20元以内（含牛奶每支3.85元）")
    args = ap.parse_args()

    if args.start:
        start = dt.date.fromisoformat(args.start)
    elif args.month:
        y, m = map(int, args.month.split("-"))
        start = dt.date(y, m, 1)
    else:
        ap.error("需要 --month 或 --start")

    if args.weeks:
        days = args.weeks * 7
    elif args.days:
        days = args.days
    else:
        y, m = start.year, start.month
        nxt = dt.date(y + (m == 12), 1 if m == 12 else m + 1, 1)
        days = (nxt - start).days

    with open(args.menus, encoding="utf-8") as f:
        menus = json.load(f)
    for v in set(CYCLE_VALUES):
        if v not in menus or not menus[v]:
            ap.error(f"菜谱缺少值 {v}")

    # 逐日算值，并按值累计已用菜谱套数（顺序取、循环）
    used = {v: 0 for v in menus}
    day_rows = []
    for i in range(days):
        d = start + dt.timedelta(days=i)
        v = value_for(d)
        sets = menus[v]
        dishes = split_set(sets[used[v] % len(sets)])
        used[v] += 1
        head = f"星期{WEEKDAY_CN[d.weekday()]}{v} {d.month}/{d.day}"
        day_rows.append({"date": d, "head": head, "value": v, "dishes": dishes})

    # 按周切 7 天
    weeks = []
    for w in range((days + 6) // 7):
        chunk = day_rows[w * 7:(w + 1) * 7]
        if not chunk:
            break
        weeks.append({
            "label": f"{start.month}月第{'一二三四五'[w]}周" if start.month == chunk[0]['date'].month
                     else f"{chunk[0]['date'].month}月第{'一二三四五'[w]}周",
            "days": [{"head": r["head"], "value": r["value"], "dishes": r["dishes"]} for r in chunk],
        })

    first, last = day_rows[0]["date"], day_rows[-1]["date"]
    title = args.title or f"运行宵夜排餐表 {first:%m.%d}-{last:%m.%d}"
    out = {
        "title": title,
        "notes": ["备注：", "1、纯牛奶和酸奶各送一半（缺货情况除外）。"],
        "people_note": args.people_note,
        "slots": DEFAULT_SLOTS,
        "weeks": weeks,
    }
    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    # 统计
    from collections import Counter
    cnt = Counter(r["value"] for r in day_rows)
    seq = "".join(r["value"] for r in day_rows)
    print(f"已生成: {args.out}")
    print(f"范围: {first} ~ {last}（{days} 天）")
    print(f"逐日: {seq}")
    print(f"各值天数: {dict(sorted(cnt.items()))}")


if __name__ == "__main__":
    main()
