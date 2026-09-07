#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
rotation_calc.py — 정리정돈 담당 청년부 계산 (rotation.json 기준)

사용법:
    python3 rotation_calc.py <rotation.json> <YYYY-MM | YYMMDD>

규칙: overrides[YYYY-MM] 이 있으면 그 값. 없으면
      order[(index(anchor.team) + (대상월 - anchor.month 개월 차)) % len(order)].
출력: JSON {month, team, source: "override"|"rotation", anchor, order}
"""
import sys, json, datetime


def main(path, when):
    r = json.load(open(path, encoding="utf-8"))
    if len(when) == 6 and when.isdigit():
        d = datetime.datetime.strptime(when, "%y%m%d").date()
        ym = "%04d-%02d" % (d.year, d.month)
    else:
        ym = when
    y, m = [int(x) for x in ym.split("-")]
    ay, am = [int(x) for x in r["anchor"]["month"].split("-")]
    order = r["order"]
    ov = r.get("overrides", {})
    if ym in ov:
        team, src = ov[ym], "override"
    else:
        diff = (y - ay) * 12 + (m - am)
        team, src = order[(order.index(r["anchor"]["team"]) + diff) % len(order)], "rotation"
    print(json.dumps({"month": ym, "team": team, "source": src,
                      "anchor": r["anchor"], "order": order}, ensure_ascii=False))


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print(__doc__); sys.exit(2)
    main(sys.argv[1], sys.argv[2])
