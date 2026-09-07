#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_master.py — 마스터 템플릿 구조 확인 (config.json 의 templates / replacements 작성용)

마스터마다 템플릿 슬라이드가 쓰는 레이아웃 번호와 텍스트 run 분할이 다르다.
치환(apply_repl)은 <a:t>OLD</a:t> run 전체 일치로만 동작하므로, 추측하지 말고
이 스크립트로 실제 run 을 확인한 뒤 치환 쌍을 만든다.

사용법:
    python3 inspect_master.py <마스터.pptx>

출력(JSON):
  slide_count, layout_count
  slides: [{n, layout, runs:[...]}]           # 슬라이드별 레이아웃 번호와 <a:t> run 목록
  templates_suggest: {cover,intro,newfamily,grid,end,layout_cover,layout_title}
     - 슬라이드 1~5 를 표지/정리정돈/새가족/사진그리드/끝 으로 보고,
       layout_cover = 표지 슬라이드의 레이아웃, layout_title = 정리정돈 슬라이드의 레이아웃.
"""
import zipfile, re, sys, json, posixpath


def main(path):
    z = zipfile.ZipFile(path)
    names = set(z.namelist())
    slides = sorted(int(m.group(1)) for n in names for m in [re.match(r'ppt/slides/slide(\d+)\.xml$', n)] if m)
    layouts = [n for n in names if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$', n)]
    out = []
    for n in slides:
        xml = z.read("ppt/slides/slide%d.xml" % n).decode("utf-8")
        rels_name = "ppt/slides/_rels/slide%d.xml.rels" % n
        lay = None
        if rels_name in names:
            m = re.search(r'Target="[^"]*slideLayout(\d+)\.xml"', z.read(rels_name).decode("utf-8"))
            lay = int(m.group(1)) if m else None
        runs = re.findall(r'<a:t>([^<]*)</a:t>', xml)
        out.append({"n": n, "layout": lay, "runs": runs})
    by_n = {s["n"]: s for s in out}
    suggest = None
    if len(slides) >= 5:
        suggest = {"cover": 1, "intro": 2, "newfamily": 3, "grid": 4, "end": 5,
                   "layout_cover": by_n[1]["layout"], "layout_title": by_n[2]["layout"]}
    print(json.dumps({"master": path, "slide_count": len(slides), "layout_count": len(layouts),
                      "slides": out, "templates_suggest": suggest}, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print(__doc__); sys.exit(2)
    main(sys.argv[1])
