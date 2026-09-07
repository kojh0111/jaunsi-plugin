#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate.py — 복구방지 검증 (모두 통과해야 산출물 인정)

PowerPoint "복구" 창이 뜨는 주요 원인들을 빌드 후 정적 검증으로 차단한다.
한 번에 열리게 하려면 이 7개 게이트가 전부 통과해야 한다.

검사 항목:
  1. 모든 part XML well-formed
  2. sldLayoutId @id 가 전(全) 마스터에서 전역 고유   ← 복구 주원인
  3. 본문 r:id / r:embed ⊆ 해당 part .rels (dangling 참조 0)
  4. .rels Target(Internal) 실제 존재(dangling 0) + 중복 rId 0
  5. [Content_Types] 가 모든 part 커버(Override/Default), svg Default 포함
  6. 슬라이드 내 cNvPr @id 중복 0
  7. sldIdLst ↔ presentation.rels 1:1 (누락/고아 0)

사용법:
    python3 validate.py <deck.pptx>
출력: JSON  { "ok": true/false, "checks":[...], "errors":[...] }
종료코드: 0 통과, 1 실패
"""
import sys, re, json, zipfile, posixpath
import xml.dom.minidom as minidom

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"


def relpath(part): return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part)+".rels")
def resolve(part, tgt): return posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt))

def parse_rels(xml):
    out=[]
    for m in re.findall(r'<Relationship\b[^>]*/>', xml):
        rid=re.search(r'Id="([^"]+)"',m).group(1)
        typ=re.search(r'Type="([^"]+)"',m).group(1)
        tgt=re.search(r'Target="([^"]+)"',m).group(1)
        mode=re.search(r'TargetMode="([^"]+)"',m)
        out.append((rid,typ,tgt,mode.group(1) if mode else "Internal"))
    return out


def validate(path):
    z = zipfile.ZipFile(path)
    data = {n: z.read(n) for n in z.namelist()}
    names = set(data)
    errors = []
    checks = []

    def add(name, ok, detail=""):
        checks.append({"check": name, "ok": ok, "detail": detail})
        if not ok:
            errors.append("%s: %s" % (name, detail))

    # ---- 1. well-formed XML ----
    bad = []
    for n, b in data.items():
        if n.endswith(".xml") or n.endswith(".rels"):
            try:
                minidom.parseString(b)
            except Exception as e:
                bad.append("%s (%s)" % (n, e))
    add("1.xml_well_formed", not bad, ("malformed: " + "; ".join(bad[:5])) if bad else "all parts parse")

    # ---- 2. sldLayoutId @id 전역 고유 ----
    seen = {}; dup = []
    for n in names:
        if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n):
            for mid in re.findall(r'<p:sldLayoutId id="(\d+)"', data[n].decode("utf-8")):
                if mid in seen and seen[mid] != n:
                    dup.append("id=%s in %s & %s" % (mid, seen[mid], n))
                seen[mid] = n
    add("2.sldLayoutId_unique", not dup, ("dup: " + "; ".join(dup[:5])) if dup else "%d ids unique" % len(seen))

    # ---- 3. 본문 r:id/r:embed ⊆ part rels ----
    dangling_use = []
    for n in names:
        if not (n.endswith(".xml")) or n.endswith(".rels"): continue
        if not (n.startswith("ppt/slides/") or n.startswith("ppt/slideLayouts/")
                or n.startswith("ppt/slideMasters/") or n == "ppt/presentation.xml"):
            continue
        body = data[n].decode("utf-8")
        used = set(re.findall(r'r:(?:id|embed|link)="([^"]+)"', body))
        rp = relpath(n)
        have = set(rid for rid,_,_,_ in parse_rels(data[rp].decode("utf-8"))) if rp in data else set()
        miss = used - have
        if miss:
            dangling_use.append("%s -> %s" % (n, sorted(miss)))
    add("3.rid_refs_resolve", not dangling_use,
        ("missing: " + "; ".join(dangling_use[:5])) if dangling_use else "all r:id/r:embed resolve")

    # ---- 4. rels Target 존재 + 중복 rId 0 ----
    dangling_tgt = []; dup_rid = []
    for n in names:
        if not n.endswith(".rels"): continue
        # a rels file <dir>/_rels/<name>.rels owns part <dir>/<name>;
        # its Internal Targets resolve relative to <dir> (package root for _rels/.rels).
        base_dir = posixpath.dirname(posixpath.dirname(n))
        rels = parse_rels(data[n].decode("utf-8"))
        ids = [r[0] for r in rels]
        for rid in set(ids):
            if ids.count(rid) > 1:
                dup_rid.append("%s in %s" % (rid, n))
        for rid, typ, tgt, mode in rels:
            if mode != "Internal": continue
            res = posixpath.normpath(posixpath.join(base_dir, tgt))
            if res not in names:
                dangling_tgt.append("%s -> %s" % (n, tgt))
    add("4.rels_targets_exist", not dangling_tgt and not dup_rid,
        ("; ".join((["dangling:"+",".join(dangling_tgt[:4])] if dangling_tgt else []) +
                   (["dup_rid:"+",".join(dup_rid[:4])] if dup_rid else [])) or "ok"))

    # ---- 5. Content_Types coverage ----
    ct = data["[Content_Types].xml"].decode("utf-8")
    defaults = dict(re.findall(r'<Default Extension="([^"]+)" ContentType="([^"]+)"', ct))
    overrides = set(re.findall(r'<Override PartName="([^"]+)"', ct))
    uncovered = []
    for n in names:
        if n == "[Content_Types].xml" or "/_rels/" in n or n.endswith(".rels"):
            continue
        ext = n.rsplit(".", 1)[-1].lower() if "." in n else ""
        if ("/" + n) in overrides: continue
        if ext in defaults: continue
        uncovered.append(n)
    svg_present = any(n.lower().endswith(".svg") for n in names)
    svg_ok = (not svg_present) or ("svg" in defaults)
    add("5.content_types_cover", not uncovered and svg_ok,
        ("uncovered: " + ", ".join(uncovered[:5])) if uncovered else
        ("svg Default 누락" if not svg_ok else "all parts covered"))

    # ---- 6. cNvPr @id 슬라이드 내 중복 0 ----
    dup_cnv = []
    for n in names:
        if not re.match(r'ppt/slides/slide\d+\.xml$', n): continue
        ids = re.findall(r'<p:cNvPr id="(\d+)"', data[n].decode("utf-8"))
        seen2 = set()
        for i in ids:
            if i in seen2: dup_cnv.append("%s id=%s" % (n, i))
            seen2.add(i)
    add("6.cNvPr_unique_per_slide", not dup_cnv,
        ("dup: " + "; ".join(dup_cnv[:5])) if dup_cnv else "no duplicate shape ids")

    # ---- 7. sldIdLst ↔ presentation.rels 1:1 ----
    pres = data["ppt/presentation.xml"].decode("utf-8")
    pr = data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    sld_rids = re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', pres)
    rel_slide = {rid: tgt for rid, typ, tgt, mode in parse_rels(pr) if typ.endswith("/slide")}
    miss_rel = [r for r in sld_rids if r not in rel_slide]
    orphan = [r for r in rel_slide if r not in set(sld_rids)]
    # also every referenced slide target must exist
    miss_file = [rel_slide[r] for r in sld_rids if r in rel_slide and resolve("ppt/presentation.xml", rel_slide[r]) not in names]
    add("7.sldIdLst_matches_rels", not miss_rel and not orphan and not miss_file,
        "; ".join((["no_rel:"+",".join(miss_rel[:4])] if miss_rel else []) +
                  (["orphan:"+",".join(orphan[:4])] if orphan else []) +
                  (["missing_file:"+",".join(miss_file[:4])] if miss_file else [])) or
        "%d slides 1:1" % len(sld_rids))

    # ---- 8. 레이아웃 소유권: 각 slideLayout 은 정확히 한 마스터의 rels 에만 등장하고,
    #         그 레이아웃의 rels 가 가리키는 slideMaster 가 바로 그 마스터여야 함.
    #         (외부 PPT 병합 시 파트 번호가 겹쳐 한 레이아웃을 두 마스터가 공유하면 다른 게이트는 통과해도 복구창이 뜬다) ----
    lay_owner = {}
    bad_own = []
    for n in names:
        if not re.match(r'ppt/slideMasters/slideMaster\d+\.xml$', n): continue
        rp = relpath(n)
        if rp not in names: continue
        for rid, typ, tgt, mode in parse_rels(data[rp].decode("utf-8")):
            if typ.endswith("/slideLayout"):
                lay = resolve(n, tgt)
                if lay in lay_owner and lay_owner[lay] != n:
                    bad_own.append("%s shared by %s & %s" % (lay, lay_owner[lay], n))
                lay_owner[lay] = n
    for lay, m in lay_owner.items():
        rp = relpath(lay)
        if rp not in names: continue
        back = [resolve(lay, tgt) for rid, typ, tgt, mode in parse_rels(data[rp].decode("utf-8")) if typ.endswith("/slideMaster")]
        if back and back[0] != m:
            bad_own.append("%s -> %s but owned by %s" % (lay, back[0], m))
    add("8.layout_ownership", not bad_own,
        ("; ".join(bad_own[:4])) if bad_own else "%d layouts, one master each" % len(lay_owner))

    ok = all(c["ok"] for c in checks)
    return {"ok": ok, "deck": path, "checks": checks, "errors": errors}


def main():
    if len(sys.argv) != 2:
        print("usage: python3 validate.py <deck.pptx>", file=sys.stderr); sys.exit(2)
    res = validate(sys.argv[1])
    print(json.dumps(res, ensure_ascii=False, indent=2))
    sys.exit(0 if res["ok"] else 1)


if __name__ == "__main__":
    main()
