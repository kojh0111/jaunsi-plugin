#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
slim_pptx.py — PPTX 용량 분석·슬림화 (stdlib + magick)

analyze: 미디어별 크기/해상도/투명도/참조처를 조사해 절감 후보를 JSON으로 보고.
apply:   안전 규칙에 따라 변환·정리한 사본(<원본>.슬림.pptx)을 만든다. 원본은 건드리지 않는다.

안전 규칙(2026-08-19 자운시 마스터 슬림화에서 검증):
  - 불투명(opaque) PNG ≥ min-size → JPEG q<quality>, 장변 max-edge 상한.
    실제 절감이 30% 미만이면 원본 유지. 투명도 있는 PNG는 절대 변환하지 않음.
  - JPEG ≥ min-size 이고 장변 > max-edge → 다운스케일 재압축.
  - 예술효과 원본 레이어(a14:imgProps/imgLayer + hdphoto .wdp)는 화면 표시에 쓰이지
    않으므로 ext 블록·rels·파일을 세트로 제거.
  - 어떤 .rels 에서도 참조되지 않는 고아 미디어 삭제.
  - --keep 으로 지정한 파일은 어떤 경우에도 손대지 않음.
  - 재압축 zip 은 디렉터리 엔트리 없이, [Content_Types].xml 을 먼저 기록.

사용법:
    python3 slim_pptx.py analyze <파일.pptx>
    python3 slim_pptx.py apply   <파일.pptx> [--keep image20.png,image4.png]
                                 [--min-size 500000] [--max-edge 3840] [--quality 92]
"""
import zipfile, re, os, sys, json, argparse, subprocess, tempfile, glob, shutil

IMG_RE = re.compile(r"ppt/media/.+\.(png|jpg|jpeg)$", re.I)


def identify(path):
    try:
        out = subprocess.run(["magick", "identify", "-format", "%w %h %[opaque]", path + "[0]"],
                             capture_output=True, text=True, check=True).stdout.split()
        return int(out[0]), int(out[1]), out[2] == "True"
    except Exception:
        return None, None, None


def collect_refs(z):
    """media 파일명 → 참조하는 파트 목록"""
    refs = {}
    for n in z.namelist():
        if n.endswith(".rels"):
            owner = n.replace("_rels/", "").replace(".rels", "")
            for m in re.findall(r'Target="(?:\.\./)?media/([^"]+)"', z.read(n).decode("utf-8", "ignore")):
                refs.setdefault(m, []).append(owner)
    return refs


def used_layouts(z):
    used = set()
    for n in z.namelist():
        if re.match(r"ppt/slides/_rels/.*\.rels$", n):
            used |= set(re.findall(r"slideLayouts/(slideLayout\d+)\.xml", z.read(n).decode()))
    return used


def analyze(path, min_size):
    z = zipfile.ZipFile(path)
    refs = collect_refs(z)
    used = used_layouts(z)
    total = os.path.getsize(path)
    rows, est = [], 0
    with tempfile.TemporaryDirectory() as td:
        for i in z.infolist():
            name = i.filename
            base = name.split("/")[-1]
            if not name.startswith("ppt/media/"):
                continue
            row = {"file": base, "bytes": i.compress_size, "refs": refs.get(base, [])}
            lay = [r.split("/")[-1].replace(".xml", "") for r in row["refs"] if "slideLayouts" in r]
            row["unused_layout_only"] = bool(row["refs"]) and all("slideLayouts" in r for r in row["refs"]) \
                and not any(l in used for l in lay)
            row["orphan"] = not row["refs"]
            if base.lower().endswith(".wdp"):
                row["action"] = "wdp 제거(예술효과 원본)"
                est += i.compress_size
            elif IMG_RE.match(name) and i.compress_size >= min_size:
                p = os.path.join(td, base)
                open(p, "wb").write(z.read(name))
                w, h, opaque = identify(p)
                row.update({"dims": f"{w}x{h}", "opaque": opaque})
                if row["orphan"]:
                    row["action"] = "고아 미디어 삭제"
                    est += i.compress_size
                elif base.lower().endswith(".png") and opaque:
                    row["action"] = "PNG→JPEG 변환 후보"
                    est += int(i.compress_size * 0.85)
                elif base.lower().endswith((".jpg", ".jpeg")) and w and max(w, h) > 3840:
                    row["action"] = "다운스케일 재압축 후보"
                    est += int(i.compress_size * 0.5)
                else:
                    row["action"] = "유지(투명도 있음)" if opaque is False else "유지"
            elif row["orphan"]:
                row["action"] = "고아 미디어 삭제"
                est += i.compress_size
            else:
                continue  # 소형·비대상 생략
            rows.append(row)
    rows.sort(key=lambda r: -r["bytes"])
    return {"file": path, "total_mb": round(total / 1e6, 1),
            "used_layouts": sorted(used), "estimated_saving_mb": round(est / 1e6, 1),
            "media": rows}


def apply(path, keep, min_size, max_edge, quality):
    keep = set(keep)
    out_path = re.sub(r"\.pptx$", ".슬림.pptx", path)
    td = tempfile.mkdtemp(prefix="pptxslim_")
    with zipfile.ZipFile(path) as z:
        z.extractall(td)
    media_dir = os.path.join(td, "ppt", "media")
    z = zipfile.ZipFile(path)
    refs = collect_refs(z)
    log, renamed = [], {}

    for f in sorted(os.listdir(media_dir)):
        p = os.path.join(media_dir, f)
        sz = os.path.getsize(p)
        if f in keep or f.startswith("."):
            continue
        if f.lower().endswith(".wdp") or (f not in refs and sz > 0):
            os.remove(p)
            log.append(f"제거: {f} ({sz//1000}KB)")
            continue
        if sz < min_size:
            continue
        w, h, opaque = identify(p)
        if f.lower().endswith(".png") and opaque:
            new = f.rsplit(".", 1)[0] + ".jpeg"
            dst = os.path.join(media_dir, new)
            subprocess.run(["magick", p, "-resize", f"{max_edge}x{max_edge}>",
                            "-quality", str(quality), dst], check=True)
            if os.path.getsize(dst) < sz * 0.7:
                os.remove(p)
                renamed[f] = new
                log.append(f"변환: {f} → {new} ({sz//1000}→{os.path.getsize(dst)//1000}KB)")
            else:
                os.remove(dst)
        elif f.lower().endswith((".jpg", ".jpeg")) and w and max(w, h) > max_edge:
            tmp = p + ".tmp"
            subprocess.run(["magick", p, "-resize", f"{max_edge}x{max_edge}>",
                            "-quality", str(max(quality - 2, 85)), tmp], check=True)
            if os.path.getsize(tmp) < sz * 0.7:
                os.replace(tmp, p)
                log.append(f"재압축: {f} ({sz//1000}→{os.path.getsize(p)//1000}KB)")
            else:
                os.remove(tmp)

    # rels 갱신: 파일명 치환 + wdp 관계 제거
    for rp in glob.glob(os.path.join(td, "**", "_rels", "*.rels"), recursive=True):
        xml = open(rp, encoding="utf-8").read()
        orig = xml
        for old, new in renamed.items():
            xml = xml.replace("media/" + old, "media/" + new)
        xml = re.sub(r'<Relationship[^>]*Target="[^"]*\.wdp"[^>]*/>', "", xml)
        if xml != orig:
            open(rp, "w", encoding="utf-8").write(xml)

    # imgLayer(예술효과 원본) ext 블록 제거 — 레이아웃·슬라이드·마스터 전체
    for xp in glob.glob(os.path.join(td, "ppt", "**", "*.xml"), recursive=True):
        xml = open(xp, encoding="utf-8").read()
        new = re.sub(r'<a:ext uri="[^"]*">(?:(?!</a:ext>).)*?imgLayer.*?</a:ext>', "", xml, flags=re.S)
        if new != xml:
            open(xp, "w", encoding="utf-8").write(new)

    # [Content_Types]
    ctp = os.path.join(td, "[Content_Types].xml")
    ct = open(ctp, encoding="utf-8").read()
    ct = re.sub(r'<Default Extension="wdp"[^>]*/>', "", ct)
    if renamed and 'Extension="jpeg"' not in ct:
        ct = ct.replace("<Default", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default', 1)
    open(ctp, "w", encoding="utf-8").write(ct)

    # 재압축: 디렉터리 엔트리 없이, [Content_Types].xml 먼저
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zo:
        zo.write(ctp, "[Content_Types].xml")
        for root, _, files in os.walk(td):
            for f in files:
                full = os.path.join(root, f)
                arc = os.path.relpath(full, td)
                if arc == "[Content_Types].xml" or f == ".DS_Store":
                    continue
                zo.write(full, arc)
    shutil.rmtree(td)
    return {"output": out_path,
            "before_mb": round(os.path.getsize(path) / 1e6, 1),
            "after_mb": round(os.path.getsize(out_path) / 1e6, 1),
            "changes": log}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("mode", choices=["analyze", "apply"])
    ap.add_argument("pptx")
    ap.add_argument("--keep", default="", help="쉼표구분: 절대 손대지 않을 미디어 파일명")
    ap.add_argument("--min-size", type=int, default=500000)
    ap.add_argument("--max-edge", type=int, default=3840)
    ap.add_argument("--quality", type=int, default=92)
    a = ap.parse_args()
    keep = [s.strip() for s in a.keep.split(",") if s.strip()]
    if a.mode == "analyze":
        out = analyze(a.pptx, a.min_size)
    else:
        out = apply(a.pptx, keep, a.min_size, a.max_edge, a.quality)
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
