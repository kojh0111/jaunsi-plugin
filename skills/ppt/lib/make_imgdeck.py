#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_imgdeck.py — 외부 PPT를 "이미지판" PPTX로 변환

캔버스가 마스터와 다르고 텍스트가 있는 외부 PPT는 scale(그룹 축소) 병합 시
글자 크기가 안 줄어 레이아웃이 깨진다(그룹 축소는 도형만 줄인다). 이럴 때:
  1) PowerPoint(AppleScript)로 원본을 PDF 저장 (임베드 폰트 그대로 렌더링)
  2) pdftoppm -png -r 160 으로 페이지별 PNG 추출
  3) 이 스크립트로 원본 pptx 골격을 재사용해 슬라이드 본문만
     풀블리드 <p:pic> 하나로 교체한 이미지판 pptx 생성
  4) 이미지판을 config.json 의 source 로 병합 (scale 은 원본 캔버스 크기 그대로)

사용법:
    python3 make_imgdeck.py <원본.pptx> <출력.pptx> <png폴더>

png폴더의 *.png 를 파일명 정렬 순으로 발표 순서에 매핑한다
(pdftoppm 의 pg-01.png .. pg-NN.png 가 그대로 발표 순서).
장수가 슬라이드 수와 다르면 에러로 중단한다.
캔버스 크기(sldSz)는 원본 presentation.xml 에서 자동으로 읽는다.
원본이 PDF 뿐이면 make_skeleton.py 로 빈 골격 pptx(캔버스는 마스터와 동일)를 만들어 <원본.pptx> 자리에 쓴다.
"""
import zipfile, re, sys, glob, os, posixpath

def main():
    if len(sys.argv) != 4:
        sys.exit("사용법: python3 make_imgdeck.py <원본.pptx> <출력.pptx> <png폴더>")
    src, dst, pngdir = sys.argv[1], sys.argv[2], sys.argv[3]

    pngs = sorted(glob.glob(os.path.join(pngdir, "*.png")))
    zin = zipfile.ZipFile(src)

    # 캔버스 크기
    pres = zin.read("ppt/presentation.xml").decode("utf-8")
    m = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres)
    CX, CY = int(m.group(1)), int(m.group(2))

    # 발표 순서: sldIdLst 의 r:id 순서 → presentation.rels → slide 파일
    prels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
    rid2tgt = dict(re.findall(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', prels))
    order_rids = re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"',
                            re.search(r'<p:sldIdLst>.*?</p:sldIdLst>', pres, re.S).group())
    ordered_slides = [posixpath.normpath(posixpath.join("ppt", rid2tgt[r])) for r in order_rids]

    if len(pngs) != len(ordered_slides):
        sys.exit(f"PNG {len(pngs)}장 != 슬라이드 {len(ordered_slides)}장 — 렌더링 누락 확인 필요")

    slide2png = {sl: i + 1 for i, sl in enumerate(ordered_slides)}  # 발표 순서 n → pgrender{n}.png

    SLIDE_TMPL = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '<p:pic><p:nvPicPr><p:cNvPr id="2" name="render"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId100"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>' % (CX, CY)
    )

    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in zin.namelist():
            mr = re.match(r"ppt/slides/_rels/(slide\d+\.xml)\.rels$", name)
            if name in slide2png:
                zout.writestr(name, SLIDE_TMPL)
            elif mr and ("ppt/slides/" + mr.group(1)) in slide2png:
                n = slide2png["ppt/slides/" + mr.group(1)]
                rels = zin.read(name).decode("utf-8")
                layout = re.search(r'<Relationship\b[^>]*slideLayout[^>]*/>', rels).group(0)
                zout.writestr(name, (
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    + layout +
                    f'<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/pgrender{n}.png"/>'
                    '</Relationships>'))
            elif name.startswith("ppt/embeddings/") or name.startswith("ppt/notesSlides/"):
                continue  # OLE 바이너리·노트 제거 (복구 원인 차단)
            elif name == "[Content_Types].xml":
                ct = zin.read(name).decode("utf-8")
                if 'Extension="png"' not in ct:
                    ct = ct.replace("<Default", '<Default Extension="png" ContentType="image/png"/><Default', 1)
                zout.writestr(name, ct)
            else:
                zout.writestr(name, zin.read(name))
        for i, p in enumerate(pngs, 1):
            zout.writestr(f"ppt/media/pgrender{i}.png", open(p, "rb").read())

    print(f"done: {dst} ({len(pngs)}장, 캔버스 {CX}x{CY})")

if __name__ == "__main__":
    main()
