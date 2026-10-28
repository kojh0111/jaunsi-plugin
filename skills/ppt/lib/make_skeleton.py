#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""빈 N장 pptx 골격 생성 (stdlib). 테마는 지정한 pptx 의 ppt/theme/theme1.xml 을 복사.
사용법: python3 make_skeleton.py <출력.pptx> <테마원본.pptx> <장수> <cx> <cy> [masterId]
"""
import sys, zipfile

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = f'xmlns:a="{A}" xmlns:r="{R}" xmlns:p="{P}"'
RELT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
XMLH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

EMPTY_TREE = (
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
    '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree>'
)

def rels(items):
    body = "".join(
        f'<Relationship Id="{rid}" Type="{RELT}{typ}" Target="{tgt}"/>' for rid, typ, tgt in items
    )
    return (XMLH + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + body + "</Relationships>")

def main():
    out, theme_src, n, cx, cy = sys.argv[1], sys.argv[2], int(sys.argv[3]), int(sys.argv[4]), int(sys.argv[5])
    master_id = int(sys.argv[6]) if len(sys.argv) > 6 else 2147483648
    theme_xml = zipfile.ZipFile(theme_src).read("ppt/theme/theme1.xml")

    sld_ids = "".join(f'<p:sldId id="{256 + i}" r:id="rId{2 + i}"/>' for i in range(n))
    presentation = (
        XMLH + f'<p:presentation {NS} saveSubsetFonts="1">'
        f'<p:sldMasterIdLst><p:sldMasterId id="{master_id}" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sld_ids}</p:sldIdLst>'
        f'<p:sldSz cx="{cx}" cy="{cy}"/><p:notesSz cx="6858000" cy="9144000"/>'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr>'
        '<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
        '<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
        '<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>'
        '</p:defaultTextStyle></p:presentation>'
    )
    pres_rels = rels(
        [("rId1", "slideMaster", "slideMasters/slideMaster1.xml")]
        + [(f"rId{2 + i}", "slide", f"slides/slide{i + 1}.xml") for i in range(n)]
        + [(f"rId{2 + n}", "theme", "theme/theme1.xml")]
    )

    def rpr(sz, mj=False):
        f = "mj" if mj else "mn"
        return (f'<a:defRPr sz="{sz}"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>'
                f'<a:latin typeface="+{f}-lt"/><a:ea typeface="+{f}-ea"/><a:cs typeface="+{f}-cs"/></a:defRPr>')

    master = (
        XMLH + f'<p:sldMaster {NS}><p:cSld>'
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
        + EMPTY_TREE + '</p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        f'<p:sldLayoutIdLst><p:sldLayoutId id="{master_id + 1}" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:txStyles>'
        f'<p:titleStyle><a:lvl1pPr algn="ctr">{rpr(4400, True)}</a:lvl1pPr></p:titleStyle>'
        f'<p:bodyStyle><a:lvl1pPr>{rpr(2800)}</a:lvl1pPr></p:bodyStyle>'
        f'<p:otherStyle><a:defPPr><a:defRPr lang="ko-KR"/></a:defPPr><a:lvl1pPr>{rpr(1800)}</a:lvl1pPr></p:otherStyle>'
        '</p:txStyles></p:sldMaster>'
    )
    master_rels = rels([("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
                        ("rId2", "theme", "../theme/theme1.xml")])

    layout = (
        XMLH + f'<p:sldLayout {NS} type="blank" preserve="1"><p:cSld name="Blank">'
        + EMPTY_TREE + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'
    )
    layout_rels = rels([("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")])

    slide = (XMLH + f'<p:sld {NS}><p:cSld>' + EMPTY_TREE
             + '</p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>')
    slide_rels = rels([("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")])

    PML = "application/vnd.openxmlformats-officedocument.presentationml."
    overrides = [
        ("/ppt/presentation.xml", PML + "presentation.main+xml"),
        ("/ppt/slideMasters/slideMaster1.xml", PML + "slideMaster+xml"),
        ("/ppt/slideLayouts/slideLayout1.xml", PML + "slideLayout+xml"),
        ("/ppt/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"),
    ] + [(f"/ppt/slides/slide{i + 1}.xml", PML + "slide+xml") for i in range(n)]
    content_types = (
        XMLH + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        + "".join(f'<Override PartName="{pn}" ContentType="{ct}"/>' for pn, ct in overrides)
        + "</Types>"
    )
    root_rels = rels([("rId1", "officeDocument", "ppt/presentation.xml")])

    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("ppt/presentation.xml", presentation)
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels)
        z.writestr("ppt/slideMasters/slideMaster1.xml", master)
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", master_rels)
        z.writestr("ppt/slideLayouts/slideLayout1.xml", layout)
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", layout_rels)
        z.writestr("ppt/theme/theme1.xml", theme_xml)
        for i in range(n):
            z.writestr(f"ppt/slides/slide{i + 1}.xml", slide)
            z.writestr(f"ppt/slides/_rels/slide{i + 1}.xml.rels", slide_rels)
    print(f"skeleton: {out} ({n}장, {cx}x{cy}, masterId={master_id})")

if __name__ == "__main__":
    main()
