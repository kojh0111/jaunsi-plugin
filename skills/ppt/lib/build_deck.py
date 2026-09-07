#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_deck.py — 자운시 주간 PPT 빌더 (순수 python stdlib: zipfile/xml/re)

config.json 하나로 마스터 템플릿 + 사진 그리드 + 외부 PPT 를 조립해
PowerPoint "복구" 창 없이 열리는 PPTX 를 생성한다.

사용법:
    python3 build_deck.py <config.json>

config.json 스키마:
{
  "master":  "/abs/master.pptx",          # 자운시 마스터 템플릿
  "output":  "/abs/<날짜>_자운시 PPT_완성.pptx",
  "photos_dir": "/abs/<날짜>/images",      # content.photos 파일이 들어있는 폴더
  "work_dir":   "/abs/<날짜>/.work_photos", # (선택) 리사이즈 임시폴더. 기본 output 옆
  "resize":  true,                          # magick 으로 사전 리사이즈 여부

  "templates": {                            # 마스터 내 템플릿 슬라이드 번호 (inspect_master.py 로 확인)
    "cover":1, "intro":2, "newfamily":3, "grid":4, "end":5,
    "layout_cover":30, "layout_title":33      # 생략 시 cover/intro 슬라이드의 rels 에서 자동 검출
  },

  "replacements": {                         # <a:t>OLD</a:t> -> <a:t>NEW</a:t>
    "cover":     [["2024.07.07.", "2026.06.07."]],
    "intro":     [["2","1",1], ["24","26"], ["7","6"]],
    "newfamily": [],
    "end":       []
  },

  "content": [                              # 제목 슬라이드(제목당 1장, 사진 1/2/4 그리드)
    {"title":"예배팀 클라이밍 모꼬지", "photos":["a.jpeg"]},
    {"title":"6월 운영위",            "photos":["b.jpeg","c.jpeg"]}
  ],

  "sources": [                              # 끼워넣을 외부 PPT
    {"path":"/abs/new_family.pptx", "scale":null},
    {"path":"/abs/care.pptx",       "scale":null},
    {"path":"/abs/ad.pptx",         "scale":[12192000,6858000]}
  ],

  "assembly": [                             # 최종 슬라이드 순서
    {"type":"cover"},
    {"type":"intro"},
    {"type":"newfamily"},
    {"type":"content"},
    {"type":"divider", "label":"새가족팀"},
    {"type":"source",  "index":0},
    {"type":"divider", "label":"양육팀"},
    {"type":"source",  "index":1, "skip_first":true},
    {"type":"source",  "index":2},
    {"type":"end"}
  ]
}
"""
import zipfile, re, os, posixpath, json, sys, subprocess

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT = {"slide":"application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
      "slideMaster":"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
      "slideLayout":"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
      "theme":"application/vnd.openxmlformats-officedocument.theme+xml"}


def esc(s): return s.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
def relpath(part): return posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part)+".rels")
def resolve(part, tgt):
    # 절대경로 Target("/ppt/...")도 허용 — 일부 생성기(Canva 등)가 이렇게 쓴다; 상대경로로만 풀면 파일을 못 찾아 복구창
    if tgt.startswith("/"): return posixpath.normpath(tgt.lstrip("/"))
    return posixpath.normpath(posixpath.join(posixpath.dirname(part), tgt))

def parse_rels(xml):
    out=[]
    for m in re.findall(r'<Relationship\b[^>]*/>', xml):
        rid=re.search(r'Id="([^"]+)"',m).group(1)
        typ=re.search(r'Type="([^"]+)"',m).group(1)
        tgt=re.search(r'Target="([^"]+)"',m).group(1)
        mode=re.search(r'TargetMode="([^"]+)"',m)
        out.append((rid,typ,tgt,mode.group(1) if mode else "Internal"))
    return out


def apply_repl(xml, pairs):
    for p in (pairs or []):
        old, new = p[0], p[1]
        cnt = p[2] if len(p) > 2 else -1
        a, b = "<a:t>%s</a:t>"%old, "<a:t>%s</a:t>"%new
        xml = xml.replace(a, b, cnt) if cnt >= 0 else xml.replace(a, b)
    return xml


def magick_resize(src, dst):
    subprocess.run(["magick", src, "-auto-orient", "-resize", "2000x2000>",
                    "-strip", dst], check=True)

def magick_dims(path):
    out = subprocess.run(["magick", "identify", "-format", "%w %h", path],
                         check=True, capture_output=True, text=True).stdout.strip()
    w, h = out.split()[:2]
    return int(w), int(h)


def main(cfg_path):
    cfg = json.load(open(cfg_path, encoding="utf-8"))
    BASE = cfg["master"]
    OUT  = cfg["output"]
    PH   = cfg["photos_dir"]
    WORK = cfg.get("work_dir") or os.path.join(os.path.dirname(OUT), ".work_photos")
    DO_RESIZE = cfg.get("resize", True)
    os.makedirs(WORK, exist_ok=True)

    T = cfg.get("templates", {})
    n_cover = T.get("cover", 1); n_intro = T.get("intro", 2)
    n_newfam = T.get("newfamily", 3); n_grid = T.get("grid", 4); n_end = T.get("end", 5)
    LAY_COVER = T.get("layout_cover"); LAY_TITLE = T.get("layout_title")

    R = cfg.get("replacements", {})
    CONTENT = cfg.get("content", [])
    SOURCES = [(s["path"], s.get("scale")) for s in cfg.get("sources", [])]
    ASSEMBLY = cfg.get("assembly", [])

    # ---- resolve photo path: prefetch & resize ----
    def photo_path(fn):
        src = fn if os.path.isabs(fn) else os.path.join(PH, fn)
        if not os.path.exists(src):
            raise FileNotFoundError("photo not found: %s" % src)
        if not DO_RESIZE:
            return src, magick_dims(src)
        dst = os.path.join(WORK, "rz_" + os.path.basename(src))
        if os.path.splitext(dst)[1].lower() not in (".jpeg", ".jpg", ".png"):
            dst = dst + ".jpeg"
        magick_resize(src, dst)
        return dst, magick_dims(dst)

    # ================= load base =================
    zin = zipfile.ZipFile(BASE)
    data = {n: zin.read(n) for n in zin.namelist()}

    # canvas size from master presentation.xml
    pres0 = data["ppt/presentation.xml"].decode("utf-8")
    msz = re.search(r'<p:sldSz cx="(\d+)" cy="(\d+)"', pres0)
    if cfg.get("canvas"):
        W, H = cfg["canvas"]
    elif msz:
        W, H = int(msz.group(1)), int(msz.group(2))
    else:
        W, H = 13823950, 7775575

    # photo grid geometry (proportional to canvas; matches validated 13823950x7775575 layout)
    TOP = int(H * 0.19549); BOT = int(H * 0.97228)
    LF  = int(W * 0.03111); RT = int(W * 0.96891)
    RW, RH = RT-LF, BOT-TOP

    def fit(cx, cy, cw, ch, pw, ph):
        sc = min(cw/pw, ch/ph); w=int(pw*sc); h=int(ph*sc)
        return int(cx+(cw-w)/2), int(cy+(ch-h)/2), w, h
    def cells(n):
        if n==1: return [(LF,TOP,RW,RH)]
        if n==2:
            g=350000; cw=(RW-g)//2
            return [(LF,TOP,cw,RH),(LF+cw+g,TOP,cw,RH)]
        if n==4:
            g=300000; cw=(RW-g)//2; ch=(RH-g)//2
            return [(LF,TOP,cw,ch),(LF+cw+g,TOP,cw,ch),(LF,TOP+ch+g,cw,ch),(LF+cw+g,TOP+ch+g,cw,ch)]
        if n==3:  # 2 위 + 1 아래(가운데)
            g=300000; cw=(RW-g)//2; ch=(RH-g)//2
            return [(LF,TOP,cw,ch),(LF+cw+g,TOP,cw,ch),(LF+(RW-cw)//2,TOP+ch+g,cw,ch)]
        raise ValueError("unsupported photo count per title: %d" % n)
    def pic_xml(pid, rid, x, y, w, h):
        return ('<p:pic><p:nvPicPr><p:cNvPr id="%d" name="Picture %d"/>'
                '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
                '<p:blipFill><a:blip r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
                '<p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
                % (pid, pid, rid, x, y, w, h))

    # ---- layout numbers: 마스터마다 다르므로 미지정 시 템플릿 슬라이드의 rels 에서 읽는다 ----
    def slide_layout(n):
        rp = "ppt/slides/_rels/slide%d.xml.rels" % n
        m = re.search(r'Target="[^"]*slideLayout(\d+)\.xml"', data[rp].decode("utf-8")) if rp in data else None
        if not m: raise RuntimeError("slide%d 의 레이아웃을 찾지 못함 — templates.layout_* 를 명시하세요" % n)
        return int(m.group(1))
    if LAY_COVER is None: LAY_COVER = slide_layout(n_cover)
    if LAY_TITLE is None: LAY_TITLE = slide_layout(n_intro)

    # ---- extract template slides ----
    s_grid = data["ppt/slides/slide%d.xml"%n_grid].decode("utf-8")
    s_grid = re.sub(r"<p:pic>.*?</p:pic>", "", s_grid, flags=re.S)
    s_grid = re.sub(r"<p:transition.*?</p:transition>", "", s_grid, flags=re.S)

    cover = apply_repl(data["ppt/slides/slide%d.xml"%n_cover].decode("utf-8"), R.get("cover"))
    intro = apply_repl(data["ppt/slides/slide%d.xml"%n_intro].decode("utf-8"), R.get("intro"))
    newfam_xml  = apply_repl(data["ppt/slides/slide%d.xml"%n_newfam].decode("utf-8"), R.get("newfamily"))
    newfam_rels = data["ppt/slides/_rels/slide%d.xml.rels"%n_newfam].decode("utf-8")
    end_xml = apply_repl(data["ppt/slides/slide%d.xml"%n_end].decode("utf-8"), R.get("end"))

    def layrel(lay):
        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{REL}slideLayout" Target="../slideLayouts/slideLayout{lay}.xml"/>')

    # drop all original slide parts (we rewrite slideN.xml from scratch)
    for n in [k for k in list(data) if k.startswith("ppt/slides/")]:
        data.pop(n, None)

    # ---- build named template slides ----
    tmpl = {}
    tmpl["cover"]     = (cover.encode(), (layrel(LAY_COVER)+"</Relationships>").encode())
    tmpl["intro"]     = (intro.encode(), (layrel(LAY_TITLE)+"</Relationships>").encode())
    tmpl["newfamily"] = (newfam_xml.encode(), newfam_rels.encode())
    tmpl["end"]       = (end_xml.encode(), (layrel(LAY_COVER)+"</Relationships>").encode())

    def divider(label, sz=None):
        body = end_xml.replace("<a:t>끝</a:t>", "<a:t>%s</a:t>"%esc(label))
        # 긴 라벨 줄바꿈 방지: 한 줄 고정 + (선택) 글자크기 축소
        body = body.replace("<a:bodyPr/>", '<a:bodyPr wrap="none" anchor="ctr"/>', 1)
        if sz:
            body = body.replace('<a:rPr lang="ko-KR" altLang="en-US" dirty="0"/><a:t>%s</a:t>'%esc(label),
                                '<a:rPr lang="ko-KR" altLang="en-US" sz="%d" dirty="0"/><a:t>%s</a:t>'%(sz, esc(label)), 1)
        return (body.encode(), (layrel(LAY_COVER)+"</Relationships>").encode())

    # ---- content (photo grid) slides ----
    content_slides = []
    pid=1000
    for ci, item in enumerate(CONTENT):
        title = item["title"]; files = item["photos"]
        para=('<a:p><a:r><a:rPr kumimoji="1" lang="ko-KR" altLang="en-US" sz="5999" dirty="0"/>'
              '<a:t>%s</a:t></a:r><a:endParaRPr kumimoji="1" lang="ko-KR" altLang="en-US" sz="5999" dirty="0"/></a:p>')%esc(title)
        body=re.sub(r"<a:p>.*?</a:p>", para, s_grid, flags=re.S, count=1)
        # if grid template has >1 paragraph, strip extras handled by count=1 leaves rest; ensure single title para:
        pics=""; rels=layrel(LAY_TITLE)
        if item.get("layout") == "row":
            # 한 줄 가로 배치: 모든 사진 높이 동일, 폭은 비율 비례(저스티파이드), 가운데 정렬
            g=300000
            dims=[photo_path(fn)[1] for fn in files]
            usable=RW-g*(len(files)-1)
            hrow=min(RH, usable/sum(pw/phh for pw,phh in dims))
            ws=[int(hrow*pw/phh) for pw,phh in dims]
            x0=LF+(RW-(sum(ws)+g*(len(files)-1)))//2; y0=TOP+int((RH-hrow)/2)
            cs=[]; xx=x0
            for wv in ws:
                cs.append((xx,y0,wv,int(hrow))); xx+=wv+g
        else:
            cs=cells(len(files))
        for k, fn in enumerate(files):
            rid="rIdP%d"%(k+1)
            rp, (pw, phh) = photo_path(fn)
            ext = os.path.splitext(rp)[1].lower().lstrip(".") or "jpeg"
            mp="ppt/media/photo_%d_%d.%s"%(ci+1,k+1,ext)
            data[mp]=open(rp,"rb").read()
            x,y,w,h=fit(*cs[k], pw, phh); pid+=1
            pics+=pic_xml(pid, rid, x,y,w,h)
            rels+='<Relationship Id="%s" Type="%simage" Target="../media/%s"/>'%(rid,REL,posixpath.basename(mp))
        body=body.replace("</p:spTree>", pics+"</p:spTree>")
        content_slides.append((body.encode(), (rels+"</Relationships>").encode()))

    # ================= import external sources =================
    # 외부 PPT 파트(마스터/레이아웃/테마/폰트) 번호는 마스터의 기존 최대값+1 부터 동적으로 부여한다.
    # 고정 시작값을 쓰면 마스터의 기존 파일을 덮어써 한 레이아웃을 두 마스터가 공유하게 되고 복구창이 뜬다.
    def _next(pattern):
        nums=[int(m.group(1)) for n in data for m in [re.match(pattern,n)] if m]
        return (max(nums)+1) if nums else 1
    n_master_file=[_next(r'ppt/slideMasters/slideMaster(\d+)\.xml$')]
    n_layout=[_next(r'ppt/slideLayouts/slideLayout(\d+)\.xml$')]
    n_theme=[_next(r'ppt/theme/theme(\d+)\.xml$')]
    n_font=[_next(r'ppt/fonts/font(\d+)\.fntdata$')]
    _mids=[int(x) for x in re.findall(r'<p:sldMasterId id="(\d+)"', data["ppt/presentation.xml"].decode("utf-8"))]
    n_master_id=[(max(_mids)+1) if _mids else 2147483674]
    tgt_pres=data["ppt/presentation.xml"].decode("utf-8")
    have_typefaces=set(re.findall(r'<p:font typeface="([^"]+)"',tgt_pres))
    ct_adds=[]; svg_needed=[False]
    pres_rel_adds=[]; master_id_adds=[]; font_embed_adds=[]
    per_source=[[] for _ in SOURCES]
    rid_counter=[700]
    def newrid():
        rid_counter[0]+=1; return "rId%d"%rid_counter[0]

    def wrap_scale(xml, cx, cy):
        # 일부 생성기는 xmlns:a/xmlns:r 를 루트가 아닌 개별 요소에 선언 → 삽입할 <a:xfrm> 이 unbound 되므로 루트에 보장
        m0=re.search(r'<p:sld\b[^>]*>', xml)
        if m0:
            root=m0.group(0); add=""
            if 'xmlns:a=' not in root: add+=' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
            if 'xmlns:r=' not in root: add+=' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
            if add: xml=xml.replace(root, root[:-1]+add+">", 1)
        grp=('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="90001" name="scalewrap"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
             '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/>'
             '<a:chOff x="0" y="0"/><a:chExt cx="%d" cy="%d"/></a:xfrm></p:grpSpPr>'%(W,H,cx,cy))
        xml=re.sub(r'(<p:spTree>.*?</p:grpSpPr>)', lambda m:m.group(1)+grp, xml, count=1, flags=re.S)
        xml=xml.replace("</p:spTree>","</p:grpSp></p:spTree>",1)
        return xml

    for si,(path,scale) in enumerate(SOURCES):
        z=zipfile.ZipFile(path); s={n:z.read(n) for n in z.namelist()}
        def srels(part):
            rp=relpath(part); return parse_rels(s[rp].decode()) if rp in s else []
        masters=sorted(n for n in s if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$',n))
        layouts=sorted(n for n in s if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$',n))
        pres=s["ppt/presentation.xml"].decode()
        prx={rid:resolve("ppt/presentation.xml",tgt) for rid,typ,tgt,mode in parse_rels(s["ppt/_rels/presentation.xml.rels"].decode())}
        order=re.findall(r'<p:sldId\b[^>]*r:id="([^"]+)"', re.search(r'<p:sldIdLst>.*?</p:sldIdLst>',pres,re.S).group())
        src_slides=[prx[r] for r in order if r in prx and '/slides/' in prx[r]]
        themes=set()
        for m in masters:
            for rid,typ,tgt,mode in srels(m):
                if typ.endswith('/theme'): themes.add(resolve(m,tgt))
        themes=sorted(themes)
        media=set()
        for part in masters+layouts+src_slides:
            for rid,typ,tgt,mode in srels(part):
                if mode=="Internal" and typ.endswith('/image'): media.add(resolve(part,tgt))
        rn={}
        for m in masters: rn[m]=f"ppt/slideMasters/slideMaster{n_master_file[0]}.xml"; n_master_file[0]+=1
        for l in layouts: rn[l]=f"ppt/slideLayouts/slideLayout{n_layout[0]}.xml"; n_layout[0]+=1
        for t in themes:  rn[t]=f"ppt/theme/theme{n_theme[0]}.xml"; n_theme[0]+=1
        for md in sorted(media):
            ext=md.rsplit('.',1)[-1].lower()
            rn[md]=f"ppt/media/s{si}_{posixpath.basename(md)}"
            if ext=="svg": svg_needed[0]=True
        for sl in src_slides: rn[sl]=f"ppt/slides/slide_imp_{si}_{src_slides.index(sl)}.xml"

        def remap_rels(xml, src_part):
            keep=[]
            for rid,typ,tgt,mode in parse_rels(xml):
                if mode!="Internal":
                    keep.append(f'<Relationship Id="{rid}" Type="{typ}" Target="{tgt}" TargetMode="{mode}"/>'); continue
                res=resolve(src_part,tgt)
                if typ.endswith('/notesSlide'):
                    continue
                if res in rn:
                    newt=posixpath.relpath(rn[res], posixpath.dirname(rn[src_part]))
                    keep.append(f'<Relationship Id="{rid}" Type="{typ}" Target="{newt}"/>')
                else:
                    continue
            return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    +"".join(keep)+"</Relationships>")

        for src_part,new_part in rn.items():
            body=s[src_part]
            if src_part in src_slides and scale:
                body=wrap_scale(body.decode("utf-8"), *scale).encode("utf-8")
            data[new_part]=body
            rp=relpath(src_part)
            if rp in s:
                data[relpath(new_part)]=remap_rels(s[rp].decode(), src_part).encode("utf-8")
            if new_part.startswith("ppt/slideMasters/"): ct_adds.append((new_part,CT["slideMaster"]))
            elif new_part.startswith("ppt/slideLayouts/"): ct_adds.append((new_part,CT["slideLayout"]))
            elif new_part.startswith("ppt/theme/"): ct_adds.append((new_part,CT["theme"]))

        for m in masters:
            rid=newrid()
            pres_rel_adds.append((rid,"slideMaster", posixpath.relpath(rn[m],"ppt")))
            master_id_adds.append((n_master_id[0],rid)); n_master_id[0]+=1

        fpres_rels={rid:(typ,resolve("ppt/presentation.xml",tgt)) for rid,typ,tgt,mode in parse_rels(s["ppt/_rels/presentation.xml.rels"].decode())}
        efl=re.search(r'<p:embeddedFontLst>(.*?)</p:embeddedFontLst>',pres,re.S)
        if efl:
            for ef in re.findall(r'<p:embeddedFont>.*?</p:embeddedFont>', efl.group(1), flags=re.S):
                tf=re.search(r'<p:font typeface="([^"]+)"',ef).group(1)
                if tf in have_typefaces: continue
                newef=ef; used=False
                for slot,rid in re.findall(r'<p:(regular|bold|italic|boldItalic)\b[^>]*r:id="([^"]+)"',ef):
                    if rid in fpres_rels and fpres_rels[rid][1] in s:
                        fnt=fpres_rels[rid][1]
                        newname=f"ppt/fonts/font{n_font[0]}.fntdata"; n_font[0]+=1
                        data[newname]=s[fnt]
                        nrid=newrid()
                        pres_rel_adds.append((nrid,"font", posixpath.relpath(newname,"ppt")))
                        newef=re.sub(r'(<p:%s\b[^>]*r:id=")[^"]+(")'%slot, r'\g<1>'+nrid+r'\g<2>', newef)
                        used=True
                if used:
                    have_typefaces.add(tf); font_embed_adds.append(newef)

        for sl in src_slides:
            new=rn[sl]
            per_source[si].append((data[new], data.get(relpath(new), (layrel(LAY_TITLE)+"</Relationships>").encode())))

    # ================= assemble per ASSEMBLY =================
    final = []
    for step in ASSEMBLY:
        t = step["type"]
        if t == "cover":     final.append(tmpl["cover"])
        elif t == "intro":   final.append(tmpl["intro"])
        elif t == "newfamily": final.append(tmpl["newfamily"])
        elif t == "end":     final.append(tmpl["end"])
        elif t == "content": final += content_slides
        elif t == "divider": final.append(divider(step["label"], step.get("sz")))
        elif t == "source":
            idx = step["index"]; sl = per_source[idx]
            if step.get("skip_first"): sl = sl[1:]
            if "slice" in step:
                a,b = step["slice"]; sl = per_source[idx][a:b]
            final += sl
        else:
            raise ValueError("unknown assembly step: %s" % t)

    # clear temp imported slide partnames
    for k in [k for k in list(data) if k.startswith("ppt/slides/slide_imp_")]:
        data.pop(k, None); data.pop(relpath(k), None)

    N=len(final)
    for i,(body,rels) in enumerate(final,1):
        data[f"ppt/slides/slide{i}.xml"]=body
        data[f"ppt/slides/_rels/slide{i}.xml.rels"]=rels

    # ---- globally-unique sldLayoutId @id across ALL masters (복구 주원인 자동교정) ----
    _ctr=[2147483800]; _used=set()
    for _m in [k for k in data if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$',k)]:
        _x=data[_m].decode("utf-8")
        def _repl(mo):
            while _ctr[0] in _used: _ctr[0]+=1
            v=_ctr[0]; _used.add(v); _ctr[0]+=1
            return mo.group(1)+str(v)+mo.group(3)
        _x=re.sub(r'(<p:sldLayoutId id=")(\d+)(")', _repl, _x)
        data[_m]=_x.encode("utf-8")

    # ---- prune media not referenced by any rels ----
    ref=set()
    for n,b in list(data.items()):
        if n.endswith(".rels"):
            for rid,typ,tgt,mode in parse_rels(b.decode("utf-8")):
                if mode=="Internal" and ("media/" in tgt):
                    ref.add(resolve(n.replace("/_rels/","/").replace(".rels",""), tgt))
    for n in [k for k in list(data) if k.startswith("ppt/media/")]:
        if n not in ref: data.pop(n,None)

    # ---- presentation.xml ----
    pres=data["ppt/presentation.xml"].decode("utf-8")
    mi="".join(f'<p:sldMasterId id="{mid}" r:id="{rid}"/>' for mid,rid in master_id_adds)
    pres=re.sub(r'(</p:sldMasterIdLst>)', mi+r'\1', pres, count=1)
    sld="".join(f'<p:sldId id="{256+i}" r:id="rIdS{i+1}"/>' for i in range(N))
    pres=re.sub(r"<p:sldIdLst>.*?</p:sldIdLst>", f"<p:sldIdLst>{sld}</p:sldIdLst>", pres, flags=re.S)
    if font_embed_adds and "<p:embeddedFontLst>" in pres:
        pres=pres.replace("</p:embeddedFontLst>", "".join(font_embed_adds)+"</p:embeddedFontLst>",1)
    data["ppt/presentation.xml"]=pres.encode("utf-8")

    # ---- presentation.xml.rels ----
    pr=data["ppt/_rels/presentation.xml.rels"].decode("utf-8")
    pr=re.sub(r'<Relationship[^>]*Type="[^"]*/relationships/slide"[^>]*/>',"",pr)
    adds=""
    for i in range(N):
        adds+=f'<Relationship Id="rIdS{i+1}" Type="{REL}slide" Target="slides/slide{i+1}.xml"/>'
    for rid,typ,tgt in pres_rel_adds:
        adds+=f'<Relationship Id="{rid}" Type="{REL}{typ}" Target="{tgt}"/>'
    pr=pr.replace("</Relationships>", adds+"</Relationships>")
    data["ppt/_rels/presentation.xml.rels"]=pr.encode("utf-8")

    # ---- [Content_Types].xml ----
    ct=data["[Content_Types].xml"].decode("utf-8")
    ct=re.sub(r'<Override PartName="/ppt/slides/slide\d+\.xml"[^>]*/>',"",ct)
    ov="".join(f'<Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="{CT["slide"]}"/>' for i in range(N))
    for part,c in ct_adds: ov+=f'<Override PartName="/{part}" ContentType="{c}"/>'
    if svg_needed[0] and 'Extension="svg"' not in ct:
        ct=ct.replace("<Default", '<Default Extension="svg" ContentType="image/svg+xml"/><Default',1)
    ct=ct.replace("</Types>", ov+"</Types>")
    data["[Content_Types].xml"]=ct.encode("utf-8")

    # ================= write =================
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    with zipfile.ZipFile(OUT,"w",zipfile.ZIP_DEFLATED) as zf:
        names=["[Content_Types].xml"]+[n for n in data if n!="[Content_Types].xml"]
        for n in names: zf.writestr(n, data[n])

    result = {
        "output": OUT,
        "slides": N,
        "media": len([n for n in data if n.startswith('ppt/media/')]),
        "masters": len([n for n in data if re.match(r'ppt/slideMasters/slideMaster\d+\.xml$',n)]),
        "layouts": len([n for n in data if re.match(r'ppt/slideLayouts/slideLayout\d+\.xml$',n)]),
        "themes": len([n for n in data if re.match(r'ppt/theme/theme\d+\.xml$',n)]),
        "size_mb": round(os.path.getsize(OUT)/1e6, 1),
    }
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return result


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("usage: python3 build_deck.py <config.json>", file=sys.stderr); sys.exit(2)
    main(sys.argv[1])
