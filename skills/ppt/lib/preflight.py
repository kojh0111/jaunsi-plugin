#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
preflight.py — 날짜 결정 + 폴더 생성 + 필수물 체크리스트(JSON)

기본 날짜 = 실행일로부터 가장 가까운(다가오는) 일요일.
  (오늘이 일요일이면 오늘, 아니면 다음 일요일.)

사용법:
    python3 preflight.py [--date YYMMDD] [--root /abs/church_pptx]   (root 기본값 = 현재 폴더)

출력: JSON
    {
      "date": "260614",
      "date_cover": "2026.06.14.",
      "root": "...", "date_dir": "...", "images_dir": "...", "ppts_dir": "...", "master_dir": "...",
      "masters": ["자운시 여름.pptx", ...],
      "checklist": {
        "master_library":  {"ok": true,  "detail": "..."},
        "chat_txt":        {"ok": false, "detail": "...", "found": []},
        "images":          {"ok": true,  "detail": "22장", "files": [...]},
        "ppts":            {"ok": true,  "detail": "3개(선택)", "files": [...]}
      },
      "ready": false,
      "guidance": ["images/ 에 사진을 넣으세요: ...", ...]
    }
"""
import os, sys, json, datetime, argparse, glob, shutil, zipfile

# 프로젝트 루트 = 현재 폴더(OS 무관). 플러그인 코드 위치와 무관하게 데이터는 항상 cwd 기준.
DEFAULT_ROOT = os.getcwd()
ROTATION_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "rotation.default.json")
IMG_EXT = (".jpg", ".jpeg", ".png", ".heic", ".webp")
PPT_EXT = (".pptx",)


def normalize_kakao_exports(date_dir, images_dir, ppts_dir):
    """날짜 폴더에 통째로 들어온 카톡 export(zip/폴더)를 규약 위치로 자동 분류.

    - Kakaotalk*/KakaoTalk* .zip → 같은 이름 폴더로 해제 후 zip 삭제
    - Kakaotalk*/KakaoTalk* 폴더 → *.txt 는 날짜 폴더 직속, 이미지는 images/,
      *.pptx 는 ppts/ 로 이동. 같은 이름이 이미 있으면 건너뜀(수동 확인).
    반환: 이동 내역 문자열 리스트.
    """
    moved = []
    for z in glob.glob(os.path.join(date_dir, "[Kk]akao*.zip")):
        folder = z[:-4]
        if not os.path.exists(folder):
            with zipfile.ZipFile(z) as zf:
                zf.extractall(folder)
            os.remove(z)
            moved.append("압축해제: %s" % os.path.basename(z))
    for d in glob.glob(os.path.join(date_dir, "[Kk]akao*")):
        if not os.path.isdir(d):
            continue
        for f in sorted(os.listdir(d)):
            if f.startswith("."):
                continue
            srcf = os.path.join(d, f)
            if not os.path.isfile(srcf):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext == ".txt":
                dest = os.path.join(date_dir, f)
            elif ext in IMG_EXT:
                dest = os.path.join(images_dir, f)
            elif ext in PPT_EXT:
                dest = os.path.join(ppts_dir, f)
            else:
                continue
            if os.path.exists(dest):
                moved.append("건너뜀(이미 있음): %s" % f)
                continue
            shutil.move(srcf, dest)
            moved.append("%s → %s/" % (f, os.path.basename(os.path.dirname(dest)) or f))
    return moved


def upcoming_sunday(today=None):
    today = today or datetime.date.today()
    # weekday(): Mon=0 .. Sun=6
    days = (6 - today.weekday()) % 7   # 0 if today is Sunday
    return today + datetime.timedelta(days=days)


def parse_date_arg(s):
    s = s.strip()
    if len(s) == 6 and s.isdigit():
        yy, mm, dd = int(s[0:2]), int(s[2:4]), int(s[4:6])
        return datetime.date(2000 + yy, mm, dd)
    raise ValueError("--date must be YYMMDD (e.g. 260614), got: %r" % s)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", default=None, help="YYMMDD (없으면 다가오는 일요일)")
    ap.add_argument("--root", default=DEFAULT_ROOT)
    args = ap.parse_args()

    d = parse_date_arg(args.date) if args.date else upcoming_sunday()
    date_yymmdd = d.strftime("%y%m%d")
    date_cover  = d.strftime("%Y.%m.%d.")

    root = os.path.abspath(args.root)
    master_dir = os.path.join(root, "master")
    date_dir   = os.path.join(root, date_yymmdd)
    images_dir = os.path.join(date_dir, "images")
    ppts_dir   = os.path.join(date_dir, "ppts")

    for p in (master_dir, date_dir, images_dir, ppts_dir):
        os.makedirs(p, exist_ok=True)

    normalized = normalize_kakao_exports(date_dir, images_dir, ppts_dir)

    # 정리정돈 로테이션 파일: 프로젝트 루트에 없으면 플러그인 기본값을 복사해 생성(사용자 데이터로 관리)
    rotation_file = os.path.join(root, "rotation.json")
    rotation_created = False
    if not os.path.exists(rotation_file) and os.path.exists(ROTATION_DEFAULT):
        shutil.copy(ROTATION_DEFAULT, rotation_file)
        rotation_created = True

    masters = sorted(os.path.basename(p) for p in glob.glob(os.path.join(master_dir, "*.pptx")))

    # chat txt (날짜 폴더 직속 자동 탐색)
    txts = sorted(os.path.basename(p) for p in glob.glob(os.path.join(date_dir, "*.txt")))
    images = sorted(os.path.basename(p) for p in os.listdir(images_dir)
                    if os.path.splitext(p)[1].lower() in IMG_EXT)
    ppts = sorted(os.path.basename(p) for p in os.listdir(ppts_dir)
                  if os.path.splitext(p)[1].lower() in PPT_EXT)

    checklist = {
        "master_library": {
            "ok": len(masters) > 0,
            "detail": ("%d개" % len(masters)) if masters else "비어있음",
        },
        "chat_txt": {
            "ok": len(txts) > 0,
            "detail": (", ".join(txts)) if txts else "없음",
            "found": txts,
        },
        "images": {
            "ok": len(images) > 0,
            "detail": ("%d장" % len(images)) if images else "없음",
            "files": images,
        },
        "ppts": {
            "ok": True,  # 선택 항목
            "detail": ("%d개" % len(ppts)) if ppts else "0개(선택)",
            "files": ppts,
        },
    }

    guidance = []
    if not checklist["master_library"]["ok"]:
        guidance.append("master 템플릿이 없습니다. %s 에 마스터 .pptx 를 넣으세요." % master_dir)
    if not checklist["chat_txt"]["ok"]:
        guidance.append("카톡 채팅 export(.txt)를 %s 직속에 넣으세요." % date_dir)
    if not checklist["images"]["ok"]:
        guidance.append("사진을 %s 에 넣으세요 (KakaoTalk_Photo_*.jpeg 등)." % images_dir)
    if not checklist["ppts"]["files"]:
        guidance.append("(선택) 끝 앞에 끼워넣을 외부 PPT 는 %s 에 넣으세요." % ppts_dir)

    # 필수: 마스터 라이브러리 + 카톡 txt + 이미지. ppts 는 선택.
    ready = all(checklist[k]["ok"] for k in ("master_library", "chat_txt", "images"))

    out = {
        "date": date_yymmdd,
        "date_cover": date_cover,
        "weekday_ok": d.weekday() == 6,
        "root": root,
        "master_dir": master_dir,
        "date_dir": date_dir,
        "images_dir": images_dir,
        "ppts_dir": ppts_dir,
        "masters": masters,
        "rotation_file": rotation_file,
        "rotation_created": rotation_created,
        "normalized": normalized,
        "checklist": checklist,
        "ready": ready,
        "guidance": guidance,
    }
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
