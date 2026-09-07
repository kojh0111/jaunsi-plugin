#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_env.py — 자운시 PPT 빌더 환경 점검 (stdlib 만 사용)

현재 폴더(프로젝트 루트)에서 실행한다. 도구·PowerPoint·데이터 폴더 상태를 JSON 으로 출력하고,
누락 항목마다 OS 별 설치 명령 힌트를 붙인다.

사용법:
    python3 check_env.py            # mac/linux
    python  check_env.py            # windows
"""
import os, sys, json, shutil, platform, subprocess, glob

IS_WIN = sys.platform.startswith("win")
IS_MAC = sys.platform == "darwin"


def which(*names):
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def version(cmd):
    try:
        out = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
        return (out.stdout or out.stderr).strip().splitlines()[0][:80]
    except Exception:
        return None


def find_powerpoint():
    if IS_MAC:
        for p in ["/Applications/Microsoft PowerPoint.app",
                  os.path.expanduser("~/Applications/Microsoft PowerPoint.app")]:
            if os.path.isdir(p):
                return p
        return None
    if IS_WIN:
        roots = [os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
        for r in roots:
            hits = glob.glob(os.path.join(r, "Microsoft Office", "**", "POWERPNT.EXE"), recursive=True)
            if hits:
                return hits[0]
        return None
    return None


def main():
    root = os.getcwd()
    checks = []

    def add(name, ok, detail, required=True, fix=None):
        checks.append({"name": name, "ok": bool(ok), "required": required,
                       "detail": detail, "fix": (None if ok else fix)})

    py = which("python3", "python", "py")
    add("python", py, "%s (%s)" % (py, platform.python_version()),
        fix=("winget install Python.Python.3.12 후 터미널 재시작" if IS_WIN else
             "brew install python (mac 은 보통 기본 제공)"))

    magick = which("magick")
    add("magick", magick, magick and version(["magick", "-version"]),
        fix=("winget install ImageMagick.ImageMagick 후 터미널 재시작" if IS_WIN else
             "brew install imagemagick"))

    pdftoppm = which("pdftoppm")
    add("pdftoppm", pdftoppm, pdftoppm or "없음(외부 PPT 이미지판 변환·렌더링 확인에만 필요)",
        required=False,
        fix=("winget install oschwartz10612.Poppler 또는 PowerPoint COM Export PNG 로 대체" if IS_WIN else
             "brew install poppler"))

    if IS_MAC:
        osa = which("osascript")
        add("osascript", osa, osa, fix="macOS 기본 도구 — 정상 설치라면 항상 있음")
        add("qlmanage", which("qlmanage"), which("qlmanage"), required=False,
            fix="macOS 기본 도구(미리보기용, 없어도 진행 가능)")

    ppt_app = find_powerpoint()
    add("powerpoint", ppt_app, ppt_app or "Microsoft PowerPoint 를 찾지 못함",
        fix="Microsoft 365 / Office 설치 후 PowerPoint 를 한 번 실행(라이선스 로그인)")

    master_dir = os.path.join(root, "master")
    masters = sorted(os.path.basename(p) for p in glob.glob(os.path.join(master_dir, "*.pptx"))
                     if not os.path.basename(p).startswith("~$"))
    add("master_dir", os.path.isdir(master_dir), master_dir,
        fix="python3 preflight.py 또는 /jaunsi:setup 이 자동 생성")
    add("master_pptx", len(masters) > 0, masters or "master/ 에 .pptx 없음",
        fix="전임자에게 마스터 pptx 를 받아 master/ 폴더에 넣기")

    rot = os.path.join(root, "rotation.json")
    add("rotation_json", os.path.exists(rot), rot, required=False,
        fix="첫 /jaunsi:ppt 실행 시 preflight 가 플러그인 기본값으로 생성")

    missing_required = [c["name"] for c in checks if c["required"] and not c["ok"]]
    print(json.dumps({
        "os": platform.system(), "os_version": platform.release(),
        "root": root, "checks": checks,
        "missing_required": missing_required,
        "ready": not missing_required,
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
