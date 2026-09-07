---
name: setup
description: 자운시 PPT 빌더 첫 설치·인수인계 시 환경 점검. python/magick/PowerPoint 존재, master/ 폴더와 마스터 pptx, rotation.json 을 확인하고 누락 항목의 설치 방법을 안내한다. "설치 점검", "환경 확인", "처음 세팅", "인수인계" 요청 시 사용.
---

# 자운시 PPT 환경 점검 (첫 설치·인수인계)

호출: `/jaunsi:setup` — **프로젝트 폴더(church_pptx)** 로 `cd` 한 뒤 Claude 를 실행한 상태에서 호출한다.
아직 폴더가 없으면 새 빈 폴더(예: `~/church_pptx`)를 만들어 그 안에서 실행해도 된다.

역할 분리: 코드는 플러그인(자동 업데이트), 데이터(마스터 pptx·rotation.json·주간 폴더)는 이 프로젝트 폴더.
데이터는 전임자에게 폴더째 받는다.

## 실행 흐름

### 1. 환경 점검
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/setup/lib/check_env.py"
```
(Windows: `python "${CLAUDE_PLUGIN_ROOT}/skills/setup/lib/check_env.py"`)

JSON 의 `checks` 를 표로 정리해 보여준다: 이름 / 필수 여부 / 상태 / 누락 시 `fix`(설치 명령).
- `python`, `magick`, `powerpoint` 는 필수. 없으면 `fix` 의 명령을 그대로 안내하고, 설치 후 **터미널을 재시작**해야
  PATH 가 반영된다고 알린다.
- `pdftoppm` 은 선택(외부 PPT 를 이미지판으로 바꿀 때만 필요). 없어도 진행 가능하되 설치를 권장.
- 도구 설치가 필요하면 AskUserQuestion 게이트 `설치했음` 으로 대기 → 다시 check_env 실행(루프).

### 2. 데이터 폴더 준비
```bash
mkdir -p master
```
- `master_pptx` 가 비어 있으면 AskUserQuestion 게이트:
  "마스터 pptx 는 전임자(이전 담당자)에게 받아 `<절대경로>/master/` 에 넣으세요. 넣었으면 선택." → `넣었음` → check_env 재실행.
  선택지: `넣었음` / `나중에 하겠음`(안내만 남기고 종료).
- `rotation.json` 은 없어도 된다 — 첫 `/jaunsi:ppt` 실행 때 플러그인 기본값으로 생성된다.
  전임자에게 받은 폴더에 이미 있으면 그것이 정본이니 덮어쓰지 않는다.

### 3. 다음 안내
`ready:true` 면 이렇게 마무리한다:
- "준비 완료. 매주 `<날짜>/` 폴더에 카톡 export 와 외부 PPT 를 넣고 `/jaunsi:ppt YYMMDD` 를 실행하세요."
- 플러그인 업데이트: `/plugin update jaunsi@jaunsi-plugin`.
- 자세한 사용법은 플러그인 README(저장소 `kojh0111/jaunsi-plugin`).

## Windows 치환표 (ppt/slim 스킬 공통)

스킬 본문의 명령 예시는 macOS 기준이다. Windows(PowerShell)에서는 아래로 바꿔 실행한다:
- `python3 ...` → `python ...` (`python` 이 없다고 뜨면 `py ...`).
- `unzip -p` → `python -c` + zipfile 등 동등한 수단. `cp` → `Copy-Item`, `pkill` → `Stop-Process -Name POWERPNT`.
- `qlmanage`(미리보기)는 mac 전용 — PowerPoint COM `Export` 로 PNG 를 뽑아 보여주거나, 미리보기를 생략하고 경로만 보고한다.
- AppleScript(PowerPoint→PDF / 열림 테스트)는 PowerShell COM 으로 대체:
  ```powershell
  $pp = New-Object -ComObject PowerPoint.Application
  $pres = $pp.Presentations.Open("<원본.pptx 절대경로>", $true)   # ReadOnly. 복구창이 뜨면 여기서 예외
  $pres.Slides.Count                                             # 열림 테스트: 슬라이드 수
  $pres.SaveAs("<출력.pdf 절대경로>", 32)                        # 32 = ppSaveAsPDF
  # 또는 페이지별 PNG 직접 추출(pdftoppm 불필요): $pres.Export("<png폴더>", "PNG", 3200, 1800)
  $pres.Close(); $pp.Quit()
  ```
  `Export` 로 PNG 를 바로 뽑았다면 `pdftoppm` 단계는 건너뛴다(파일명 순서만 pg-01.. 형식으로 정리).
- 파이썬 lib 스크립트(preflight/parse_chat/inspect_master/build_deck/make_imgdeck/validate/slim_pptx)는 stdlib 만 쓰므로 OS 무관.
