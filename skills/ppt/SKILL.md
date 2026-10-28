---
name: ppt
description: 자운시 주간 PPT 빌더. 카톡 사진/제목 + 외부 PPT를 자운시 마스터 템플릿(page-4 사진그리드)에 끼워 PowerPoint "복구" 창 없이 한 번에 열리는 주간 PPTX를 생성한다. 매주 날짜만 바꿔 동일 품질로 반복 실행. "만들래" / "자운시 ppt" / "주간 ppt" 요청 시 사용.
argument-hint: "[날짜 예:260614]"
---

# 자운시 주간 PPT 빌더

호출: `/jaunsi:ppt [YYMMDD]` — 반드시 **프로젝트 폴더(church_pptx, `master/` 가 있는 곳)** 에서 Claude 를 실행한 상태여야 한다.
코드는 플러그인(`${CLAUDE_PLUGIN_ROOT}`)에, 데이터(마스터·로테이션·주간 폴더)는 프로젝트 폴더에 둔다.

매주 카톡 채팅 export(사진+제목)와 외부 PPT 몇 개를 자운시 마스터 템플릿에 합쳐
**한 번에(복구 없이) 열리는** 주간 PPTX를 만든다. 출력은 **PPTX만**(PDF 생략).

핵심 목표: "복구해줘 → 폰트 그대로 → 다시 수정" 같은 반복 단계 **없이 한 방에** 완성.
그걸 위해 빌더가 복구 원인을 자동 교정하고 `validate.py` 8개 게이트로 차단한다.

---

## 폴더 규약
```
<프로젝트 루트>/                                      ← church_pptx 폴더 (여기서 Claude 실행)
  ├─ rotation.json                ← 정리정돈 로테이션(없으면 preflight 가 플러그인 기본값으로 생성)
  ├─ master/                      ← 마스터 템플릿 라이브러리(.pptx 모음)
  └─ <날짜>/                       예: 260614/
       ├─ *.txt                   (필수, 카톡 채팅 export — 자동 탐색)
       ├─ images/                 (사진들)
       └─ ppts/                   (끝 앞에 끼워넣을 외부 PPT, 선택)
출력: <프로젝트 루트>/<날짜>/<날짜>_자운시 PPT_완성.pptx
```
> 반드시 프로젝트 루트(`master/` 가 있는 폴더)로 `cd` 후 Claude 실행. preflight.py 는 **현재 폴더를 루트로** 쓴다(OS 무관).
> 처음 쓰는 기기라면 먼저 `/jaunsi:setup` 으로 도구·폴더를 점검한다.

도구: `python3`(stdlib만), `magick`(이미지 리사이즈/식별), `qlmanage`(미리보기, mac 전용). pip/uv 불필요.
모든 lib 경로는 `${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/` 기준(플러그인 설치 폴더 — Claude 가 실제 경로로 치환해 준다).

**Windows 에서 실행할 때**: 본문 명령 예시는 macOS 기준이다. `python3`→`python`, AppleScript→PowerPoint COM,
`qlmanage`→COM `Export` PNG 로 치환한다. 치환표는 `/jaunsi:setup` 스킬(`skills/setup/SKILL.md`)의
"Windows 치환표" 절에 있다. 파이썬 lib 스크립트는 stdlib 만 쓰므로 OS 무관.

---

## 실행 흐름 (이 순서를 지킬 것)

### 1. 날짜 결정 + 폴더/필수물 체크
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/preflight.py" [--date <YYMMDD>]
```
- 인자 없으면 **다가오는 일요일**로 자동 결정(오늘이 일요일이면 오늘). 인자 주면 그 날짜.
- `<날짜>/`, `images/`, `ppts/`, `master/` 생성 + 체크리스트 JSON 출력. 프로젝트에 `rotation.json` 이 없으면
  플러그인의 `rotation.default.json` 을 복사해 만든다(JSON 의 `rotation_created:true`).
- 날짜 폴더에 카톡 export 가 **zip/폴더째로** 들어와 있으면 자동 분류한다
  (txt → 직속, 이미지 → `images/`, pptx → `ppts/`; 결과는 JSON 의 `normalized` 에 표시).
  사용자에게는 "압축 풀 필요 없이 통째로 넣어도 된다"고 안내.
- `ppts/` 에 **pptx 대신 PDF** 가 들어오면(카톡에 pdf 광고가 올라온 경우) preflight 는 외부 PPT 0개로 세지만
  "외부 PPT 없음"으로 넘기지 말고 4.5절의 **PDF 경로**로 이미지판을 만들어 병합한다.

**누락이 있으면 멈추지 말고 "넣었음" 게이트(AskUserQuestion)로 대기할 것.**
스킬이 폴더를 이미 만들어 뒀으니, 사용자는 그 폴더에 파일을 채우기만 하면 된다.
누락 항목별 정확한 경로를 안내하고, 아래처럼 AskUserQuestion 으로 선택지를 만든다:
- 질문: "아래 폴더에 파일을 넣어주세요. 다 넣으셨으면 선택하세요."
  - 본문에 누락 항목과 넣을 **정확한 절대경로**를 나열
    (예: `카톡 txt → <날짜>/`, `사진 → <날짜>/images/`, `외부 PPT(선택) → <날짜>/ppts/`)
- 선택지:
  - **`넣었음`** — 다시 `preflight.py` 를 실행해 재확인. 여전히 누락이면 같은 게이트를 다시 띄운다(루프).
  - **`외부 PPT 없이 진행`** — ppts 가 비어도 됨(선택 항목)일 때.
  - **`취소`** — 중단.
> 핵심 UX: "폴더는 내가(스킬) 만들어 둔다 → 너가(사용자) 채운다 → `넣었음` 누르면 이어서 진행."
> 절대 사용자에게 "명령을 다시 입력하라"고 떠넘기지 말 것 — `넣었음` 게이트로 그 자리에서 재개.
- `ready:true` 가 되면 다음 단계로.

### 2. 마스터 선택 (AskUserQuestion) → 마스터 구조 확인
- `preflight` 의 `masters` 목록을 **선택지로** 제시 → 사용자가 하나 선택.
  (`.bak`·`원본백업` 같은 백업 파일은 선택지에서 제외하고, 이름에 계절이 있으면 현재 계절을 추천.)
- 목록에 없고 외부 경로를 주면 그 파일을 `master/` 로 **복사**한 뒤 선택.
  ```bash
  cp "<외부경로>" master/
  ```
- `master/` 가 비어 있으면 1단계와 동일한 **`넣었음` 게이트**:
  "`master/` 폴더에 템플릿 .pptx 를 넣어주세요" 안내 → `넣었음` 누르면 다시 `ls`/preflight 로 목록 갱신.

**마스터 구조 확인(필수)** — 마스터마다 레이아웃 번호와 텍스트 run 분할이 다르다. 추측하지 말고 확인한다:
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/inspect_master.py" "master/<선택한 마스터>.pptx"
```
- `templates_suggest` 를 config.json 의 `templates` 로 그대로 쓴다
  (슬라이드 1~5 = 표지/정리정돈/새가족/사진그리드/끝, `layout_cover`=표지 레이아웃, `layout_title`=정리정돈 레이아웃).
- `slides[0].runs`(표지) 와 `slides[1].runs`(정리정돈) 의 run 목록을 메모해 둔다 — 3.5·5단계 치환 쌍은
  **이 run 문자열과 정확히 일치**해야 적용된다(예: 표지 날짜가 `2024.10` + `.20.` 두 run 으로 나뉘어 있으면
  치환도 두 쌍으로 만든다).

### 3. 카톡 파싱 → 제목·사진수 후보 제시
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/parse_chat.py" "<날짜>/<그.txt>" --date <날짜> [--since <YYYY.M.D>]
```
- 채팅 txt 는 수년치 누적이다. `--date <날짜>` 를 **반드시** 넘긴다 — `--since` 생략 시 대상 주일의 7일 전(직전 일요일)부터
  필터한다. 실행하는 날을 기준으로 잡으면 전날(토요일) 사진이 잘려 나가므로 실행일 기준을 쓰지 않는다.
- 후보가 0건이거나 카톡 원문의 사진 메시지보다 적으면 `--since` 를 며칠 앞당겨 다시 실행한다.
- "사진"/"사진 N장" + 다음 줄 제목 규칙으로 후보 추출(JSON).
- 사진은 `magick`+`montage` 로 몽타주 이미지를 만들어 **SendUserFile 로 먼저 보여준다**.
- `maybe:true`(무제·다른발신자) 후보는 채팅 원문 앞뒤 줄을 직접 읽어 교차확인
  (예: "사진" 다음 줄이 제목, 그 다음 "파일:" 줄은 외부 PPT — 삭제된 메시지가 끼면 파서가 어긋남).
- **주의**: AskUserQuestion 본문에는 마크다운 표가 렌더링되지 않는다. 매핑 확인은
  몽타주 전송 + 질문 문장 안에 `① 제목: 사진1 / ② 제목: 사진2·3` 식 나열로 할 것.

### 3.5 정리정돈 로테이션 계산 (intro 슬라이드)

intro(2번) 슬라이드의 "정리정돈 담당 청년부"는 **매월 바뀐다**. 사용자에게 묻기 전에 스킬이 먼저 계산한다.
정리정돈 슬라이드는 **2번과 맨 마지막(끝 뒤) 두 곳**에 같은 내용으로 들어가는 것이 정상이다 —
`assembly` 를 `{"type":"end"}` 뒤 `{"type":"intro"}` 로 끝내야 한다(5단계). 둘 다 같은 `replacements.intro` 가 적용된다.

1. `<프로젝트 루트>/rotation.json` 을 읽는다 — `order`(순환 순서), `anchor`(기준 월=담당), `overrides`(예외).
2. 대상 월의 담당을 스크립트로 계산한다(손 계산 금지):
   ```bash
   python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/rotation_calc.py" rotation.json <날짜>
   ```
   규칙: `overrides[YYYY-MM]` 우선, 없으면 `order[(anchor.team의 index + anchor.month부터의 개월 차) % len(order)]`.
   예: anchor 2026-06=1청년부, order 가 1청년부→2청년부→3청년부→신혼청 이면 7월=2청년부, 8월=3청년부, 9월=신혼청, 10월=1청년부.
3. `replacements.intro` 치환 쌍을 구성한다. 치환은 `build_deck.py`의 `apply_repl`이
   **`<a:t>OLD</a:t>` run 전체 일치**로 처리하므로 2단계에서 확인한 run 단위로 짝을 만든다.
   intro 슬라이드의 전형적 run 구조: `N`(청년부 번호) · `청년부` · `YY`(연도) · `년 ` · `M`(월) · `월 정리정돈`
   - 예) 마스터 run 이 `2`,`청년부`,`24`,`년 `,`7`,`월 정리정돈` 이고 2026년 8월 = 3청년부:
     `[["2","3"], ["24","26"], ["7","8"]]`
   - 예) 담당이 신혼청인 달(2026년 9월): `[["2","신혼"], ["청년부","청"], ["24","26"], ["7","9"]]`
   연도 치환(`24`→`26`)만 하고 월·청년부 번호를 안 바꾸는 실수 금지.
   짧은 run(`1`,`2` 같은 한 글자)은 같은 슬라이드의 다른 run 과 겹칠 수 있으니 치환 쌍 3번째 값으로 횟수를 제한할 수 있다(`["1","3",1]`).
4. 계산 결과를 사용자에게 **확인만** 받는다: "이번 달(7월) 정리정돈 = 2청년부 맞나요?"
   - 맞으면 진행. 틀리면 사용자가 말한 담당으로 반영하고, **프로젝트 `rotation.json`의 `overrides`에
     `"YYYY-MM": "N청년부"` 로 기록**해 다음에 같은 오류가 반복되지 않게 한다.
   - 로테이션 자체가 바뀌었다면(순서/기준 변경) `anchor`를 이번 달 기준으로 갱신한다.

### 4. 사진↔제목 매핑·순서 + 외부 PPT 알약 확인 (AskUserQuestion)
필수로 물어볼 것:
1. **제목별 사진 매핑/순서** — 각 제목에 어떤 사진(파일명)이 1/2/3/4장 들어가는지, 슬라이드 순서.
   카톡에 없는 추가 슬라이드(예: 사용자가 별도로 준 운영위 사진)도 content 항목으로 그냥 추가하면 된다.
   - 그리드는 **최대 4장**. 5장 이상이면 **같은 제목으로 4장씩 분할**(마지막 슬라이드에 나머지).
   - 세로사진 위주 슬라이드는 `layout:"row"`(한 줄 가로 배치)를 제안할 것.
   - 가로 2장 + 세로 1장인 3장 슬라이드는 `layout:"side"`(가로 2장 왼쪽에 위아래, 세로 1장 오른쪽)를 제안할 것.
     photos 순서가 곧 배치다: `[가로, 가로, 세로]`.
   - **한 줄 원칙**: 콘텐츠 제목·알약 라벨은 한 줄이 기본. 긴 제목이 줄바꿈/넘침될 것 같으면
     **글자 크기를 임의로 줄이지 말고** 사용자에게 먼저 보고해 문구를 조정받는다.
2. **외부 PPT 병합 순서** — 외부 PPT 의 **기본 위치는 콘텐츠(사진) 슬라이드 뒤, `end` 앞**이다.
   사용자에게는 외부 PPT 들끼리의 순서만 확인하고, 다른 위치를 명시적으로 요구할 때만 바꾼다.
3. **PPT마다 구분 알약(divider) 넣을지 + 텍스트** — 각 외부 PPT 앞에 알약 슬라이드를
   넣을지 말지, 넣으면 라벨(예: `새가족팀`, `양육팀`)을 무엇으로 할지 **각각** 확인.
   (알약은 마스터의 layout_cover(끝/표지 스타일)로 텍스트만 교체해 통일된다.)
4. **캔버스 크기 다른 외부 PPT** — 빌더가 자동 감지한다. `scale` 은 `null` 로 두면 된다.
   - 마스터와 같으면 그대로, 다르면 **확대**거나 **이미지·사진뿐**일 때 자동 축척(빌드 로그에 "자동 축척" 출력).
   - **텍스트가 있는데 축소**해야 하면 빌드가 **중단**된다 — 그룹 축소는 도형만 줄이고 **글자 크기를
     안 줄여서 레이아웃이 전부 깨지기** 때문. 아래 "이미지판 변환" 절차로 바꿔 넣을 것.
   - OLE 임베드(엑셀 표 등, `ppt/embeddings/*.bin`)가 있는 PPT 도 rel 이 끊겨 복구 원인이 되므로
     이미지판 경로가 안전하다.

### 4.5 외부 PPT 이미지판 변환 (캔버스 다르고 텍스트 있을 때)

원본을 픽셀 그대로 보존하면서 병합하는 경로. 편집 가능성만 포기한다(발표용이므로 무해).

**원본이 PDF 일 때**(카톡에 pdf 광고만 올라온 경우): 1단계를 건너뛰고 2단계(`pdftoppm`)부터 간다.
원본 pptx 골격이 없으므로 4단계 전에 `make_skeleton.py` 로 **빈 N장 골격**을 만든다:
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/make_skeleton.py" "<골격.pptx>" "master/<마스터>.pptx" <장수> <마스터 cx> <마스터 cy>
```
캔버스(cx/cy)는 마스터와 똑같이 준다(`unzip -p "master/<마스터>.pptx" ppt/presentation.xml | grep -o '<p:sldSz[^/]*/>'`) —
그래야 축척 없이 풀블리드로 병합된다. 4단계의 `<원본.pptx>` 자리에 이 골격을 넣고, 병합 후 골격 파일은
**삭제**한다(`ppts/` 에 남기면 preflight 가 외부 PPT 로 센다). PowerPoint AppleScript 로 빈 골격을 만드는 방법은
실패한다(`page setup` -10006, `save as presentation` 이 구형 .ppt 로 저장) — 쓰지 말 것.

1. **PowerPoint(AppleScript)로 PDF 저장** — 임베드 폰트를 그대로 렌더링한다(로컬 폰트 설치 불필요):
   ```applescript
   tell application "Microsoft PowerPoint"
     activate
     open POSIX file "<원본.pptx 절대경로>"
     delay 3
     set pres to active presentation
     save pres in POSIX file "<출력.pdf 절대경로>" as save as PDF
     close pres saving no
   end tell
   ```
   주의: `save as PNG`(폴더 저장)는 macOS 에서 조용히 묵살된다 — **PDF 만 사용**.
   대용량(100MB+) 파일은 저장이 조용히 실패할 수 있다 — 저장 후 파일 존재를 반드시 확인.
2. **PDF → PNG**: `pdftoppm -png -r 160 <출력.pdf> <png폴더>/pg` (3200x1800 수준).
3. **추출 PNG 를 Read 로 직접 열어 눈으로 확인** — 폰트·표가 원본과 같은지. 여기서 깨졌으면
   임베드 폰트 문제이니 원인부터 해결.
4. **이미지판 pptx 생성**:
   ```bash
   python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/make_imgdeck.py" "<원본.pptx>" "<이미지판.pptx>" "<png폴더>"
   ```
   원본 골격을 재사용해 슬라이드 본문만 풀블리드 이미지로 교체한다(캔버스 크기 자동 감지).
5. config.json 의 source 를 이미지판으로 교체하고 `scale` 은 `null`(자동) — 이미지판은 텍스트가 없으므로
   빌더가 원본 캔버스 크기로 자동 축척한다(이미지는 그룹 축소로도 올바르게 줄어든다).
6. 빌드 후 중간 산출물(PDF, png 폴더)은 사용자에게 물어보고 정리.

### 5. config.json 작성 → 빌드
4단계 결과로 `<날짜>/config.json` 작성(스키마는 `lib/build_deck.py` 상단 docstring 참고).
`templates` 는 2단계 `templates_suggest`, `replacements.cover` 는 2단계 표지 run 기준으로 만든다
(표지 원본 날짜 run → 이번 주 일요일 날짜; run 이 둘로 나뉘어 있으면 두 쌍).
`assembly` 는 반드시 `{"type":"end"}` 뒤에 `{"type":"intro"}` 를 두어 **정리정돈이 2번과 맨 끝 두 곳**에 들어가게 한다
(끝에 없으면 빌더가 경고를 낸다).
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/build_deck.py" "<날짜>/config.json"
```
- 사진은 빌더가 `magick -auto-orient -resize 2000x2000> -strip` 로 사전 리사이즈
  (EXIF 회전 보정 + 용량↓) 후 실제 비율로 fit-contain 배치.
- 빌드 후 `unzip -p <완성.pptx> ppt/slides/slide1.xml | grep -o '<a:t>[^<]*</a:t>'` 로 표지·intro(slide2)·마지막 슬라이드(정리정돈) 치환이
  실제로 반영됐는지 확인한다. 치환이 안 됐으면 run 문자열이 다른 것이니 2단계 결과로 되돌아가 쌍을 고친다.

### 6. 복구방지 검증 (모두 통과해야 산출 인정)
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/validate.py" "<날짜>/<날짜>_자운시 PPT_완성.pptx"
```
- 8개 게이트 전부 `ok:true` 여야 함. 하나라도 실패면 **중단 + 리포트** 후 원인 수정 재빌드.
- **게이트는 "복구 없이 열림"만 보장하고 "제대로 보임"은 보장하지 않는다**.
  외부 PPT 를 scale 병합했다면 병합된 슬라이드 1~2장을 렌더링해 **직접 눈으로 확인**할 것:
  4.5절 1~2단계(PowerPoint→PDF→pdftoppm)로 해당 페이지만 뽑아 Read 로 검사.
  완성본이 대용량이라 PDF 저장이 실패하면, 4.5절 3단계에서 소스 이미지판을
  이미 눈으로 확인했는지로 갈음한다.
  PDF 저장이 시간초과(-1712)면 PowerPoint 가 저장 폴더 권한을 묻는 **"파일 액세스 부여" 창**에 막힌 것이다
  (`System Events` 로 창 이름 확인). 권한 창은 대신 누르지 말고, 확인할 슬라이드만 1장씩 담은 변형 config
  (`content`/`assembly` 를 그 슬라이드 하나로)를 빌드해 `qlmanage -t -s 800` 으로 첫 장을 렌더링해 확인한다.
- **PowerPoint 실제 열림 테스트(필수, 게이트와 별도)** — 게이트가 못 잡는 손상이 있을 수 있으므로
  산출물을 PowerPoint 로 직접 열어 판정한다. 복구창이 뜨면 아래 스크립트가 **시간초과(exit 124)**
  또는 **`-50 매개변수 오류`** 로 실패한다. 슬라이드 수가 출력되고 exit 0 이어야 통과.
  ```bash
  pkill -x "Microsoft PowerPoint"; sleep 3
  osascript -e 'tell application "Microsoft PowerPoint" to activate'; sleep 5
  timeout 60 osascript -e 'tell application "Microsoft PowerPoint"
  open POSIX file "<완성.pptx 절대경로>"
  delay 8
  set n to count of slides of active presentation
  close active presentation saving no
  return n
  end tell'; echo "exit=$?"
  ```
  `-609 연결이 유효하지 않습니다` 는 앱이 아직 뜨기 전에 명령을 보낸 것이므로 파일 문제가 아니다 — 몇 초 기다렸다가 재시도한다.
  실패하면 **이분법으로 원인 분리**: 외부 PPT 제외 / 알약만 / 소스 하나만 변형 config 를 만들어 각각 빌드·열림
  테스트해 범위를 좁힌 뒤 원인을 고친다(원본 외부 PPT 단독 열림 여부도 같이 확인).
  Windows 는 PowerPoint COM `Presentations.Open` 이 예외를 던지는지로 판정.

### 7. 산출물 보고 + 미리보기
- 경로 보고. 핵심 슬라이드(표지/사진그리드/알약/끝) `qlmanage -t -s 1200` 로 미리보기 생성해 제시.
- **PPTX만.** PDF는 만들지 않는다(내부 검증용 렌더링 PDF 는 산출물이 아님 — 보고 후 정리 대상).
- **휴대폰 전달(선택)**: 대화 첨부(SendUserFile)는 30MB 한도라 완성본(보통 30MB+)은 못 보낸다.
  iCloud Drive 복사는 **사용자가 요청할 때만** 한다. 요청이 없으면 **macOS 에서만** "iCloud 로 복사할까요?" 한 줄로 묻고,
  Windows 에서는 묻지 않는다(iCloud Drive 경로가 없다). 복사할 때는 `~/Library/Mobile Documents/com~apple~CloudDocs/자운시 PPT/` 에
  넣고 `cmp` 로 원본과 같은지 확인한다 — 휴대폰 '파일' 앱 → iCloud Drive → 자운시 PPT 에서 받는다.
  재빌드하면 다시 복사해 덮어쓴다.

---

## 복구가 왜 뜨고 무엇으로 막는가 (반드시 기억)

PowerPoint "복구" 창은 패키지 내부 참조가 어긋날 때 뜬다. 빌더/검증기가 막는 원인과 처리:

| 현상(복구 원인) | 원인 | 빌더 처리 | 게이트 |
|---|---|---|---|
| **여러 마스터의 `sldLayoutId @id` 충돌** — 가장 흔한 원인 | 외부 PPT 병합 시 레이아웃 ID 가 마스터와 겹침 | 전 마스터 레이아웃 ID를 2147483800+ 로 **전역 재번호** | 2 |
| 본문 `r:id`/`r:embed` 가 .rels 에 없음(dangling) | 소스 import 시 rId 만 옮기고 rels 를 안 옮김 | .rels 리매핑, notesSlide/printerSettings 등 미복사 타겟 drop | 3 |
| .rels Target 파일이 실제로 없음 / 중복 rId | 미디어·레이아웃·테마를 부분만 복사 | 통째 복사 + 미참조 미디어 prune | 4 |
| `[Content_Types]` 미커버(특히 svg Default 누락) | 새 파트 추가 후 타입 등록 누락 | slide/master/layout/theme Override 추가, svg 있으면 Default 추가 | 5 |
| 슬라이드 내 `cNvPr @id` 중복 | 사진 도형을 같은 id 로 여러 개 삽입 | 사진 도형 id 를 1000+ 증가 카운터로 부여 | 6 |
| `sldIdLst` ↔ presentation.rels 불일치(고아/누락 슬라이드) | 슬라이드 rId 를 손으로 관리 | 슬라이드 rId 를 rIdS1..N 으로 일괄 재배선 | 7 |
| **한 레이아웃/테마/마스터 파일을 두 마스터가 공유** | 외부 PPT 파트 번호를 고정 시작값으로 부여해 마스터의 기존 파일을 덮어씀 | 시작 번호를 마스터의 최대값+1 로 **동적 계산** | 8 |

폰트: 외부 PPT의 **임베드 폰트(embeddedFontLst)** 를 `.fntdata` 통째로 import 하고
presentation.xml 에 다시 연결 → "폰트 그대로" 유지(재설치/치환 불필요).

이 8개 보정은 `lib/build_deck.py` 에 들어 있다. 절대 이 8개 보정을 빼지 말 것 —
하나라도 빠지면 복구 창이 돌아온다.

---

## config.json 핵심 필드 (요약)
- `master`, `output`, `photos_dir` — 절대경로.
- `templates` — 2단계 `inspect_master.py` 의 `templates_suggest`. `layout_cover`/`layout_title` 를 생략하면
  빌더가 표지/정리정돈 슬라이드의 rels 에서 자동 검출한다.
- `replacements.cover` — 표지 날짜: 마스터 원본 날짜 run → 이번 주 날짜. run 이 하나면 `[["2024.07.07.","2026.06.07."]]`,
  둘로 나뉘어 있으면 `[["2024.10","2026.09"], [".20.",".06."]]` 처럼 run 마다 한 쌍.
- `replacements.intro` — 정리정돈 슬라이드 치환. **연도 숫자만이 아니라 3.5단계에서 계산한
  월·담당 청년부까지 반영할 것**.
- `content` — `[{title, photos:[파일명...]}]` (제목당 1슬라이드, 사진 1/2/3/4 그리드).
  항목에 `"layout":"row"` 를 주면 한 줄 가로(저스티파이드) 배치 — 세로사진 3~4장일 때 권장.
  `"layout":"side"` 는 3장 전용 — photos 앞 2장을 왼쪽에 위아래로 쌓고 3번째(세로사진)를 오른쪽에 둔다.
- `sources` — `[{path, scale|null}]` 외부 PPT. `scale` 은 `null`(자동 감지)이 기본. 직접 `[cx,cy]` 를 주면
  그 값을 쓰되 실제 캔버스와 다르면 경고한다. 텍스트 있는 PPT 를 축소해야 하면 빌드가 중단되니
  4.5절의 이미지판을 path 로 지정할 것.
- `assembly` — 최종 순서. 예:
  ```json
  [{"type":"cover"},{"type":"intro"},{"type":"newfamily"},{"type":"content"},
   {"type":"divider","label":"새가족팀"},{"type":"source","index":0},
   {"type":"divider","label":"양육팀"},{"type":"source","index":1,"skip_first":true},
   {"type":"source","index":2},{"type":"end"},{"type":"intro"}]
  ```
  - `intro`(정리정돈)는 **2번과 맨 끝(`end` 뒤) 두 곳** — 같은 슬라이드가 두 번 들어가는 것이 정상이다.
    끝에 없으면 빌더가 경고를 낸다.
  - `source` 의 **기본 위치는 `content` 뒤, `end` 앞** — 사용자가 명시적으로 다른 위치를
    요구할 때만 바꾼다. `content` 앞에 두지 말 것.
  - `divider` 는 **선택** — 알약 안 넣을 PPT 앞에는 넣지 않으면 됨(4단계에서 사용자 확인).
  - `skip_first:true` 는 외부 PPT 의 자체 표지를 버릴 때.

---

## 검증(스킬 자체 재현 테스트)
임의의 한 주 자료(마스터 1개, 카톡 txt, 사진 여러 장, 외부 PPT 1~3개)로 위 흐름을 실행해
장수가 `5(표지·정리정돈·새가족·끝·정리정돈) + 콘텐츠 슬라이드 + 알약 + 외부 PPT 슬라이드` 와 일치하고, `validate.py` 8게이트 전부 통과,
PowerPoint 열림 테스트 exit 0 이면 재현 성공이다.
