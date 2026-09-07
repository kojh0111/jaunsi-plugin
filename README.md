# 자운시 주간 PPT 빌더 (Claude Code 플러그인 `jaunsi`)

높은뜻광성교회 청년마을 자운시 주간 PPT 를 만드는 Claude Code 플러그인입니다.
카톡 채팅 export(사진 + 제목)와 외부 PPT 를 마스터 템플릿에 합쳐,
PowerPoint 에서 **복구 창 없이 한 번에 열리는** 주간 PPTX 를 만듭니다.

| 스킬 | 용도 |
|---|---|
| `/jaunsi:setup` | 첫 설치·인수인계 때 환경(python, ImageMagick, PowerPoint)·데이터 폴더 점검 |
| `/jaunsi:ppt YYMMDD` | 매주 주간 PPT 빌드 |
| `/jaunsi:slim <pptx>` | 마스터 파일 용량이 비정상적으로 클 때 슬림화(요청 시에만) |

코드(이 플러그인)와 데이터(마스터 pptx, rotation.json, 매주 자료)는 분리돼 있습니다.
데이터는 GitHub 에 올리지 않고, `church_pptx` 폴더째 전임자에게 받습니다.

---

## 새 담당자 시작하기 (3단계)

### 1. Claude Code 설치
https://claude.com/claude-code 안내대로 설치하고 로그인합니다. 본인 Claude 계정을 사용합니다.

### 2. 플러그인 설치
터미널에서 `claude` 를 실행한 뒤 Claude 안에서:
```
/plugin marketplace add kojh0111/jaunsi-plugin
/plugin install jaunsi@jaunsi-plugin
```
업데이트는 `/plugin update jaunsi@jaunsi-plugin` 한 줄입니다. 코드는 항상 이 저장소의 최신 버전을 따릅니다.

### 3. 데이터 폴더 준비 → 환경 점검
전임자에게 `church_pptx` 폴더(아래 구조)를 받아 원하는 위치에 둡니다.
받은 폴더가 없으면 빈 폴더를 만들고 마스터 pptx 한 개만 구해 `master/` 에 넣으면 됩니다
(`/jaunsi:setup` 이 폴더를 만들어 주고 파일을 넣을 때까지 기다립니다). 마스터가 전혀 없으면 지난주 완성본 PPT 에서
표지·정리정돈·새가족·사진 그리드·끝 5장만 남긴 파일을 마스터로 씁니다.
```
church_pptx/
  ├─ master/            ← 마스터 템플릿 .pptx (계절별)
  ├─ rotation.json      ← 정리정돈 담당 청년부 로테이션 (정본)
  └─ 260906/ ...        ← 지난 주간 자료 (참고용, 없어도 됨)
```
그 폴더에서 Claude 를 실행하고 환경 점검을 합니다:
```bash
cd <church_pptx 폴더>
claude
```
```
/jaunsi:setup
```
누락된 도구가 있으면 설치 명령을 알려 줍니다(예: mac `brew install imagemagick poppler`,
Windows `winget install ImageMagick.ImageMagick`). Microsoft PowerPoint 는 열림 테스트에 필요하므로 설치돼 있어야 합니다.

---

## 매주 하는 일

1. `church_pptx/<날짜>/` 폴더를 만들고(스킬이 만들어 주기도 함) 자료를 넣습니다. `<날짜>` 는 그 주 일요일 `YYMMDD`.
   ```
   church_pptx/260614/
   ├─ *.txt      ← 카톡 채팅 export (필수). zip 째로 넣어도 자동 분류됨
   ├─ images/    ← 사진
   └─ ppts/      ← 끝 앞에 끼워넣을 외부 PPT (선택)
   ```
2. `cd church_pptx` → `claude` → `/jaunsi:ppt 260614` (날짜를 빼면 다가오는 일요일).
3. Claude 가 순서대로 묻습니다: 마스터 선택 → 이번 달 정리정돈 담당 확인 → 사진↔제목 매핑·순서 →
   외부 PPT 순서와 구분 알약(새가족팀/양육팀 등) 문구.
4. 빌드 후 자동 검증(구조 게이트 8개 + PowerPoint 실제 열림 테스트)을 통과하면 완성:
   ```
   church_pptx/260614/260614_자운시 PPT_완성.pptx
   ```

### 알아둘 것
- **마스터가 바뀌면** `.pptx` 를 `master/` 에 넣기만 하면 선택지에 나옵니다. 레이아웃 번호·문구 구조는 스킬이 파일에서 직접 읽습니다.
- **정리정돈 로테이션**은 `rotation.json` 이 정본입니다. 계산이 틀리면 Claude 에게 말하면 `overrides` 에 기록해 재발을 막습니다.
- **외부 PPT 가 마스터와 캔버스 크기가 다르고 글자가 있으면** 스킬이 PowerPoint 로 렌더링한 이미지판으로 병합합니다(글자 깨짐 방지).
- 마스터 파일이 100MB 급으로 커졌을 때만 `/jaunsi:slim master/<파일>.pptx` 를 씁니다. 매주 빌드에는 쓰지 않습니다.

---

## 인수인계 체크리스트

| 항목 | 어디에 | 넘기는 방법 |
|---|---|---|
| 코드 | 이 GitHub 저장소 | 새 담당자가 `/plugin install` (항상 최신) |
| 데이터 | `church_pptx/` (master, rotation.json) | 폴더째 전달 (USB·드라이브). 저장소에는 올리지 않음 |
| 계정 | 본인 Claude 계정 | 각자 로그인 |

코드에 문제가 있으면 이 저장소에 이슈를 남기거나, `church_pptx` 폴더에서 Claude 에게 증상을 설명해 고치게 한 뒤 저장소로 PR 을 보냅니다.

## 폴더 구조 (플러그인)
```
.claude-plugin/plugin.json     플러그인 정보
skills/setup/                  환경 점검 (check_env.py)
skills/ppt/                    주간 빌더 (preflight, parse_chat, inspect_master, build_deck, make_imgdeck, validate)
skills/slim/                   pptx 슬림화 (slim_pptx.py)
```
모든 스크립트는 Python 표준 라이브러리만 사용합니다(pip 불필요). 이미지 처리에 ImageMagick(`magick`)이 필요합니다.
