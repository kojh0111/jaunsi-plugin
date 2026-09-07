---
name: slim
description: PPTX(주로 마스터 템플릿) 용량 검증·슬림화. 파일이 비정상적으로 클 때 원인을 분석하고, 화면 픽셀 변화 없이 안전하게 용량을 줄인다. 사용자가 "마스터 검증", "용량 왜 커", "용량 줄여줘" 등으로 **명시적으로 요청할 때만** 사용. 주간 PPT 제작 흐름(jaunsi:ppt)에서는 자동으로 실행하지 않는다.
argument-hint: "[대상.pptx 경로]"
---

# PPTX 용량 검증·슬림화

호출: `/jaunsi:slim <대상.pptx>`

마스터 템플릿 같은 pptx가 너무 클 때, **어디서 용량이 나오는지 검증**하고
**화면에 보이는 픽셀은 그대로 둔 채** 안전하게 줄인다.
마스터가 100MB 급으로 커진 경우에도 20MB 안팎까지 줄어드는 절차다.

> 검증 요청 시에만 사용. 매주 PPT 빌드에 끼워넣지 말 것.
> Windows 에서는 `python3`→`python`, AppleScript 열림 확인→PowerPoint COM(`New-Object -ComObject PowerPoint.Application`)으로 치환. `magick` 필요(`winget install ImageMagick.ImageMagick`).

## 원리 (왜 안전한가)

pptx가 비대해지는 전형적 원인과 대응:

| 원인 | 대응 | 화면 영향 |
|---|---|---|
| 사진을 PNG로 저장 (수십 MB) | 불투명 확인 후 JPEG q92 변환 | 없음(눈으로 구분 불가) |
| 초고해상도 원본 (6000px+) | 장변 3840px 상한 다운스케일 | 없음(4K 스크린 기준 충분) |
| 예술효과 원본 레이어(.wdp) | ext 블록+rels+파일 세트 제거 | 없음(표시는 메인 blip 사용) |
| 고아 미디어(참조 0) | 삭제 | 없음 |
| 미사용 레이아웃의 배경 | 위 규칙으로 자동 축소됨 | 없음(화면에 안 나옴) |

**절대 하지 않는 것**: 투명도가 실제로 있는 PNG의 JPEG 변환(투명도 소실),
폰트(.fntdata) 제거, 레이아웃 삭제(디자인 자산 보존), 원본 덮어쓰기(사본 생성 방식).

## 실행 흐름

### 1. 분석
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/slim/lib/slim_pptx.py" analyze "<대상.pptx>"
```
- 미디어별 크기·해상도·투명도·참조처·예상 절감량 JSON.
- 사용자에게 요약 보고: 총 용량, 주범 top 3, 예상 절감량, 각각 화면에 나오는 그림인지 여부.

### 2. 보호 목록 결정 + 사용자 확인 (AskUserQuestion)
- **슬라이드가 실제 사용하는 레이아웃**(`used_layouts`)의 미디어는 보수적으로 `--keep` 에 넣는다
  (표지·제목 레이아웃의 배경 이미지처럼 매주 화면에 나오는 것).
- 진행 여부를 확인받고 시작. 원본은 절대 덮어쓰지 않고 사본을 만든다고 안내.

### 3. 적용
```bash
python3 "${CLAUDE_PLUGIN_ROOT}/skills/slim/lib/slim_pptx.py" apply "<대상.pptx>" --keep "<보호목록>"
```
- `<대상>.슬림.pptx` 사본 생성. 변환 내역과 전후 용량이 JSON으로 나온다.
- 절감이 30% 미만인 변환은 스크립트가 자동으로 원본을 유지한다.

### 4. 3중 검증 (모두 통과해야 교체 제안)
1. **구조**: `python3 "${CLAUDE_PLUGIN_ROOT}/skills/ppt/lib/validate.py" "<슬림.pptx>"` — 8게이트 전부 ok.
2. **열림**: PowerPoint AppleScript로 열어 슬라이드 수 확인(복구창 없이 열리는지).
   ```applescript
   tell application "Microsoft PowerPoint"
     activate
     open POSIX file "<슬림.pptx 절대경로>"
     delay 5
     return count of slides of active presentation
   end tell
   ```
   시간초과(-1712)/오류(-9074)가 나면 남은 대화상자가 원인일 수 있으니 PowerPoint 를
   완전히 종료(`pkill -x "Microsoft PowerPoint"`)하고 재시도. scratchpad 경로는
   샌드박스 때문에 못 여니 **대상 파일이 있던 폴더에서** 연다.
3. **픽셀**: 변환된 이미지 2~3장을 원본(zip에서 추출)과 나란히 `magick`으로 축소해
   Read 로 직접 눈 비교. (PSNR 수치는 IM7 출력 파싱이 불안정하니 눈 비교가 기준.)

### 5. 교체 + 보고
- 사용자 확인 후: 원본을 `<이름>.원본백업.pptx` 로 이동하고 슬림판을 원본 이름으로 교체.
- 보고: 전후 용량, 변환 내역 요약, 백업 경로, "화질 문제 시 백업으로 복원" 안내.
- 중간 산출물 정리는 사용자에게 물어본 뒤 수행.

## 주의사항 (재발 방지 지식)

- **미디어 파일만 삭제 금지** — .rels 가 가리키는 파일이 없어지면 그 즉시 PowerPoint 복구창.
  반드시 rels·[Content_Types] 까지 세트로 정리 (스크립트가 처리함).
- 재압축 zip 은 **디렉터리 엔트리 없이** 만들어야 함(있으면 content-types 게이트 실패).
  `zip -r` 을 쓸 거면 `-D` 필수. 스크립트는 python zipfile 로 안전하게 처리.
- PowerPoint 의 PDF 저장은 조용히 실패할 때가 있다 — 파일 존재를 반드시 확인하고,
  실패해도 열림 확인(슬라이드 수)으로 갈음 가능.
- 주간 완성본(자운시)이 큰 원인이 마스터라면 이 스킬로 **마스터를 고치는 게 정답**
  (빌더에 매번 prune 로직을 넣는 것보다 소스를 고치는 쪽; 사용자 결정 사항).
