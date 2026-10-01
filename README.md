# 비고컴퍼니 매출현황 (bgo-sales-dashboard)

비고컴퍼니 7개 지점의 월간 총매출 Excel 파일을 자동으로 읽고 취합하여
**"비고컴퍼니 통합 매출현황"** 웹 대시보드를 만드는 프로젝트입니다.

> 이 저장소는 상진 대표님의 기존 프로젝트(`koc9262-jpg/bgocompany-sales`)와
> 완전히 별개의 독립 프로젝트입니다. 기존 저장소는 일절 수정하지 않았으며,
> 정보구조/사용성 참고 목적으로만 확인했습니다.

배포 주소: `https://<내-GitHub-계정>.github.io/bgo-sales-dashboard/`

---

## 1. 프로젝트 목적

- 7개 지점의 총매출 Excel을 `data` 폴더에 넣으면(수동 또는 아래 Drive 연동으로 자동)
- 자동으로 최신 월 파일을 찾아 데이터를 추출/검증/집계하고
- `index.html` 대시보드에 반영합니다.

2026-10부터는 Google Drive 연동(아래 1.1절)으로 매일 자동 수집이 추가되었습니다.
Gmail API(대표 보고 자동발송)는 아직 사용하지 않습니다.

### 1.1 Google Drive 연동 (2026-10~)

사용자의 Google Drive "1. 매출보고서" 폴더 아래 지점별 하위 폴더(1.역삼점~7.삼전점)에
각 지점이 올리는 **"YYMM_b_N_지점명_총매출.xlsx"** 파일(폴더당 1개, 계속 덮어써짐)만
읽습니다. 같은 폴더의 "YYMMDD_..._일일매출.xlsx" 파일은 무시합니다.

- 지점별 Drive 폴더 ID는 `src/drive_config.py`에서 한 곳으로 관리합니다. 폴더가
  재생성되면 이 파일만 고치면 됩니다.
- 이 연동은 **Claude Code에 연결된 Google Drive 커넥터**를 통해서만 동작합니다.
  일반 `python` 스크립트는 Drive API를 직접 호출하지 않으며, `update.py`는 지금처럼
  `data` 폴더만 바라봅니다 — Drive에서 받은 파일을 `data` 폴더에 내려놓는 역할만
  추가된 것입니다.
- 매일 14:00(로컬 시간)에 Claude Code 예약 작업(`bgo-sales-drive-sync`)이 실행되어
  7개 파일을 받고, `update.py` 실행과 원본 대조 검증까지 통과하면 **로컬 git
  커밋까지만** 수행합니다. **`git push`는 자동으로 하지 않습니다** — 검증 결과를
  알림으로 받은 뒤 "push해줘"라고 확인해야 GitHub/Pages에 반영됩니다.
- 지점 폴더에 13:00까지 파일이 안 올라오는 경우를 대비해 14:00까지 여유를 두었습니다.
  그래도 7개 중 하나라도 없으면 기존 Dashboard는 그대로 유지되고 실패 사실만 알립니다.
(향후 확장 예정 — [8. 향후 확장](#8-향후-확장) 참고)

---

## 2. 폴더 구조

```
bgo-sales-dashboard/
├── data/                 # 매월 Excel 7개를 넣는 곳
├── src/
│   ├── parser.py         # Excel 파일 탐색 및 데이터 추출
│   ├── validator.py      # 데이터 검증 (필수값, 교차검증)
│   ├── aggregator.py     # 전 지점 집계, 순위 계산, JSON 생성
│   └── utils.py          # 지점 설정(BRANCHES) 등 공통 유틸
├── output/
│   └── dashboard_data.json  # 표준 데이터 (대시보드가 이 파일을 읽음)
├── backup/               # index.html 자동 백업
├── index.html            # 대시보드 (순수 HTML/CSS/JS)
├── update.py             # 실행 스크립트 (이것만 실행하면 됨)
├── requirements.txt
└── .github/workflows/    # (필요 시) 배포 자동화
```

역할 분리: Excel 파싱(`parser.py`) / 검증(`validator.py`) / 집계(`aggregator.py`) /
화면 표시(`index.html`)는 서로 독립적으로 동작하며, 중간에 표준 JSON
(`output/dashboard_data.json`)을 거칩니다. 이 JSON은 향후 Gmail 보고,
Google Drive 연동, 모바일 앱에서도 동일하게 재사용할 수 있습니다.

---

## 3. Excel 파일명 규칙

```
{YYMM}_b_{지점번호}_{지점명}점_총매출.xlsx
```

예:
```
2609_b_1_역삼점_총매출.xlsx
2609_b_2_선정릉점_총매출.xlsx
2609_b_3_대치점_총매출.xlsx
2609_b_4_송파점_총매출.xlsx
2609_b_5_장안점_총매출.xlsx
2609_b_6_가락점_총매출.xlsx
2609_b_7_삼전점_총매출.xlsx
```

- `YYMM`은 매월 바뀝니다 (2609, 2610, 2611 …). 특정 월이 코드에 고정되어 있지 않고,
  `data` 폴더 안에서 **가장 최신 월의 파일 7개**를 자동으로 선택합니다.
- 같은 월/같은 지점 파일이 여러 개 있으면 파일명 정확도 → 수정시간 순으로
  자동 선택하고, 화면/터미널에 WARNING을 표시합니다.

## 4. 7개 지점 번호

| 번호 | 지점 |
|---|---|
| 1 | 역삼 |
| 2 | 선정릉 |
| 3 | 대치 |
| 4 | 송파 |
| 5 | 장안 |
| 6 | 가락 |
| 7 | 삼전 |

지점 설정은 `src/utils.py`의 `BRANCHES` 한 곳에서만 관리합니다.

---

## 5. 매월 업데이트 방법

1. `data` 폴더에 이번 달 총매출 Excel 7개를 저장합니다.
2. 터미널(명령 프롬프트)에서 프로젝트 폴더로 이동합니다.
3. 아래 명령을 실행합니다.

   ```bash
   python update.py
   ```

4. 실행 결과에서 아래 문구를 확인합니다.

   ```
   7 / 7 FILES OK
   ...
   UPDATE COMPLETE
   ```

5. 문제가 없으면 `git add`, `git commit`, `git push`로 GitHub에 반영합니다.
   (Pages는 몇 초~몇 분 내에 자동 갱신됩니다.)

---

## 6. 오류가 발생했을 때

`update.py`는 오류가 있으면 **조용히 종료되지 않고** 화면에 아래처럼
무엇이 문제인지 표시합니다.

```
==================================================
UPDATE FAILED
==================================================

ERROR:

삼전점 파일에서 PT 현재매출 데이터를 찾을 수 없습니다.

FILE:
2609_b_7_삼전점_총매출.xlsx

Dashboard는 업데이트하지 않았습니다.
기존 정상 Dashboard는 유지됩니다.
==================================================
```

이 경우 **기존 대시보드는 그대로 유지**되며, `index.html`이 잘못된 숫자로
덮어써지지 않습니다. 메시지에 나온 지점/파일을 확인한 뒤 다시 실행하면 됩니다.

> 만약 터미널에 한글이 깨져 보인다면, 실행 전 `chcp 65001`을 입력해
> 콘솔 인코딩을 UTF-8로 바꿔주세요. (결과 파일 자체는 항상 정상 UTF-8입니다.)

검증 통과 후 정상적으로 갱신되기 직전에는 기존 `index.html`이
`backup/index_YYYYMMDD_HHMMSS.html`로 자동 백업됩니다.

---

## 7. GitHub Pages 주소

```
https://<내-GitHub-계정>.github.io/bgo-sales-dashboard/
```

`main` 브랜치 루트에서 `index.html`을 바로 서비스하도록 설정되어 있습니다.
별도의 빌드 과정이 없으므로 `git push` 후 Pages가 자동으로 최신 내용을 반영합니다.

---

## 8. 향후 확장 (현재 버전에는 없음)

현재 1차 버전에서는 아래 기능을 구현하지 않았습니다.

- Google Drive API 연동 (지점별 Excel 자동 업로드/수신)
- Gmail API 연동 (대표 보고 자동 발송)
- OAuth 인증, 대표 이메일 주소 저장

추후 아래 흐름으로 확장할 수 있도록 Excel 파싱 로직을 `index.html`과
분리해 두었습니다.

```
Google Drive → Excel 자동 업로드 → 최신 파일 자동 감지
→ Parser → Validator → Aggregator → dashboard_data.json
→ Web Dashboard / Gmail 대표 보고
```

---

## 9. 현재 상태에 대한 안내

`src/parser.py`의 `extract_branch_data()` / `extract_pt_ranking()`은
**실제 지점 Excel 원본 파일을 분석하기 전까지는 셀 위치를 임의로
추측하지 않기 위해 의도적으로 미구현 상태**입니다. 실제 Excel 7개가
`data` 폴더에 제공되면 구조를 분석하여 이 부분을 완성합니다.
