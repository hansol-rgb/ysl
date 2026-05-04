# CLAUDE.md

이 저장소는 BubbleShare가 진행 중인 **YSL Beauty Korea GEO(Generative Engine Optimization) 컨설팅 프로젝트의 산출물 공유본**입니다. AI 검색 엔진(ChatGPT, Google AI Overview, Naver 등)에서 YSL 브랜드 가시성을 어떻게 분석했는지, 어떤 데이터를 추출했는지 함께 검토할 수 있도록 정리되어 있습니다.

원본 작업 저장소(raw 데이터 전체 + 작업 히스토리 포함)는 별도로 운영되며, 이 저장소는 **검토·논의용 스냅샷**입니다.

## 빠른 시작 — 데이터 어떻게 보면 되나요

가장 먼저 볼 파일 3개:

1. **`projects/ysl/final/Bubbleshare_YSL_Question_List_수정.xlsx`** — 분석에 사용된 질문 360개와 시드 키워드 120개 (6시트 구조)
2. **`projects/ysl/final/Bubbleshare_YSL_VIVI_Monitoring_0504.xlsx`** — AI 응답 raw 통합본 (15,120 응답 / 40,089 인용)
3. **`projects/ysl/final/Bubbleshare_YSL_AI_Visibility_Report_0504.xlsx`** — 위 raw에서 도출한 가시성 분석 (10시트, 한국어 인사이트 포함)

## 프로젝트 한눈에

- **분석 카테고리 3개**: 향수(Perfume) / 기프팅(Gifting) / 쿠션(Cushion)
- **인텐트 4단계**: 니즈 인식 → 정보 탐색 → 대안 비교 → 구매 결정
- **3축 Intent Type**:
  - 축1: 사용 상황/TPO (계절·기념일·선물 상황)
  - 축2: 소비자 프로필 (연령·취향·피부타입·남성 구매자)
  - 축3: 구매 허들 (성능·브랜드 비교·가격·채널·실패 방지·형태·유지비)
- **경쟁사 11개**: Dior, Chanel, Hera, MAC, Jo Malone, Nars, Estee Lauder, Lancome, Tom Ford, Sulwhasoo, Prada
- **데이터 채널**: ChatGPT (질문 기반) + Google AI Overview·Naver (키워드 기반)

## 핵심 결과 (2026-05-04 시점)

- AI 응답 생성률: **65.3%** (ChatGPT 100% / Google 81.2% / Naver 14.6%)
- **YSL Mention Rate (Combined): 9.06%** (1,370 / 15,120)
- 채널별 YSL Mention: Google AIO 13.2% / ChatGPT 12.6% / Naver 1.4%
- Top 경쟁사 평균: Dior 18.8% / Chanel 14.5% / Jo Malone 11.3% / Hera 9.3%
- 카테고리별 YSL Mention: 기프팅 12.0% / 쿠션 9.8% / 향수 5.4%
- **YSL 자사몰 인용: 0.27%** (yslbeauty 도메인 매칭) — Mention 9.06% → Citation 0.27% 누수가 가장 큰 갭

## 디렉토리 구조

```
projects/ysl/
├── final/      # 클라이언트 전달용 최종 산출물
├── scripts/    # 산출물을 만든 빌더 스크립트 (재현 가능)
├── template/   # 인풋 정의 (브랜드 매칭 룰, 프롬프트 등)
└── assets/     # 로고 등 PPT 삽입용 이미지
```

## 산출물 가이드 (`projects/ysl/final/`)

| 파일 | 내용 |
|------|------|
| `Bubbleshare_YSL_Question_List_수정.xlsx` | 6시트: Seed KW(120) / Prompt setting / Questions v1~Final(360) / Questions Final + 1:1 검색 키워드 + Google·Naver MSV |
| `Bubbleshare_YSL_Intent_Category_Research.xlsx` | 카테고리별 포지셔닝 리서치 (커뮤니티 URL + 콘텐츠 요약) |
| `Bubbleshare_YSL_VIVI_Monitoring_0504.xlsx` | 통합 모니터링 raw (4시트, 15,120 응답 / 40,089 인용) |
| `Bubbleshare_YSL_AI_Visibility_Report_0504.xlsx` | AI 가시성 분석 v2 (10시트: Background + Combined / Q-side / KW-side 각 Funnel·Competitive·Content) |
| `Bubbleshare_YSL_AI_Visibility_Report_0504_compare.xlsx` | 휴리스틱 분류 vs BackOffice 분류 비교 |
| `Bubbleshare_YSL_URL_Citation_Check_0504.xlsx` | 특정 URL이 인용에 잡히는지 점검 (6시트) |
| `Bubbleshare_YSL_GEO_고객질문리스트.xlsx` | 클라이언트 전달용 24개 질문 (5개 테마) |
| `ysl_questions.csv` / `ysl_kw.csv` | BackOffice 인풋용 — 질문 360 / 키워드 360 |

## 시트 구조 (자주 보게 되는 것)

### `Bubbleshare_YSL_VIVI_Monitoring_0504.xlsx` (4시트)
- **00. Keywords&MSV** (360행) — 키워드별 Google·Naver 월간 검색량
- **01. Query List** (360행) — 메타 마스터 (Category/Sub Category/Axis/Intent/Question/1:1 KW/MSV)
- **02. Mention** (15,120행) — 응답 long-format. AI 응답 텍스트 + 12개 브랜드 멘션 Y/N + 횟수
- **03. Citation** (40,089행) — 인용 long-format. URL + Domain Type / Content Type + 메타데이터

### `Bubbleshare_YSL_AI_Visibility_Report_0504.xlsx` (10시트)
- **01. Analysis Background** — 표본 크기, 사이클 범위, 분석 기준 요약
- **02~04. Combined** — Q+KW 합본 Funnel / Competitive / Content
- **05~07. Q-side (ChatGPT)** — 질문 채널 한정
- **08~10. KW-side (Google + Naver)** — 키워드 채널 한정

## 분석 방법론 핵심

### 인용 분류 — BackOffice 우선 + 휴리스틱 fallback
0504 버전부터 BackOffice가 인용 URL별로 분류값(Domain Type / Content Type)을 제공합니다. 이 값이 있으면 그대로 쓰고, 없으면 키워드 매칭 휴리스틱으로 채웁니다.

이전 휴리스틱-only 방식과 비교 (compare.xlsx 참고):
- **공식몰 카테고리: 휴리스틱 0.9% → BackOffice 15.3%** (17배 차이, 휴리스틱이 자사·경쟁사 공식몰 인용을 크게 과소평가)
- **"AI/검색" 카테고리는 거의 가짜였음**: 휴리스틱 27.2% → BackOffice 0.8% (chatgpt/perplexity 키워드 다른 맥락 매칭)
- 휴리스틱 "기타" 6,754건 중 97%가 BackOffice로 재분류 가능

### Reference ID 스키마
응답·인용 raw와 메타(Question List)는 Reference ID로 join됩니다.
- 카테고리 prefix: 향수 `PF` / 기프팅 `GF` / 쿠션 `CS`
- 질문: `{prefix}-{NNN:03d}` (예: `PF-001`)
- 키워드: `{prefix}-KW-{NNN:03d}` (예: `PF-KW-019`)

## 산출물 재생성 방법

빌더 스크립트로 raw → 분석 산출물 흐름을 재현할 수 있습니다. 단, 이 저장소엔 raw 응답·인용 CSV가 들어있지 않으므로 재생성은 원본 작업 저장소에서만 가능합니다. 스크립트는 **로직 검토용**으로 포함했습니다.

```bash
# 의존성
python3 -m pip install openpyxl pandas python-pptx

# VIVI Monitoring 통합 xlsx 생성
python3 projects/ysl/scripts/build_vivi_monitoring.py

# AI Visibility Report v2 생성
python3 projects/ysl/scripts/build_visibility_report_v2.py

# 1차 PPT 갱신
python3 projects/ysl/scripts/generate_ppt_v2.py
```

## 컨벤션

- 파일명 prefix: `Bubbleshare_YSL_`
- 파일 버저닝: `_v2`/`_v3` 신규 파일 만들지 않음. 같은 이름 유지, 백업은 `_bak_YYYYMMDD` 1개만
- VIVI / Visibility Report 파일명 suffix: `_N차` 또는 `_MMDD` (현재 latest = `_0504`)
- 빌더는 인풋 suffix를 출력에 자동 미러링 (`_0504.xlsx` 인풋 → `_0504.xlsx` 출력)
- 한국어 기본, 메트릭명(Mention Rate, SOV 등)은 영문 유지
- MSV 기준: Google 또는 Naver 중 하나라도 > 10이면 채택
- 카테고리 비율: 향수 40% / 기프팅 40% / 쿠션 20%
- 질문은 전부 제네릭(브랜드명 배제), 존댓말, ?/. 종결

## 함께 검토할 때 좋은 질문 예시

- "Mention 9.06% → Citation 0.27% 누수 구간이 어디서 발생하나요?" → Visibility Report Funnel 시트
- "경쟁사 대비 YSL 포지션이 어디인가요?" → Visibility Report Competitive 시트
- "어떤 카테고리에 자원을 집중해야 할까요?" → 카테고리별 Mention Rate 비교
- "특정 URL이 실제 인용에 잡히나요?" → URL_Citation_Check 시트

문의·논의 사항은 `hansol@bubbleshare.io`.
