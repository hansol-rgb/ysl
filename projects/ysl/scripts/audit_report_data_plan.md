# Audit Report Data Extraction Plan

## Context

`projects/ysl/context/Audit_Report_Template.pdf` 가 정의한 17페이지짜리 슬라이드 템플릿에 들어갈 데이터를, `projects/ysl/report/` 의 raw (VIVI Monitoring 4시트) 에서 추출해 **하나의 xlsx** 로 형식화한다. **PDF의 sub-block 한 개 = xlsx 시트 한 장**. 다음 사이클 raw가 들어왔을 때 같은 빌더로 같은 시트 구조의 새 xlsx 가 재생성되어 PPT 슬라이드를 같은 자리/같은 형식으로 갱신할 수 있게 만드는 것이 목표.

원본 PDF 본문 텍스트와 페이지별 PNG 렌더본은 `/tmp/audit_pdf/`(audit.txt, page-01.png ~ page-17.png, hires-{11,14,15,16}-*.png)에 캐시됨.

---

## 1. 정확도 / 한계

PDF 만 보고 설계했기 때문에 다음 항목은 PPT 와 cross-check 후 1차 보정 필요:

- 표 행/열 방향 (PPT가 transpose 했을 가능성, 신뢰도 ~75%)
- 경쟁사 표시 개수 (raw 12 vs PDF 예시 8개 — 모두 노출이 기본)

데이터 자체는 raw에서 정직하게 집계하므로 시트명/컬럼명 보정만으로 빠르게 수정 가능.

## 1-1. 용어 매핑 (PDF/PPT ↔ Raw)

### 핵심 차원
- **Topic** (Topic 1/2/3) = `Category` (향수 / 기프팅 / 쿠션)
- **Brand** (자사 브랜드) = `YSL Beauty 언급` 컬럼
- **Competitors** (A, B, …) = Dior / Chanel / Hera / MAC / Jo Malone / Nars / Estee Lauder / Lancome / Tom Ford / Sulwhasoo / Prada Beauty (총 11개)
- **AI Engine / AI Platform** = `Channel` (`chatgpt` / `google` / `naver`) — Citation 시트는 `Platform`

### Funnel 5단계 (응답 단위)
- **Query Set**: Mention 시트 전체 행수
- **AI Existence**: `Has AI Overview == 'Yes'`
- **Commercial**: 12 브랜드 중 1개라도 `{Brand} 언급 == 'Y'`
- **Mention**: `YSL Beauty 언급 == 'Y'`
- **Citation**: 해당 응답에 yslbeauty 도메인 URL이 1개라도 인용된 경우 (Citation 시트 `Domain LIKE '%yslbeauty%'` 의 (Reference ID, Cycle, Channel) 매칭)

### Position / Intent (PDF 안에서 Intent Type 두 의미 충돌 정리)
- **Position** (1-2, 3-4) = raw Axis 번호 (Position 1 = Axis 1)
- **Position 라벨 / 축 이름** = 사용상황/TPO (Axis 1) / 소비자 프로필 (Axis 2) / 구매 허들 (Axis 3)
- **Intent Type** (1-2 / 3-4) = raw Axis 컬럼의 **셀 값** (예: "① 연령/취향/입문")
- **Intent (3-3 Customer Decision Making Journey)** = `Intent` 컬럼 (니즈 인식 / 정보 탐색 / 대안 비교 / 구매 결정)

### 도메인 / 자사몰 / E-commerce
- **자사 도메인 / 자사몰** = `Domain` LIKE `%yslbeauty%` (raw 7개 도메인: yslbeautykr/yslbeautyus/yslbeauty/.co.uk/.th/.com.sg/.com.au)
  - raw `YSL 공홈` 컬럼은 일부 누락이 있어 도메인 키워드 매칭이 더 정확
- **경쟁사 자사몰 매핑**

| Brand | 자사몰 도메인 키워드 |
|---|---|
| Dior | dior.com |
| Chanel | chanel.com |
| Hera | hera.com |
| MAC | maccosmetics |
| Jo Malone | jomalone |
| Nars | narscosmetics |
| Estee Lauder | esteelauder |
| Lancome | lancome |
| Sulwhasoo | sulwhasoo.com |
| Tom Ford | tomford (raw 미등장 → 0) |
| Prada Beauty | prada (raw 미등장 → 0) |

- **E-commerce 리테일러 화이트리스트**

| 리테일러 | 매칭 도메인 |
|---|---|
| 쿠팡 | coupang.com (+pages/guide.coupang.com) |
| SSG/이마트/신세계몰 | ssg.com 계열 (m-shinsegaemall, emart.ssg, department.ssg, shinsegaemall) |
| 카카오 선물하기 | gift.kakao.com |
| 카카오 쇼핑하우 | shoppinghow.kakao.com (+m.) |
| 올리브영 | oliveyoung.co.kr (+m.) |
| 세포라 | sephora.com (+.hk) |
| 무신사 | musinsa.com |
| 롯데온 | lotteon.com (+story.) |
| 네이버 스마트스토어 | smartstore.naver.com |
| 11번가 | 11st.co.kr (+m./search.) |
| 컬리 | kurly.com |
| SSG 면세점 | ssgdfs.com |
| 롯데 면세점 | lottedfs.com (+m.kor.) |
| G마켓 | gmarket.co.kr (+cjgmarket) |
| 기타 해외/특수 이커머스 | harrods.com, idus.com, kream.co.kr 등 — 그룹 "기타" |

- **Glowpick** (1,241건) = E-commerce 가 아닌 **정보 사이트 (뷰티 큐레이션)** 로 분류
- **자사몰 vs E-commerce 중복 처리**: 자사몰 매칭 우선 → 그 외 E-commerce 매칭

### 페이지 유형 URL 패턴 (3-7)
URL 문자열 키워드 휴리스틱 (1차):
- **검색결과**: `/search`, `/display/`, `category`, `?query=`, `q=`
- **상품 상세 (PDP)**: `/product/`, `/goods/`, `/prd/`, `/p/`
- **기획전 / 매거진**: `/promotion/`, `/magazine/`, `/event/`, `/curation/`, `/story/`, `/ranking/`
- **기타**: 위에 안 잡힘

### Domain Type 분류
PDF의 6분류는 무시하고 **raw `Domain Type` 10개 그대로 사용**:
external_blog / official / news / video / ecommerce / forum / official_blog / social_media / others / wiki

### 경쟁사 그룹
화장품 업계 표준 분류(Prestige vs Mass)에서 우리 11 경쟁사 모두 Prestige 라 그룹 분류 무의미 → **시트 25 의 Group 컬럼 제거**.

### 분석 단위
- **Combined**: Q-side(질문) + KW-side(키워드) 합본 = raw 전체 (Mention 15,120 / Citation 40,089). 1차는 Combined 만, Q/KW 분리는 추후 옵션.

### Raw에 없는 컬럼 처리
- 시트 21/22 `콘텐츠 주제`, `소구 메시지 특징` → **빈 칸** (수동 입력용)
- 시트 24 `인용 특징` → **빈 칸**
- 시트 24 `자사 브랜드 관련` (Y/N) → **자동** (해당 인용이 YSL 언급 응답에서 발생했는지)
- 시트 24 `PDP 예시 URL` → **자동** (해당 리테일러의 PDP 1건 샘플)
- 시트 25 `핵심 인용 콘텐츠 유형` → **자동** (브랜드 자사몰 인용 중 Domain Type 분포 1위 라벨)
- 시트 25 `시사점` → **빈 칸**
- 시트 03 `목적` → `Bubbleshare_YSL_Question_List_수정.xlsx` 의 `Prompt setting` 시트 R28-R31 에서 자동 추출

---

## 2. PDF 17페이지 = 15개 번호 섹션 = **26개 sub-block (시트)**

PDF 의 모든 sub-block 을 1 시트씩 분리:

### Section 1. Analysis Overview (3 sheets)
| # | Sheet | PDF Sec | Page |
|--:|-------|---------|-----:|
| 01 | `01_Data_Overview` | 1-1 | 1 |
| 02 | `02_Position_Intent` | 1-2 | 2 |
| 03 | `03_Category` | 1-3 | 3 |

### Section 2. Overall Analysis Rate (9 sheets)
| # | Sheet | PDF Sec | Page |
|--:|-------|---------|-----:|
| 04 | `04_Overall_Funnel` | 2-1 A | 4 |
| 05 | `05_Channel_Funnel` | 2-1 B | 4 |
| 06 | `06_Category_Funnel` | 2-1 C | 4 |
| 07 | `07_Intent_Funnel` | 2-1 D | 5 |
| 08 | `08_Mention_by_Competitors` | 2-2 | 5–6 |
| 09 | `09_Citation_YSL_Domain` | 2-3 A (자사 도메인) | 6 |
| 10 | `10_Citation_Ecommerce` | 2-3 B (E-Commerce Channel) | 6–7 |
| 11 | `11_Citation_Domain_Rank` | 2-4 A (Chart Table 도메인 순위) | 7 |
| 12 | `12_Citation_DomainType_Share` | 2-4 B (Bar Table 도메인 타입 분포) | 8 |

### Section 3. Topic Analysis (14 sheets)
모든 시트는 `Topic` 컬럼으로 향수/기프팅/쿠션 long-format.

| # | Sheet | PDF Sec | Page |
|--:|-------|---------|-----:|
| 13 | `13_Topic_Funnel_Overall` | 3-1 A | 8–9 |
| 14 | `14_Topic_Funnel_Channel` | 3-1 B | 9 |
| 15 | `15_Topic_Funnel_Intent` | 3-1 C | 9 |
| 16 | `16_Topic_Mention_by_Competitors` | 3-2 | 9–10 |
| 17 | `17_Topic_Customer_Journey` | 3-3 | 10–11 |
| 18 | `18_Topic_Positioning` | 3-4 | 11–12 |
| 19 | `19_Topic_Citation_Domain_Rank` | 3-5 A | 12 |
| 20 | `20_Topic_Citation_DomainType_Share` | 3-5 B | 13 |
| 21 | `21_Topic_YouTube_Top5` | 3-6 (YouTube) | 14 |
| 22 | `22_Topic_Blog_Top5` | 3-6 (Blog) | 14 |
| 23 | `23_Topic_Ecommerce_Pages` | 3-7 A (페이지 유형별 인용) | 14–15 |
| 24 | `24_Topic_Ecommerce_Summary` | 3-7 B (인용 특징 요약) | 15 |
| 25 | `25_Topic_Brand_Citation` | 3-8 A (브랜드 자사몰 Citation 현황) | 16 |
| 26 | `26_Topic_Brand_OwnSite_Pages` | 3-8 B (자사몰 페이지 유형 분석) | 16–17 |

**총 26 시트.**

---

## 3. Raw 데이터 컬럼 인벤토리

`projects/ysl/report/Bubbleshare_YSL_VIVI_Monitoring_0504.xlsx`

### `01. Query List` — 360행
- `Reference ID (Q)` / `Reference ID (KW)` / `Category` / `Sub Category (Seed KW)`
- `Category (Axis 1) - 사용상황/TPO` / `Category (Axis 2) - 소비자 프로필` / `Category (Axis 3) - 구매 허들`
- `Intent` / `Question` / `1:1 Keyword` / `Google MSV` / `Naver MSV`

### `02. Mention` — 15,120행 × 43컬럼
- 메타: `Index, Date, Cycle, Reference ID, Keyword/Query, Category, Sub Category, Question, 1:1 Keyword`
- 핵심 차원: `Brand / Non-Brand`, `Channel` (`google` / `chatgpt` / `naver`), `Has AI Overview` (`Yes`/`No`), `AI Response Text`, `Source URLs/Titles/Domains`
- 브랜드 멘션 12 × 2: `{Brand} 언급` (Y/N) / `{Brand} 언급#` (count)
  - YSL Beauty, Dior, Chanel, Hera, MAC, Jo Malone, Nars, Estee Lauder, Lancome, Tom Ford, Sulwhasoo, Prada Beauty

### `03. Citation` — 40,089행 × 29컬럼
- 메타: `Index, Date, Cycle, Reference ID, Keyword/Query, Category, Sub Category, Intent, Platform, Query, Rank`
- URL/도메인: `Title, URL, Domain, Domain Type, Content Type` (+ CI 신뢰도)
- 메타데이터: `Published Date, Publisher Name, Meta Author/Site Name/Language/Description/OG Type/Keywords`
- 자사 플래그: `YSL 공홈` (Y/N) — BackOffice 분류 결과

### `00. Keywords&MSV` — 360행
- `No., Keyword, Google MSV, Naver MSV`

### Reference ID prefix → Category
- `PF-` 향수 / `GF-` 기프팅 / `CS-` 쿠션
- `PF-001` 형식이면 Question, `PF-KW-001` 형식이면 Keyword

---

## 4. Funnel 5단계 정의 (PDF 2-1 A 기준)

| 단계 | 정의 | 집계 |
|-----|------|------|
| **Query Set** | 수집된 응답 전체 | `len(Mention)` |
| **AI Existence** | AI 답변이 생성된 응답 | `Has AI Overview == 'Yes'` |
| **Commercial** | 12개 브랜드 중 하나라도 언급된 응답 | `any({Brand} 언급 == 'Y')` |
| **Mention** | YSL Beauty 가 언급된 응답 | `YSL Beauty 언급 == 'Y'` |
| **Citation** | YSL 공홈이 1개라도 인용된 응답 | Citation 시트 `YSL 공홈 == 'Y'` 의 (Reference ID, Cycle, Channel) 매칭 |

---

## 5. 시트별 Row × Column 스키마 (26시트)

각 시트 첫 행에 메타 4줄(Section / Source / Period / Brand) → 빈 행 → 헤더 → 데이터.

---

### `01_Data_Overview` (1-1)
| Metric | Value |
|--------|-------|
| Unique Queries | 360 |
| Total Responses | 15120 |
| Total Citations | 40089 |
| Channels | ChatGPT, Google AIO, Naver |
| Responses Per Query (avg) | Total Responses / Unique Queries |
| Analysis Period | min(Date) ~ max(Date) |
| Brand | YSL Beauty |
| Cycles | unique Cycle |
| Category Mix (향수 / 기프팅 / 쿠션) | 144 / 144 / 72 |

### `02_Position_Intent` (1-2)
3축 × 각 축 Value × 4 메트릭. long-format.

| Position (Axis) | Intent Type Value | 논브랜드 응답수 | 논브랜드 쿼리수 | 브랜드 응답수 | 브랜드 쿼리수 |
|---|---|---:|---:|---:|---:|
| Axis 1 (사용상황/TPO) | (각 Value) | … | … | … | … |
| Axis 1 | Total | … | … | … | … |
| Axis 2 (소비자 프로필) | … | | | | |
| Axis 3 (구매 허들) | … | | | | |
| Grand Total | | … | … | … | … |

집계: Mention 시트 `Brand / Non-Brand` 컬럼 + Reference ID join Query List Axis 컬럼.

### `03_Category` (1-3)
| Category | 목적 | 질문 수 | 키워드 수 | 응답 수 | 비율 |
|---|---|---:|---:|---:|---:|
| 향수 | (template/audit_report_meta.json) | 144 | 144 | 6048 | 40% |
| 기프팅 | … | 144 | 144 | 6048 | 40% |
| 쿠션 | … | 72 | 72 | 3024 | 20% |

---

### `04_Overall_Funnel` (2-1 A)
| Stage | Count | Rate |
|---|---:|---:|
| Query Set | 15120 | 1.000 |
| AI Existence | 9869 | 0.6527 |
| Commercial | 5897 | 0.3900 |
| Mention | 1370 | 0.0906 |
| Citation | (count) | (rate) |

### `05_Channel_Funnel` (2-1 B)
| Channel | Stage | Count | Rate |
|---|---|---:|---:|
| ChatGPT | Query Set | … | 1.000 |
| ChatGPT | AI Existence | … | … |
| … | … | … | … |
| Google AIO | … | | |
| Naver AI Briefing | … | | |

### `06_Category_Funnel` (2-1 C)
| Category | Stage | Count | Rate |
|---|---|---:|---:|
| 향수 | Query Set | … | 1.000 |
| … | … | | |

### `07_Intent_Funnel` (2-1 D)
| Intent | Stage | Count | Rate |
|---|---|---:|---:|
| 니즈 인식 | Query Set | … | 1.000 |
| 정보 탐색 | … | | |
| 대안 비교 | … | | |
| 구매 결정 | … | | |

---

### `08_Mention_by_Competitors` (2-2)
| AI Engine | Questions | Total Brand Mentions | Brand | Mention Count | Rate |
|---|---:|---:|---|---:|---:|
| ChatGPT | 5040 | … | YSL Beauty | … | … |
| ChatGPT | 5040 | … | Dior | … | … |
| ChatGPT | … | … | (각 브랜드 12개) | … | … |
| ChatGPT | 5040 | … | Grand Total | … | … |
| Google AIO | … | … | … | … | … |
| Naver AI | … | … | … | … | … |
| Grand Total | 15120 | … | … | … | … |

`Rate = Mention Count / Questions`. Brand 12개 + Total.

---

### `09_Citation_YSL_Domain` (2-3 A)
| No. | AI Platform | Domain | URL | 인용 수 |
|---:|---|---|---|---:|
| 1 | … | yslbeautykr.com | (구체 URL) | … |

→ Citation `YSL 공홈 == 'Y'` 필터, (Platform, Domain, URL) groupby.

### `10_Citation_Ecommerce` (2-3 B)
| No. | Channel | Target URL | 페이지 설명 | 인용 여부 |
|---:|---|---|---|---|
| 1 | Lotte on | … | (auto: 검색결과/PDP/기획전) | Y/N |
| 2 | Naver Brand Store | … | … | … |
| 3 | Olive Young | … | … | … |

→ template 메타의 e-commerce 도메인 매칭 + URL 패턴으로 페이지 유형 자동 분류.

---

### `11_Citation_Domain_Rank` (2-4 A)
| Rank | AI Platform | Domain | Domain Type | Number of Citation | Citation Rate |
|---:|---|---|---|---:|---:|
| 1 | ChatGPT | blog.naver.com | external_blog | … | … |
| 2 | … | … | … | … | … |
| (Top N: 전체 + Platform별) | | | | | |

→ Citation 시트 (Platform, Domain, Domain Type) groupby 카운트 내림차순.

### `12_Citation_DomainType_Share` (2-4 B)
| AI Platform | Domain Type Group | Citation Count | Share % |
|---|---|---:|---:|
| ChatGPT | 일반 웹사이트 | … | 25% |
| ChatGPT | 네이버 블로그·티스토리 | … | 59% |
| ChatGPT | 포럼 | … | 1% |
| ChatGPT | 비디오 | … | … |
| ChatGPT | 뉴스 | … | … |
| ChatGPT | 기타 | … | 7% |
| Naver AI | … | … | … |
| Google AIO | … | … | … |
| Total | … | … | … |

→ raw `Domain Type` 값을 6개 그룹으로 매핑(template 메타).

---

### `13_Topic_Funnel_Overall` (3-1 A)
| Topic | Stage | Count | Rate |
|---|---|---:|---:|
| 향수 | Query Set | … | 1.000 |
| 향수 | AI Existence | … | … |
| 향수 | Commercial | … | … |
| 향수 | Mention | … | … |
| 향수 | Citation | … | … |
| 기프팅 | … | | |
| 쿠션 | … | | |

### `14_Topic_Funnel_Channel` (3-1 B)
| Topic | Channel | Stage | Count | Rate |
|---|---|---|---:|---:|

### `15_Topic_Funnel_Intent` (3-1 C)
| Topic | Intent | Stage | Count | Rate |
|---|---|---|---:|---:|

---

### `16_Topic_Mention_by_Competitors` (3-2)
`08` 와 동일 구조에 `Topic` 컬럼 prepend.

| Topic | AI Engine | Questions | Total Brand Mentions | Brand | Mention Count | Rate |
|---|---|---:|---:|---|---:|---:|

---

### `17_Topic_Customer_Journey` (3-3)
| Topic | Intent | Brand | Mention Count | Mention Rate |
|---|---|---|---:|---:|
| 향수 | 니즈 인식 | YSL Beauty | … | … |
| 향수 | 니즈 인식 | Dior | … | … |
| … | (4 Intent × 12 Brand × 3 Topic) | | | |

분모 = (Topic, Intent) 의 Mention 행수. 분자 = `{Brand} 언급 == 'Y'` 카운트.

---

### `18_Topic_Positioning` (3-4)
| Topic | Position (Axis) | Axis Value | Brand | 전체 응답 수 | Brand Mention 수 | Mention Rate |
|---|---|---|---|---:|---:|---:|
| 향수 | Axis 1 (사용상황/TPO) | (Value) | YSL Beauty | … | … | … |
| 향수 | Axis 1 | (Value) | Dior | … | … | … |
| … | | | | | | |

(Axis 1/2/3 × Axis Value × Brand 12) per Topic 3.

---

### `19_Topic_Citation_Domain_Rank` (3-5 A)
| Topic | Rank | AI Platform | Domain | Domain Type | Number of Citation | Citation Rate |

### `20_Topic_Citation_DomainType_Share` (3-5 B)
| Topic | AI Platform | Domain Type Group | Citation Count | Share % |

---

### `21_Topic_YouTube_Top5` (3-6 YouTube)
| Topic | Rank | Title | URL | 인용 수 | 콘텐츠 주제 | 소구 메시지 특징 |
|---|---:|---|---|---:|---|---|
| 향수 | 1 | … | … | … | (빈 칸 / LLM 추출) | (빈 칸 / LLM 추출) |
| 향수 | 2 | … | … | … | … | … |

→ Citation 필터: Domain Type 이 `youtube` 계열 + Topic. URL groupby 카운트 내림차순 Top 5.

### `22_Topic_Blog_Top5` (3-6 Blog)
| Topic | Rank | Title | URL | 인용 수 | 콘텐츠 주제 | 소구 메시지 특징 |

→ Domain Type 이 `external_blog` / `네이버 블로그·티스토리` 계열 + Topic.

---

### `23_Topic_Ecommerce_Pages` (3-7 A)
| Topic | 리테일러 | 검색결과 페이지 | 상품 상세 (PDP) | 기획전/매거진 | 합계 |
|---|---|---:|---:|---:|---:|
| 향수 | 올리브영 | … | … | … | … |
| 향수 | 롯데온 | … | … | … | … |
| 향수 | 네이버 브랜드스토어 | … | … | … | … |
| 향수 | 신세계몰 | … | … | … | … |
| 향수 | … | | | | |
| 기프팅 | … | | | | |
| 쿠션 | … | | | | |

→ Citation 도메인 매칭 + URL 패턴(`/display/`, `/product/`, `/promotion/` 등) 분류.

### `24_Topic_Ecommerce_Summary` (3-7 B)
| Topic | 리테일러 | 주요 인용 페이지 유형 | 인용 특징 | 자사 브랜드 관련 (Y/N) | PDP 예시 URL |

`인용 특징` 은 1차 룰 기반 + 빈 칸 (수동 보완용).

---

### `25_Topic_Brand_Citation` (3-8 A)
| Topic | Brand | Group | Mention 건수 | 자사몰 Citation | Citation Rate | 핵심 인용 콘텐츠 유형 | 시사점 |
|---|---|---|---:|---:|---:|---|---|
| 향수 | YSL Beauty | Demo | … | … | … | (auto + 빈 칸) | (빈 칸) |
| 향수 | Dior | Demo | … | … | … | … | … |
| 향수 | … | (12 brands) | … | … | … | … | … |

`자사몰 Citation` = 각 브랜드 공식 도메인 매칭 (template/brands_config 활용). Citation Rate = 자사몰 Citation / Mention 건수.

### `26_Topic_Brand_OwnSite_Pages` (3-8 B)
| Topic | Brand | 페이지 유형 | 인용 수 | 비율 | 대표 콘텐츠 예시 (Title) | 대표 URL |
|---|---|---|---:|---:|---|---|
| 향수 | YSL Beauty | Expert Advice | … | … | … | … |
| 향수 | YSL Beauty | PDP | … | … | … | … |
| 향수 | YSL Beauty | 진단형 | … | … | … | … |
| 향수 | Dior | … | | | | |

→ 각 브랜드 자사몰 도메인 필터 + URL 패턴 분류.

---

## 6. 산출물

- **경로**: `projects/ysl/report/Bubbleshare_YSL_Audit_Report_Data_0504.xlsx`
- **빌더**: `projects/ysl/scripts/build_audit_report_data.py`
  - 입력 raw: `projects/ysl/report/Bubbleshare_YSL_VIVI_Monitoring_0504.xlsx`
  - 입력 메타: `projects/ysl/template/brands_config.json`, `projects/ysl/template/audit_report_meta.json` (신규 — Category 목적 / E-commerce 도메인 매핑 / Domain Type 6분류 그룹핑 / 자사몰 페이지 유형 룰)
  - 출력: 위 xlsx (26 시트)
  - suffix mirroring: 입력 `_0504` → 출력 `_0504`

---

## 7. Builder 구조 (개요)

```
build_audit_report_data.py
├── load_raw(monitoring_xlsx) -> {query_list, mention, citation, kw_msv}
├── load_meta(template_dir)   -> {brand_config, ecom_domains, axis_labels, domain_type_groups, category_purpose, brand_official_domains, page_type_patterns}
├── compute_funnel(df, group_keys=()) -> rows of (Stage, Count, Rate)
├── compute_competitive_matrix(df, dim_col=None) -> matrix
├── build_sheets()
│   ├── s01_data_overview()
│   ├── s02_position_intent()
│   ├── s03_category()
│   ├── s04_overall_funnel()
│   ├── s05_channel_funnel()
│   ├── s06_category_funnel()
│   ├── s07_intent_funnel()
│   ├── s08_mention_by_competitors()
│   ├── s09_citation_ysl_domain()
│   ├── s10_citation_ecommerce()
│   ├── s11_citation_domain_rank()
│   ├── s12_citation_domain_type_share()
│   ├── s13_topic_funnel_overall()
│   ├── s14_topic_funnel_channel()
│   ├── s15_topic_funnel_intent()
│   ├── s16_topic_mention_by_competitors()
│   ├── s17_topic_customer_journey()
│   ├── s18_topic_positioning()
│   ├── s19_topic_citation_domain_rank()
│   ├── s20_topic_citation_domain_type_share()
│   ├── s21_topic_youtube_top5()
│   ├── s22_topic_blog_top5()
│   ├── s23_topic_ecommerce_pages()
│   ├── s24_topic_ecommerce_summary()
│   ├── s25_topic_brand_citation()
│   └── s26_topic_brand_ownsite_pages()
├── write_xlsx(output_path, sheets)
└── main(input_xlsx, output_xlsx)
```

기존 `build_visibility_report_v2.py` 의 funnel/competitive 집계 함수 재사용 가능하면 import.

---

## 8. 검증 방법

1. `python3 projects/ysl/scripts/build_audit_report_data.py` 실행 → 26 시트 모두 채워졌는지 확인.
2. **교차 검증** (`Bubbleshare_YSL_AI_Visibility_Report_0504.xlsx` 와 대조):
   - `04_Overall_Funnel` 5행 ↔ Visibility Report `02. Funnel (Combined)` Section A
   - `05_Channel_Funnel` ↔ `02. Funnel (Combined)` Section B
   - `08_Mention_by_Competitors` × ChatGPT × YSL Beauty Rate ↔ `03. Competitive (Combined)` Section A
   - `09_Citation_YSL_Domain` 합계 ↔ `04. Content (Combined)` Section A
3. PDF 페이지 1, 4, 5, 6 의 핵심 수치 5~10개를 샘플링 → xlsx 셀 값 대조.
4. 다음 사이클 raw (`_0606.xlsx`) 들어왔을 때 같은 스크립트로 빌드 → 시트 26개·시트명·헤더 변동 없음 확인.
5. PPT 옆에 두고 시트별 행/열 방향, 컬럼명 라벨링 보정 (1차 빌드 후).

---

## 9. 후속 / 추후 보강

- 1차 빌드 후 PPT 슬라이드와 셀 단위 cross-check → 행/열 방향, 컬럼 순서 보정
- Q-side / KW-side 분리 시트는 1차 빌드 검증 통과 후 옵션으로 추가
- 자사몰 페이지 유형 분류 (Expert Advice / 진단형 / PDP) 룰은 raw URL 샘플링 후 `template/audit_report_meta.json` 에 보강
- 시트 첫 4행 메타 라인(Section/Source/Period/Brand) 포함 — PPT 빌더에서 `header_row=5` 명시적 지정으로 충돌 회피
