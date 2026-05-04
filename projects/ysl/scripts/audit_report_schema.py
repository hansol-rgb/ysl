"""
Audit Report Schema — single source of truth for the 26 sheets.

이 모듈은 PDF 템플릿 (Audit_Report_Template.pdf) 의 17 페이지 / 15 번호 섹션을
26 개 sub-block (= xlsx 시트 = HTML 슬라이드 = sheets/*.md) 으로 정규화한 메타를 제공한다.

사용처:
- `build_audit_report_data.py`  — schema 의 sheet_id / sheet_name 을 그대로 사용
- `build_audit_report_html.py`  — schema 의 pdf_section / purpose / 헤더 를 슬라이드 메타로 사용
- `build_audit_report_docs.py`  — schema 전체를 sheets/*.md 로 직렬화
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass(frozen=True)
class SheetSpec:
    sheet_id: str                 # "01", "02", ...
    sheet_name: str               # 실제 xlsx 시트명 (Excel 31자 제한 안)
    pdf_section: str              # "1-1. Data Overview"
    pdf_pages: str                # "1" 또는 "5-6"
    title_ko: str                 # 한국어 슬라이드 제목
    purpose: str                  # 이 시트가 PPT 슬라이드에서 답하는 질문 / 의의
    headers: list[str]            # 컬럼 헤더 (long-format)
    dimensions: list[str]         # 어떤 차원으로 쪼개는지 (예: Topic / Channel / Brand)
    data_rules: str               # raw → 시트 변환 룰 (한국어, 멀티라인 OK)
    validation: str               # 검증 포인트 (Visibility Report 등 cross-check 대상)
    blank_columns: list[str] = field(default_factory=list)  # 수동 입력용 빈 컬럼


# ---------------------------------------------------------------------------
# 공통 텍스트
# ---------------------------------------------------------------------------

_FUNNEL_RULE = """\
Funnel 5단계 (응답 단위):
- Query Set: Mention 시트 전체 행수
- AI Existence: `Has AI Overview == 'Yes'`
- Commercial: 12 브랜드 중 1개라도 `{Brand} 언급 == 'Y'`
- Mention: `YSL Beauty 언급 == 'Y'`
- Citation: 해당 응답의 (Reference ID, Cycle, Channel) 쌍이 Citation 시트 `Domain LIKE '%yslbeauty%'` 에 매칭되는 응답"""


# ---------------------------------------------------------------------------
# 26 시트 메타
# ---------------------------------------------------------------------------

SHEETS: list[SheetSpec] = [
    SheetSpec(
        sheet_id="01",
        sheet_name="01_Data_Overview",
        pdf_section="1-1. Data Overview",
        pdf_pages="1",
        title_ko="분석 개요 / 데이터 메타",
        purpose="분석에 사용한 데이터 규모와 메타 정보 (분석 기간, 응답 수, 카테고리 분포)를 한 슬라이드로 요약. 모든 후속 슬라이드의 신뢰도 근거.",
        headers=["Metric", "Value"],
        dimensions=["Metric"],
        data_rules=("""\
- Unique Queries = `01. Query List` 행수
- Total Responses = `02. Mention` 행수
- Total Citations = `03. Citation` 행수
- Channels = `02. Mention.Channel` unique 값
- Responses Per Query (avg) = Total Responses / Unique Queries
- Analysis Period = `02. Mention.Date` min ~ max
- Brand = "YSL Beauty" (constant)
- Cycles = `02. Mention.Cycle` unique 값
- Category Mix (Questions/Responses) = Category 별 카운트
- YSL Own-Site Citation URL count = `03. Citation` 중 `Domain LIKE '%yslbeauty%'` 행수"""),
        validation="Total Responses 가 Visibility Report `02. Funnel (Combined)` Section A 'Query Set' 값과 일치 (15,120).",
    ),
    SheetSpec(
        sheet_id="02",
        sheet_name="02_Position_Intent",
        pdf_section="1-2. Position and Intent Type",
        pdf_pages="2",
        title_ko="포지션 / 인텐트 타입 매트릭스",
        purpose="3 축(사용상황/소비자 프로필/구매 허들)별로 응답 분포를 보여주고, 브랜드/논브랜드 응답 수 / 쿼리 수를 함께 표시.",
        headers=["Position (Axis)", "Intent Type Value",
                 "논브랜드 응답수", "논브랜드 쿼리수",
                 "브랜드 응답수", "브랜드 쿼리수"],
        dimensions=["Position (Axis)", "Intent Type Value"],
        data_rules=("""\
용어 매핑:
- Position = raw Axis 번호 (Position 1 = Axis 1, Position 2 = Axis 2, Position 3 = Axis 3)
- Position 라벨 = Axis 이름 (Axis 1 = 사용상황/TPO, Axis 2 = 소비자 프로필, Axis 3 = 구매 허들)
- Intent Type Value = Query List 의 Axis 컬럼 셀 값 (예: "① 연령/취향/입문")

집계:
- `02. Mention.Reference ID` 로 `01. Query List` 와 join → Axis 1/2/3 컬럼 셀 값 추출
- 각 Axis × Value × {브랜드, 논브랜드} 별 응답 수(행 카운트) / 쿼리 수(unique Reference ID)
- Axis 별 Total + Grand Total 행 추가"""),
        validation="Grand Total `브랜드 응답수 + 논브랜드 응답수` 의 합이 Total Responses (15120) 의 일정 비율 (Axis 가 정의된 행만 포함).",
    ),
    SheetSpec(
        sheet_id="03",
        sheet_name="03_Category",
        pdf_section="1-3. Category",
        pdf_pages="3",
        title_ko="카테고리 (분석 주제) 정의",
        purpose="향수 / 기프팅 / 쿠션 3개 카테고리의 정의, 비중, 질문 수, 응답 수를 한 슬라이드로 정리.",
        headers=["Category", "목적", "질문 수", "키워드 수", "응답 수", "비율"],
        dimensions=["Category"],
        data_rules=("""\
- Category = 향수 / 기프팅 / 쿠션
- 목적 = `Bubbleshare_YSL_Question_List_수정.xlsx` 의 `Prompt setting` 시트 R28-R31 에서 자동 추출
- 질문 수 = `01. Query List.Category` 카운트
- 키워드 수 = `01. Query List.Reference ID (KW)` 가 있는 행 카운트 (카테고리별)
- 응답 수 = `02. Mention.Category` 카운트
- 비율 = 질문 수 / 360"""),
        validation="질문 수 합 = 360 / 비율 합 = 1.0",
    ),
    SheetSpec(
        sheet_id="04",
        sheet_name="04_Overall_Funnel",
        pdf_section="2-1 A. Overall Funnel",
        pdf_pages="4",
        title_ko="전체 Funnel (5단계)",
        purpose="전체 데이터에서 Query → AI Existence → Commercial → Mention → Citation 5단계 누수 구간 시각화. YSL 가시성 핵심 지표.",
        headers=["Stage", "Count", "Rate"],
        dimensions=["Stage"],
        data_rules=_FUNNEL_RULE,
        validation="Visibility Report `02. Funnel (Combined)` Section A 와 1:1 일치 — Query Set 15120 / AI Existence 9869 (65.27%) / Commercial 5897 (39.00%) / Mention 1370 (9.06%) / Citation 41 (0.27%).",
    ),
    SheetSpec(
        sheet_id="05",
        sheet_name="05_Channel_Funnel",
        pdf_section="2-1 B. Channel Funnel",
        pdf_pages="4",
        title_ko="채널별 Funnel — AI 플랫폼별 가시성",
        purpose="ChatGPT / Google AIO / Naver AI 3개 채널 각각의 Funnel 비교. 어느 채널에서 누수가 큰지.",
        headers=["Channel", "Stage", "Count", "Rate"],
        dimensions=["Channel", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n그룹: `02. Mention.Channel` ∈ {chatgpt, google, naver}",
        validation="채널별 Query Set 합 = 15120. 채널별 Mention Rate 가 ChatGPT 12.6% / Google 13.2% / Naver 1.4% 와 일치.",
    ),
    SheetSpec(
        sheet_id="06",
        sheet_name="06_Category_Funnel",
        pdf_section="2-1 C. Category Funnel",
        pdf_pages="4",
        title_ko="카테고리별 Funnel",
        purpose="향수 / 기프팅 / 쿠션 3개 카테고리 각각의 Funnel 비교. 어느 카테고리에서 YSL 약세인지.",
        headers=["Category", "Stage", "Count", "Rate"],
        dimensions=["Category", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n그룹: `02. Mention.Category` ∈ {향수, 기프팅, 쿠션}",
        validation="카테고리별 Mention Rate: 기프팅 12.0% / 쿠션 9.8% / 향수 5.4% (CLAUDE.md).",
    ),
    SheetSpec(
        sheet_id="07",
        sheet_name="07_Intent_Funnel",
        pdf_section="2-1 D. Intent Funnel",
        pdf_pages="5",
        title_ko="구매 여정 Intent 4단계별 Funnel",
        purpose="니즈 인식 → 정보 탐색 → 대안 비교 → 구매 결정 단계별 가시성. 어느 여정 단계에서 YSL 노출이 약한지.",
        headers=["Intent", "Stage", "Count", "Rate"],
        dimensions=["Intent", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n그룹: `02. Mention.Reference ID` 로 `01. Query List` 와 join → `Intent` ∈ {니즈 인식, 정보 탐색, 대안 비교, 구매 결정}",
        validation="Intent 4단계 Query Set 합이 15120 의 약 1.0 비율.",
    ),
    SheetSpec(
        sheet_id="08",
        sheet_name="08_Mention_by_Competitors",
        pdf_section="2-2. Overall Mention Rate by Competitors",
        pdf_pages="5-6",
        title_ko="경쟁사 대비 멘션율 (채널별)",
        purpose="3개 채널 × 12개 브랜드 매트릭스. 자사 vs 경쟁사 SOV 한 슬라이드.",
        headers=["AI Engine", "Questions", "Total Brand Mentions",
                 "Brand", "Mention Count", "Rate"],
        dimensions=["AI Engine", "Brand"],
        data_rules=("""\
- AI Engine = `02. Mention.Channel` (ChatGPT / Google AIO / Naver AI Briefing)
- Questions = 채널별 Mention 행수
- Total Brand Mentions = 12 브랜드 중 1개라도 언급된 응답 수
- Brand = YSL Beauty + 11 경쟁사 + Grand Total
- Mention Count = `{Brand} 언급 == 'Y'` 행수
- Rate = Mention Count / Questions
- 채널별 블록 마지막에 'Grand Total' 행, 마지막에 'All' 채널 블록"""),
        validation="ChatGPT × YSL = 12.6% / Google × YSL = 13.2% / Naver × YSL = 1.4%. Dior 채널평균 18.8%, Chanel 14.5%, Jo Malone 11.3%.",
    ),
    SheetSpec(
        sheet_id="09",
        sheet_name="09_Citation_YSL_Domain",
        pdf_section="2-3 A. 자사 도메인 인용",
        pdf_pages="6",
        title_ko="자사 도메인 (YSL 공홈) 인용 현황",
        purpose="yslbeauty 계열 도메인이 어떤 페이지가 어느 채널에서 몇 번 인용됐는지 Top 30. 자사몰 콘텐츠 효과 측정.",
        headers=["No.", "AI Platform", "Domain", "URL", "인용 수"],
        dimensions=["AI Platform", "URL"],
        data_rules=("""\
- 필터: `03. Citation.Domain LIKE '%yslbeauty%'`
- (Platform, Domain, URL) groupby 카운트
- 인용 수 내림차순, Top 30
- AI Platform 라벨: chatgpt → ChatGPT / google → Google AIO / naver → Naver AI Briefing"""),
        validation="총 합계 = `Bubbleshare_YSL_AI_Visibility_Report_0504.xlsx` `04. Content (Combined)` Section A 합계 (61).",
    ),
    SheetSpec(
        sheet_id="10",
        sheet_name="10_Citation_Ecommerce",
        pdf_section="2-3 B. E-Commerce Channel 인용",
        pdf_pages="6-7",
        title_ko="E-Commerce 리테일러 인용 현황",
        purpose="14개 리테일러 (쿠팡/SSG/카카오선물하기/올리브영 등) 각각의 어떤 URL 이 어느 채널에서 인용됐는지. 리테일 채널 영향력.",
        headers=["No.", "AI Platform", "Channel (Retailer)", "Target URL",
                 "페이지 유형", "YSL 언급 응답 (Y/N)", "인용 수"],
        dimensions=["AI Platform", "Channel (Retailer)", "URL"],
        data_rules=("""\
- 필터: yslbeauty 도메인 제외 (자사몰은 시트 09)
- 리테일러 매칭: schema 의 `ECOMMERCE_RETAILERS` 화이트리스트 (Glowpick 은 정보사이트로 제외)
- 페이지 유형: URL 패턴 매칭 (검색결과 / 상품 상세 PDP / 기획전·매거진 / 기타)
- YSL 언급 응답: 해당 인용이 발생한 응답에서 `YSL Beauty 언급 == 'Y'` 였는지
- 인용 수 내림차순"""),
        validation="자사몰 (yslbeauty) 도메인이 결과에 없어야 함. Glowpick 도 없어야 함.",
    ),
    SheetSpec(
        sheet_id="11",
        sheet_name="11_Citation_Domain_Rank",
        pdf_section="2-4 A. Citation Chart Table (도메인 순위)",
        pdf_pages="7",
        title_ko="채널별 인용 도메인 순위 (Top 30)",
        purpose="각 AI 채널이 어떤 도메인을 가장 많이 인용하는지 Top 30. 채널별 정보 출처 패턴.",
        headers=["Rank", "AI Platform", "Domain", "Domain Type",
                 "Number of Citation", "Citation Rate"],
        dimensions=["AI Platform", "Domain"],
        data_rules=("""\
- (Platform, Domain, Domain Type) groupby 카운트
- Platform 별 인용 수 내림차순 → Rank 부여, Top 30
- Citation Rate = Domain 인용 수 / Platform 전체 인용 수"""),
        validation="ChatGPT 1위 도메인이 blog.naver.com 또는 youtube.com 일 가능성 높음 (raw 분포 기준).",
    ),
    SheetSpec(
        sheet_id="12",
        sheet_name="12_Citation_DomainType_Share",
        pdf_section="2-4 B. Citation Bar Table (도메인 타입 분포)",
        pdf_pages="8",
        title_ko="채널별 Domain Type 분포",
        purpose="ChatGPT / Naver / Google 각각 어떤 종류의 콘텐츠 (블로그/영상/뉴스 등)를 선호하는지. 콘텐츠 전략 시사점.",
        headers=["AI Platform", "Domain Type", "Citation Count", "Share"],
        dimensions=["AI Platform", "Domain Type"],
        data_rules=("""\
- Domain Type = `03. Citation.Domain Type` (raw 10종 분류 그대로 사용)
  - external_blog / official / news / video / ecommerce / forum / official_blog / social_media / others / wiki
- (Platform, Domain Type) 카운트
- Share = Type 카운트 / Platform 전체 인용 수
- Total 행: 전체 채널 합산"""),
        validation="Platform 별 Share 합 = 1.0",
    ),
    SheetSpec(
        sheet_id="13",
        sheet_name="13_Topic_Funnel_Overall",
        pdf_section="3-1 A. Topic Overall Funnel",
        pdf_pages="8-9",
        title_ko="카테고리별 전체 Funnel",
        purpose="향수 / 기프팅 / 쿠션 카테고리 각각의 Funnel 5단계. 카테고리별 가시성 절대값.",
        headers=["Topic", "Stage", "Count", "Rate"],
        dimensions=["Topic", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n필터: `02. Mention.Category` 별로 분리",
        validation="Topic 별 Query Set 합 = 15120 (향수 5184 + 기프팅 5184 + 쿠션 4752).",
    ),
    SheetSpec(
        sheet_id="14",
        sheet_name="14_Topic_Funnel_Channel",
        pdf_section="3-1 B. Topic Channel Funnel",
        pdf_pages="9",
        title_ko="카테고리 × 채널 Funnel",
        purpose="향수 ChatGPT, 향수 Google, ... 9개 조합 Funnel. 카테고리×채널 교차 분석.",
        headers=["Topic", "Channel", "Stage", "Count", "Rate"],
        dimensions=["Topic", "Channel", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n필터: (Category, Channel) 조합 9개",
        validation="(Topic, Channel) 별 Query Set 합 = 해당 Topic 의 Query Set.",
    ),
    SheetSpec(
        sheet_id="15",
        sheet_name="15_Topic_Funnel_Intent",
        pdf_section="3-1 C. Topic Intent Funnel",
        pdf_pages="9",
        title_ko="카테고리 × Intent Funnel",
        purpose="향수의 니즈 인식, 향수의 정보 탐색, ... 12개 조합 Funnel. 카테고리×구매여정 교차.",
        headers=["Topic", "Intent", "Stage", "Count", "Rate"],
        dimensions=["Topic", "Intent", "Stage"],
        data_rules=_FUNNEL_RULE + "\n\n필터: Reference ID 로 Query List join → (Category, Intent) 조합",
        validation="(Topic, Intent) 별 Query Set 합 = 해당 Topic 의 Query Set.",
    ),
    SheetSpec(
        sheet_id="16",
        sheet_name="16_Topic_Mention_by_Competitors",
        pdf_section="3-2. Topic Mention Rate by Competitors",
        pdf_pages="9-10",
        title_ko="카테고리 × 채널 × 경쟁사 매트릭스",
        purpose="시트 08 의 카테고리 분리 버전. 향수 ChatGPT 에서 YSL vs Dior, ... 카테고리별 SOV.",
        headers=["Topic", "AI Engine", "Questions", "Total Brand Mentions",
                 "Brand", "Mention Count", "Rate"],
        dimensions=["Topic", "AI Engine", "Brand"],
        data_rules="시트 08 과 동일 로직, 카테고리 필터 추가. 각 카테고리별로 채널 × 브랜드 매트릭스 + Grand Total + All 블록.",
        validation="카테고리별 채널 × Brand 매트릭스 가 시트 08 의 부분합과 일치.",
    ),
    SheetSpec(
        sheet_id="17",
        sheet_name="17_Topic_Customer_Journey",
        pdf_section="3-3. Topic Mention Rate by Customer Decision Making Journey",
        pdf_pages="10-11",
        title_ko="카테고리 × 구매여정 × 브랜드별 멘션율",
        purpose="구매 여정 4단계 (니즈 인식~구매 결정) 에서 카테고리별 12 브랜드 점유. 여정 단계별 SOV 추이.",
        headers=["Topic", "Intent", "Brand", "Mention Count", "Mention Rate"],
        dimensions=["Topic", "Intent", "Brand"],
        data_rules=("""\
- Reference ID 로 Query List join → (Category, Intent) 그룹핑
- 그룹별 분모 = Mention 행수, 분자 = `{Brand} 언급 == 'Y'` 행수
- 출력: 3 Topic × 4 Intent × 12 Brand = 144 행"""),
        validation="(Topic, Intent) 별 12 브랜드 합이 100% 초과 (한 응답에서 여러 브랜드 동시 언급 가능).",
    ),
    SheetSpec(
        sheet_id="18",
        sheet_name="18_Topic_Positioning",
        pdf_section="3-4. Positioning Decision Mention Rate",
        pdf_pages="11-12",
        title_ko="카테고리 × Position(Axis) × Axis Value × 브랜드 멘션율",
        purpose="3 축 × Axis Value × 12 브랜드 매트릭스 (Topic 별). 어떤 상황/프로필/허들에서 어느 브랜드가 강한지.",
        headers=["Topic", "Position (Axis)", "Axis Value", "Brand",
                 "전체 응답 수", "Brand Mention 수", "Mention Rate"],
        dimensions=["Topic", "Position", "Axis Value", "Brand"],
        data_rules=("""\
- Reference ID 로 Query List join → 각 Axis 컬럼의 값 추출
- (Topic, Axis, Value) 그룹핑 → 분모 = 그룹 행수, 분자 = `{Brand} 언급 == 'Y'` 행수
- 출력 행 = 3 Topic × 3 Axis × ~7 Value × 12 Brand ≈ 750"""),
        validation="Axis Value 가 '-' 인 응답은 제외 (Query List 에서 해당 축 적용 안 된 질문).",
    ),
    SheetSpec(
        sheet_id="19",
        sheet_name="19_Topic_Citation_Domain_Rank",
        pdf_section="3-5 A. Topic Citation Chart Table",
        pdf_pages="12",
        title_ko="카테고리 × 채널별 인용 도메인 Top 20",
        purpose="시트 11 의 카테고리 분리 버전. 향수에서 ChatGPT 가 가장 인용한 도메인 Top 20, ...",
        headers=["Topic", "Rank", "AI Platform", "Domain", "Domain Type",
                 "Number of Citation", "Citation Rate"],
        dimensions=["Topic", "AI Platform", "Domain"],
        data_rules="시트 11 동일 로직, Category 필터 추가. (Topic, Platform) 별 Top 20.",
        validation="(Topic, Platform) 별 Citation Rate 합 = Top 20 까지의 누적 비율.",
    ),
    SheetSpec(
        sheet_id="20",
        sheet_name="20_Topic_Citation_DomainType_Sh",
        pdf_section="3-5 B. Topic Citation Bar Table",
        pdf_pages="13",
        title_ko="카테고리 × 채널별 Domain Type 분포",
        purpose="시트 12 의 카테고리 분리 버전. 카테고리별 콘텐츠 유형 선호도 차이 비교.",
        headers=["Topic", "AI Platform", "Domain Type", "Citation Count", "Share"],
        dimensions=["Topic", "AI Platform", "Domain Type"],
        data_rules="시트 12 동일 로직, Category 필터 추가.",
        validation="(Topic, Platform) 별 Share 합 = 1.0",
    ),
    SheetSpec(
        sheet_id="21",
        sheet_name="21_Topic_YouTube_Top5",
        pdf_section="3-6. Topic YouTube Top 5 Citation",
        pdf_pages="14",
        title_ko="카테고리별 YouTube 영상 Top 5 인용",
        purpose="각 카테고리에서 AI 가 가장 많이 인용한 YouTube 영상 Top 5. 영상 콘텐츠 영향력.",
        headers=["Topic", "Rank", "Title", "URL", "인용 수", "콘텐츠 주제", "소구 메시지 특징"],
        dimensions=["Topic", "URL"],
        data_rules=("""\
- 필터: `Domain LIKE '%youtube.com%'`
- (Topic, URL) 카운트, 카테고리별 Top 5
- Title = `03. Citation.Title` 첫 매칭 값"""),
        validation="카테고리당 5 행 = 총 15 행.",
        blank_columns=["콘텐츠 주제", "소구 메시지 특징"],
    ),
    SheetSpec(
        sheet_id="22",
        sheet_name="22_Topic_Blog_Top5",
        pdf_section="3-6. Topic Blog Top 5 Citation",
        pdf_pages="14",
        title_ko="카테고리별 블로그 Top 5 인용",
        purpose="각 카테고리에서 AI 가 가장 많이 인용한 네이버 블로그 / 티스토리 Top 5. 블로그 콘텐츠 영향력.",
        headers=["Topic", "Rank", "Title", "URL", "인용 수", "콘텐츠 주제", "소구 메시지 특징"],
        dimensions=["Topic", "URL"],
        data_rules=("""\
- 필터: `Domain LIKE '%blog.naver.com%' OR '%tistory.com%'`
- (Topic, URL) 카운트, 카테고리별 Top 5"""),
        validation="카테고리당 5 행 = 총 15 행.",
        blank_columns=["콘텐츠 주제", "소구 메시지 특징"],
    ),
    SheetSpec(
        sheet_id="23",
        sheet_name="23_Topic_Ecommerce_Pages",
        pdf_section="3-7 A. Topic E-Commerce 페이지 유형별 인용",
        pdf_pages="14-15",
        title_ko="카테고리 × 리테일러 × 페이지 유형 매트릭스",
        purpose="각 리테일러에서 어떤 페이지 (검색결과/PDP/기획전) 가 인용됐는지 카테고리별 분포. E-commerce 채널 인용 패턴.",
        headers=["Topic", "리테일러", "검색결과", "상품 상세 (PDP)",
                 "기획전/매거진", "기타", "합계"],
        dimensions=["Topic", "리테일러"],
        data_rules=("""\
- 필터: yslbeauty 제외 + ECOMMERCE_RETAILERS 매칭
- 페이지 유형 분류 (URL 패턴):
  - 상품 상세 (PDP): /product/, /goods/, /prd/, /p/
  - 기획전/매거진: /promotion/, /magazine/, /event/, /curation/, /story/, /ranking/
  - 검색결과: /search, /display/, ?query=, ?q=
  - 기타: 위에 안 잡히는 것
- (Topic, 리테일러, 페이지 유형) 카운트
- Wide-format: 페이지 유형이 컬럼"""),
        validation="합계 행이 페이지 유형 4개 합과 일치.",
    ),
    SheetSpec(
        sheet_id="24",
        sheet_name="24_Topic_Ecommerce_Summary",
        pdf_section="3-7 B. Topic E-Commerce 인용 특징 요약",
        pdf_pages="15",
        title_ko="카테고리 × 리테일러 인용 특징 요약",
        purpose="각 리테일러의 주요 인용 페이지 유형, YSL 관련 여부, PDP 예시. 리테일 채널 시사점.",
        headers=["Topic", "리테일러", "주요 인용 페이지 유형", "인용 특징",
                 "자사 브랜드 관련 (Y/N)", "PDP 예시 URL"],
        dimensions=["Topic", "리테일러"],
        data_rules=("""\
- 주요 인용 페이지 유형 = 시트 23 의 페이지 유형 중 1위 (자동)
- 자사 브랜드 관련 (Y/N) = 해당 리테일러 인용이 YSL 언급 응답에서 발생했는지 (자동)
- PDP 예시 URL = 페이지 유형이 'PDP' 인 첫 매칭 URL (자동)
- 인용 특징 = 빈 칸 (수동 입력)"""),
        validation="시트 23 의 (Topic, 리테일러) 와 행 수 일치.",
        blank_columns=["인용 특징"],
    ),
    SheetSpec(
        sheet_id="25",
        sheet_name="25_Topic_Brand_Citation",
        pdf_section="3-8 A. Topic 브랜드 자사몰 Citation 현황",
        pdf_pages="16",
        title_ko="카테고리 × 12 브랜드 자사몰 Citation 현황",
        purpose="각 카테고리에서 12 브랜드의 자사몰이 얼마나 인용됐는지. 자사몰 콘텐츠 효과 비교.",
        headers=["Topic", "Brand", "Mention 건수", "자사몰 Citation",
                 "Citation Rate", "핵심 인용 콘텐츠 유형", "시사점"],
        dimensions=["Topic", "Brand"],
        data_rules=("""\
- Mention 건수 = 카테고리 내 `{Brand} 언급 == 'Y'` 응답 수
- 자사몰 Citation = 브랜드별 공식 도메인 키워드 (BRAND_OWN_DOMAIN_KEYWORDS) 매칭된 인용 수
- Citation Rate = 자사몰 Citation / Mention 건수
- 핵심 인용 콘텐츠 유형 = 자사몰 인용 중 Domain Type 1위 (자동)
- 시사점 = 빈 칸 (수동 입력)
- Tom Ford / Prada Beauty 는 raw 미등장 → 자사몰 Citation 0"""),
        validation="YSL Beauty 자사몰 Citation 합 (3 카테고리) ≈ 시트 09 합계.",
        blank_columns=["시사점"],
    ),
    SheetSpec(
        sheet_id="26",
        sheet_name="26_Topic_Brand_OwnSite_Pages",
        pdf_section="3-8 B. Topic 브랜드 자사몰 페이지 유형 분석",
        pdf_pages="16-17",
        title_ko="카테고리 × 브랜드 × 자사몰 페이지 유형 분석",
        purpose="각 브랜드 자사몰의 어떤 페이지 (PDP/기획전 등) 가 인용됐는지. 자사몰 콘텐츠 구조 시사점.",
        headers=["Topic", "Brand", "페이지 유형", "인용 수", "비율",
                 "대표 콘텐츠 예시 (Title)", "대표 URL"],
        dimensions=["Topic", "Brand", "페이지 유형"],
        data_rules=("""\
- 필터: `Domain` 이 BRAND_OWN_DOMAIN_KEYWORDS 매칭
- (Topic, Brand, 페이지 유형) 카운트
- 비율 = 페이지 유형 카운트 / (Topic, Brand) 자사몰 인용 합
- 대표 콘텐츠 예시 / URL = 첫 매칭"""),
        validation="(Topic, Brand) 별 비율 합 = 1.0",
    ),
]


def get_sheet(sheet_id: str) -> SheetSpec:
    for s in SHEETS:
        if s.sheet_id == sheet_id:
            return s
    raise KeyError(sheet_id)


def all_sheets() -> list[SheetSpec]:
    return list(SHEETS)
