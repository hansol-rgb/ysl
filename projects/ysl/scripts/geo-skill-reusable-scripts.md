# GEO 스킬 재사용 가능 스크립트 정리

YSL Beauty 프로젝트에서 작성된 Python 스크립트 중 `.claude/skills/` 의 GEO 파이프라인 스킬에 재사용할 만한 것들을 단계별로 정리.

## 단계 1 — geo-seed-keywords (온보딩 → 커뮤니티 리서치 → 시드키워드)

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `extract_seed_kw.py` | `ysl/` | 브랜드 키워드 PDF (월별 검색량) + `research_data_intermediate.md` 파싱 → `seed_keywords.json`. 브랜디드 + 제네릭 시드키워드 도출 핵심 로직 |
| `gen_excel.py` | `ysl/` | `research_data_intermediate.md` → `Intent_Category_Research.xlsx` (5시트: 포지셔닝 가설 / 카테고리별 리서치 / 요약+방법론). geo-community-research 산출물 포맷 |
| `gen_mdx.py` | `ysl/` | 동일 소스 → 마크다운 버전. Lancome 템플릿 구조 기반 |

## 단계 1 보조 — geo-community-research (다채널 리서치 → 정리)

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `scrape_urls.py` | `ysl/` | `research_data_intermediate.md` 내 URL 추출 → 본문 크롤링 → `scraped_data.json`. SSL 우회·타임아웃·boilerplate 제거 포함 |
| `build_summaries.py` | `ysl/` | 스크래핑 본문 → 깔끔한 Content Summary (`summaries.json`). 한국 커뮤니티 boilerplate 패턴 제거 룰 보유 |
| `update_files.py` | `ysl/` | `scraped_data.json` → 기존 Excel/MDX의 `Description + User Needs` 컬럼을 `Content Summary` 단일 컬럼으로 교체 |
| `apply_summaries.py` | `ysl/` | `summaries.json` 기반 최종 업데이트 — Pre-existing Content 제거 / Source 중복 제거 |

## 단계 2 — geo-question-builder (시드키워드 → 질문 + Excel)

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `gen_questions_perfume.py` / `gen_questions_gifting.py` / `gen_questions_cushion.py` | `ysl/` | 카테고리별 질문 JSON 생성 V1. 시드키워드 × Intent × 3축 Intent Type 룰 |
| `gen_q_*_v3.py` (3종) | `ysl/` | V3 — 제네릭 전환·비율 조정 후 최종 버전. **재사용은 V3 우선** |
| `compile_questions_excel.py` | `ysl/` | 3개 카테고리 JSON + Seed KW JSON → 통합 Excel (`Seed KW` / `Prompt Setting` / `Questions` 시트) |
| `reformat_excel.py` | `ysl/` | Skinceuticals 포맷 정확히 맞춰 Question List 재생성. 색상·헤더·시트 구조 표준화 |

> 참고: `gen_q_*_v2.py`, `_tmp_v2_*.py` 는 중간 산출 — **V3와 V1만 보관 권장**, V2/tmp는 스킬 패키징 시 제외.

## 단계 3 — geo-backoffice-runner (수집 후 데이터 통합)

BackOffice 자체는 외부 도구. 응답·인용 CSV 처리 단계에 다음 두 스크립트가 코어.

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `build_ysl_audit_ppt.py` | `archive/` | 4 사이클 KW + Q CSV 통합 → `Reference ID` join → 멘션 추출 (`extract_mentions`) + 도메인 분류 (`classify_domain` + `DOMAIN_RULES`) → `/tmp/ysl_metrics.json`. **브랜드 매칭·도메인 룰은 그대로 재사용 가능** |
| `build_vivi_monitoring.py` | `ysl/scripts/` | (NEW) 4 사이클 응답 + 인용을 long-format VIVI Monitoring xlsx 4시트로 통합. 키워드/쿼리 양쪽 + 12 브랜드 멘션 카운트 + 자사 도메인 플래그. **고객 모니터링 파일 표준 빌더** |

핵심 함수 (스킬에 발췌해 쓸 만한 것):
- `extract_mentions(text, brands)` — variations substring 매칭 (대소문자 무시)
- `extract_mentions_with_count(text, brands)` — Y/N + 매칭 횟수 (VIVI Monitoring용)
- `classify_domain(domain)` — Blog/Social/Commerce/Media/Brand-owned/Reference/Other
- `parse_epoch_from_filename(fname)` + `epoch_to_dt(ms)` — CSV 파일명에서 사이클 날짜 자동 추출
- `load_meta()` — Question List `(Final)` + `(Final) + KW` 두 시트 join (No → 카테고리/인텐트/축1·2·3/MSV/Reference ID)

## 단계 4 — geo-final-report (전체 데이터 → PPT 보고서)

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `generate_ppt_v2.py` | `ysl/scripts/` | **권장**. SKC 템플릿 미러링 PPT 생성. 셀 구조·병합·폰트·컬럼 폭 절대 건드리지 않고 텍스트만 교체. 슬라이드 단위 삭제 + 로고 교체. PPT 템플릿 보존 원칙의 표준 |
| `generate_ppt.py` | `ysl/scripts/` | V1. 전역 텍스트 치환 + 표/차트 교체 + TBD 박스 오버레이. **V2보다 공격적** — V2 우선 사용 |
| `build_ysl_audit_ppt.py` | `archive/` | 데이터 파이프라인 (위 단계 3 참조). PPT 생성 직전 `/tmp/ysl_metrics.json` + `/tmp/ysl_*.pkl` 산출 |

## 공통 유틸 (모든 단계 공유)

| 스크립트 | 위치 | 역할 |
|---|---|---|
| `apply_colors.py` | `ysl/` | Skinceuticals/BubbleShare 색상 체계(`#7030A0` / `#002060` / `#ECD1ED` 등) 일괄 적용. Excel 헤더·서식 표준화 |

## 제외 (스킬 패키징 시 빼도 됨)

- `ysl/_tmp_v2_cushion.py`, `_tmp_v2_gifting.py`, `_tmp_v2_perfume.py` — 중간 작업 파일
- `gen_q_*_v2.py` — V3로 대체됨

## 스킬 매핑 요약

| 스킬 | 핵심 스크립트 (이식 우선순위 순) |
|---|---|
| **geo-seed-keywords** | `extract_seed_kw.py` → `gen_excel.py` / `gen_mdx.py` |
| **geo-community-research** | `scrape_urls.py` → `build_summaries.py` → `update_files.py` / `apply_summaries.py` |
| **geo-question-builder** | `gen_q_*_v3.py` (3종) → `compile_questions_excel.py` → `reformat_excel.py` |
| **geo-backoffice-runner** | `build_ysl_audit_ppt.py` (브랜드 멘션 + 도메인 분류 함수) + `build_vivi_monitoring.py` |
| **geo-final-report** | `build_ysl_audit_ppt.py` (데이터 파이프라인) + `generate_ppt_v2.py` (PPT 빌더) |
| **공통** | `apply_colors.py` |
