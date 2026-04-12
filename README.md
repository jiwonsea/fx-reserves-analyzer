# fx-reserves-analyzer

한국 외환보유고(BOK ECOS)와 원/달러 환율(yfinance)의 관계를 계량분석하는 Python 도구.

한국은행 국제·금융부서 인턴십 포트폴리오 프로젝트.

---

## 분석 내용

| 분석 | 방법 |
|------|------|
| 단위근 | ADF 검정 → 1차 차분 적용 |
| 선형 관계 | 피어슨 상관계수 (외환보유고 변화량 vs 환율 변동률) |
| 인과성 | Granger 인과성 검정 (양방향, lag 1~6개월) |
| 동태 반응 | VAR + IRF (외환보유고 충격 → 환율 반응 경로) |
| 분산 분해 | FEVD (환율 변동성에서 외환보유고 기여 비율) |

---

## 실제 분석 결과 인사이트

분석 기간: 1995-01 ~ 2026-03 (375개월, BOK ECOS 월말 환율 기준)

1. **피어슨 r = −0.4752 (p<0.0001, n=374)**: 외환보유고 MoM 증가 시 환율 하락 경향이 중간 강도(moderate)로 확인 — 단순 상관만으로도 개입 효과가 유의미하게 포착됨.
2. **Granger 인과성 비대칭**: 보유고→환율은 lag **3개월**에서 첫 유의(p=0.020), 환율→보유고는 lag **1개월**에서 유의(p=0.013) — 환율 충격이 BOK 개입을 즉각 유발하지만, 개입의 환율 안정 효과는 3개월의 지연을 가짐.
3. **FEVD 21.9% (12개월 기준)**: 환율 분산의 약 22%가 외환보유고 충격으로 설명 — BOK 개입이 단기 변동성의 상당 부분을 통제하나, 나머지 78%는 글로벌 달러 강세·경상수지·투자심리 등 외부 요인이 주도함.

---

## 빠른 시작

### 1. 환경 설정

```bash
# 의존성 설치
pip install -r requirements.txt

# API 키 설정
cp .env.example .env
# .env 파일에서 BOK_API_KEY 값을 실제 키로 교체
# 발급: https://ecos.bok.or.kr/api/#/user/login
```

### 2. 실행

```bash
cd F:\dev\Portfolio\fx-reserves-analyzer
python main.py
```

### 3. 출력 파일

| 파일 | 설명 |
|------|------|
| `output/fx_reserves_analysis.png` | 5-panel 분석 차트 |
| `output/fx_reserves_report.xlsx` | 요약/Granger/VAR계수 3-sheet Excel |

### 4. 테스트

```bash
pytest tests/ -v
```

---

## 디렉터리 구조

```
fx-reserves-analyzer/
├── config.py                 # 환경변수·상수·로깅
├── main.py                   # 메인 진입점
├── pipeline/
│   ├── ecos_collector.py     # BOK ECOS API (외환보유고)
│   └── fx_collector.py       # yfinance (USDKRW=X)
├── engine/
│   ├── unit_root.py          # ADF 단위근 검정
│   ├── correlation.py        # 피어슨 상관계수
│   ├── granger.py            # Granger 인과성 검정
│   ├── var_model.py          # VAR + IRF + FEVD
│   └── events.py             # 이벤트 기간 상수
├── output/
│   ├── chart_generator.py    # matplotlib 5-panel
│   └── excel_reporter.py     # openpyxl 리포트
└── tests/
    └── test_pipeline.py      # pytest 6케이스
```

---

## 주요 이벤트 기간 (차트 음영)

| 기간 | 이벤트 |
|------|--------|
| 1997-07 ~ 1998-12 | IMF 외환위기 |
| 2008-09 ~ 2009-03 | 글로벌 금융위기 |
| 2020-03 ~ 2020-06 | 코로나 쇼크 |
| 2022-01 ~ 2022-10 | Fed 긴축(킹달러) |
| 2026-03 ~ 현재    | 중동 원유 공급 충격 |

---

## 의존성

- pandas ≥ 2.2, numpy, statsmodels ≥ 0.14
- yfinance, scipy, matplotlib, openpyxl
- python-dotenv, requests
