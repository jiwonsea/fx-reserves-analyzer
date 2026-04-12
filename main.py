"""fx-reserves-analyzer 메인 진입점.

실행: python main.py

단계:
  1. 환경 설정 검증
  2. 데이터 수집 (BOK ECOS + yfinance)
  3. 분석 엔진 실행 (ADF → Pearson → Granger → VAR)
  4. 출력 생성 (차트 + Excel)
  5. 터미널 핵심 수치 출력
"""

import sys

# config를 가장 먼저 import → dotenv 로드 + 로깅 설정 선행
import config  # noqa: F401

import logging

import pandas as pd

logger = logging.getLogger(__name__)


def _fmt_granger(granger: dict, direction: str) -> str:
    """Granger 결과를 자소서용 문자열로 포맷한다."""
    if direction == "x_to_y":
        lag = granger["sig_lag_x_to_y"]
        p = granger["sig_p_x_to_y"]
        label = "보유고→환율"
    else:
        lag = granger["sig_lag_y_to_x"]
        p = granger["sig_p_y_to_x"]
        label = "환율→보유고"

    if lag is not None:
        return f"lag {lag}개월에서 유의 (p={p:.4f})"
    return "유의하지 않음 (p≥0.05)"


def main() -> None:
    # ------------------------------------------------------------------
    # Step 1: 환경 검증
    # ------------------------------------------------------------------
    if not config.BOK_API_KEY:
        sys.exit(
            "\n[ERROR] BOK_API_KEY가 설정되지 않았습니다.\n"
            "  1. cp .env.example .env\n"
            "  2. .env 파일에 BOK_API_KEY=<발급받은 키> 입력\n"
            "  발급: https://ecos.bok.or.kr/api/#/user/login\n"
        )

    logger.info("=" * 60)
    logger.info("fx-reserves-analyzer 분석 시작")
    logger.info("=" * 60)

    # ------------------------------------------------------------------
    # Step 2: 데이터 수집
    # ------------------------------------------------------------------
    from pipeline.ecos_collector import fetch_fx_reserves
    from pipeline.fx_collector import fetch_usdkrw

    logger.info("[1/5] 데이터 수집 중...")
    reserves = fetch_fx_reserves(config.BOK_API_KEY)
    usdkrw = fetch_usdkrw()

    # 공통 인덱스 정렬 (PeriodIndex 자동 정렬)
    merged = pd.DataFrame({"reserves": reserves, "usdkrw": usdkrw}).dropna()
    reserves = merged["reserves"]
    usdkrw = merged["usdkrw"]

    # 변화량·변동률 계산
    reserves_delta = reserves.diff().dropna()       # MoM 변화량 (억달러)
    fx_return = usdkrw.pct_change().dropna() * 100  # MoM 변동률 (%)

    # 두 파생 시계열도 동기화
    common = reserves_delta.index.intersection(fx_return.index)
    reserves_delta = reserves_delta.loc[common]
    fx_return = fx_return.loc[common]

    data = {
        "reserves": reserves,
        "usdkrw": usdkrw,
        "reserves_delta": reserves_delta,
        "fx_return": fx_return,
    }
    logger.info("데이터 준비 완료: %s ~ %s (%d개월)", reserves.index[0], reserves.index[-1], len(reserves))

    # ------------------------------------------------------------------
    # Step 3: 분석 엔진
    # ------------------------------------------------------------------
    logger.info("[2/5] ADF 단위근 검정 중...")
    from engine.unit_root import run_adf

    adf_reserves = run_adf(reserves, name="외환보유고")
    adf_fx = run_adf(usdkrw, name="원/달러 환율")

    logger.info("[3/5] 피어슨 상관계수 계산 중...")
    from engine.correlation import run_pearson

    pearson = run_pearson(reserves_delta, fx_return)

    logger.info("[4/5] Granger 인과성 검정 중...")
    from engine.granger import run_granger

    granger = run_granger(reserves_delta, fx_return)

    logger.info("[5/5] VAR 모델 추정 중 (IRF bootstrap 포함)...")
    from engine.var_model import RESERVES_COL, FX_COL, run_var

    var_df = pd.DataFrame(
        {RESERVES_COL: reserves_delta, FX_COL: fx_return}
    ).dropna()
    var_res = run_var(var_df)

    results = {
        "adf_reserves": adf_reserves,
        "adf_fx": adf_fx,
        "pearson": pearson,
        "granger": granger,
        "var": var_res,
    }

    # ------------------------------------------------------------------
    # Step 4: 출력 생성
    # ------------------------------------------------------------------
    logger.info("차트 생성 중...")
    from output.chart_generator import generate_charts

    chart_path = generate_charts(data, results)

    logger.info("Excel 리포트 생성 중...")
    from output.excel_reporter import generate_excel

    excel_path = generate_excel(data, results, chart_path)

    # ------------------------------------------------------------------
    # Step 5: 터미널 핵심 수치 출력
    # ------------------------------------------------------------------
    n_months = len(reserves)
    start_str = str(reserves.index[0])
    end_str = str(reserves.index[-1])

    print("\n" + "=" * 55)
    print("  === 핵심 분석 결과 ===")
    print("=" * 55)
    print(f"  분석 기간    : {start_str} ~ {end_str} ({n_months}개월)")
    print(f"  피어슨 상관계수: {pearson['r']:.4f} (p={pearson['p_value']:.4f})")
    print(f"  Granger (보유고→환율): {_fmt_granger(granger, 'x_to_y')}")
    print(f"  Granger (환율→보유고): {_fmt_granger(granger, 'y_to_x')}")
    print(f"  VAR 최적 lag : {var_res['optimal_lag']}개월 (AIC)")
    print(f"  IRF peak     : 외환보유고 충격 후 {var_res['peak_month']}개월에 환율 최대 반응")
    print(f"  FEVD         : 환율 변동성의 {var_res['fevd_pct']:.1f}%가 외환보유고로 설명")
    print("=" * 55)
    print(f"\n  [차트]  {chart_path}")
    print(f"  [Excel] {excel_path}\n")


if __name__ == "__main__":
    main()
