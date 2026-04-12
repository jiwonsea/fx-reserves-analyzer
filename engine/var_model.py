"""VAR 모델 + IRF + FEVD.

컬럼 순서: ['reserves_delta', 'fx_return']
- reserves_delta: 외환보유고 MoM 변화량 (억달러)
- fx_return:      원/달러 환율 MoM 변동률 (%)
"""

import logging

import numpy as np
import pandas as pd
from statsmodels.tsa.api import VAR

import config

logger = logging.getLogger(__name__)

RESERVES_COL = "reserves_delta"
FX_COL = "fx_return"


def run_var(df: pd.DataFrame) -> dict:
    """VAR 모델을 적합하고 IRF·FEVD를 계산한다.

    Args:
        df: 두 컬럼(reserves_delta, fx_return)을 가진 정상 시계열 DataFrame.

    Returns:
        dict with keys:
            optimal_lag  — AIC 선택 lag 수
            aic          — 해당 AIC 값
            irf_values   — IRF: 외환보유고 충격 → 환율 반응 (길이 irf_periods+1)
            irf_lower    — 95% 신뢰구간 하한 (bootstrap)
            irf_upper    — 95% 신뢰구간 상한 (bootstrap)
            peak_month   — |IRF| 최대값의 기간 (0 = 충격 시점)
            fevd_pct     — 12개월 시점 환율 분산에서 외환보유고 설명 비율 (%)
            var_params   — VAR 계수 DataFrame

    Raises:
        ValueError: 관측치 부족 또는 컬럼 누락.
    """
    for col in (RESERVES_COL, FX_COL):
        if col not in df.columns:
            raise ValueError(f"DataFrame에 컬럼 '{col}'이 없습니다.")

    clean = df[[RESERVES_COL, FX_COL]].dropna()
    if len(clean) < 50:
        raise ValueError(
            f"VAR 추정에 필요한 최소 관측치(50)보다 적습니다 (n={len(clean)})."
        )

    model = VAR(clean)
    lag_order = model.select_order(maxlags=config.VAR_MAX_LAG)
    selected = lag_order.aic if lag_order.aic is not None else 0
    optimal_lag = max(1, int(selected))
    if selected != optimal_lag:
        logger.warning("AIC 선택 lag=%s → 최소값 1로 조정", selected)
    logger.info("VAR 최적 lag (AIC): %d", optimal_lag)

    results = model.fit(optimal_lag)

    fx_idx = list(clean.columns).index(FX_COL)
    res_idx = list(clean.columns).index(RESERVES_COL)

    # IRF 계산
    irf_periods = config.IRF_PERIODS
    logger.info("IRF 계산 중... (bootstrap repl=%d, ~20초 소요)", config.IRF_BOOTSTRAP_REPL)
    irf_obj = results.irf(irf_periods)
    irf_values = irf_obj.irfs[:, fx_idx, res_idx]

    # 신뢰구간 (bootstrap)
    try:
        lower, upper = irf_obj.errband_mc(
            orth=True,
            svar=False,
            repl=config.IRF_BOOTSTRAP_REPL,
            signif=0.05,
        )
        irf_lower = lower[:, fx_idx, res_idx]
        irf_upper = upper[:, fx_idx, res_idx]
    except Exception as exc:
        logger.warning("IRF bootstrap 실패 (%s) → 단순 ±1.5σ 대체", exc)
        std = float(np.std(irf_values))
        irf_lower = irf_values - 1.5 * std
        irf_upper = irf_values + 1.5 * std

    peak_month = int(np.argmax(np.abs(irf_values)))

    # FEVD
    fevd_obj = results.fevd(irf_periods)
    # decomp shape: (n_vars, periods, n_vars)
    fevd_pct = float(fevd_obj.decomp[fx_idx, -1, res_idx]) * 100

    logger.info(
        "IRF peak: %d개월 후, FEVD(12개월): %.1f%%", peak_month, fevd_pct
    )

    return {
        "optimal_lag": optimal_lag,
        "aic": float(lag_order.aic),
        "irf_values": irf_values,
        "irf_lower": irf_lower,
        "irf_upper": irf_upper,
        "peak_month": peak_month,
        "fevd_pct": fevd_pct,
        "var_params": results.params,
        "fevd_obj": fevd_obj,
    }
