"""ADF 단위근 검정 — 차분 필요 여부를 판단한다."""

import logging

import pandas as pd
from statsmodels.tsa.stattools import adfuller

logger = logging.getLogger(__name__)


def run_adf(series: pd.Series, name: str = "series") -> dict:
    """Augmented Dickey-Fuller 단위근 검정.

    Args:
        series: 검정할 시계열 (결측값 포함 가능).
        name:   로그/출력용 변수명.

    Returns:
        dict with keys:
            adf_stat    — ADF 통계량
            p_value     — p-value (H0: 단위근 있음 = 비정상)
            needs_diff  — True이면 1차 차분 필요 (p > 0.05)
            diff_series — 1차 차분 시계열 (dropna 적용)
            critical_values — {"1%": ..., "5%": ..., "10%": ...}

    Raises:
        ValueError: 유효 관측치 < 20.
    """
    clean = series.dropna()
    if len(clean) < 20:
        raise ValueError(
            f"{name}: ADF 검정에 필요한 최소 관측치(20)보다 적습니다 (n={len(clean)})."
        )

    adf_stat, p_value, _, _, critical_values, _ = adfuller(clean, autolag="AIC")
    needs_diff = p_value > 0.05
    diff_series = series.diff().dropna()

    logger.info(
        "%s ADF: 통계량=%.4f, p=%.4f → %s",
        name,
        adf_stat,
        p_value,
        "차분 필요" if needs_diff else "정상 시계열",
    )

    return {
        "adf_stat": adf_stat,
        "p_value": p_value,
        "needs_diff": needs_diff,
        "diff_series": diff_series,
        "critical_values": critical_values,
    }
