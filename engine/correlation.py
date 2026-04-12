"""피어슨 상관계수 — 외환보유고 전월 변화량 vs 당월 환율 변동률."""

import logging

import pandas as pd
from scipy.stats import pearsonr

logger = logging.getLogger(__name__)


def run_pearson(reserves_delta: pd.Series, fx_return: pd.Series) -> dict:
    """외환보유고 MoM 변화량과 환율 MoM 변동률의 피어슨 상관계수를 계산한다.

    Args:
        reserves_delta: 외환보유고 전월 대비 변화량 (억달러).
        fx_return:      환율 전월 대비 변동률 (%).

    Returns:
        dict with keys:
            r       — 피어슨 상관계수
            p_value — 양측 검정 p-value
            n       — 유효 관측치 수
    """
    # 공통 인덱스로 정렬 후 결측 제거
    aligned = (
        pd.DataFrame({"delta": reserves_delta, "return": fx_return})
        .dropna()
    )

    if len(aligned) < 10:
        raise ValueError(
            f"피어슨 상관계수 계산에 필요한 최소 관측치(10)보다 적습니다 (n={len(aligned)})."
        )

    r, p_value = pearsonr(aligned["delta"], aligned["return"])
    n = len(aligned)

    logger.info("피어슨 r=%.4f, p=%.4f, n=%d", r, p_value, n)
    return {"r": r, "p_value": p_value, "n": n}
