"""Granger 인과성 검정 — 양방향, lag 1~6개월."""

import contextlib
import io
import logging

import numpy as np
import pandas as pd
from statsmodels.tsa.stattools import grangercausalitytests

import config

logger = logging.getLogger(__name__)


def run_granger(
    x: pd.Series,
    y: pd.Series,
    max_lag: int = config.GRANGER_MAX_LAG,
    alpha: float = 0.05,
) -> dict:
    """양방향 Granger 인과성 검정.

    Args:
        x:       예측 변수 (외환보유고 변화량 등 정상 시계열).
        y:       결과 변수 (환율 변동률 등 정상 시계열).
        max_lag: 검정할 최대 lag (기본값: config.GRANGER_MAX_LAG).
        alpha:   유의수준 (기본값: 0.05).

    Returns:
        dict with keys:
            x_to_y        — {lag: p_value}  x → y 방향
            y_to_x        — {lag: p_value}  y → x 방향
            sig_lag_x_to_y — x→y 유의 최소 lag (없으면 None)
            sig_p_x_to_y   — 해당 lag의 p-value (없으면 None)
            sig_lag_y_to_x — y→x 유의 최소 lag (없으면 None)
            sig_p_y_to_x   — 해당 lag의 p-value (없으면 None)
    """
    aligned = pd.DataFrame({"x": x, "y": y}).dropna()

    if len(aligned) < max_lag * 3 + 10:
        raise ValueError(
            f"Granger 검정에 필요한 관측치 부족 (n={len(aligned)}, max_lag={max_lag})."
        )

    x_arr = aligned["x"].values
    y_arr = aligned["y"].values

    def _granger(response: np.ndarray, predictor: np.ndarray) -> dict[int, float]:
        """grangercausalitytests wrapper: predictor → response 방향."""
        data = np.column_stack([response, predictor])
        # verbose 출력 억제 (statsmodels 일부 버전에서 verbose=False여도 출력됨)
        with contextlib.redirect_stdout(io.StringIO()):
            results = grangercausalitytests(data, maxlag=max_lag)
        return {lag: results[lag][0]["ssr_ftest"][1] for lag in range(1, max_lag + 1)}

    x_to_y = _granger(response=y_arr, predictor=x_arr)
    y_to_x = _granger(response=x_arr, predictor=y_arr)

    def _first_sig(p_dict: dict[int, float]) -> tuple[int | None, float | None]:
        for lag in range(1, max_lag + 1):
            if p_dict[lag] < alpha:
                return lag, p_dict[lag]
        return None, None

    sig_lag_xy, sig_p_xy = _first_sig(x_to_y)
    sig_lag_yx, sig_p_yx = _first_sig(y_to_x)

    logger.info(
        "Granger x→y: %s",
        f"lag {sig_lag_xy}개월 유의 (p={sig_p_xy:.4f})" if sig_lag_xy else "유의하지 않음",
    )
    logger.info(
        "Granger y→x: %s",
        f"lag {sig_lag_yx}개월 유의 (p={sig_p_yx:.4f})" if sig_lag_yx else "유의하지 않음",
    )

    return {
        "x_to_y": x_to_y,
        "y_to_x": y_to_x,
        "sig_lag_x_to_y": sig_lag_xy,
        "sig_p_x_to_y": sig_p_xy,
        "sig_lag_y_to_x": sig_lag_yx,
        "sig_p_y_to_x": sig_p_yx,
    }
