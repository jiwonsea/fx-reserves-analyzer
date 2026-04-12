"""matplotlib 5-panel 차트 생성.

Panel 1: 외환보유고 + 원/달러 환율 시계열 (dual axis, 이벤트 음영)
Panel 2: 외환보유고 변화량 vs 환율 변동률 산점도 (회귀선)
Panel 3: Granger p-value by lag (양방향)
Panel 4: VAR IRF — 외환보유고 충격 → 환율 반응
Panel 5: FEVD — 환율 분산 설명 비율 (12개월)
"""

import logging

import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import config
from engine.events import EVENTS, apply_event_shading

logger = logging.getLogger(__name__)

# 한글 폰트 & minus 부호 설정
plt.rcParams["font.family"] = "Malgun Gothic"
plt.rcParams["axes.unicode_minus"] = False


def generate_charts(
    data: dict,
    results: dict,
    output_path: str | None = None,
) -> str:
    """5개 패널 분석 차트를 생성하고 PNG로 저장한다.

    Args:
        data: {
            "reserves": pd.Series (PeriodIndex, 억달러),
            "usdkrw":   pd.Series (PeriodIndex, 원),
            "reserves_delta": pd.Series,
            "fx_return": pd.Series,
        }
        results: {
            "pearson": dict,
            "granger": dict,
            "var": dict,
        }
        output_path: 저장 경로. None이면 config.CHART_PATH 사용.

    Returns:
        저장된 PNG 파일 경로 (str).
    """
    if output_path is None:
        output_path = str(config.CHART_PATH)

    reserves = data["reserves"]
    usdkrw = data["usdkrw"]
    res_delta = data["reserves_delta"]
    fx_ret = data["fx_return"]

    granger = results["granger"]
    var_res = results["var"]
    pearson = results["pearson"]

    # Timestamp 변환 (axvspan 호환)
    ts_reserves = reserves.index.to_timestamp()
    ts_usdkrw = usdkrw.index.to_timestamp()

    # ---------- 레이아웃 ----------
    from matplotlib.gridspec import GridSpec

    fig = plt.figure(figsize=(18, 20))
    gs = GridSpec(3, 2, figure=fig, hspace=0.38, wspace=0.32)
    ax1 = fig.add_subplot(gs[0, 0])
    ax2 = fig.add_subplot(gs[0, 1])
    ax3 = fig.add_subplot(gs[1, 0])
    ax4 = fig.add_subplot(gs[1, 1])
    ax5 = fig.add_subplot(gs[2, :])  # 전폭

    fig.suptitle(
        "한국 외환보유고 ↔ 원/달러 환율 관계 분석",
        fontsize=16,
        fontweight="bold",
        y=0.98,
    )

    # =========================================================
    # Panel 1: 시계열 dual-axis + 이벤트 음영
    # =========================================================
    ax1r = ax1.twinx()

    ax1.plot(ts_reserves, reserves.values, color="steelblue", linewidth=1.4, label="외환보유고")
    ax1r.plot(ts_usdkrw, usdkrw.values, color="firebrick", linewidth=1.0, alpha=0.8, label="원/달러")

    apply_event_shading(ax1)

    ax1.set_ylabel("외환보유고 (억달러)", color="steelblue", fontsize=9)
    ax1r.set_ylabel("원/달러 환율", color="firebrick", fontsize=9)
    ax1.set_title("외환보유고 & 원/달러 환율 시계열", fontsize=10, fontweight="bold")
    ax1.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
    ax1.xaxis.set_major_locator(mdates.YearLocator(5))
    plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45, fontsize=7)

    # 범례 합치기
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax1r.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, fontsize=7, loc="upper left")

    # =========================================================
    # Panel 2: 산점도 + 회귀선
    # =========================================================
    aligned2 = pd.DataFrame({"delta": res_delta, "ret": fx_ret}).dropna()
    x_vals = aligned2["delta"].values
    y_vals = aligned2["ret"].values

    ax2.scatter(x_vals, y_vals, s=12, alpha=0.5, color="steelblue", edgecolors="none")

    coeffs = np.polyfit(x_vals, y_vals, 1)
    x_line = np.linspace(x_vals.min(), x_vals.max(), 200)
    ax2.plot(x_line, np.polyval(coeffs, x_line), color="firebrick", linewidth=1.5)

    r = pearson["r"]
    p = pearson["p_value"]
    ax2.annotate(
        f"r = {r:.3f}\np = {p:.4f}\nn = {pearson['n']}",
        xy=(0.05, 0.88),
        xycoords="axes fraction",
        fontsize=8,
        bbox={"boxstyle": "round,pad=0.3", "facecolor": "lightyellow", "edgecolor": "gray"},
    )
    ax2.axhline(0, color="gray", linewidth=0.7, linestyle="--")
    ax2.axvline(0, color="gray", linewidth=0.7, linestyle="--")
    ax2.set_xlabel("외환보유고 변화량 (억달러)", fontsize=9)
    ax2.set_ylabel("환율 변동률 (%)", fontsize=9)
    ax2.set_title("외환보유고 변화량 vs 환율 변동률", fontsize=10, fontweight="bold")

    # =========================================================
    # Panel 3: Granger p-value by lag
    # =========================================================
    lags = list(range(1, config.GRANGER_MAX_LAG + 1))
    p_xy = [granger["x_to_y"][lag] for lag in lags]
    p_yx = [granger["y_to_x"][lag] for lag in lags]

    ax3.plot(lags, p_xy, "o-", color="steelblue", linewidth=1.5, label="보유고→환율")
    ax3.plot(lags, p_yx, "s--", color="firebrick", linewidth=1.5, label="환율→보유고")
    ax3.axhline(0.05, color="black", linewidth=1.0, linestyle=":", label="α=0.05")
    ax3.set_xlabel("Lag (개월)", fontsize=9)
    ax3.set_ylabel("p-value", fontsize=9)
    ax3.set_title("Granger 인과성 검정 p-value", fontsize=10, fontweight="bold")
    ax3.set_xticks(lags)
    ax3.legend(fontsize=8)
    ax3.set_ylim(0, max(max(p_xy), max(p_yx)) * 1.15)

    # =========================================================
    # Panel 4: VAR IRF
    # =========================================================
    irf_periods = config.IRF_PERIODS
    irf_x = list(range(irf_periods + 1))
    irf_vals = var_res["irf_values"]
    irf_lower = var_res["irf_lower"]
    irf_upper = var_res["irf_upper"]

    ax4.plot(irf_x, irf_vals, color="steelblue", linewidth=1.8, label="IRF")
    ax4.fill_between(irf_x, irf_lower, irf_upper, alpha=0.2, color="steelblue", label="95% CI")
    ax4.axhline(0, color="black", linewidth=0.8, linestyle="--")

    peak = var_res["peak_month"]
    ax4.axvline(peak, color="firebrick", linewidth=1.0, linestyle=":", alpha=0.7)
    ax4.annotate(
        f"peak: {peak}개월",
        xy=(peak, irf_vals[peak]),
        xytext=(peak + 0.5, irf_vals[peak] * 1.2 if irf_vals[peak] != 0 else 0.001),
        fontsize=7,
        arrowprops={"arrowstyle": "->", "color": "firebrick"},
        color="firebrick",
    )
    ax4.set_xlabel("기간 (개월)", fontsize=9)
    ax4.set_ylabel("환율 반응", fontsize=9)
    ax4.set_title(f"IRF: 외환보유고 1단위 충격 → 환율 반응\n(VAR lag={var_res['optimal_lag']})", fontsize=10, fontweight="bold")
    ax4.set_xticks(irf_x)
    ax4.legend(fontsize=8)

    # =========================================================
    # Panel 5: FEVD stacked area
    # =========================================================
    fevd_obj = var_res["fevd_obj"]
    from engine.var_model import FX_COL, RESERVES_COL
    cols = [RESERVES_COL, FX_COL]
    fx_idx_fevd = cols.index(FX_COL)
    res_idx_fevd = cols.index(RESERVES_COL)

    periods = list(range(1, irf_periods + 1))
    fevd_res_share = [
        fevd_obj.decomp[fx_idx_fevd, t - 1, res_idx_fevd] * 100 for t in periods
    ]
    fevd_other = [100 - v for v in fevd_res_share]

    ax5.stackplot(
        periods,
        fevd_res_share,
        fevd_other,
        labels=["외환보유고 기여", "기타 변수 기여"],
        colors=["steelblue", "lightgray"],
        alpha=0.85,
    )
    ax5.set_xlabel("기간 (개월)", fontsize=10)
    ax5.set_ylabel("분산 설명 비율 (%)", fontsize=10)
    ax5.set_title(
        f"FEVD: 환율 변동성 분산 분해 (12개월 기준 외환보유고 기여: {var_res['fevd_pct']:.1f}%)",
        fontsize=10,
        fontweight="bold",
    )
    ax5.set_xticks(periods)
    ax5.set_ylim(0, 100)
    ax5.legend(fontsize=9, loc="upper right")

    # 저장
    plt.savefig(output_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    logger.info("차트 저장 완료: %s", output_path)
    return output_path
