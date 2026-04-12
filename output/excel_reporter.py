"""openpyxl Excel 리포트 — 데이터 시트 + 네이티브 Excel 차트.

시트 구성:
  Summary        — 핵심 통계 요약
  TimeSeries     — 월별 외환보유고/환율 원시 데이터 + 이중축 꺾은선 차트
  DeltaReturn    — MoM 변화량/변동률 + 산점도 차트 (회귀선 포함)
  Granger        — lag × 방향 p-value 표 + 꺾은선 차트
  IRF            — 충격반응함수 데이터 + 신뢰구간 포함 차트
  FEVD           — 분산분해 데이터 + 누적 막대 차트
  VAR계수        — VAR 회귀계수 행렬
"""

import logging
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, ScatterChart, Series
from openpyxl.chart.axis import NumericAxis
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config

logger = logging.getLogger(__name__)

_HEADER_FONT = Font(bold=True, size=10)
_HEADER_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
_TITLE_FONT = Font(bold=True, size=12)
_SIG_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")


def _style_header(ws, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def _auto_width(ws, padding: int = 4, max_width: int = 40) -> None:
    for col_cells in ws.columns:
        max_len = max(
            (len(str(c.value)) for c in col_cells if c.value is not None),
            default=0,
        )
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(
            max_len + padding, max_width
        )


# ---------------------------------------------------------------------------
# Sheet builders
# ---------------------------------------------------------------------------

def _sheet_summary(wb: Workbook, data: dict, results: dict) -> None:
    ws = wb.create_sheet("Summary")
    reserves = data["reserves"]
    pearson = results["pearson"]
    granger = results["granger"]
    var_res = results["var"]
    adf_r = results.get("adf_reserves", {})
    adf_f = results.get("adf_fx", {})

    def _granger_str(direction: str) -> str:
        lag = granger[f"sig_lag_{direction}"]
        p = granger[f"sig_p_{direction}"]
        return f"lag {lag}개월 유의 (p={p:.4f})" if lag else "유의하지 않음"

    rows = [
        ("항목", "값"),
        ("분석 기간 시작", str(reserves.index[0])),
        ("분석 기간 종료", str(reserves.index[-1])),
        ("총 관측 월수", f"{len(reserves)}개월"),
        ("", ""),
        ("외환보유고 ADF p-value", f"{adf_r.get('p_value', 0):.4f}" if adf_r else ""),
        ("환율 ADF p-value", f"{adf_f.get('p_value', 0):.4f}" if adf_f else ""),
        ("차분 적용", "1차 차분"),
        ("", ""),
        ("피어슨 상관계수 (r)", f"{pearson['r']:.4f}"),
        ("피어슨 p-value", f"{pearson['p_value']:.4f}"),
        ("피어슨 n", str(pearson["n"])),
        ("", ""),
        ("Granger 보유고→환율", _granger_str("x_to_y")),
        ("Granger 환율→보유고", _granger_str("y_to_x")),
        ("", ""),
        ("VAR 최적 lag (AIC)", f"{var_res['optimal_lag']}개월"),
        ("IRF peak 기간", f"{var_res['peak_month']}개월 후"),
        ("FEVD 외환보유고 기여 (12개월)", f"{var_res['fevd_pct']:.2f}%"),
    ]

    for r_idx, (key, val) in enumerate(rows, start=1):
        ws.cell(row=r_idx, column=1, value=key)
        ws.cell(row=r_idx, column=2, value=val)
        if r_idx == 1:
            _style_header(ws, r_idx, 2)

    _auto_width(ws)


def _sheet_timeseries(wb: Workbook, data: dict) -> None:
    """TimeSeries 시트: 월별 데이터 + 이중축 꺾은선 차트."""
    ws = wb.create_sheet("TimeSeries")
    reserves = data["reserves"]
    usdkrw = data["usdkrw"]

    headers = ["기간", "외환보유고(억달러)", "원/달러(월말)"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for r_idx, period in enumerate(reserves.index, start=2):
        ws.cell(row=r_idx, column=1, value=str(period))
        ws.cell(row=r_idx, column=2, value=round(float(reserves[period]), 2))
        ws.cell(row=r_idx, column=3, value=round(float(usdkrw[period]), 2) if period in usdkrw.index else None)

    n = len(reserves)
    _auto_width(ws)

    # 이중축 꺾은선 차트
    c1 = LineChart()
    c1.title = "외환보유고 & 원/달러 환율"
    c1.style = 10
    c1.height = 18
    c1.width = 36
    c1.y_axis.title = "외환보유고 (억달러)"
    c1.y_axis.axId = 200

    data_res = Reference(ws, min_col=2, min_row=1, max_row=n + 1)
    c1.add_data(data_res, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n + 1)
    c1.set_categories(cats)

    # 환율 → 보조 Y축
    c2 = LineChart()
    c2.y_axis.axId = 100
    c2.y_axis.crosses = "max"
    c2.y_axis.title = "원/달러 환율"
    data_fx = Reference(ws, min_col=3, min_row=1, max_row=n + 1)
    c2.add_data(data_fx, titles_from_data=True)
    c2.set_categories(cats)
    c1 += c2

    # 시리즈 색상
    c1.series[0].graphicalProperties.line.solidFill = "4472C4"   # 파랑 (보유고)
    if len(c1.series) > 1:
        c1.series[1].graphicalProperties.line.solidFill = "FF0000"  # 빨강 (환율)

    ws.add_chart(c1, f"A{n + 4}")


def _sheet_delta_return(wb: Workbook, data: dict, results: dict) -> None:
    """DeltaReturn 시트: MoM 변화량/변동률 + 산점도."""
    ws = wb.create_sheet("DeltaReturn")
    res_delta = data["reserves_delta"]
    fx_ret = data["fx_return"]

    aligned = pd.DataFrame({"delta": res_delta, "ret": fx_ret}).dropna()

    # 회귀선 계산 (2점: min/max x)
    x_vals = aligned["delta"].values
    y_vals = aligned["ret"].values
    slope, intercept = np.polyfit(x_vals, y_vals, 1)
    x_min, x_max = float(x_vals.min()), float(x_vals.max())

    # 데이터 시트: 기간 | delta | return
    headers = ["기간", "보유고변화(억$)", "환율변동률(%)"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for r_idx, (period, row) in enumerate(aligned.iterrows(), start=2):
        ws.cell(row=r_idx, column=1, value=str(period))
        ws.cell(row=r_idx, column=2, value=round(float(row["delta"]), 3))
        ws.cell(row=r_idx, column=3, value=round(float(row["ret"]), 4))

    n = len(aligned)

    # 회귀선 2점 테이블 (E열)
    ws.cell(row=1, column=5, value="회귀_X").font = _HEADER_FONT
    ws.cell(row=1, column=6, value="회귀_Y").font = _HEADER_FONT
    ws.cell(row=2, column=5, value=round(x_min, 3))
    ws.cell(row=2, column=6, value=round(slope * x_min + intercept, 4))
    ws.cell(row=3, column=5, value=round(x_max, 3))
    ws.cell(row=3, column=6, value=round(slope * x_max + intercept, 4))

    _auto_width(ws)

    # 산점도
    sc = ScatterChart()
    sc.title = f"보유고변화 vs 환율변동률  (r={results['pearson']['r']:.3f})"
    sc.style = 13
    sc.height = 15
    sc.width = 28
    sc.x_axis.title = "외환보유고 변화량 (억달러)"
    sc.y_axis.title = "환율 변동률 (%)"

    # 데이터 포인트
    xvals = Reference(ws, min_col=2, min_row=2, max_row=n + 1)
    yvals = Reference(ws, min_col=3, min_row=2, max_row=n + 1)
    s1 = Series(yvals, xvals, title="월별 데이터")
    s1.marker.symbol = "circle"
    s1.marker.size = 4
    s1.graphicalProperties.line.noFill = True
    sc.series.append(s1)

    # 회귀선
    rx = Reference(ws, min_col=5, min_row=2, max_row=3)
    ry = Reference(ws, min_col=6, min_row=2, max_row=3)
    s2 = Series(ry, rx, title="회귀선")
    s2.graphicalProperties.line.solidFill = "FF0000"
    s2.graphicalProperties.line.width = 20000
    sc.series.append(s2)

    ws.add_chart(sc, f"A{n + 4}")


def _sheet_granger(wb: Workbook, results: dict) -> None:
    """Granger 시트: lag × 방향 p-value 표 + 꺾은선 차트."""
    ws = wb.create_sheet("Granger")
    granger = results["granger"]
    max_lag = config.GRANGER_MAX_LAG

    # 넓은 포맷 (차트용): Lag | 보유고→환율 | 환율→보유고
    wide_headers = ["Lag(개월)", "보유고→환율 p", "환율→보유고 p"]
    for col, h in enumerate(wide_headers, start=1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(wide_headers))

    for lag in range(1, max_lag + 1):
        r = lag + 1
        p_xy = granger["x_to_y"][lag]
        p_yx = granger["y_to_x"][lag]
        ws.cell(row=r, column=1, value=lag)
        ws.cell(row=r, column=2, value=round(p_xy, 4))
        ws.cell(row=r, column=3, value=round(p_yx, 4))
        # 유의한 셀 초록 음영
        if p_xy < 0.05:
            ws.cell(row=r, column=2).fill = _SIG_FILL
        if p_yx < 0.05:
            ws.cell(row=r, column=3).fill = _SIG_FILL

    # 긴 포맷 (읽기용): E~H열
    long_headers = ["Lag", "방향", "p-value", "유의"]
    for col, h in enumerate(long_headers, start=5):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, 8)

    r2 = 2
    for lag in range(1, max_lag + 1):
        for direction, p_val in [
            ("보유고 → 환율", granger["x_to_y"][lag]),
            ("환율 → 보유고", granger["y_to_x"][lag]),
        ]:
            ws.cell(row=r2, column=5, value=lag)
            ws.cell(row=r2, column=6, value=direction)
            ws.cell(row=r2, column=7, value=round(p_val, 4))
            ws.cell(row=r2, column=8, value="Y" if p_val < 0.05 else "N")
            if p_val < 0.05:
                ws.cell(row=r2, column=7).fill = _SIG_FILL
            r2 += 1

    _auto_width(ws)

    # 꺾은선 차트
    lc = LineChart()
    lc.title = "Granger 인과성 p-value by Lag"
    lc.style = 10
    lc.height = 14
    lc.width = 28
    lc.y_axis.title = "p-value"
    lc.x_axis.title = "Lag (개월)"
    lc.y_axis.scaling.min = 0
    lc.y_axis.scaling.max = 1.0

    data_ref = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=max_lag + 1)
    lc.add_data(data_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=2, max_row=max_lag + 1)
    lc.set_categories(cats)

    lc.series[0].graphicalProperties.line.solidFill = "4472C4"
    lc.series[1].graphicalProperties.line.solidFill = "FF0000"
    lc.series[1].graphicalProperties.line.dashDot = "dash"

    # α=0.05 기준선 — E열 이후 별도 컬럼에 0.05 값 배치
    alpha_col = 4
    ws.cell(row=1, column=alpha_col, value="α=0.05")
    ws.cell(row=1, column=alpha_col).font = _HEADER_FONT
    for lag in range(1, max_lag + 1):
        ws.cell(row=lag + 1, column=alpha_col, value=0.05)

    data_alpha = Reference(ws, min_col=alpha_col, min_row=1, max_row=max_lag + 1)
    lc.add_data(data_alpha, titles_from_data=True)
    lc.series[2].graphicalProperties.line.solidFill = "000000"
    lc.series[2].graphicalProperties.line.dashDot = "dot"

    ws.add_chart(lc, f"A{max_lag + 5}")


def _sheet_irf(wb: Workbook, results: dict) -> None:
    """IRF 시트: 충격반응함수 + 신뢰구간 꺾은선 차트."""
    ws = wb.create_sheet("IRF")
    var_res = results["var"]
    irf_vals = var_res["irf_values"]
    irf_lower = var_res["irf_lower"]
    irf_upper = var_res["irf_upper"]
    periods = config.IRF_PERIODS

    headers = ["기간(개월)", "IRF", "하한(95%)", "상한(95%)"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for i in range(periods + 1):
        r = i + 2
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=round(float(irf_vals[i]), 6))
        ws.cell(row=r, column=3, value=round(float(irf_lower[i]), 6))
        ws.cell(row=r, column=4, value=round(float(irf_upper[i]), 6))

    _auto_width(ws)

    n = periods + 1
    lc = LineChart()
    lc.title = f"IRF: 외환보유고 충격 → 환율 반응 (peak={var_res['peak_month']}개월)"
    lc.style = 10
    lc.height = 15
    lc.width = 28
    lc.y_axis.title = "환율 반응"
    lc.x_axis.title = "기간 (개월)"

    data_ref = Reference(ws, min_col=2, max_col=4, min_row=1, max_row=n + 1)
    lc.add_data(data_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=2, max_row=n + 1)
    lc.set_categories(cats)

    # IRF 실선 파랑, 신뢰구간 회색 점선
    lc.series[0].graphicalProperties.line.solidFill = "4472C4"
    lc.series[0].graphicalProperties.line.width = 25000
    lc.series[1].graphicalProperties.line.solidFill = "A0A0A0"
    lc.series[1].graphicalProperties.line.dashDot = "dash"
    lc.series[2].graphicalProperties.line.solidFill = "A0A0A0"
    lc.series[2].graphicalProperties.line.dashDot = "dash"

    ws.add_chart(lc, f"A{n + 4}")


def _sheet_fevd(wb: Workbook, results: dict) -> None:
    """FEVD 시트: 분산분해 + 누적 막대 차트."""
    ws = wb.create_sheet("FEVD")
    var_res = results["var"]
    fevd_obj = var_res["fevd_obj"]
    from engine.var_model import FX_COL, RESERVES_COL
    cols = [RESERVES_COL, FX_COL]
    fx_idx = cols.index(FX_COL)
    res_idx = cols.index(RESERVES_COL)
    periods = config.IRF_PERIODS

    headers = ["기간(개월)", "외환보유고 기여(%)", "기타 기여(%)"]
    for col, h in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for t in range(1, periods + 1):
        r = t + 1
        res_pct = float(fevd_obj.decomp[fx_idx, t - 1, res_idx]) * 100
        other_pct = 100 - res_pct
        ws.cell(row=r, column=1, value=t)
        ws.cell(row=r, column=2, value=round(res_pct, 2))
        ws.cell(row=r, column=3, value=round(other_pct, 2))

    _auto_width(ws)

    bc = BarChart()
    bc.type = "col"
    bc.grouping = "stacked"
    bc.overlap = 100
    bc.title = f"FEVD: 환율 변동성 분산 분해 (12개월 기여: {var_res['fevd_pct']:.1f}%)"
    bc.style = 10
    bc.height = 15
    bc.width = 28
    bc.y_axis.title = "기여 비율 (%)"
    bc.x_axis.title = "기간 (개월)"
    bc.y_axis.scaling.min = 0
    bc.y_axis.scaling.max = 100

    data_ref = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=periods + 1)
    bc.add_data(data_ref, titles_from_data=True)
    cats = Reference(ws, min_col=1, min_row=2, max_row=periods + 1)
    bc.set_categories(cats)

    bc.series[0].graphicalProperties.solidFill = "4472C4"   # 파랑
    bc.series[1].graphicalProperties.solidFill = "D3D3D3"   # 회색

    ws.add_chart(bc, f"A{periods + 5}")


def _sheet_var_params(wb: Workbook, results: dict) -> None:
    """VAR계수 시트."""
    ws = wb.create_sheet("VAR계수")
    params_df: pd.DataFrame = results["var"]["var_params"]

    ws.cell(row=1, column=1, value="VAR 모델 계수 행렬").font = _TITLE_FONT
    ws.cell(row=2, column=1, value="변수")
    for col_idx, col_name in enumerate(params_df.columns, start=2):
        ws.cell(row=2, column=col_idx, value=col_name)
    _style_header(ws, 2, len(params_df.columns) + 1)

    for r_idx, (idx_val, row_vals) in enumerate(params_df.iterrows(), start=3):
        ws.cell(row=r_idx, column=1, value=str(idx_val))
        for col_idx, val in enumerate(row_vals, start=2):
            ws.cell(row=r_idx, column=col_idx, value=round(float(val), 6))

    _auto_width(ws)


# ---------------------------------------------------------------------------
# 메인 진입점
# ---------------------------------------------------------------------------

def generate_excel(
    data: dict,
    results: dict,
    chart_path: str | None = None,  # 더 이상 사용 안 함 (네이티브 차트로 대체)
    output_path: str | None = None,
) -> str:
    """분석 결과를 네이티브 Excel 차트 포함 파일로 내보낸다."""
    if output_path is None:
        output_path = str(config.EXCEL_PATH)

    wb = Workbook()
    # 기본 시트 제거
    if "Sheet" in wb.sheetnames:
        del wb["Sheet"]

    _sheet_summary(wb, data, results)
    _sheet_timeseries(wb, data)
    _sheet_delta_return(wb, data, results)
    _sheet_granger(wb, results)
    _sheet_irf(wb, results)
    _sheet_fevd(wb, results)
    _sheet_var_params(wb, results)

    wb.save(output_path)
    logger.info("Excel 저장 완료 (네이티브 차트 포함): %s", output_path)
    return output_path
