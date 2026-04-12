"""openpyxl Excel 리포트 생성 — 3개 시트 + 차트 PNG 삽입."""

import logging
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import config

logger = logging.getLogger(__name__)

# 헤더 스타일
_HEADER_FONT = Font(bold=True, size=10)
_HEADER_FILL = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
_TITLE_FONT = Font(bold=True, size=12)


def _style_header_row(ws, row: int, ncols: int) -> None:
    for col in range(1, ncols + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = Alignment(horizontal="center")


def _auto_col_width(ws) -> None:
    for col_cells in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value is not None:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)


def generate_excel(
    data: dict,
    results: dict,
    chart_path: str,
    output_path: str | None = None,
) -> str:
    """분석 결과를 Excel로 내보낸다.

    Args:
        data:        pipeline 수집 데이터 (reserves, usdkrw, etc.)
        results:     engine 분석 결과 (pearson, granger, var, adf_*)
        chart_path:  삽입할 차트 PNG 경로.
        output_path: 저장 경로. None이면 config.EXCEL_PATH 사용.

    Returns:
        저장된 Excel 파일 경로 (str).
    """
    if output_path is None:
        output_path = str(config.EXCEL_PATH)

    wb = Workbook()

    # =========================================================
    # Sheet 1: 요약 통계
    # =========================================================
    ws1 = wb.active
    ws1.title = "Summary"

    reserves = data["reserves"]
    usdkrw = data["usdkrw"]
    pearson = results["pearson"]
    granger = results["granger"]
    var_res = results["var"]
    adf_res = results.get("adf_reserves", {})
    adf_fx = results.get("adf_fx", {})

    start_str = str(reserves.index[0])
    end_str = str(reserves.index[-1])

    summary_rows = [
        ("항목", "값"),
        ("분석 기간 시작", start_str),
        ("분석 기간 종료", end_str),
        ("총 관측 월수", f"{len(reserves)}개월"),
        ("", ""),
        ("외환보유고 ADF p-value", f"{adf_res.get('p_value', 'N/A'):.4f}" if adf_res else "N/A"),
        ("환율 ADF p-value", f"{adf_fx.get('p_value', 'N/A'):.4f}" if adf_fx else "N/A"),
        ("차분 여부", "1차 차분 적용"),
        ("", ""),
        ("피어슨 상관계수 (r)", f"{pearson['r']:.4f}"),
        ("피어슨 p-value", f"{pearson['p_value']:.4f}"),
        ("피어슨 n", str(pearson["n"])),
        ("", ""),
        (
            "Granger: 보유고→환율 유의 최소 lag",
            f"{granger['sig_lag_x_to_y']}개월 (p={granger['sig_p_x_to_y']:.4f})"
            if granger["sig_lag_x_to_y"]
            else "유의하지 않음",
        ),
        (
            "Granger: 환율→보유고 유의 최소 lag",
            f"{granger['sig_lag_y_to_x']}개월 (p={granger['sig_p_y_to_x']:.4f})"
            if granger["sig_lag_y_to_x"]
            else "유의하지 않음",
        ),
        ("", ""),
        ("VAR 최적 lag (AIC 기준)", f"{var_res['optimal_lag']}개월"),
        ("IRF peak 기간", f"{var_res['peak_month']}개월 후"),
        ("FEVD 외환보유고 기여 (12개월)", f"{var_res['fevd_pct']:.2f}%"),
    ]

    # 헤더 스타일
    ws1.cell(row=1, column=1, value="항목").font = _HEADER_FONT
    ws1.cell(row=1, column=1).fill = _HEADER_FILL
    ws1.cell(row=1, column=2, value="값").font = _HEADER_FONT
    ws1.cell(row=1, column=2).fill = _HEADER_FILL

    for r_idx, (key, val) in enumerate(summary_rows, start=1):
        ws1.cell(row=r_idx, column=1, value=key)
        ws1.cell(row=r_idx, column=2, value=val)
        if r_idx == 1:
            _style_header_row(ws1, r_idx, 2)

    _auto_col_width(ws1)

    # 차트 이미지 삽입
    img_row = len(summary_rows) + 3
    try:
        img = XLImage(str(chart_path))
        img.width = 900
        img.height = 560
        ws1.add_image(img, f"A{img_row}")
    except Exception as exc:
        logger.warning("차트 이미지 삽입 실패: %s", exc)

    # =========================================================
    # Sheet 2: Granger 검정 결과
    # =========================================================
    ws2 = wb.create_sheet("Granger")
    headers2 = ["Lag (개월)", "방향", "p-value", "유의 (α=0.05)"]
    for col_idx, h in enumerate(headers2, start=1):
        ws2.cell(row=1, column=col_idx, value=h)
    _style_header_row(ws2, 1, len(headers2))

    row2 = 2
    for lag in range(1, config.GRANGER_MAX_LAG + 1):
        p_xy = granger["x_to_y"][lag]
        p_yx = granger["y_to_x"][lag]
        for direction, p_val in [("보유고 → 환율", p_xy), ("환율 → 보유고", p_yx)]:
            ws2.cell(row=row2, column=1, value=lag)
            ws2.cell(row=row2, column=2, value=direction)
            ws2.cell(row=row2, column=3, value=round(p_val, 4))
            ws2.cell(row=row2, column=4, value="Y" if p_val < 0.05 else "N")
            row2 += 1

    _auto_col_width(ws2)

    # =========================================================
    # Sheet 3: VAR 계수
    # =========================================================
    ws3 = wb.create_sheet("VAR_Coefficients")
    params_df: pd.DataFrame = var_res["var_params"]

    ws3.cell(row=1, column=1, value="VAR 모델 계수 행렬").font = _TITLE_FONT

    # 컬럼 헤더
    ws3.cell(row=2, column=1, value="변수")
    for col_idx, col_name in enumerate(params_df.columns, start=2):
        ws3.cell(row=2, column=col_idx, value=col_name)
    _style_header_row(ws3, 2, len(params_df.columns) + 1)

    # 데이터
    for r_idx, (idx_val, row_vals) in enumerate(params_df.iterrows(), start=3):
        ws3.cell(row=r_idx, column=1, value=str(idx_val))
        for col_idx, val in enumerate(row_vals, start=2):
            ws3.cell(row=r_idx, column=col_idx, value=round(float(val), 6))

    _auto_col_width(ws3)

    wb.save(output_path)
    logger.info("Excel 저장 완료: %s", output_path)
    return output_path
