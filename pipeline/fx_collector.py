"""BOK ECOS API → 원/달러 환율 월별 시계열 수집.

지표: 3.1.2.1. 주요국 통화 대원화환율 (731Y004, 항목 0000001)
단위: 원/달러 (월별 평균 매매기준율)
"""

import logging

import pandas as pd
import requests

import config

logger = logging.getLogger(__name__)


def fetch_usdkrw(
    api_key: str | None = None,
    start: str = config.ANALYSIS_START,
    end: str | None = None,
) -> pd.Series:
    """BOK ECOS에서 원/달러 환율 월별 데이터를 수집한다.

    Args:
        api_key: BOK ECOS API 키. None이면 config.BOK_API_KEY 사용.
        start:   시작 연월 YYYYMM (기본값: config.ANALYSIS_START).
        end:     종료 연월 YYYYMM. None이면 직전 완성 월로 자동 설정.

    Returns:
        pd.Series — index: pd.PeriodIndex(freq='M'), values: 원/달러 (float)

    Raises:
        RuntimeError: API 오류.
    """
    if api_key is None:
        api_key = config.BOK_API_KEY

    if not api_key:
        raise RuntimeError("BOK_API_KEY가 설정되지 않았습니다.")

    if end is None:
        end = str(pd.Period.now("M") - 1).replace("-", "")

    row_count = 1000
    url = (
        f"{config.ECOS_BASE_URL}/StatisticSearch"
        f"/{api_key}/json/kr/1/{row_count}"
        f"/{config.ECOS_FX_STAT_CODE}/M/{start}/{end}"
        f"/{config.ECOS_FX_ITEM_CODE}/{config.ECOS_FX_ITEM_CODE2}"
    )

    logger.info("ECOS 원/달러 환율 수집 중: %s ~ %s", start, end)
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as exc:
        raise RuntimeError(f"ECOS 환율 API 요청 실패: {exc}") from exc

    if "RESULT" in data:
        code = data["RESULT"].get("CODE", "")
        msg = data["RESULT"].get("MESSAGE", "")
        raise RuntimeError(f"ECOS API 오류 [{code}]: {msg}")

    if "StatisticSearch" not in data:
        raise RuntimeError(f"ECOS API 예상치 못한 응답: {list(data.keys())}")

    rows = data["StatisticSearch"].get("row", [])
    if not rows:
        raise RuntimeError("ECOS 환율 응답에 데이터가 없습니다.")

    times, values = [], []
    for row in rows:
        raw_val = row.get("DATA_VALUE", "").strip()
        if not raw_val:
            continue
        try:
            times.append(pd.Period(row["TIME"], "M"))
            values.append(float(raw_val))
        except (KeyError, ValueError) as exc:
            logger.debug("환율 행 파싱 건너뜀 (%s): %s", row, exc)

    series = pd.Series(values, index=pd.PeriodIndex(times, freq="M"), name="usdkrw")
    series = series.sort_index()

    # 결측 확인 후 ffill
    n_missing = int(series.isna().sum())
    if n_missing:
        logger.warning("원/달러 환율 %d개월 결측 → forward-fill 적용", n_missing)
        series = series.ffill()

    logger.info(
        "원/달러 환율 수집 완료: %s ~ %s (%d개월)",
        series.index[0],
        series.index[-1],
        len(series),
    )
    return series
