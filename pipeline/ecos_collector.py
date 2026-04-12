"""BOK ECOS API → 외환보유액 월별 시계열 수집.

엔드포인트 형식:
  {ECOS_BASE_URL}/StatisticSearch/{api_key}/json/kr/1/{row_count}/{stat_code}/M/{start}/{end}

응답:
  {"StatisticSearch": {"list_total_count": "N", "row": [{TIME, DATA_VALUE}, ...]}}
인증 실패 시 HTTP 200 + {"RESULT": {"CODE": "AUTH-...", "MESSAGE": "..."}} 반환
"""

import logging

import pandas as pd
import requests

import config

logger = logging.getLogger(__name__)


def fetch_fx_reserves(
    api_key: str,
    start: str = config.ANALYSIS_START,
    end: str | None = None,
) -> pd.Series:
    """BOK ECOS에서 외환보유액 월별 데이터를 수집한다.

    Args:
        api_key: BOK ECOS API 키.
        start:   시작 연월 YYYYMM (기본값: config.ANALYSIS_START).
        end:     종료 연월 YYYYMM. None이면 직전 완성 월(당월 제외)로 자동 설정.

    Returns:
        pd.Series — index: pd.PeriodIndex(freq='M'), values: 억달러 (float)

    Raises:
        RuntimeError: API 키 미설정 또는 ECOS 응답 오류.
    """
    if not api_key:
        raise RuntimeError(
            "BOK_API_KEY가 설정되지 않았습니다. "
            ".env.example을 .env로 복사하고 API 키를 입력해 주세요.\n"
            "발급: https://ecos.bok.or.kr/api/#/user/login"
        )

    if end is None:
        end = str(pd.Period.now("M") - 1).replace("-", "")  # e.g. "202503"

    row_count = 1000  # ECOS 1회 요청 최대 행 수
    url = (
        f"{config.ECOS_BASE_URL}/StatisticSearch"
        f"/{api_key}/json/kr/1/{row_count}"
        f"/{config.ECOS_STAT_CODE}/M/{start}/{end}"
        f"/{config.ECOS_ITEM_CODE}"
    )

    logger.info("ECOS 외환보유액 수집 중: %s ~ %s", start, end)
    try:
        resp = requests.get(url, timeout=30)
        resp.raise_for_status()
        data = resp.json()
    except requests.RequestException as exc:
        raise RuntimeError(f"ECOS API 요청 실패: {exc}") from exc

    # 인증 실패 / 데이터 없음 감지 (HTTP 200이어도 RESULT 키로 에러 반환)
    if "RESULT" in data:
        code = data["RESULT"].get("CODE", "")
        msg = data["RESULT"].get("MESSAGE", "")
        raise RuntimeError(f"ECOS API 오류 [{code}]: {msg}")

    if "StatisticSearch" not in data:
        raise RuntimeError(f"ECOS API 예상치 못한 응답 구조: {list(data.keys())}")

    rows = data["StatisticSearch"].get("row", [])
    if not rows:
        raise RuntimeError(
            f"ECOS 응답에 데이터가 없습니다. "
            f"stat_code={config.ECOS_STAT_CODE}, start={start}, end={end}\n"
            "config.py의 ECOS_STAT_CODE를 확인해 주세요."
        )

    # 파싱: TIME(YYYYMM) → Period, DATA_VALUE(str) → float
    # ECOS 732Y001 단위: 천달러 → 억달러로 변환 (/100_000)
    times, values = [], []
    for row in rows:
        raw_val = row.get("DATA_VALUE", "").strip()
        if not raw_val:
            continue  # 결측 월 스킵
        try:
            times.append(pd.Period(row["TIME"], "M"))
            values.append(float(raw_val) / 100_000)  # 천달러 → 억달러
        except (KeyError, ValueError) as exc:
            logger.debug("ECOS 행 파싱 건너뜀 (%s): %s", row, exc)

    series = pd.Series(values, index=pd.PeriodIndex(times, freq="M"), name="fx_reserves")
    series = series.sort_index()

    logger.info(
        "외환보유액 수집 완료: %s ~ %s (%d개월)",
        series.index[0],
        series.index[-1],
        len(series),
    )
    return series
