"""pytest 테스트 — 6개 케이스.

외부 API 호출은 mock 처리, engine 함수는 합성 데이터로 검증한다.
"""

import io
import unittest.mock as mock

import numpy as np
import pandas as pd
import pytest


# --------------------------------------------------------------------------
# Test 1: ECOS 컬렉터 — API 키 없을 때 RuntimeError
# --------------------------------------------------------------------------
def test_ecos_missing_key():
    from pipeline.ecos_collector import fetch_fx_reserves

    with pytest.raises(RuntimeError, match="BOK_API_KEY"):
        fetch_fx_reserves(api_key="")


# --------------------------------------------------------------------------
# Test 2: FX 컬렉터 — 반환값이 PeriodIndex(freq='M') 인지 확인
# --------------------------------------------------------------------------
def test_fx_collector_period_index():
    import pandas as pd
    import yfinance as yf

    # 2년치 가짜 일별 데이터 생성
    dates = pd.date_range("2022-01-01", "2023-12-31", freq="B")
    fake_close = pd.Series(
        np.random.uniform(1200, 1400, len(dates)), index=dates, name="Close"
    )
    # yfinance는 MultiIndex 컬럼 DataFrame 반환
    fake_df = pd.DataFrame({"Close": fake_close.values}, index=dates)
    fake_df.columns = pd.MultiIndex.from_tuples([("Close", "USDKRW=X")])

    with mock.patch("yfinance.download", return_value=fake_df):
        from pipeline.fx_collector import fetch_usdkrw

        result = fetch_usdkrw(start="202201", end="202312")

    assert isinstance(result.index, pd.PeriodIndex), "인덱스가 PeriodIndex여야 합니다"
    assert result.index.freqstr == "M", "주기가 M(월별)이어야 합니다"
    assert len(result) > 0, "데이터가 비어있으면 안 됩니다"


# --------------------------------------------------------------------------
# Test 3: ADF — 랜덤워크(비정상)는 needs_diff=True
# --------------------------------------------------------------------------
def test_adf_needs_diff():
    np.random.seed(42)
    random_walk = pd.Series(np.cumsum(np.random.randn(150)))

    from engine.unit_root import run_adf

    result = run_adf(random_walk, name="random_walk")
    assert result["needs_diff"], (
        f"랜덤워크는 needs_diff=True여야 합니다 (p={result['p_value']:.4f})"
    )
    assert "diff_series" in result
    assert len(result["diff_series"]) == len(random_walk) - 1


# --------------------------------------------------------------------------
# Test 4: 피어슨 — 합성 음의 상관 데이터
# --------------------------------------------------------------------------
def test_pearson_negative():
    np.random.seed(0)
    n = 100
    x = pd.Series(np.random.randn(n))
    y = -2.0 * x + 0.5 * np.random.randn(n)  # 강한 음의 상관

    from engine.correlation import run_pearson

    result = run_pearson(x, y)
    assert result["r"] < -0.8, f"r은 -0.8 미만이어야 합니다 (실제: {result['r']:.4f})"
    assert result["p_value"] < 0.001
    assert result["n"] == n


# --------------------------------------------------------------------------
# Test 5: Granger — y = lag(x, 2) + noise → x→y 방향 유의
# --------------------------------------------------------------------------
def test_granger_synthetic():
    np.random.seed(7)
    n = 120
    x = np.random.randn(n)
    y = np.zeros(n)
    for i in range(2, n):
        y[i] = 0.6 * x[i - 2] + 0.15 * np.random.randn()

    x_s = pd.Series(x)
    y_s = pd.Series(y)

    from engine.granger import run_granger

    result = run_granger(x_s, y_s, max_lag=4)

    assert result["sig_lag_x_to_y"] is not None, (
        "x→y 방향에서 유의한 lag가 있어야 합니다 "
        f"(p-values: {result['x_to_y']})"
    )
    assert result["sig_p_x_to_y"] < 0.05


# --------------------------------------------------------------------------
# Test 6: VAR — 반환 dict 키 완전성 검증
# --------------------------------------------------------------------------
def test_var_keys():
    np.random.seed(99)
    n = 80
    # 간단한 VAR(1) 프로세스
    data = np.zeros((n, 2))
    for t in range(1, n):
        data[t] = 0.3 * data[t - 1] + np.random.randn(2) * 0.5

    df = pd.DataFrame(data, columns=["reserves_delta", "fx_return"])

    from engine.var_model import run_var

    result = run_var(df)

    required_keys = {"optimal_lag", "aic", "irf_values", "irf_lower", "irf_upper", "peak_month", "fevd_pct", "var_params"}
    missing = required_keys - set(result.keys())
    assert not missing, f"반환 dict에 키가 없습니다: {missing}"

    assert isinstance(result["optimal_lag"], int)
    assert 0 <= result["peak_month"] <= 12
    assert 0.0 <= result["fevd_pct"] <= 100.0
