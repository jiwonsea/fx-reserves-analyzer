"""Central configuration: env vars, constants, logging bootstrap.

Import this module first in main.py so that dotenv and logging are
initialised before any pipeline/engine module is imported.
"""

import logging
import os
from pathlib import Path

from dotenv import load_dotenv

# Load .env from project root (safe to call multiple times)
load_dotenv(Path(__file__).parent / ".env")

# ---------------------------------------------------------------------------
# API credentials
# ---------------------------------------------------------------------------
BOK_API_KEY: str | None = os.getenv("BOK_API_KEY")

# ---------------------------------------------------------------------------
# ECOS API constants
# ---------------------------------------------------------------------------
ECOS_BASE_URL = "https://ecos.bok.or.kr/api"

# 외환보유액 월별 — 3.5. 외환보유액 합계 (ECOS 코드: 732Y001, 항목코드: 99)
# 단위: 천달러 → 변환 후 억달러로 사용 (ecos_collector에서 /100_000)
ECOS_STAT_CODE = "732Y001"
ECOS_ITEM_CODE = "99"   # 합계 (금+SDR+IMF포지션+외환)

# 원/달러 환율 월별 — 3.1.2.1. 주요국 통화 대원화환율 (ECOS 코드: 731Y004, 항목코드: 0000001)
# 단위: 원/달러, 월별 평균
ECOS_FX_STAT_CODE = "731Y004"
ECOS_FX_ITEM_CODE = "0000001"   # 원/미국달러(매매기준율)
ECOS_FX_ITEM_CODE2 = "0000200"  # 0000100=월평균, 0000200=월말종가 (원래 spec: 월말 기준)

# ---------------------------------------------------------------------------
# Analysis parameters
# ---------------------------------------------------------------------------
ANALYSIS_START = "199501"   # 1995-01 (ECOS 가용 시작 기준)
GRANGER_MAX_LAG = 6         # 최대 lag 6개월
VAR_MAX_LAG = 6             # VAR order 탐색 상한
IRF_PERIODS = 12            # IRF 반응 기간 (개월)
IRF_BOOTSTRAP_REPL = 200    # bootstrap 반복 횟수

# ---------------------------------------------------------------------------
# Output paths
# ---------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

CHART_PATH = OUTPUT_DIR / "fx_reserves_analysis.png"
EXCEL_PATH = OUTPUT_DIR / "fx_reserves_report.xlsx"

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    level=logging.INFO,
    datefmt="%Y-%m-%d %H:%M:%S",
)
