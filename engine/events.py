"""주요 환율 이벤트 기간 상수 및 차트 음영 헬퍼."""

import matplotlib.pyplot as plt
import pandas as pd

# 이벤트 기간 정의 (start/end: YYYY-MM 형식)
EVENTS: list[dict] = [
    {
        "start": "1997-07",
        "end": "1998-12",
        "label": "IMF 외환위기",
        "color": "red",
        "alpha": 0.12,
    },
    {
        "start": "2008-09",
        "end": "2009-03",
        "label": "글로벌 금융위기",
        "color": "orange",
        "alpha": 0.15,
    },
    {
        "start": "2020-03",
        "end": "2020-06",
        "label": "코로나 쇼크",
        "color": "green",
        "alpha": 0.15,
    },
    {
        "start": "2022-01",
        "end": "2022-10",
        "label": "Fed 긴축(킹달러)",
        "color": "purple",
        "alpha": 0.12,
    },
    {
        "start": "2026-03",
        "end": "2026-04",
        "label": "중동 원유 충격",
        "color": "saddlebrown",
        "alpha": 0.15,
    },
]


def apply_event_shading(ax: plt.Axes) -> None:
    """주어진 Axes에 이벤트 기간 음영을 적용한다.

    ax.xaxis가 Timestamp 기반이라고 가정한다.
    """
    for evt in EVENTS:
        start_ts = pd.Period(evt["start"], "M").to_timestamp()
        end_ts = pd.Period(evt["end"], "M").to_timestamp("D", "E")
        ax.axvspan(
            start_ts,
            end_ts,
            color=evt["color"],
            alpha=evt["alpha"],
            label=evt["label"],
        )
