# -*- coding: utf-8 -*-
"""
18_QUANTILE_RANK_IC_5BIN_ANALYSIS.py
===================================

부가 연구: 기존 1~N(약 200종목) 연속 Rank IC를 5개 분위로 압축한
"5-Quantile Rank IC (QIC5)"를 계산한다.

핵심 아이디어
-------------
기존:
    Signal 종목순위 1~200
    vs
    Future Return 종목순위 1~200
    -> Daily Spearman Rank IC

신규 부가연구:
    Signal 순위를 Q1~Q5로 압축
    Future Return 순위도 Q1~Q5로 압축
    -> Q1~Q5 라벨끼리 Spearman correlation

정확한 방향
-----------
Reversal Signal = - 과거 누적수익률

Signal:
    Q1 = Signal 가장 낮음 = 최근 winner
    Q5 = Signal 가장 높음 = 최근 loser

Future Return:
    Q1 = 미래수익률 가장 낮음
    Q5 = 미래수익률 가장 높음

따라서:
    Signal Q1 -> Future Q1
    Signal Q5 -> Future Q5
처럼 "같은 분위"에 위치할수록 reversal ranking이 맞은 것이다.

예:
    과거 winner(Q1)가 이후 상대적으로 낮은 수익률(Q1)
    과거 loser(Q5)가 이후 상대적으로 높은 수익률(Q5)
    -> 높은 양(+)의 QIC5

중요
----
이 QIC5는 기존 Continuous Rank IC를 대체하기보다 "부가 연구"로 쓴다.
5개 분위로 압축하므로 정보가 줄고 ties가 많이 생기며,
기존 Rank IC보다 극단/구간 수준의 순위 유지 여부를 더 직접적으로 본다.

같이 저장하는 보조지표
----------------------
1) ContinuousRankIC
   - 기존 방식과 동일한 1~N Spearman
   - 같은 pair-universe에서 재계산하여 직접 비교 가능

2) QIC5
   - Signal Q1~Q5 vs Future Q1~Q5 Spearman

3) SameQ_Hit
   - Signal 분위 == Future 분위 비율
   - 독립적이고 5분위가 균등하면 기대값 약 20%

4) Q1_to_Q1
   - Signal Q1 종목 중 Future Q1이 된 비율

5) Q5_to_Q5
   - Signal Q5 종목 중 Future Q5가 된 비율

6) MeanAbsQDistance
   - 평균 |SignalQ - FutureQ|
   - 작을수록 분위 일치가 강함

7) 5x5 Transition Matrix
   - 각 Signal Q가 미래에 Q1~Q5 중 어디로 이동했는지 집계

8) DeltaQIC5 vs Raw
   - PCA QIC5 - 같은 Formation_N의 Raw QIC5
   - 예: W252_PC123_N5 - RAW_N5

통계
----
Daily QIC5 시계열 평균에 대해 Newey-West HAC t/p를 계산한다.
보수적으로:
    HAC lag = max(Formation_N, Future_Horizon) - 1

입력 Excel 5개
--------------
1) 00_COMMON252_BASE.xlsx
   sheets:
     COMMON252_ELIGIBLE
     RAW_CUM_LOG_5
     RAW_CUM_LOG_10
     RAW_CUM_LOG_30
     RAW_CUM_LOG_60

2) 02_PCA_W60_CUMULATIVE.xlsx
   sheets:
     PC1_CUM5/10/30/60
     PC12_CUM5/10/30/60
     PC123_CUM5/10/30/60

3) 04_PCA_W120_CUMULATIVE.xlsx
4) 06_PCA_W252_CUMULATIVE.xlsx
   same sheet structure

5) FUTURE_LOG_RETURNS_5_10_20_60.xlsx
   sheets:
     FWD_LOG_5
     FWD_LOG_10
     FWD_LOG_20
     FWD_LOG_60

불필요한 파일
-------------
- STATE_IC_TRAIN_LIBRARY.xlsx
- PCA DAILY Excel
- 수정주가 Full History
- 기존 IC_FINAL_ANALYSIS.xlsx
(이번 계산은 stock-level signal + stock-level future return에서 새로 직접 계산)

출력
----
QUANTILE_RANK_IC_5BIN_ANALYSIS/
  QUANTILE_RANK_IC_5BIN_ANALYSIS.xlsx

Sheets
------
PARAMETERS
CASE_DEFINITION
QIC_SUMMARY
TRANSITION_MATRIX
DAILY_FWD5
DAILY_FWD10
DAILY_FWD20
DAILY_FWD60

Daily sheet 주요 열
-------------------
Date
CaseID
PCA_Window
PC_Remove
Formation_N
Horizon
N
ContinuousRankIC
QIC5
SameQ_Hit
Q1_to_Q1
Q5_to_Q5
MeanAbsQDistance
Raw_Case
Raw_QIC5
DeltaQIC5_vs_Raw

Case 구성
---------
Raw 4:
    RAW_N5 / RAW_N10 / RAW_N30 / RAW_N60

PCA 36:
    W={60,120,252}
    x Remove={PC1,PC12,PC123}
    x N={5,10,30,60}

Total = 40 cases
Horizon = {5,10,20,60}
Total summary combinations = 160

실행
----
python 18_QUANTILE_RANK_IC_5BIN_ANALYSIS.py

또는 입력파일들이 다른 상위폴더에 있으면:
python 18_QUANTILE_RANK_IC_5BIN_ANALYSIS.py --data-dir "C:\\path\\to\\folder"

"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import numpy as np


# =============================================================================
# CONSTANTS
# =============================================================================

FORMATION_NS = (5, 10, 30, 60)
HORIZONS = (5, 10, 20, 60)
PCA_WINDOWS = (60, 120, 252)
PC_REMOVES = ("PC1", "PC12", "PC123")

MIN_STOCKS = 25
N_QUANTILES = 5

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EPS = 1e-15


# =============================================================================
# BASIC HELPERS
# =============================================================================

def idx_to_col(idx_zero_based: int) -> str:
    n = idx_zero_based + 1
    s = ""

    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s

    return s


def col_to_idx(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref)

    if not m:
        raise ValueError(cell_ref)

    n = 0

    for ch in m.group(1):
        n = n * 26 + ord(ch) - 64

    return n - 1


def normalize_ticker(value) -> str:
    s = str(value).strip()

    if s.lower() in {"", "none", "nan"}:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    if s.isdigit():
        s = s.zfill(6)

    return s.upper()


def normalize_filename_token(s: str) -> str:
    s = s.lower()
    s = re.sub(r"\(\d+\)", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def find_input_file(
    base: Path,
    key: str,
) -> Path:
    key_norm = normalize_filename_token(
        key
    )

    candidates = []

    for p in base.rglob(
        "*.xlsx"
    ):
        name = normalize_filename_token(
            p.name
        )

        if key_norm in name:
            if (
                "quantile_rank_ic_5bin_analysis"
                in name
            ):
                continue

            candidates.append(
                p
            )

    if not candidates:
        raise FileNotFoundError(
            f"\n필수 파일을 찾지 못했습니다: {key}\n"
            f"검색폴더: {base}\n"
        )

    return max(
        candidates,
        key=lambda p: p.stat().st_mtime,
    )


def excel_serial_to_yyyymmdd(
    value,
) -> int:
    from datetime import datetime, timedelta

    d = (
        datetime(
            1899,
            12,
            30,
        )
        + timedelta(
            days=float(
                value
            )
        )
    )

    return int(
        d.strftime(
            "%Y%m%d"
        )
    )


def parse_date_value(
    value,
) -> int:
    if (
        value is None
        or value == ""
    ):
        raise ValueError(
            "empty date"
        )

    if isinstance(
        value,
        (
            int,
            float,
            np.integer,
            np.floating,
        ),
    ):
        x = float(
            value
        )

        if (
            20000
            <= x
            <= 80000
        ):
            return excel_serial_to_yyyymmdd(
                x
            )

        if (
            19000101
            <= x
            <= 21001231
        ):
            return int(
                x
            )

    s = str(
        value
    ).strip()

    if re.fullmatch(
        r"\d{8}",
        s,
    ):
        return int(
            s
        )

    if re.fullmatch(
        r"\d+(\.\d+)?",
        s,
    ):
        x = float(
            s
        )

        if (
            20000
            <= x
            <= 80000
        ):
            return excel_serial_to_yyyymmdd(
                x
            )

    parts = re.findall(
        r"\d+",
        s,
    )

    if len(
        parts
    ) >= 3:
        y, m, d = map(
            int,
            parts[
                :3
            ],
        )

        return (
            y
            * 10000
            + m
            * 100
            + d
        )

    raise ValueError(
        f"날짜 파싱 실패: {value!r}"
    )


def yyyymmdd_text(
    x: int,
) -> str:
    s = str(
        int(
            x
        )
    )

    return (
        f"{s[:4]}-"
        f"{s[4:6]}-"
        f"{s[6:8]}"
    )


# =============================================================================
# XLSX STREAM READER
# =============================================================================

def workbook_sheet_targets(
    path: Path,
) -> dict[str, str]:
    with ZipFile(
        path
    ) as z:
        wb = ET.fromstring(
            z.read(
                "xl/workbook.xml"
            )
        )

        rels = ET.fromstring(
            z.read(
                "xl/_rels/workbook.xml.rels"
            )
        )

        relmap = {
            r.attrib[
                "Id"
            ]: r.attrib[
                "Target"
            ]
            for r in rels
        }

        out = {}

        sheets = wb.find(
            f"{{{NS_MAIN}}}sheets"
        )

        for s in sheets:
            rid = s.attrib[
                f"{{{NS_REL}}}id"
            ]

            target = relmap[
                rid
            ]

            if target.startswith(
                "/"
            ):
                target = target.lstrip(
                    "/"
                )
            elif not target.startswith(
                "xl/"
            ):
                target = (
                    "xl/"
                    + target
                )

            out[
                s.attrib[
                    "name"
                ]
            ] = target

        return out


def load_shared_strings(
    z: ZipFile,
) -> list[str]:
    if (
        "xl/sharedStrings.xml"
        not in z.namelist()
    ):
        return []

    root = ET.fromstring(
        z.read(
            "xl/sharedStrings.xml"
        )
    )

    return [
        "".join(
            (
                t.text
                or ""
            )
            for t in si.iter(
                f"{{{NS_MAIN}}}t"
            )
        )
        for si in root
    ]


def cell_value(
    cell: ET.Element,
    shared: list[str],
):
    typ = cell.attrib.get(
        "t"
    )

    if typ == "inlineStr":
        t = cell.find(
            f".//{{{NS_MAIN}}}t"
        )

        return (
            t.text
            if t is not None
            else ""
        )

    v = cell.find(
        f"{{{NS_MAIN}}}v"
    )

    if v is None:
        return None

    if typ == "s":
        return shared[
            int(
                v.text
            )
        ]

    if typ == "b":
        return int(
            v.text
        )

    return v.text


def iter_sheet_rows(
    xlsx_path: Path,
    sheet_name: str,
):
    targets = workbook_sheet_targets(
        xlsx_path
    )

    if sheet_name not in targets:
        raise KeyError(
            f"{xlsx_path.name}: 시트 없음: {sheet_name}\n"
            f"실제 시트={list(targets)}"
        )

    target = targets[
        sheet_name
    ]

    with ZipFile(
        xlsx_path
    ) as z:
        shared = load_shared_strings(
            z
        )

        with z.open(
            target
        ) as f:
            for _, elem in ET.iterparse(
                f,
                events=(
                    "end",
                ),
            ):
                if (
                    elem.tag
                    != f"{{{NS_MAIN}}}row"
                ):
                    continue

                cells = {}

                for c in elem:
                    if (
                        c.tag
                        != f"{{{NS_MAIN}}}c"
                    ):
                        continue

                    ref = c.attrib.get(
                        "r"
                    )

                    if not ref:
                        continue

                    cells[
                        col_to_idx(
                            ref
                        )
                    ] = cell_value(
                        c,
                        shared,
                    )

                yield cells
                elem.clear()


def read_wide_matrix(
    path: Path,
    sheet_name: str,
) -> tuple[
    np.ndarray,
    list[str],
    np.ndarray,
]:
    it = iter_sheet_rows(
        path,
        sheet_name,
    )

    try:
        header = next(
            it
        )
    except StopIteration:
        raise RuntimeError(
            f"{path.name}/{sheet_name}: 빈 시트"
        )

    max_col = max(
        header.keys()
    )

    tickers = [
        normalize_ticker(
            header.get(
                c,
                "",
            )
        )
        for c in range(
            1,
            max_col + 1,
        )
    ]

    dates = []
    rows = []

    for cells in it:
        if 0 not in cells:
            continue

        try:
            d = parse_date_value(
                cells[
                    0
                ]
            )
        except Exception:
            continue

        row = np.full(
            len(
                tickers
            ),
            np.nan,
            dtype=np.float64,
        )

        for c, raw in cells.items():
            if c == 0:
                continue

            j = c - 1

            if (
                j < 0
                or j
                >= len(
                    tickers
                )
                or raw
                in (
                    None,
                    "",
                )
            ):
                continue

            try:
                row[
                    j
                ] = float(
                    raw
                )
            except Exception:
                pass

        dates.append(
            d
        )

        rows.append(
            row
        )

    if not rows:
        raise RuntimeError(
            f"{path.name}/{sheet_name}: 유효 데이터 없음"
        )

    return (
        np.asarray(
            dates,
            dtype=np.int32,
        ),
        tickers,
        np.vstack(
            rows
        ),
    )


def align_matrix(
    ref_dates: np.ndarray,
    ref_tickers: list[str],
    dates: np.ndarray,
    tickers: list[str],
    arr: np.ndarray,
    label: str,
) -> np.ndarray:
    date_map = {
        int(
            d
        ): i
        for i, d in enumerate(
            dates
        )
    }

    ticker_map = {
        t: j
        for j, t in enumerate(
            tickers
        )
        if t
    }

    out = np.full(
        (
            len(
                ref_dates
            ),
            len(
                ref_tickers
            ),
        ),
        np.nan,
        dtype=np.float64,
    )

    src_cols = []
    dst_cols = []

    for j, t in enumerate(
        ref_tickers
    ):
        sj = ticker_map.get(
            t
        )

        if sj is not None:
            src_cols.append(
                sj
            )
            dst_cols.append(
                j
            )

    if not src_cols:
        raise RuntimeError(
            f"{label}: 공통 ticker가 없습니다."
        )

    src_cols = np.asarray(
        src_cols,
        dtype=int,
    )

    dst_cols = np.asarray(
        dst_cols,
        dtype=int,
    )

    matched_dates = 0

    for di, d in enumerate(
        ref_dates
    ):
        si = date_map.get(
            int(
                d
            )
        )

        if si is None:
            continue

        out[
            di,
            dst_cols,
        ] = arr[
            si,
            src_cols,
        ]

        matched_dates += 1

    if matched_dates == 0:
        raise RuntimeError(
            f"{label}: 공통 날짜가 없습니다."
        )

    return out


# =============================================================================
# CASES
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    pca_window: int | None
    pc_remove: str
    formation_n: int
    source_key: str
    source_sheet: str


def build_cases() -> list[CaseDef]:
    out = []

    # Raw 4
    for n in FORMATION_NS:
        out.append(
            CaseDef(
                case_id=f"RAW_N{n}",
                pca_window=None,
                pc_remove="RAW",
                formation_n=n,
                source_key="base",
                source_sheet=f"RAW_CUM_LOG_{n}",
            )
        )

    # PCA 36
    for W in PCA_WINDOWS:
        for pc in PC_REMOVES:
            for n in FORMATION_NS:
                out.append(
                    CaseDef(
                        case_id=(
                            f"W{W}_{pc}_N{n}"
                        ),
                        pca_window=W,
                        pc_remove=pc,
                        formation_n=n,
                        source_key=f"cum{W}",
                        source_sheet=(
                            f"{pc}_CUM{n}"
                        ),
                    )
                )

    return out


# =============================================================================
# RANK / QUINTILE STATISTICS
# =============================================================================

def average_rank(
    x: np.ndarray,
) -> np.ndarray:
    """
    Tie-aware average rank, 1..N.
    Equivalent to scipy.stats.rankdata(method='average') without scipy.
    """
    x = np.asarray(
        x,
        dtype=np.float64,
    )

    n = len(
        x
    )

    order = np.argsort(
        x,
        kind="mergesort",
    )

    sorted_x = x[
        order
    ]

    ranks_sorted = np.empty(
        n,
        dtype=np.float64,
    )

    i = 0

    while i < n:
        j = i + 1

        while (
            j < n
            and sorted_x[
                j
            ]
            == sorted_x[
                i
            ]
        ):
            j += 1

        # positions i..j-1 correspond to ranks i+1..j
        avg = (
            (
                i + 1
            )
            + j
        ) / 2.0

        ranks_sorted[
            i:j
        ] = avg

        i = j

    out = np.empty(
        n,
        dtype=np.float64,
    )

    out[
        order
    ] = ranks_sorted

    return out


def pearson_corr(
    x: np.ndarray,
    y: np.ndarray,
) -> float:
    x = np.asarray(
        x,
        dtype=np.float64,
    )

    y = np.asarray(
        y,
        dtype=np.float64,
    )

    if (
        len(
            x
        ) < 2
        or len(
            y
        ) != len(
            x
        )
    ):
        return np.nan

    x0 = (
        x
        - np.mean(
            x
        )
    )

    y0 = (
        y
        - np.mean(
            y
        )
    )

    den = math.sqrt(
        float(
            np.dot(
                x0,
                x0,
            )
        )
        * float(
            np.dot(
                y0,
                y0,
            )
        )
    )

    if den <= EPS:
        return np.nan

    return float(
        np.dot(
            x0,
            y0,
        )
        / den
    )


def spearman_corr(
    x: np.ndarray,
    y: np.ndarray,
) -> float:
    return pearson_corr(
        average_rank(
            x
        ),
        average_rank(
            y
        ),
    )


def exact_five_quantiles(
    x: np.ndarray,
) -> np.ndarray:
    """
    Match the existing quintile-portfolio logic:
    stable sort -> np.array_split into 5 near-equal groups.

    Returns integer labels 1..5.
    """
    x = np.asarray(
        x,
        dtype=np.float64,
    )

    n = len(
        x
    )

    if n < N_QUANTILES:
        raise ValueError(
            "5분위 생성 불가: N<5"
        )

    order = np.argsort(
        x,
        kind="mergesort",
    )

    labels = np.empty(
        n,
        dtype=np.int8,
    )

    groups = np.array_split(
        order,
        N_QUANTILES,
    )

    for q, inds in enumerate(
        groups,
        start=1,
    ):
        labels[
            inds
        ] = q

    return labels


def daily_metrics(
    signal: np.ndarray,
    future: np.ndarray,
    eligible: np.ndarray,
):
    valid = (
        eligible
        & np.isfinite(
            signal
        )
        & np.isfinite(
            future
        )
    )

    idx = np.flatnonzero(
        valid
    )

    N = len(
        idx
    )

    if N < MIN_STOCKS:
        return None

    s = signal[
        idx
    ]

    f = future[
        idx
    ]

    # Existing continuous Rank IC on the exact same pair sample.
    continuous_ic = spearman_corr(
        s,
        f,
    )

    # New 5-bin quantile labels.
    qs = exact_five_quantiles(
        s
    )

    qf = exact_five_quantiles(
        f
    )

    # Spearman on 1..5 labels with ties.
    qic5 = spearman_corr(
        qs.astype(
            np.float64
        ),
        qf.astype(
            np.float64
        ),
    )

    same_q = float(
        np.mean(
            qs == qf
        )
    )

    q1_mask = (
        qs == 1
    )

    q5_mask = (
        qs == 5
    )

    q1_to_q1 = (
        float(
            np.mean(
                qf[
                    q1_mask
                ]
                == 1
            )
        )
        if q1_mask.any()
        else np.nan
    )

    q5_to_q5 = (
        float(
            np.mean(
                qf[
                    q5_mask
                ]
                == 5
            )
        )
        if q5_mask.any()
        else np.nan
    )

    mean_abs_dist = float(
        np.mean(
            np.abs(
                qs.astype(
                    np.float64
                )
                - qf.astype(
                    np.float64
                )
            )
        )
    )

    transition = np.zeros(
        (
            5,
            5,
        ),
        dtype=np.int64,
    )

    for a, b in zip(
        qs,
        qf,
    ):
        transition[
            int(
                a
            )
            - 1,
            int(
                b
            )
            - 1,
        ] += 1

    return {
        "N": N,
        "ContinuousRankIC": continuous_ic,
        "QIC5": qic5,
        "SameQ_Hit": same_q,
        "Q1_to_Q1": q1_to_q1,
        "Q5_to_Q5": q5_to_q5,
        "MeanAbsQDistance": mean_abs_dist,
        "Transition": transition,
    }


# =============================================================================
# HAC
# =============================================================================

def normal_two_sided_p(
    t: float,
) -> float:
    if not np.isfinite(
        t
    ):
        return np.nan

    # 2 * (1 - Phi(|t|)) = erfc(|t|/sqrt(2))
    return float(
        math.erfc(
            abs(
                t
            )
            / math.sqrt(
                2.0
            )
        )
    )


def hac_mean_test(
    values,
    lag: int,
):
    x = np.asarray(
        values,
        dtype=np.float64,
    )

    x = x[
        np.isfinite(
            x
        )
    ]

    T = len(
        x
    )

    if T < 3:
        return (
            np.nan,
            np.nan,
            np.nan,
            T,
        )

    mu = float(
        np.mean(
            x
        )
    )

    u = (
        x
        - mu
    )

    L = min(
        max(
            0,
            int(
                lag
            ),
        ),
        T - 1,
    )

    gamma0 = float(
        np.dot(
            u,
            u,
        )
        / T
    )

    lrv = gamma0

    for ell in range(
        1,
        L + 1,
    ):
        gamma = float(
            np.dot(
                u[
                    ell:
                ],
                u[
                    :-ell
                ],
            )
            / T
        )

        weight = (
            1.0
            - ell
            / (
                L + 1.0
            )
        )

        lrv += (
            2.0
            * weight
            * gamma
        )

    lrv = max(
        lrv,
        0.0,
    )

    se = math.sqrt(
        lrv
        / T
    )

    if se <= EPS:
        return (
            mu,
            np.nan,
            np.nan,
            T,
        )

    t = (
        mu
        / se
    )

    p = normal_two_sided_p(
        t
    )

    return (
        mu,
        float(
            t
        ),
        p,
        T,
    )


# =============================================================================
# XLSX WRITER
# =============================================================================

CONTENT_TYPES_TEMPLATE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
{sheet_overrides}
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

STYLES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2">
<font><sz val="11"/><name val="Calibri"/></font>
<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
<fill><patternFill patternType="none"/></fill>
<fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill>
</fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def xml_inline_cell(
    ref: str,
    value,
    header: bool = False,
) -> str:
    style = (
        ' s="1"'
        if header
        else ""
    )

    return (
        f'<c r="{ref}" t="inlineStr"{style}>'
        f'<is><t>{escape(str(value))}</t></is>'
        f'</c>'
    )


def xml_number_cell(
    ref: str,
    value,
) -> str:
    return (
        f'<c r="{ref}" s="2">'
        f'<v>{float(value):.12g}</v>'
        f'</c>'
    )


def write_table_sheet(
    path: Path,
    header: list[str],
    rows,
    widths: list[float] | None = None,
):
    with path.open(
        "w",
        encoding="utf-8",
        newline="",
    ) as out:
        out.write(
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        )

        out.write(
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        )

        out.write(
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" state="frozen"/>'
            '</sheetView></sheetViews>'
        )

        if widths:
            out.write(
                "<cols>"
            )

            for i, w in enumerate(
                widths,
                start=1,
            ):
                out.write(
                    f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                )

            out.write(
                "</cols>"
            )

        out.write(
            "<sheetData>"
        )

        out.write(
            '<row r="1">'
        )

        for c, name in enumerate(
            header,
            start=1,
        ):
            out.write(
                xml_inline_cell(
                    f"{idx_to_col(c-1)}1",
                    name,
                    header=True,
                )
            )

        out.write(
            "</row>"
        )

        last_row = 1

        for last_row, row in enumerate(
            rows,
            start=2,
        ):
            out.write(
                f'<row r="{last_row}">'
            )

            for c, value in enumerate(
                row,
                start=1,
            ):
                if value is None:
                    continue

                if (
                    isinstance(
                        value,
                        float,
                    )
                    and not math.isfinite(
                        value
                    )
                ):
                    continue

                ref = (
                    f"{idx_to_col(c-1)}"
                    f"{last_row}"
                )

                if isinstance(
                    value,
                    (
                        int,
                        float,
                        np.integer,
                        np.floating,
                    ),
                ):
                    out.write(
                        xml_number_cell(
                            ref,
                            value,
                        )
                    )
                else:
                    out.write(
                        xml_inline_cell(
                            ref,
                            value,
                        )
                    )

            out.write(
                "</row>"
            )

        out.write(
            "</sheetData>"
        )

        if header:
            last_col = idx_to_col(
                len(
                    header
                )
                - 1
            )

            out.write(
                f'<autoFilter ref="A1:{last_col}{last_row}"/>'
            )

        out.write(
            "</worksheet>"
        )


def package_xlsx(
    output_path: Path,
    sheet_specs: list[
        tuple[
            str,
            Path,
        ]
    ],
):
    sheets_xml = "".join(
        f'<sheet name="{escape(name)}" '
        f'sheetId="{i}" r:id="rId{i}"/>'
        for i, (
            name,
            _,
        ) in enumerate(
            sheet_specs,
            start=1,
        )
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook '
        'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        f'<sheets>{sheets_xml}</sheets>'
        '</workbook>'
    )

    rels = []

    for i, _ in enumerate(
        sheet_specs,
        start=1,
    ):
        rels.append(
            f'<Relationship Id="rId{i}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i}.xml"/>'
        )

    rels.append(
        f'<Relationship Id="rId{len(sheet_specs)+1}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    workbook_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(
            rels
        )
        + "</Relationships>"
    )

    overrides = []

    for i in range(
        1,
        len(
            sheet_specs
        )
        + 1,
    ):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    content_types = CONTENT_TYPES_TEMPLATE.format(
        sheet_overrides="".join(
            overrides
        )
    )

    with ZipFile(
        output_path,
        "w",
        compression=ZIP_DEFLATED,
        compresslevel=6,
    ) as z:
        z.writestr(
            "[Content_Types].xml",
            content_types,
        )

        z.writestr(
            "_rels/.rels",
            ROOT_RELS,
        )

        z.writestr(
            "xl/workbook.xml",
            workbook_xml,
        )

        z.writestr(
            "xl/_rels/workbook.xml.rels",
            workbook_rels,
        )

        z.writestr(
            "xl/styles.xml",
            STYLES,
        )

        for i, (
            _,
            p,
        ) in enumerate(
            sheet_specs,
            start=1,
        ):
            z.write(
                p,
                f"xl/worksheets/sheet{i}.xml",
            )


# =============================================================================
# ANALYSIS
# =============================================================================

def analyze_case_horizon(
    case: CaseDef,
    signal: np.ndarray,
    future: np.ndarray,
    common_bool: np.ndarray,
    dates: np.ndarray,
):
    daily_rows = []
    qic_by_date = {}
    continuous_values = []
    qic_values = []
    same_values = []
    q1_values = []
    q5_values = []
    dist_values = []

    transition_total = np.zeros(
        (
            5,
            5,
        ),
        dtype=np.int64,
    )

    for di, d in enumerate(
        dates
    ):
        m = daily_metrics(
            signal=signal[
                di
            ],
            future=future[
                di
            ],
            eligible=common_bool[
                di
            ],
        )

        if m is None:
            continue

        qic_by_date[
            int(
                d
            )
        ] = m[
            "QIC5"
        ]

        continuous_values.append(
            m[
                "ContinuousRankIC"
            ]
        )
        qic_values.append(
            m[
                "QIC5"
            ]
        )
        same_values.append(
            m[
                "SameQ_Hit"
            ]
        )
        q1_values.append(
            m[
                "Q1_to_Q1"
            ]
        )
        q5_values.append(
            m[
                "Q5_to_Q5"
            ]
        )
        dist_values.append(
            m[
                "MeanAbsQDistance"
            ]
        )

        transition_total += m[
            "Transition"
        ]

        daily_rows.append(
            [
                int(
                    d
                ),
                case.case_id,
                (
                    case.pca_window
                    if case.pca_window is not None
                    else None
                ),
                case.pc_remove,
                case.formation_n,
                m[
                    "N"
                ],
                m[
                    "ContinuousRankIC"
                ],
                m[
                    "QIC5"
                ],
                m[
                    "SameQ_Hit"
                ],
                m[
                    "Q1_to_Q1"
                ],
                m[
                    "Q5_to_Q5"
                ],
                m[
                    "MeanAbsQDistance"
                ],
            ]
        )

    return {
        "daily_rows": daily_rows,
        "qic_by_date": qic_by_date,
        "continuous_values": np.asarray(
            continuous_values,
            dtype=np.float64,
        ),
        "qic_values": np.asarray(
            qic_values,
            dtype=np.float64,
        ),
        "same_values": np.asarray(
            same_values,
            dtype=np.float64,
        ),
        "q1_values": np.asarray(
            q1_values,
            dtype=np.float64,
        ),
        "q5_values": np.asarray(
            q5_values,
            dtype=np.float64,
        ),
        "dist_values": np.asarray(
            dist_values,
            dtype=np.float64,
        ),
        "transition_total": transition_total,
    }


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--data-dir",
        default=None,
        help=(
            "필수 Excel 5개가 있는 상위폴더. "
            "생략 시 이 py 파일 폴더 이하를 재귀탐색."
        ),
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과 저장 폴더",
    )

    args = parser.parse_args()

    try:
        script_dir = Path(
            __file__
        ).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    data_dir = (
        Path(
            args.data_dir
        ).resolve()
        if args.data_dir
        else script_dir
    )

    output_dir = (
        Path(
            args.output_dir
        ).resolve()
        if args.output_dir
        else (
            script_dir
            / "QUANTILE_RANK_IC_5BIN_ANALYSIS"
        )
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "QUANTILE_RANK_IC_5BIN_ANALYSIS.xlsx"
    )

    files = {
        "base": find_input_file(
            data_dir,
            "00_COMMON252_BASE",
        ),
        "cum60": find_input_file(
            data_dir,
            "02_PCA_W60_CUMULATIVE",
        ),
        "cum120": find_input_file(
            data_dir,
            "04_PCA_W120_CUMULATIVE",
        ),
        "cum252": find_input_file(
            data_dir,
            "06_PCA_W252_CUMULATIVE",
        ),
        "future": find_input_file(
            data_dir,
            "FUTURE_LOG_RETURNS_5_10_20_60",
        ),
    }

    print(
        "="
        * 100
    )
    print(
        "5-QUANTILE RANK IC (QIC5) ANALYSIS"
    )
    print(
        "="
        * 100
    )

    for k, p in files.items():
        print(
            f"{k:<8}: {p}"
        )

    print()

    # -------------------------------------------------------------------------
    # Common252 base
    # -------------------------------------------------------------------------
    print(
        "[1/5] COMMON252_ELIGIBLE 읽기...",
        flush=True,
    )

    (
        dates,
        tickers,
        common,
    ) = read_wide_matrix(
        files[
            "base"
        ],
        "COMMON252_ELIGIBLE",
    )

    common_bool = (
        np.nan_to_num(
            common,
            nan=0.0,
        )
        == 1
    )

    print(
        f"      dates={len(dates):,} | tickers={len(tickers):,}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Future returns
    # -------------------------------------------------------------------------
    print(
        "[2/5] FWD 5/10/20/60 읽기...",
        flush=True,
    )

    future_map = {}

    for h in HORIZONS:
        d, t, arr = read_wide_matrix(
            files[
                "future"
            ],
            f"FWD_LOG_{h}",
        )

        future_map[
            h
        ] = align_matrix(
            dates,
            tickers,
            d,
            t,
            arr,
            f"FWD_LOG_{h}",
        )

        print(
            f"      FWD{h}: finite={int(np.isfinite(future_map[h]).sum()):,}",
            flush=True,
        )

    # -------------------------------------------------------------------------
    # Cases
    # -------------------------------------------------------------------------
    cases = build_cases()

    case_map = {
        c.case_id: c
        for c in cases
    }

    # Store Raw daily QIC for DeltaQIC calculation.
    raw_qic_map = {
        (
            n,
            h,
        ): {}
        for n in FORMATION_NS
        for h in HORIZONS
    }

    # Temporarily store all per-horizon daily rows.
    daily_by_horizon = {
        h: []
        for h in HORIZONS
    }

    summary_records = []
    transition_records = []

    print(
        "[3/5] 40 Cases x 4 horizons QIC5 계산...",
        flush=True,
    )

    # Analyze Raw first so DeltaQIC can be populated immediately for PCA.
    ordered_cases = (
        [
            c
            for c in cases
            if c.pc_remove
            == "RAW"
        ]
        + [
            c
            for c in cases
            if c.pc_remove
            != "RAW"
        ]
    )

    for ci, case in enumerate(
        ordered_cases,
        start=1,
    ):
        print(
            f"      [{ci:02d}/{len(ordered_cases):02d}] {case.case_id}",
            flush=True,
        )

        d, t, cumulative = read_wide_matrix(
            files[
                case.source_key
            ],
            case.source_sheet,
        )

        cumulative = align_matrix(
            dates,
            tickers,
            d,
            t,
            cumulative,
            case.case_id,
        )

        # Higher reversal signal = larger loser score.
        signal = -cumulative

        for h in HORIZONS:
            res = analyze_case_horizon(
                case=case,
                signal=signal,
                future=future_map[
                    h
                ],
                common_bool=common_bool,
                dates=dates,
            )

            lag = (
                max(
                    case.formation_n,
                    h,
                )
                - 1
            )

            (
                mean_qic,
                hac_t,
                hac_p,
                valid_dates,
            ) = hac_mean_test(
                res[
                    "qic_values"
                ],
                lag=lag,
            )

            (
                mean_cont_ic,
                cont_hac_t,
                cont_hac_p,
                _,
            ) = hac_mean_test(
                res[
                    "continuous_values"
                ],
                lag=lag,
            )

            if case.pc_remove == "RAW":
                raw_qic_map[
                    (
                        case.formation_n,
                        h,
                    )
                ] = res[
                    "qic_by_date"
                ]

            # Daily delta vs same-n Raw.
            raw_case_id = (
                f"RAW_N{case.formation_n}"
            )

            raw_daily = raw_qic_map[
                (
                    case.formation_n,
                    h,
                )
            ]

            delta_values = []

            for row in res[
                "daily_rows"
            ]:
                date_int = int(
                    row[
                        0
                    ]
                )

                qic = row[
                    7
                ]

                raw_qic = (
                    raw_daily.get(
                        date_int
                    )
                    if case.pc_remove
                    != "RAW"
                    else qic
                )

                delta = (
                    qic
                    - raw_qic
                    if (
                        case.pc_remove
                        != "RAW"
                        and raw_qic
                        is not None
                        and np.isfinite(
                            raw_qic
                        )
                        and np.isfinite(
                            qic
                        )
                    )
                    else None
                )

                if delta is not None:
                    delta_values.append(
                        delta
                    )

                # Append Horizon / Raw comparison fields.
                full_row = [
                    row[
                        0
                    ],
                    row[
                        1
                    ],
                    row[
                        2
                    ],
                    row[
                        3
                    ],
                    row[
                        4
                    ],
                    h,
                    row[
                        5
                    ],
                    row[
                        6
                    ],
                    row[
                        7
                    ],
                    row[
                        8
                    ],
                    row[
                        9
                    ],
                    row[
                        10
                    ],
                    row[
                        11
                    ],
                    raw_case_id,
                    raw_qic,
                    delta,
                ]

                daily_by_horizon[
                    h
                ].append(
                    full_row
                )

            if delta_values:
                (
                    mean_delta,
                    delta_t,
                    delta_p,
                    delta_n,
                ) = hac_mean_test(
                    delta_values,
                    lag=lag,
                )
            else:
                mean_delta = np.nan
                delta_t = np.nan
                delta_p = np.nan
                delta_n = 0

            qic_vals = res[
                "qic_values"
            ]

            summary_records.append(
                (
                    case.case_id,
                    case.pca_window,
                    case.pc_remove,
                    case.formation_n,
                    h,
                    valid_dates,
                    float(
                        np.mean(
                            res[
                                "continuous_values"
                            ]
                        )
                    )
                    if len(
                        res[
                            "continuous_values"
                        ]
                    )
                    else None,
                    cont_hac_t,
                    cont_hac_p,
                    mean_qic,
                    float(
                        np.median(
                            qic_vals
                        )
                    )
                    if len(
                        qic_vals
                    )
                    else None,
                    float(
                        np.mean(
                            qic_vals
                            > 0
                        )
                    )
                    if len(
                        qic_vals
                    )
                    else None,
                    lag,
                    hac_t,
                    hac_p,
                    float(
                        np.mean(
                            res[
                                "same_values"
                            ]
                        )
                    )
                    if len(
                        res[
                            "same_values"
                        ]
                    )
                    else None,
                    float(
                        np.mean(
                            res[
                                "q1_values"
                            ]
                        )
                    )
                    if len(
                        res[
                            "q1_values"
                        ]
                    )
                    else None,
                    float(
                        np.mean(
                            res[
                                "q5_values"
                            ]
                        )
                    )
                    if len(
                        res[
                            "q5_values"
                        ]
                    )
                    else None,
                    float(
                        np.mean(
                            res[
                                "dist_values"
                            ]
                        )
                    )
                    if len(
                        res[
                            "dist_values"
                        ]
                    )
                    else None,
                    raw_case_id,
                    mean_delta,
                    delta_n,
                    delta_t,
                    delta_p,
                )
            )

            # Transition matrix: one row per SignalQ.
            M = res[
                "transition_total"
            ]

            for sq in range(
                1,
                6,
            ):
                counts = M[
                    sq
                    - 1
                ]

                total = int(
                    counts.sum()
                )

                shares = (
                    counts
                    / total
                    if total > 0
                    else np.full(
                        5,
                        np.nan,
                    )
                )

                transition_records.append(
                    (
                        case.case_id,
                        case.pca_window,
                        case.pc_remove,
                        case.formation_n,
                        h,
                        sq,
                        total,
                        int(
                            counts[
                                0
                            ]
                        ),
                        int(
                            counts[
                                1
                            ]
                        ),
                        int(
                            counts[
                                2
                            ]
                        ),
                        int(
                            counts[
                                3
                            ]
                        ),
                        int(
                            counts[
                                4
                            ]
                        ),
                        float(
                            shares[
                                0
                            ]
                        )
                        if np.isfinite(
                            shares[
                                0
                            ]
                        )
                        else None,
                        float(
                            shares[
                                1
                            ]
                        )
                        if np.isfinite(
                            shares[
                                1
                            ]
                        )
                        else None,
                        float(
                            shares[
                                2
                            ]
                        )
                        if np.isfinite(
                            shares[
                                2
                            ]
                        )
                        else None,
                        float(
                            shares[
                                3
                            ]
                        )
                        if np.isfinite(
                            shares[
                                3
                            ]
                        )
                        else None,
                        float(
                            shares[
                                4
                            ]
                        )
                        if np.isfinite(
                            shares[
                                4
                            ]
                        )
                        else None,
                        float(
                            shares[
                                sq
                                - 1
                            ]
                        )
                        if (
                            total > 0
                            and np.isfinite(
                                shares[
                                    sq
                                    - 1
                                ]
                            )
                        )
                        else None,
                    )
                )

        del cumulative
        del signal

    # Sort daily rows by date then case for readability.
    for h in HORIZONS:
        daily_by_horizon[
            h
        ].sort(
            key=lambda r: (
                int(
                    r[
                        0
                    ]
                ),
                str(
                    r[
                        1
                    ]
                ),
            )
        )

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------
    print(
        "[4/5] Excel 시트 생성...",
        flush=True,
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(
            td
        )

        specs = []

        # PARAMETERS ----------------------------------------------------------
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Research_Type",
                "Supplemental / 부가연구",
            ),
            (
                "Primary_Metric",
                "5-Quantile Rank IC (QIC5)",
            ),
            (
                "Continuous_Benchmark",
                "Daily Spearman Rank IC on the same pair universe",
            ),
            (
                "Signal",
                "Reversal Signal = - past cumulative return",
            ),
            (
                "Signal_Q1",
                "Lowest reversal score = recent winners",
            ),
            (
                "Signal_Q5",
                "Highest reversal score = recent losers",
            ),
            (
                "Future_Q1",
                "Lowest future return",
            ),
            (
                "Future_Q5",
                "Highest future return",
            ),
            (
                "Positive_QIC5",
                "Signal low/high quintiles align with future low/high quintiles",
            ),
            (
                "Pair_Universe",
                "COMMON252_ELIGIBLE & finite Signal & finite FWD return",
            ),
            (
                "Quintile_Method",
                "Stable sort then numpy.array_split into five near-equal groups",
            ),
            (
                "SameQ_Hit_Random_Benchmark",
                0.20,
            ),
            (
                "DeltaQIC5",
                "PCA QIC5 - RAW QIC5 with same Formation_N",
            ),
            (
                "HAC_Lag",
                "max(Formation_N, Horizon)-1",
            ),
            (
                "Formation_N",
                "5,10,30,60",
            ),
            (
                "Future_Horizon",
                "5,10,20,60",
            ),
            (
                "Cases",
                "4 Raw + 36 PCA = 40",
            ),
            (
                "Minimum_Stocks",
                MIN_STOCKS,
            ),
        ]

        write_table_sheet(
            p,
            [
                "Item",
                "Value",
            ],
            parameter_rows,
            [
                34,
                110,
            ],
        )

        specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # CASE DEFINITION -----------------------------------------------------
        p = td / "case_definition.xml"

        case_rows = [
            (
                c.case_id,
                c.pca_window,
                c.pc_remove,
                c.formation_n,
                c.source_key,
                c.source_sheet,
            )
            for c in cases
        ]

        write_table_sheet(
            p,
            [
                "CaseID",
                "PCA_Window",
                "PC_Remove",
                "Formation_N",
                "Source",
                "SourceSheet",
            ],
            case_rows,
            [
                28,
                14,
                14,
                14,
                16,
                24,
            ],
        )

        specs.append(
            (
                "CASE_DEFINITION",
                p,
            )
        )

        # SUMMARY -------------------------------------------------------------
        p = td / "summary.xml"

        summary_header = [
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "ValidDates",
            "Mean_ContinuousRankIC",
            "Continuous_HAC_t",
            "Continuous_HAC_p",
            "Mean_QIC5",
            "Median_QIC5",
            "Positive_QIC5_Ratio",
            "HAC_Lag",
            "QIC5_HAC_t",
            "QIC5_HAC_p",
            "Mean_SameQ_Hit",
            "Mean_Q1_to_Q1",
            "Mean_Q5_to_Q5",
            "Mean_Abs_Q_Distance",
            "Raw_Case",
            "Mean_DeltaQIC5_vs_Raw",
            "Delta_ValidDates",
            "Delta_HAC_t",
            "Delta_HAC_p",
        ]

        write_table_sheet(
            p,
            summary_header,
            summary_records,
            [
                28,
                14,
                14,
                14,
                12,
                14,
                22,
                18,
                18,
                16,
                16,
                20,
                12,
                16,
                16,
                18,
                18,
                18,
                20,
                18,
                24,
                18,
                18,
                18,
            ],
        )

        specs.append(
            (
                "QIC_SUMMARY",
                p,
            )
        )

        # TRANSITION MATRIX ---------------------------------------------------
        p = td / "transition.xml"

        transition_header = [
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "SignalQ",
            "N_Transitions",
            "FutureQ1_Count",
            "FutureQ2_Count",
            "FutureQ3_Count",
            "FutureQ4_Count",
            "FutureQ5_Count",
            "FutureQ1_Share",
            "FutureQ2_Share",
            "FutureQ3_Share",
            "FutureQ4_Share",
            "FutureQ5_Share",
            "Diagonal_SameQ_Share",
        ]

        write_table_sheet(
            p,
            transition_header,
            transition_records,
            [
                28,
                14,
                14,
                14,
                12,
                11,
                16,
                18,
                18,
                18,
                18,
                18,
                18,
                18,
                18,
                18,
                18,
                22,
            ],
        )

        specs.append(
            (
                "TRANSITION_MATRIX",
                p,
            )
        )

        # DAILY sheets --------------------------------------------------------
        daily_header = [
            "Date",
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "N",
            "ContinuousRankIC",
            "QIC5",
            "SameQ_Hit",
            "Q1_to_Q1",
            "Q5_to_Q5",
            "MeanAbsQDistance",
            "Raw_Case",
            "Raw_QIC5",
            "DeltaQIC5_vs_Raw",
        ]

        for h in HORIZONS:
            p = td / f"daily_{h}.xml"

            rows = (
                (
                    yyyymmdd_text(
                        int(
                            r[
                                0
                            ]
                        )
                    ),
                    *r[
                        1:
                    ],
                )
                for r in daily_by_horizon[
                    h
                ]
            )

            write_table_sheet(
                p,
                daily_header,
                rows,
                [
                    14,
                    28,
                    14,
                    14,
                    14,
                    12,
                    10,
                    20,
                    15,
                    16,
                    16,
                    16,
                    20,
                    18,
                    16,
                    22,
                ],
            )

            specs.append(
                (
                    f"DAILY_FWD{h}",
                    p,
                )
            )

        package_xlsx(
            output_path,
            specs,
        )

    print(
        "[5/5] 생성 파일 검증...",
        flush=True,
    )

    with ZipFile(
        output_path
    ) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성 XLSX ZIP 오류: {bad}"
            )

    print()
    print(
        "="
        * 100
    )
    print(
        "DONE"
    )
    print(
        "="
        * 100
    )
    print(
        "OUTPUT:",
        output_path,
    )
    print()
    print(
        "먼저 볼 시트:"
    )
    print(
        "  1) QIC_SUMMARY"
    )
    print(
        "  2) TRANSITION_MATRIX"
    )
    print(
        "  3) DAILY_FWD5"
    )
    print(
        "  4) DAILY_FWD10/20/60"
    )


if __name__ == "__main__":
    main()
