# -*- coding: utf-8 -*-
"""
12_QUINTILE_REVERSAL_SPREAD_ANALYSIS_DAILY.py
=======================================

연구 본체용 5분위 단기반전 분석.

핵심 질문
---------
PCA 공통요인을 제거한 단기반전 신호로 종목을 5분위로 나눴을 때,

    Q5 = Reversal Signal 가장 큼 = 과거 loser 20%  -> Long
    Q1 = Reversal Signal 가장 작음 = 과거 winner 20% -> Short

의 미래수익률 차이:

    Reversal Spread = Return(Q5) - Return(Q1)

가 양수이고 유의한가?

그리고 같은 formation n에서:

    PCA Spread - Raw Spread

가 양수이고 유의한가?

이를 통해:

1) 어떤 PCA window / PC 제거 수준이 좋은가?
2) 과거 formation n = 5/10/30/60 중 어디가 좋은가?
3) 미래 holding/effect horizon k = 5/10/20/60 중 어디까지 반전효과가 유지되는가?

를 확인한다.

------------------------------------------------------------------------------
필요 입력 Excel 5개
------------------------------------------------------------------------------

1. 00_COMMON252_BASE.xlsx
   필요한 시트:
   - COMMON252_ELIGIBLE
   - RAW_CUM_LOG_5
   - RAW_CUM_LOG_10
   - RAW_CUM_LOG_30
   - RAW_CUM_LOG_60

2. 02_PCA_W60_CUMULATIVE.xlsx
   - PC1_CUM5/10/30/60
   - PC12_CUM5/10/30/60
   - PC123_CUM5/10/30/60

3. 04_PCA_W120_CUMULATIVE.xlsx

4. 06_PCA_W252_CUMULATIVE.xlsx

5. FUTURE_LOG_RETURNS_5_10_20_60.xlsx
   - FWD_LOG_5
   - FWD_LOG_10
   - FWD_LOG_20
   - FWD_LOG_60

필요하지 않은 파일
------------------
- PCA DAILY 파일
- STATE_IC_TRAIN_LIBRARY.xlsx
- Ridge / 1-NN 결과파일

------------------------------------------------------------------------------
중요한 계산 규칙
------------------------------------------------------------------------------

[1] Signal
    Signal = - cumulative past return

따라서 Signal이 클수록 과거 loser.

[2] 5분위 형성
미래수익률을 절대 보지 않고,
그날:

    Common252 eligible = 1
    + Signal finite

인 종목만 Signal 순으로 정렬한다.

낮은 Signal -> Q1
높은 Signal -> Q5

종목 수가 5로 정확히 나누어지지 않으면 최대한 균등하게 분할한다.

[3] 미래수익률
입력 FWD는 cumulative LOG return이므로
포트폴리오 수익률 계산 전에:

    simple_return = exp(FWD_LOG) - 1

로 변환한다.

각 분위는 equal-weight 평균 미래 단순수익률.

[4] Reversal Spread
    Spread_T = MeanFwdReturn(Q5,T) - MeanFwdReturn(Q1,T)

양수이면:
    과거 loser Long / winner Short가 미래에 반전했다는 뜻.

[5] HAC
Daily spread가 overlapping horizon 때문에 자기상관을 가질 수 있으므로
Newey-West HAC 사용.

보수적 lag:
    lag = max(Formation_N, Horizon) - 1

[6] PCA 개선효과
같은 formation n, 같은 horizon k에서:

    DeltaSpread_T
      = PCA_Spread_T - RAW_Spread_same_n_T

를 계산하고 HAC 검정한다.

------------------------------------------------------------------------------
출력
------------------------------------------------------------------------------

QUINTILE_REVERSAL_ANALYSIS/
    QUINTILE_REVERSAL_ANALYSIS.xlsx

시트 4개:
1. PARAMETERS
2. DAILY_SPREAD
3. SPREAD_SUMMARY
4. BEST_COMPARISON

SPREAD_SUMMARY 주요 열
----------------------
CaseID
PCA_Window
PC_Remove
Formation_N
Horizon
ValidDates

Q1_MeanFwdReturn
Q2_MeanFwdReturn
Q3_MeanFwdReturn
Q4_MeanFwdReturn
Q5_MeanFwdReturn

Q5_Q1_Spread
Spread_HAC_t
Spread_HAC_p
Q5_Beats_Q1_Ratio

DeltaSpread_vs_Raw
Delta_HAC_t
Delta_HAC_p

BEST_COMPARISON
---------------
같은 n / k에서:
Raw spread와 Mean Spread가 가장 높은 PCA Case를 바로 비교.

------------------------------------------------------------------------------
실행
------------------------------------------------------------------------------

python 12_QUINTILE_REVERSAL_SPREAD_ANALYSIS.py

기간 지정:
python 12_QUINTILE_REVERSAL_SPREAD_ANALYSIS.py --start 2010-01-01 --end 2026-07-31

필수 외부 패키지:
    numpy
"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import numpy as np


PCA_WINDOWS = (60, 120, 252)
FORMATION_HORIZONS = (5, 10, 30, 60)
FUTURE_HORIZONS = (5, 10, 20, 60)

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EPS = 1e-15
MIN_FORMATION_STOCKS = 50
MIN_RETURN_STOCKS_PER_QUINTILE = 5


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


def find_input_file(base: Path, key: str) -> Path:
    """
    base 이하를 재귀 탐색.
    파일명에 (1), (2)가 붙어 있어도 key로 찾음.
    여러 개면 최근 수정 파일 선택.
    """
    key_norm = normalize_filename_token(key)

    candidates = []

    for p in base.rglob("*.xlsx"):
        name = normalize_filename_token(p.name)

        if key_norm in name:
            if "quintile_reversal_analysis" in name:
                continue

            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            f"\n필수 파일을 찾지 못했습니다: {key}\n"
            f"검색폴더: {base}\n"
        )

    return max(
        candidates,
        key=lambda p: p.stat().st_mtime,
    )


def excel_serial_to_yyyymmdd(value) -> int:
    d = datetime(1899, 12, 30) + timedelta(days=float(value))
    return int(d.strftime("%Y%m%d"))


def parse_date_value(value) -> int:
    if value is None or value == "":
        raise ValueError("empty date")

    if isinstance(value, (int, float, np.integer, np.floating)):
        x = float(value)

        if 20000 <= x <= 80000:
            return excel_serial_to_yyyymmdd(x)

        if 19000101 <= x <= 21001231:
            return int(x)

    s = str(value).strip()

    if re.fullmatch(r"\d{8}", s):
        return int(s)

    if re.fullmatch(r"\d+(\.\d+)?", s):
        x = float(s)

        if 20000 <= x <= 80000:
            return excel_serial_to_yyyymmdd(x)

    parts = re.findall(r"\d+", s)

    if len(parts) >= 3:
        y, m, d = map(int, parts[:3])
        return y * 10000 + m * 100 + d

    raise ValueError(f"날짜 파싱 실패: {value!r}")


def parse_cli_date(s: str | None) -> int | None:
    if not s:
        return None

    return parse_date_value(s)


def yyyymmdd_text(x: int) -> str:
    s = str(int(x))
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}"


# =============================================================================
# XLSX STREAM READER
# =============================================================================

def workbook_sheet_targets(path: Path) -> dict[str, str]:
    with ZipFile(path) as z:
        wb = ET.fromstring(
            z.read("xl/workbook.xml")
        )

        rels = ET.fromstring(
            z.read("xl/_rels/workbook.xml.rels")
        )

        relmap = {
            r.attrib["Id"]: r.attrib["Target"]
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

            target = relmap[rid]

            if target.startswith("/"):
                target = target.lstrip("/")
            elif not target.startswith("xl/"):
                target = "xl/" + target

            out[s.attrib["name"]] = target

        return out


def load_shared_strings(z: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []

    root = ET.fromstring(
        z.read("xl/sharedStrings.xml")
    )

    return [
        "".join(
            (t.text or "")
            for t in si.iter(
                f"{{{NS_MAIN}}}t"
            )
        )
        for si in root
    ]


def cell_value(cell: ET.Element, shared: list[str]):
    typ = cell.attrib.get("t")

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
            int(v.text)
        ]

    if typ == "b":
        return int(v.text)

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

    with ZipFile(xlsx_path) as z:
        shared = load_shared_strings(z)

        with z.open(target) as f:
            for _, elem in ET.iterparse(
                f,
                events=("end",),
            ):
                if elem.tag != f"{{{NS_MAIN}}}row":
                    continue

                cells = {}

                for c in elem:
                    if c.tag != f"{{{NS_MAIN}}}c":
                        continue

                    ref = c.attrib.get("r")

                    if not ref:
                        continue

                    cells[
                        col_to_idx(ref)
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
        header = next(it)
    except StopIteration:
        raise RuntimeError(
            f"{path.name}/{sheet_name}: 빈 시트"
        )

    max_col = max(
        header.keys()
    )

    tickers = [
        normalize_ticker(
            header.get(c, "")
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
                cells[0]
            )
        except Exception:
            continue

        row = np.full(
            len(tickers),
            np.nan,
            dtype=np.float64,
        )

        for c, raw in cells.items():
            if c == 0:
                continue

            j = c - 1

            if (
                j < 0
                or j >= len(tickers)
                or raw in (None, "")
            ):
                continue

            try:
                row[j] = float(raw)
            except Exception:
                pass

        dates.append(d)
        rows.append(row)

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
        np.vstack(rows),
    )


def require_same_dates(
    reference: np.ndarray,
    other: np.ndarray,
    label: str,
):
    if not np.array_equal(
        reference,
        other,
    ):
        raise RuntimeError(
            f"날짜 불일치: {label}"
        )


def require_same_tickers(
    reference: list[str],
    other: list[str],
    label: str,
):
    if reference != other:
        raise RuntimeError(
            f"ticker 순서 불일치: {label}"
        )


# =============================================================================
# CASE DEFINITIONS
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    pca_window: str
    pc_remove: str
    formation_n: int
    source_key: str
    source_sheet: str


def make_cases() -> list[CaseDef]:
    cases = []

    # Raw: PCA window와 무관한 4개 unique cases
    for n in FORMATION_HORIZONS:
        cases.append(
            CaseDef(
                case_id=f"RAW_N{n}",
                pca_window="RAW",
                pc_remove="RAW",
                formation_n=n,
                source_key="base",
                source_sheet=f"RAW_CUM_LOG_{n}",
            )
        )

    # PCA adjusted: 3 x 3 x 4 = 36
    for W in PCA_WINDOWS:
        for pc in (
            "PC1",
            "PC12",
            "PC123",
        ):
            for n in FORMATION_HORIZONS:
                cases.append(
                    CaseDef(
                        case_id=f"W{W}_{pc}_N{n}",
                        pca_window=str(W),
                        pc_remove=pc,
                        formation_n=n,
                        source_key=f"cum{W}",
                        source_sheet=f"{pc}_CUM{n}",
                    )
                )

    if len(cases) != 40:
        raise RuntimeError(
            f"Case 수 오류: {len(cases)}"
        )

    return cases


# =============================================================================
# QUINTILES / DAILY PORTFOLIO
# =============================================================================

def form_quintile_indices(
    signal: np.ndarray,
    eligible: np.ndarray,
):
    """
    미래수익률을 전혀 보지 않고:
      eligible & finite(signal)
    만으로 Q1~Q5를 형성.

    Q1 = smallest signal = past winner
    Q5 = largest signal = past loser
    """
    formation_mask = (
        eligible
        & np.isfinite(signal)
    )

    idx = np.flatnonzero(
        formation_mask
    )

    n = len(idx)

    if n < MIN_FORMATION_STOCKS:
        return None

    # stable ascending signal order
    order = np.argsort(
        signal[idx],
        kind="mergesort",
    )

    sorted_idx = idx[
        order
    ]

    # 가능한 균등하게 5분할
    groups = np.array_split(
        sorted_idx,
        5,
    )

    if any(
        len(g) < MIN_RETURN_STOCKS_PER_QUINTILE
        for g in groups
    ):
        return None

    return groups


def evaluate_quintile_returns(
    groups,
    future_log: np.ndarray,
):
    """
    미래 LOG return -> SIMPLE return으로 변환 후
    각 분위 equal-weight 평균.

    분위 형성은 이미 미래수익률과 독립적으로 끝난 상태.
    """
    q_returns = []

    q_counts = []

    for g in groups:
        logret = future_log[g]

        valid = np.isfinite(
            logret
        )

        count = int(
            valid.sum()
        )

        if count < MIN_RETURN_STOCKS_PER_QUINTILE:
            return None

        simple = np.expm1(
            logret[valid]
        )

        q_returns.append(
            float(
                np.mean(simple)
            )
        )

        q_counts.append(
            count
        )

    spread = (
        q_returns[4]
        - q_returns[0]
    )

    return {
        "q_returns": q_returns,
        "q_counts": q_counts,
        "spread": float(spread),
        "q5_beats_q1": int(
            q_returns[4]
            > q_returns[0]
        ),
    }


# =============================================================================
# HAC
# =============================================================================

def hac_mean_test(
    values: np.ndarray,
    lag: int,
):
    """
    H0: E[x] = 0
    Newey-West HAC with Bartlett weights.
    two-sided asymptotic normal p-value.
    """
    x = np.asarray(
        values,
        dtype=np.float64,
    )

    x = x[
        np.isfinite(x)
    ]

    T = len(x)

    if T < 5:
        return (
            np.nan,
            np.nan,
            np.nan,
        )

    mean = float(
        np.mean(x)
    )

    u = x - mean

    gamma0 = float(
        np.dot(
            u,
            u,
        )
        / T
    )

    lrv = gamma0

    L = min(
        max(
            0,
            int(lag),
        ),
        T - 1,
    )

    for ell in range(
        1,
        L + 1,
    ):
        gamma = float(
            np.dot(
                u[ell:],
                u[:-ell],
            )
            / T
        )

        weight = (
            1.0
            - ell / (L + 1.0)
        )

        lrv += (
            2.0
            * weight
            * gamma
        )

    if lrv <= EPS:
        return (
            mean,
            np.nan,
            np.nan,
        )

    se = math.sqrt(
        lrv / T
    )

    t = mean / se

    p = math.erfc(
        abs(t)
        / math.sqrt(2.0)
    )

    return (
        mean,
        float(t),
        float(p),
    )


# =============================================================================
# ANALYSIS
# =============================================================================

def analyze_case(
    case: CaseDef,
    signal: np.ndarray,
    common_bool: np.ndarray,
    future: dict[int, np.ndarray],
    analysis_idx: np.ndarray,
    dates: np.ndarray,
):
    """
    return:
      horizon -> {
          dates: [],
          q1..q5,
          spread,
          q5beats
      }
    """
    result = {}

    for h in FUTURE_HORIZONS:
        fwd = future[h]

        out_dates = []
        q_values = [
            []
            for _ in range(5)
        ]
        q_counts = [
            []
            for _ in range(5)
        ]
        spread_values = []
        beats = []

        for ti in analysis_idx:
            groups = form_quintile_indices(
                signal=signal[ti],
                eligible=common_bool[ti],
            )

            if groups is None:
                continue

            rec = evaluate_quintile_returns(
                groups=groups,
                future_log=fwd[ti],
            )

            if rec is None:
                continue

            out_dates.append(
                int(dates[ti])
            )

            for q in range(5):
                q_values[q].append(
                    rec["q_returns"][q]
                )

                q_counts[q].append(
                    rec["q_counts"][q]
                )

            spread_values.append(
                rec["spread"]
            )

            beats.append(
                rec["q5_beats_q1"]
            )

        result[h] = {
            "dates": np.asarray(
                out_dates,
                dtype=np.int32,
            ),
            "q": [
                np.asarray(
                    v,
                    dtype=np.float64,
                )
                for v in q_values
            ],
            "q_counts": [
                np.asarray(
                    v,
                    dtype=np.float64,
                )
                for v in q_counts
            ],
            "spread": np.asarray(
                spread_values,
                dtype=np.float64,
            ),
            "beats": np.asarray(
                beats,
                dtype=np.float64,
            ),
        }

    return result


def align_delta_spread(
    pca_dates: np.ndarray,
    pca_spread: np.ndarray,
    raw_dates: np.ndarray,
    raw_spread: np.ndarray,
):
    raw_map = {
        int(d): float(v)
        for d, v in zip(
            raw_dates,
            raw_spread,
        )
        if np.isfinite(v)
    }

    delta = []

    for d, v in zip(
        pca_dates,
        pca_spread,
    ):
        if not np.isfinite(v):
            continue

        rv = raw_map.get(
            int(d)
        )

        if rv is None:
            continue

        delta.append(
            float(v - rv)
        )

    return np.asarray(
        delta,
        dtype=np.float64,
    )


def summarize_case_results(
    cases: list[CaseDef],
    case_results,
):
    raw_by_n = {
        c.formation_n: c.case_id
        for c in cases
        if c.pc_remove == "RAW"
    }

    summary_rows = []
    daily_rows = []
    summary_lookup = {}

    # -------------------------------------------------------------------------
    # Daily rows + summary rows
    # -------------------------------------------------------------------------
    for case in cases:
        for h in FUTURE_HORIZONS:
            r = case_results[
                case.case_id
            ][h]

            spread = r[
                "spread"
            ]

            if len(spread) == 0:
                continue

            q_means = [
                float(
                    np.mean(
                        r["q"][q]
                    )
                )
                for q in range(5)
            ]

            lag = max(
                case.formation_n,
                h,
            ) - 1

            spread_mean, spread_t, spread_p = hac_mean_test(
                spread,
                lag=lag,
            )

            beats_ratio = float(
                np.mean(
                    r["beats"]
                )
            )

            delta_mean = None
            delta_t = None
            delta_p = None

            raw_case_id = raw_by_n[
                case.formation_n
            ]

            raw_result = case_results[
                raw_case_id
            ][h]

            raw_map = {
                int(d): float(v)
                for d, v in zip(
                    raw_result["dates"],
                    raw_result["spread"],
                )
                if np.isfinite(v)
            }

            if case.pc_remove != "RAW":
                delta = align_delta_spread(
                    pca_dates=r["dates"],
                    pca_spread=r["spread"],
                    raw_dates=raw_result["dates"],
                    raw_spread=raw_result["spread"],
                )

                if len(delta):
                    dmean, dt, dp = hac_mean_test(
                        delta,
                        lag=lag,
                    )

                    delta_mean = (
                        float(dmean)
                        if np.isfinite(dmean)
                        else None
                    )

                    delta_t = (
                        float(dt)
                        if np.isfinite(dt)
                        else None
                    )

                    delta_p = (
                        float(dp)
                        if np.isfinite(dp)
                        else None
                    )

            avg_q1_n = float(
                np.mean(
                    r["q_counts"][0]
                )
            )

            avg_q5_n = float(
                np.mean(
                    r["q_counts"][4]
                )
            )

            row = (
                case.case_id,
                case.pca_window,
                case.pc_remove,
                case.formation_n,
                h,
                int(
                    len(spread)
                ),
                q_means[0],
                q_means[1],
                q_means[2],
                q_means[3],
                q_means[4],
                (
                    float(spread_mean)
                    if np.isfinite(
                        spread_mean
                    )
                    else None
                ),
                (
                    float(spread_t)
                    if np.isfinite(
                        spread_t
                    )
                    else None
                ),
                (
                    float(spread_p)
                    if np.isfinite(
                        spread_p
                    )
                    else None
                ),
                beats_ratio,
                avg_q1_n,
                avg_q5_n,
                delta_mean,
                delta_t,
                delta_p,
            )

            summary_rows.append(
                row
            )

            summary_lookup[
                (
                    case.case_id,
                    h,
                )
            ] = {
                "Spread": spread_mean,
                "Delta": delta_mean,
                "Delta_t": delta_t,
                "Delta_p": delta_p,
            }

            # -------------------------------------------------------------
            # DAILY_SPREAD
            # -------------------------------------------------------------
            for i, d in enumerate(
                r["dates"]
            ):
                q1 = float(
                    r["q"][0][i]
                )
                q2 = float(
                    r["q"][1][i]
                )
                q3 = float(
                    r["q"][2][i]
                )
                q4 = float(
                    r["q"][3][i]
                )
                q5 = float(
                    r["q"][4][i]
                )
                sp = float(
                    r["spread"][i]
                )

                raw_spread = raw_map.get(
                    int(d)
                )

                if case.pc_remove == "RAW":
                    # Raw case는 자기 자신이 benchmark이므로 Delta=0
                    raw_spread_out = sp
                    delta_spread = 0.0
                else:
                    raw_spread_out = (
                        float(raw_spread)
                        if raw_spread is not None
                        else None
                    )
                    delta_spread = (
                        float(
                            sp - raw_spread
                        )
                        if raw_spread is not None
                        else None
                    )

                daily_rows.append(
                    (
                        int(d),
                        case.case_id,
                        case.pca_window,
                        case.pc_remove,
                        case.formation_n,
                        h,
                        q1,
                        q2,
                        q3,
                        q4,
                        q5,
                        sp,
                        raw_case_id,
                        raw_spread_out,
                        delta_spread,
                    )
                )

    summary_rows.sort(
        key=lambda r: (
            r[4],        # horizon
            r[3],        # formation
            -(r[11] if r[11] is not None else -999),
        )
    )

    daily_rows.sort(
        key=lambda r: (
            r[0],        # date
            r[5],        # horizon
            r[4],        # formation
            r[1],        # case
        )
    )

    # -------------------------------------------------------------------------
    # Same n, same h: Raw vs best PCA spread
    # -------------------------------------------------------------------------
    best_rows = []

    for n in FORMATION_HORIZONS:
        raw_case = raw_by_n[
            n
        ]

        for h in FUTURE_HORIZONS:
            raw_info = summary_lookup.get(
                (
                    raw_case,
                    h,
                )
            )

            if raw_info is None:
                continue

            candidates = [
                c
                for c in cases
                if (
                    c.pc_remove != "RAW"
                    and c.formation_n == n
                    and (
                        c.case_id,
                        h,
                    )
                    in summary_lookup
                )
            ]

            if not candidates:
                continue

            best = max(
                candidates,
                key=lambda c: (
                    summary_lookup[
                        (
                            c.case_id,
                            h,
                        )
                    ]["Spread"]
                    if summary_lookup[
                        (
                            c.case_id,
                            h,
                        )
                    ]["Spread"] is not None
                    else -np.inf
                ),
            )

            b = summary_lookup[
                (
                    best.case_id,
                    h,
                )
            ]

            best_rows.append(
                (
                    n,
                    h,
                    raw_case,
                    raw_info[
                        "Spread"
                    ],
                    best.case_id,
                    b[
                        "Spread"
                    ],
                    b[
                        "Delta"
                    ],
                    b[
                        "Delta_t"
                    ],
                    b[
                        "Delta_p"
                    ],
                )
            )

    best_rows.sort(
        key=lambda r: (
            r[1],
            r[0],
        )
    )

    return (
        daily_rows,
        summary_rows,
        best_rows,
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
            out.write("<cols>")

            for i, w in enumerate(
                widths,
                start=1,
            ):
                out.write(
                    f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                )

            out.write("</cols>")

        out.write("<sheetData>")
        out.write('<row r="1">')

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

        out.write("</row>")

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

            out.write("</row>")

        out.write("</sheetData>")

        if header:
            last_col = idx_to_col(
                len(header) - 1
            )

            out.write(
                f'<autoFilter ref="A1:{last_col}{last_row}"/>'
            )

        out.write("</worksheet>")


def package_xlsx(
    output_path: Path,
    sheet_specs: list[
        tuple[str, Path]
    ],
):
    sheets_xml = "".join(
        f'<sheet name="{escape(name)}" '
        f'sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, _) in enumerate(
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
        + "".join(rels)
        + "</Relationships>"
    )

    # workbook/styles는 template에 이미 1회 등록됨.
    overrides = []

    for i in range(
        1,
        len(sheet_specs) + 1,
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

        for i, (_, p) in enumerate(
            sheet_specs,
            start=1,
        ):
            z.write(
                p,
                f"xl/worksheets/sheet{i}.xml",
            )


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--data-dir",
        default=None,
        help="5개 입력 Excel이 있는 상위폴더. 생략 시 이 py 폴더 이하 재귀탐색.",
    )

    parser.add_argument(
        "--start",
        default=None,
        help="분석 시작일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--end",
        default=None,
        help="분석 종료일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과 폴더",
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
        else script_dir
        / "QUINTILE_REVERSAL_ANALYSIS"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "QUINTILE_REVERSAL_ANALYSIS_DAILY.xlsx"
    )

    start_date = parse_cli_date(
        args.start
    )

    end_date = parse_cli_date(
        args.end
    )

    # -------------------------------------------------------------------------
    # Required files
    # -------------------------------------------------------------------------
    file_map = {
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

    print("=" * 100)
    print("QUINTILE REVERSAL SPREAD ANALYSIS")
    print("=" * 100)

    for k, p in file_map.items():
        print(
            f"{k:<8}: {p}"
        )

    print()

    # -------------------------------------------------------------------------
    # Common252
    # -------------------------------------------------------------------------
    print(
        "[1/5] Common252 universe 읽기...",
        flush=True,
    )

    (
        dates,
        tickers,
        common,
    ) = read_wide_matrix(
        file_map["base"],
        "COMMON252_ELIGIBLE",
    )

    common_bool = (
        np.nan_to_num(
            common,
            nan=0.0,
        )
        == 1
    )

    T, N = common.shape

    date_mask = np.ones(
        T,
        dtype=bool,
    )

    if start_date is not None:
        date_mask &= (
            dates >= start_date
        )

    if end_date is not None:
        date_mask &= (
            dates <= end_date
        )

    analysis_idx = np.flatnonzero(
        date_mask
    )

    if len(
        analysis_idx
    ) == 0:
        raise RuntimeError(
            "분석 날짜가 0개입니다."
        )

    print(
        f"      T={T:,}, N={N:,}",
        flush=True,
    )

    print(
        f"      분석기간={yyyymmdd_text(dates[analysis_idx[0]])} ~ "
        f"{yyyymmdd_text(dates[analysis_idx[-1]])}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Future
    # -------------------------------------------------------------------------
    print(
        "[2/5] 미래수익률 5/10/20/60 읽기...",
        flush=True,
    )

    future = {}

    for h in FUTURE_HORIZONS:
        d, t, arr = read_wide_matrix(
            file_map["future"],
            f"FWD_LOG_{h}",
        )

        require_same_dates(
            dates,
            d,
            f"FWD_LOG_{h}",
        )

        require_same_tickers(
            tickers,
            t,
            f"FWD_LOG_{h}",
        )

        future[h] = arr

    # -------------------------------------------------------------------------
    # Cases
    # -------------------------------------------------------------------------
    cases = make_cases()

    case_results = {}

    print(
        "[3/5] 40개 Case의 5분위 Q5-Q1 Spread 계산...",
        flush=True,
    )

    for case_no, case in enumerate(
        cases,
        start=1,
    ):
        source_path = file_map[
            case.source_key
        ]

        d, t, cumulative = read_wide_matrix(
            source_path,
            case.source_sheet,
        )

        require_same_dates(
            dates,
            d,
            case.case_id,
        )

        require_same_tickers(
            tickers,
            t,
            case.case_id,
        )

        # Signal higher = stronger loser
        signal = -cumulative

        case_results[
            case.case_id
        ] = analyze_case(
            case=case,
            signal=signal,
            common_bool=common_bool,
            future=future,
            analysis_idx=analysis_idx,
            dates=dates,
        )

        if (
            case_no % 5 == 0
            or case_no == 1
            or case_no == len(cases)
        ):
            print(
                f"      [{case_no:02d}/40] {case.case_id}",
                flush=True,
            )

        del cumulative
        del signal

    # -------------------------------------------------------------------------
    # Summary
    # -------------------------------------------------------------------------
    print(
        "[4/5] Raw 대비 PCA Spread 개선효과 + HAC 계산...",
        flush=True,
    )

    (
        daily_rows,
        summary_rows,
        best_rows,
    ) = summarize_case_results(
        cases=cases,
        case_results=case_results,
    )

    # -------------------------------------------------------------------------
    # Output
    # -------------------------------------------------------------------------
    print(
        "[5/5] Excel 저장...",
        flush=True,
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs = []

        # PARAMETERS ---------------------------------------------------------
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Analysis",
                "5-quintile short-term reversal portfolio sort",
            ),
            (
                "Start",
                yyyymmdd_text(
                    int(
                        dates[
                            analysis_idx[0]
                        ]
                    )
                ),
            ),
            (
                "End",
                yyyymmdd_text(
                    int(
                        dates[
                            analysis_idx[-1]
                        ]
                    )
                ),
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
                "Q1",
                "Lowest reversal signal = past winner 20% = Short",
            ),
            (
                "Q5",
                "Highest reversal signal = past loser 20% = Long",
            ),
            (
                "Spread",
                "Q5 equal-weight future simple return - Q1 equal-weight future simple return",
            ),
            (
                "FutureReturn",
                "exp(FWD_LOG)-1 before equal-weight portfolio averaging",
            ),
            (
                "QuantileFormation",
                "Common252 + finite Signal only; future return is NOT used for sorting",
            ),
            (
                "DeltaSpread",
                "PCA Spread - same Formation_N Raw Spread",
            ),
            (
                "HAC_Lag",
                "max(Formation_N, Horizon)-1",
            ),
        ]

        write_table_sheet(
            p,
            ["Item", "Value"],
            parameter_rows,
            [27, 84],
        )

        specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # DAILY SPREAD -------------------------------------------------------
        p = td / "daily_spread.xml"

        daily_header = [
            "Date",
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "Q1_Return",
            "Q2_Return",
            "Q3_Return",
            "Q4_Return",
            "Q5_Return",
            "Q5_Q1_Spread",
            "Raw_Case",
            "Raw_Spread",
            "DeltaSpread_vs_Raw",
        ]

        daily_out = (
            (
                yyyymmdd_text(
                    r[0]
                ),
                *r[1:],
            )
            for r in daily_rows
        )

        write_table_sheet(
            p,
            daily_header,
            daily_out,
            [
                13,
                24,
                14,
                14,
                14,
                12,
                16,
                16,
                16,
                16,
                16,
                18,
                20,
                18,
                20,
            ],
        )

        specs.append(
            (
                "DAILY_SPREAD",
                p,
            )
        )

        # SPREAD SUMMARY -----------------------------------------------------
        p = td / "spread_summary.xml"

        summary_header = [
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "ValidDates",
            "Q1_MeanFwdReturn",
            "Q2_MeanFwdReturn",
            "Q3_MeanFwdReturn",
            "Q4_MeanFwdReturn",
            "Q5_MeanFwdReturn",
            "Q5_Q1_Spread",
            "Spread_HAC_t",
            "Spread_HAC_p",
            "Q5_Beats_Q1_Ratio",
            "Avg_Q1_Stocks",
            "Avg_Q5_Stocks",
            "DeltaSpread_vs_Raw",
            "Delta_HAC_t",
            "Delta_HAC_p",
        ]

        write_table_sheet(
            p,
            summary_header,
            summary_rows,
            [
                24,
                14,
                14,
                14,
                12,
                14,
                18,
                18,
                18,
                18,
                18,
                17,
                15,
                15,
                20,
                15,
                15,
                20,
                16,
                16,
            ],
        )

        specs.append(
            (
                "SPREAD_SUMMARY",
                p,
            )
        )

        # BEST COMPARISON ----------------------------------------------------
        p = td / "best_comparison.xml"

        best_header = [
            "Formation_N",
            "Horizon",
            "Raw_Case",
            "Raw_Q5_Q1_Spread",
            "Best_PCA_Case",
            "Best_PCA_Q5_Q1_Spread",
            "DeltaSpread_vs_Raw",
            "Delta_HAC_t",
            "Delta_HAC_p",
        ]

        write_table_sheet(
            p,
            best_header,
            best_rows,
            [
                14,
                12,
                20,
                21,
                26,
                25,
                20,
                16,
                16,
            ],
        )

        specs.append(
            (
                "BEST_COMPARISON",
                p,
            )
        )

        package_xlsx(
            output_path,
            specs,
        )

    # structural validation
    with ZipFile(
        output_path
    ) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성 XLSX ZIP 오류: {bad}"
            )

    print()
    print("=" * 100)
    print("DONE")
    print("=" * 100)
    print("OUTPUT:", output_path)
    print()
    print("먼저 볼 시트:")
    print("  1) DAILY_SPREAD")
    print("  2) BEST_COMPARISON")
    print("  3) SPREAD_SUMMARY")
    print()
    print("핵심 열:")
    print("  Q5_Q1_Spread")
    print("  Spread_HAC_p")
    print("  DeltaSpread_vs_Raw")
    print("  Delta_HAC_p")
    print()
    print("해석:")
    print("  Q5_Q1_Spread > 0  -> loser Long / winner Short reversal")
    print("  DeltaSpread_vs_Raw > 0 -> PCA 제거가 Raw reversal보다 개선")


if __name__ == "__main__":
    main()
