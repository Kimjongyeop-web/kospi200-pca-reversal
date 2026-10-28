# -*- coding: utf-8 -*-
"""
20_DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST.py
=================================================

부가연구용 통합 백테스트: DeltaQIC5 기준.

입력 QUANTILE_RANK_IC_5BIN_ANALYSIS.xlsx의 DAILY_FWD5 시트에 저장된:

    DeltaQIC5_vs_Raw
      = QIC5(PCA Case) - QIC5(RAW Case with the same Formation_N)

를 매일의 PCA Case 선택점수로 사용한다.

예:
    W252_PC123_N5  -> QIC5(W252_PC123_N5) - QIC5(RAW_N5)
    W120_PC1_N10   -> QIC5(W120_PC1_N10)  - QIC5(RAW_N10)

시차
----
source date = s의 DeltaQIC5_FWD5는 s+1~s+5 미래수익률을 사용하므로
s+5 종가 후에야 알 수 있다.

현재 decision date t=s+5에서:
    1) PCA 36개 Case 중 DeltaQIC5_s가 가장 높은 Case 선택
    2) 과거 s일 종목을 사용하지 않고 현재 t일의 선택 Case Signal 계산
    3) 현재 Signal Q5(상위 20%, residual loser)와 Q1(하위 20%, winner) 생성
    4) 다음 거래일 t+1 종가 진입
    5) 5거래일 보유 후 종가 청산

즉:
    DeltaQIC5_{t-5}
      -> 현재 t의 PCA Case
      -> 현재 t의 동일 Q1/Q5
      -> t+1 종가 진입
      -> 5거래일 보유

세 전략은 매일 완전히 동일한 선택 Case와 동일한 Q1/Q5 종목을 공유한다.
따라서 성과 차이는 Case 선택 차이가 아니라 실제 사용하는 leg 차이이다.

전략
----
1) LONG_SHORT
   - 신규 cohort 총 gross 20%
   - Q5 Long 10% + Q1 Short 10%

2) LONG_ONLY
   - 신규 cohort 총 gross 20%
   - Q5 Long 20%

3) SHORT_ONLY
   - 신규 cohort 총 gross 20%
   - Q1 Short 20%

공통 규칙
---------
- 매일 신규 cohort 20%
- 5거래일 보유
- 정상상태에서 5개 cohort가 겹쳐 약 100% gross exposure
- 최고 lagged DeltaQIC5가 0 이하이면 신규 cohort는 Cash (기본값)
- 종목별 equal weight
- 매수비용 5bp / 매도비용 20bp
- 만기와 신규 주문을 종목별 signed-share 변화로 netting 후 순거래에만 비용 부과
- Short borrow fee는 포함하지 않음

필요 입력 Excel 6개
-------------------
1) QUANTILE_RANK_IC_5BIN_ANALYSIS.xlsx
   - CASE_DEFINITION
   - DAILY_FWD5의 DeltaQIC5_vs_Raw

2) 00_COMMON252_BASE.xlsx
   - COMMON252_ELIGIBLE

3) 02_PCA_W60_CUMULATIVE.xlsx
4) 04_PCA_W120_CUMULATIVE.xlsx
5) 06_PCA_W252_CUMULATIVE.xlsx
   - 현재 decision date의 PCA Signal

6) KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx
   - 수정주가

FUTURE_LOG_RETURNS 파일은 필요하지 않다. DeltaQIC5는 1번 파일에 이미 저장돼 있고,
실제 PnL은 수정주가에서 직접 계산한다.

출력
----
DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST/
  DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST.xlsx
  graphs/*.png

Excel sheets
------------
PARAMETERS
COMPARISON_SUMMARY
YEARLY_COMPARISON
DAILY_COMPARISON
DAILY_LONG_SHORT
DAILY_LONG_ONLY
DAILY_SHORT_ONLY
COHORT_SELECTION
COHORTS_LONG_SHORT
COHORTS_LONG_ONLY
COHORTS_SHORT_ONLY
CASE_USAGE

기본 실행
---------
python 20_DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST.py

최고 DeltaQIC5가 0 이하인 날도 진입:
python 20_DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST.py --allow-nonpositive-delta-qic5

필수 외부 패키지
----------------
numpy
matplotlib
"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import numpy as np
import matplotlib.pyplot as plt


# =============================================================================
# CONSTANTS
# =============================================================================

QIC_HORIZON = 5
HOLD_DAYS = 5
COHORT_GROSS_FRACTION = 0.20
DEFAULT_BUY_COST_BPS = 5.0
DEFAULT_SELL_COST_BPS = 20.0
TOP_FRACTION = 0.20
MIN_FORMATION_STOCKS = 50
MIN_LEG_STOCKS = 5

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


def find_input_file(base: Path, key: str) -> Path:
    key_norm = normalize_filename_token(key)

    candidates = []

    for p in base.rglob("*.xlsx"):
        name = normalize_filename_token(p.name)

        if key_norm in name:
            if "ic_k5_pca_portfolio_backtest" in name:
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


def find_full_history_price_file(base: Path) -> Path:
    candidates = []

    for p in base.rglob("*.xlsx"):
        token = normalize_filename_token(p.name)

        if (
            "full_history_pca_ready" in token
            and "qc" not in token
            and "분석데이터" in p.name
        ):
            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            "\nKOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026*.xlsx "
            "파일을 찾지 못했습니다.\n"
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


def align_matrix(
    ref_dates: np.ndarray,
    ref_tickers: list[str],
    dates: np.ndarray,
    tickers: list[str],
    arr: np.ndarray,
    label: str,
) -> np.ndarray:
    """
    날짜/티커 순서가 달라도 reference에 맞춰 정렬.
    """
    date_map = {
        int(d): i
        for i, d in enumerate(dates)
    }

    ticker_map = {
        t: j
        for j, t in enumerate(tickers)
        if t
    }

    out = np.full(
        (
            len(ref_dates),
            len(ref_tickers),
        ),
        np.nan,
        dtype=np.float64,
    )

    src_cols = []
    dst_cols = []

    for j, t in enumerate(ref_tickers):
        sj = ticker_map.get(t)

        if sj is not None:
            src_cols.append(sj)
            dst_cols.append(j)

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

    for di, d in enumerate(ref_dates):
        si = date_map.get(
            int(d)
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
# CASE / IC
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    pca_window: int
    pc_remove: str
    formation_n: int
    source_key: str
    source_sheet: str


def parse_pca_case_id(case_id: str) -> CaseDef | None:
    m = re.fullmatch(
        r"W(60|120|252)_(PC1|PC12|PC123)_N(5|10|30|60)",
        case_id,
    )

    if not m:
        return None

    W = int(
        m.group(1)
    )

    pc = m.group(2)

    n = int(
        m.group(3)
    )

    return CaseDef(
        case_id=case_id,
        pca_window=W,
        pc_remove=pc,
        formation_n=n,
        source_key=f"cum{W}",
        source_sheet=f"{pc}_CUM{n}",
    )


def read_pca_case_definitions(
    path: Path,
) -> list[CaseDef]:
    it = iter_sheet_rows(
        path,
        "CASE_DEFINITION",
    )

    header_cells = next(it)

    max_col = max(
        header_cells
    )

    headers = [
        str(
            header_cells.get(
                c,
                "",
            )
            or ""
        ).strip()
        for c in range(
            max_col + 1
        )
    ]

    if "CaseID" not in headers:
        raise KeyError(
            "CASE_DEFINITION: CaseID 열 없음"
        )

    case_col = headers.index(
        "CaseID"
    )

    cases = []
    seen = set()

    for cells in it:
        case_id = str(
            cells.get(
                case_col,
                "",
            )
            or ""
        ).strip()

        if (
            not case_id
            or case_id in seen
        ):
            continue

        case = parse_pca_case_id(
            case_id
        )

        if case is None:
            continue

        seen.add(
            case_id
        )

        cases.append(
            case
        )

    if not cases:
        raise RuntimeError(
            "PCA Case를 찾지 못했습니다."
        )

    cases.sort(
        key=lambda c: (
            c.pca_window,
            c.pc_remove,
            c.formation_n,
        )
    )

    return cases


def read_delta_qic5_fwd5_matrix(
    path: Path,
    ref_dates: np.ndarray,
    cases: list[CaseDef],
) -> np.ndarray:
    """Read PCA-case DeltaQIC5_vs_Raw from DAILY_FWD5."""
    it = iter_sheet_rows(
        path,
        "DAILY_FWD5",
    )

    header_cells = next(it)
    max_col = max(header_cells)

    headers = [
        str(header_cells.get(c, "") or "").strip()
        for c in range(max_col + 1)
    ]

    required = [
        "Date",
        "CaseID",
        "DeltaQIC5_vs_Raw",
    ]

    missing = [x for x in required if x not in headers]
    if missing:
        raise KeyError(
            f"DAILY_FWD5 필수열 누락: {missing}"
        )

    date_col = headers.index("Date")
    case_col = headers.index("CaseID")
    delta_col = headers.index("DeltaQIC5_vs_Raw")

    date_map = {
        int(d): i
        for i, d in enumerate(ref_dates)
    }

    case_map = {
        c.case_id: j
        for j, c in enumerate(cases)
    }

    out = np.full(
        (len(ref_dates), len(cases)),
        np.nan,
        dtype=np.float64,
    )

    rows = 0

    for cells in it:
        case_id = str(
            cells.get(case_col, "") or ""
        ).strip()

        # read_pca_case_definitions() returns PCA 36 cases only.
        cj = case_map.get(case_id)
        if cj is None:
            continue

        raw_date = cells.get(date_col)
        raw_delta = cells.get(delta_col)

        if raw_date in (None, "") or raw_delta in (None, ""):
            continue

        try:
            d = parse_date_value(raw_date)
            delta_qic5 = float(raw_delta)
        except Exception:
            continue

        di = date_map.get(d)

        if di is not None and np.isfinite(delta_qic5):
            out[di, cj] = delta_qic5
            rows += 1

    print(
        f"      DAILY_FWD5 PCA DeltaQIC5 rows read={rows:,}",
        flush=True,
    )

    return out


# =============================================================================
# SIGNAL SELECTION
# =============================================================================

@dataclass
class DecisionRecord:
    delta_qic_source_idx: int
    decision_idx: int
    entry_idx: int
    exit_idx: int
    case_idx: int
    case_id: str
    lagged_delta_qic5: float


@dataclass
class CohortSpec:
    delta_qic_source_idx: int
    decision_idx: int
    entry_idx: int
    exit_idx: int
    case_id: str
    lagged_delta_qic5: float
    long_idx: np.ndarray
    short_idx: np.ndarray
    entry_missing_filtered: int


def build_decision_records(
    dates: np.ndarray,
    delta_qic_matrix: np.ndarray,
    cases: list[CaseDef],
    start_date: int | None,
    end_date: int | None,
    require_positive_delta_qic5: bool,
):
    records = []
    skipped_nonpositive_delta_qic5 = 0
    skipped_no_delta_qic5 = 0

    T = len(dates)

    # entry = decision + 1
    # qic_source = decision - 5
    # exit = entry + 5
    for decision_idx in range(
        QIC_HORIZON,
        T - 1,
    ):
        entry_idx = (
            decision_idx + 1
        )

        exit_idx = (
            entry_idx
            + HOLD_DAYS
        )

        if exit_idx >= T:
            continue

        entry_date = int(
            dates[
                entry_idx
            ]
        )

        if (
            start_date is not None
            and entry_date < start_date
        ):
            continue

        if (
            end_date is not None
            and entry_date > end_date
        ):
            continue

        delta_qic_source_idx = (
            decision_idx
            - QIC_HORIZON
        )

        row = delta_qic_matrix[
            delta_qic_source_idx
        ]

        finite = np.isfinite(
            row
        )

        if not finite.any():
            skipped_no_delta_qic5 += 1
            continue

        score = np.where(
            finite,
            row,
            -np.inf,
        )

        case_idx = int(
            np.argmax(
                score
            )
        )

        best_delta_qic5 = float(
            row[
                case_idx
            ]
        )

        if (
            require_positive_delta_qic5
            and best_delta_qic5 <= 0
        ):
            skipped_nonpositive_delta_qic5 += 1
            continue

        records.append(
            DecisionRecord(
                delta_qic_source_idx=delta_qic_source_idx,
                decision_idx=decision_idx,
                entry_idx=entry_idx,
                exit_idx=exit_idx,
                case_idx=case_idx,
                case_id=cases[
                    case_idx
                ].case_id,
                lagged_delta_qic5=best_delta_qic5,
            )
        )

    return (
        records,
        skipped_nonpositive_delta_qic5,
        skipped_no_delta_qic5,
    )


def make_quintile_legs(
    signal_row: np.ndarray,
    eligible_row: np.ndarray,
):
    mask = (
        eligible_row
        & np.isfinite(
            signal_row
        )
    )

    idx = np.flatnonzero(
        mask
    )

    if len(idx) < MIN_FORMATION_STOCKS:
        return None

    order = np.argsort(
        signal_row[idx],
        kind="mergesort",
    )

    sorted_idx = idx[
        order
    ]

    groups = np.array_split(
        sorted_idx,
        5,
    )

    if (
        len(
            groups[0]
        ) < MIN_LEG_STOCKS
        or len(
            groups[4]
        ) < MIN_LEG_STOCKS
    ):
        return None

    # Q1 = low signal = winner = Short
    # Q5 = high signal = loser = Long
    return (
        groups[4].copy(),
        groups[0].copy(),
    )


def build_cohort_specs(
    decisions: list[DecisionRecord],
    cases: list[CaseDef],
    file_map: dict[str, Path],
    dates: np.ndarray,
    tickers: list[str],
    common_bool: np.ndarray,
    raw_prices: np.ndarray,
):
    by_case = defaultdict(
        list
    )

    for d in decisions:
        by_case[
            d.case_id
        ].append(
            d
        )

    case_lookup = {
        c.case_id: c
        for c in cases
    }

    specs = []
    skipped_signal = 0
    skipped_entry = 0

    for no, (
        case_id,
        recs,
    ) in enumerate(
        sorted(
            by_case.items()
        ),
        start=1,
    ):
        case = case_lookup[
            case_id
        ]

        print(
            f"      Signal load [{no:02d}/{len(by_case):02d}] {case_id} | uses={len(recs):,}",
            flush=True,
        )

        source_path = file_map[
            case.source_key
        ]

        d2, t2, cumulative = read_wide_matrix(
            source_path,
            case.source_sheet,
        )

        cumulative = align_matrix(
            dates,
            tickers,
            d2,
            t2,
            cumulative,
            case_id,
        )

        signal = -cumulative

        for rec in recs:
            legs = make_quintile_legs(
                signal_row=signal[
                    rec.decision_idx
                ],
                eligible_row=common_bool[
                    rec.decision_idx
                ],
            )

            if legs is None:
                skipped_signal += 1
                continue

            long_idx, short_idx = legs

            entry_price = raw_prices[
                rec.entry_idx
            ]

            long_ok = (
                np.isfinite(
                    entry_price[
                        long_idx
                    ]
                )
                & (
                    entry_price[
                        long_idx
                    ]
                    > 0
                )
            )

            short_ok = (
                np.isfinite(
                    entry_price[
                        short_idx
                    ]
                )
                & (
                    entry_price[
                        short_idx
                    ]
                    > 0
                )
            )

            missing_count = int(
                (~long_ok).sum()
                + (~short_ok).sum()
            )

            long_idx = long_idx[
                long_ok
            ]

            short_idx = short_idx[
                short_ok
            ]

            if (
                len(
                    long_idx
                ) < MIN_LEG_STOCKS
                or len(
                    short_idx
                ) < MIN_LEG_STOCKS
            ):
                skipped_entry += 1
                continue

            specs.append(
                CohortSpec(
                    delta_qic_source_idx=rec.delta_qic_source_idx,
                    decision_idx=rec.decision_idx,
                    entry_idx=rec.entry_idx,
                    exit_idx=rec.exit_idx,
                    case_id=rec.case_id,
                    lagged_delta_qic5=rec.lagged_delta_qic5,
                    long_idx=long_idx,
                    short_idx=short_idx,
                    entry_missing_filtered=missing_count,
                )
            )

        del cumulative
        del signal

    specs.sort(
        key=lambda x: x.entry_idx
    )

    return (
        specs,
        skipped_signal,
        skipped_entry,
    )


# =============================================================================
# PRICE / MARKING
# =============================================================================

def forward_fill_prices(
    raw: np.ndarray,
) -> np.ndarray:
    mark = raw.copy()

    T, N = mark.shape

    for j in range(N):
        last = np.nan

        for t in range(T):
            x = mark[
                t,
                j
            ]

            if (
                np.isfinite(x)
                and x > 0
            ):
                last = x
            elif np.isfinite(
                last
            ):
                mark[
                    t,
                    j
                ] = last

    return mark


# =============================================================================
# PORTFOLIO SIMULATION
# =============================================================================

@dataclass
class ActiveCohort:
    spec: CohortSpec
    initial_gross: float
    long_shares: np.ndarray
    short_shares: np.ndarray
    gross_pnl: float
    entry_equity: float


def exposure_of_cohort(
    c: ActiveCohort,
    mark_row: np.ndarray,
):
    long_val = (
        c.long_shares
        * mark_row[
            c.spec.long_idx
        ]
    )

    short_val = (
        c.short_shares
        * mark_row[
            c.spec.short_idx
        ]
    )

    gross = float(
        np.nansum(
            np.abs(
                long_val
            )
        )
        + np.nansum(
            np.abs(
                short_val
            )
        )
    )

    net = float(
        np.nansum(
            long_val
        )
        + np.nansum(
            short_val
        )
    )

    return (
        gross,
        net,
    )



@dataclass
class StrategyActiveCohort:
    spec: CohortSpec
    initial_gross: float
    long_idx: np.ndarray
    short_idx: np.ndarray
    long_shares: np.ndarray
    short_shares: np.ndarray
    gross_pnl: float
    entry_equity: float


STRATEGY_MODES = (
    "LONG_SHORT",
    "LONG_ONLY",
    "SHORT_ONLY",
)


def strategy_leg_fractions(
    mode: str,
) -> tuple[float, float]:
    if mode == "LONG_SHORT":
        return 0.50, 0.50

    if mode == "LONG_ONLY":
        return 1.00, 0.00

    if mode == "SHORT_ONLY":
        return 0.00, 1.00

    raise ValueError(
        f"지원하지 않는 strategy mode: {mode}"
    )


def active_exposure(
    cohort: StrategyActiveCohort,
    mark_row: np.ndarray,
) -> tuple[float, float]:
    long_value = 0.0
    short_value = 0.0

    if len(cohort.long_idx):
        long_value = float(
            np.nansum(
                cohort.long_shares
                * mark_row[
                    cohort.long_idx
                ]
            )
        )

    if len(cohort.short_idx):
        short_value = float(
            np.nansum(
                cohort.short_shares
                * mark_row[
                    cohort.short_idx
                ]
            )
        )

    gross = (
        abs(long_value)
        + abs(short_value)
    )

    net = (
        long_value
        + short_value
    )

    return float(gross), float(net)


def simulate_strategy(
    mode: str,
    specs: list[CohortSpec],
    dates: np.ndarray,
    tickers: list[str],
    raw_prices: np.ndarray,
    mark_prices: np.ndarray,
    buy_cost_rate: float,
    sell_cost_rate: float,
):
    """Simulate one of LONG_SHORT, LONG_ONLY, SHORT_ONLY."""
    long_fraction, short_fraction = strategy_leg_fractions(
        mode
    )

    entries_by_day = defaultdict(list)

    for spec in specs:
        entries_by_day[
            spec.entry_idx
        ].append(
            spec
        )

    if not specs:
        raise RuntimeError(
            "진입 가능한 cohort가 0개입니다."
        )

    first_entry = min(
        spec.entry_idx
        for spec in specs
    )

    last_exit = max(
        spec.exit_idx
        for spec in specs
    )

    equity = 1.0
    active: list[StrategyActiveCohort] = []

    daily_rows = []
    cohort_rows = []

    total_buy_notional = 0.0
    total_sell_notional = 0.0
    total_netted_trade_notional = 0.0
    total_unnetted_trade_notional = 0.0
    total_trading_cost = 0.0
    total_unnetted_cost = 0.0
    total_netting_cost_savings = 0.0
    total_exit_missing_raw = 0

    N = len(tickers)

    for t in range(
        first_entry,
        last_exit + 1,
    ):
        equity_start = equity
        mtm_pnl = 0.0

        # 1) Close-to-close PnL from positions held overnight.
        if t > 0:
            for cohort in active:
                pnl = 0.0

                if len(cohort.long_idx):
                    long_prev = mark_prices[
                        t - 1,
                        cohort.long_idx,
                    ]
                    long_now = mark_prices[
                        t,
                        cohort.long_idx,
                    ]

                    pnl += float(
                        np.nansum(
                            cohort.long_shares
                            * (
                                long_now
                                - long_prev
                            )
                        )
                    )

                if len(cohort.short_idx):
                    short_prev = mark_prices[
                        t - 1,
                        cohort.short_idx,
                    ]
                    short_now = mark_prices[
                        t,
                        cohort.short_idx,
                    ]

                    pnl += float(
                        np.nansum(
                            cohort.short_shares
                            * (
                                short_now
                                - short_prev
                            )
                        )
                    )

                cohort.gross_pnl += pnl
                mtm_pnl += pnl

            equity += mtm_pnl

        equity_before_trade = equity
        mark_row = mark_prices[t]

        # Positive = BUY shares; negative = SELL shares.
        net_trade_qty = np.zeros(
            N,
            dtype=np.float64,
        )

        unnetted_buy_notional = 0.0
        unnetted_sell_notional = 0.0

        # 2) Expiring cohorts create closing orders.
        still_active = []

        for cohort in active:
            if cohort.spec.exit_idx != t:
                still_active.append(
                    cohort
                )
                continue

            if len(cohort.long_idx):
                # Close Long = SELL.
                net_trade_qty[
                    cohort.long_idx
                ] -= cohort.long_shares

                long_exit_notional = float(
                    np.nansum(
                        np.abs(
                            cohort.long_shares
                            * mark_row[
                                cohort.long_idx
                            ]
                        )
                    )
                )

                unnetted_sell_notional += (
                    long_exit_notional
                )

            if len(cohort.short_idx):
                # Close Short = BUY. short_shares are negative.
                net_trade_qty[
                    cohort.short_idx
                ] -= cohort.short_shares

                short_exit_notional = float(
                    np.nansum(
                        np.abs(
                            cohort.short_shares
                            * mark_row[
                                cohort.short_idx
                            ]
                        )
                    )
                )

                unnetted_buy_notional += (
                    short_exit_notional
                )

            raw_exit = raw_prices[t]

            missing_exit = 0

            if len(cohort.long_idx):
                missing_exit += int(
                    (
                        ~np.isfinite(
                            raw_exit[
                                cohort.long_idx
                            ]
                        )
                    ).sum()
                )

            if len(cohort.short_idx):
                missing_exit += int(
                    (
                        ~np.isfinite(
                            raw_exit[
                                cohort.short_idx
                            ]
                        )
                    ).sum()
                )

            total_exit_missing_raw += (
                missing_exit
            )

            gross_return = (
                cohort.gross_pnl
                / cohort.initial_gross
                if cohort.initial_gross > EPS
                else np.nan
            )

            cohort_rows.append(
                (
                    int(
                        dates[
                            cohort.spec.delta_qic_source_idx
                        ]
                    ),
                    int(
                        dates[
                            cohort.spec.decision_idx
                        ]
                    ),
                    int(
                        dates[
                            cohort.spec.entry_idx
                        ]
                    ),
                    int(
                        dates[
                            cohort.spec.exit_idx
                        ]
                    ),
                    mode,
                    cohort.spec.case_id,
                    float(
                        cohort.spec.lagged_delta_qic5
                    ),
                    int(
                        len(
                            cohort.long_idx
                        )
                    ),
                    int(
                        len(
                            cohort.short_idx
                        )
                    ),
                    float(
                        cohort.entry_equity
                    ),
                    float(
                        cohort.initial_gross
                    ),
                    float(
                        cohort.gross_pnl
                    ),
                    (
                        float(
                            gross_return
                        )
                        if np.isfinite(
                            gross_return
                        )
                        else None
                    ),
                    int(
                        cohort.spec.entry_missing_filtered
                    ),
                    missing_exit,
                    ",".join(
                        tickers[i]
                        for i in cohort.long_idx
                    ),
                    ",".join(
                        tickers[i]
                        for i in cohort.short_idx
                    ),
                )
            )

        active = still_active

        # 3) New cohort(s) at today's close.
        entered_cases = []
        entered_qics = []
        entered_sources = []
        entered_decisions = []

        for spec in entries_by_day.get(
            t,
            [],
        ):
            if equity_before_trade <= EPS:
                continue

            gross_target = (
                COHORT_GROSS_FRACTION
                * equity_before_trade
            )

            long_target = (
                gross_target
                * long_fraction
            )

            short_target = (
                gross_target
                * short_fraction
            )

            entry_row = raw_prices[t]

            long_idx = (
                spec.long_idx.copy()
                if long_target > 0
                else np.asarray(
                    [],
                    dtype=int,
                )
            )

            short_idx = (
                spec.short_idx.copy()
                if short_target > 0
                else np.asarray(
                    [],
                    dtype=int,
                )
            )

            long_shares = np.asarray(
                [],
                dtype=np.float64,
            )
            short_shares = np.asarray(
                [],
                dtype=np.float64,
            )

            if long_target > 0:
                long_price = entry_row[
                    long_idx
                ]

                if not np.all(
                    np.isfinite(
                        long_price
                    )
                ):
                    continue

                long_each = (
                    long_target
                    / len(
                        long_idx
                    )
                )

                long_shares = (
                    long_each
                    / long_price
                )

                # Open Long = BUY.
                net_trade_qty[
                    long_idx
                ] += long_shares

                unnetted_buy_notional += (
                    long_target
                )

            if short_target > 0:
                short_price = entry_row[
                    short_idx
                ]

                if not np.all(
                    np.isfinite(
                        short_price
                    )
                ):
                    continue

                short_each = (
                    short_target
                    / len(
                        short_idx
                    )
                )

                short_shares = -(
                    short_each
                    / short_price
                )

                # Open Short = SELL.
                net_trade_qty[
                    short_idx
                ] += short_shares

                unnetted_sell_notional += (
                    short_target
                )

            active.append(
                StrategyActiveCohort(
                    spec=spec,
                    initial_gross=gross_target,
                    long_idx=long_idx,
                    short_idx=short_idx,
                    long_shares=long_shares,
                    short_shares=short_shares,
                    gross_pnl=0.0,
                    entry_equity=equity_before_trade,
                )
            )

            entered_cases.append(
                spec.case_id
            )
            entered_qics.append(
                float(
                    spec.lagged_delta_qic5
                )
            )
            entered_sources.append(
                yyyymmdd_text(
                    int(
                        dates[
                            spec.delta_qic_source_idx
                        ]
                    )
                )
            )
            entered_decisions.append(
                yyyymmdd_text(
                    int(
                        dates[
                            spec.decision_idx
                        ]
                    )
                )
            )

        # 4) Net entry and exit orders by ticker before costs.
        trade_notional_signed = (
            net_trade_qty
            * mark_row
        )

        valid_trade = (
            np.isfinite(
                trade_notional_signed
            )
            & np.isfinite(
                mark_row
            )
            & (
                mark_row > 0
            )
        )

        buy_notional = float(
            np.sum(
                trade_notional_signed[
                    valid_trade
                    & (
                        trade_notional_signed > 0
                    )
                ]
            )
        )

        sell_notional = float(
            np.sum(
                -trade_notional_signed[
                    valid_trade
                    & (
                        trade_notional_signed < 0
                    )
                ]
            )
        )

        netted_trade_notional = (
            buy_notional
            + sell_notional
        )

        unnetted_trade_notional = (
            unnetted_buy_notional
            + unnetted_sell_notional
        )

        trading_cost_today = (
            buy_cost_rate
            * buy_notional
            + sell_cost_rate
            * sell_notional
        )

        unnetted_cost_today = (
            buy_cost_rate
            * unnetted_buy_notional
            + sell_cost_rate
            * unnetted_sell_notional
        )

        netting_cost_savings = max(
            0.0,
            unnetted_cost_today
            - trading_cost_today,
        )

        equity -= trading_cost_today

        total_buy_notional += buy_notional
        total_sell_notional += sell_notional
        total_netted_trade_notional += (
            netted_trade_notional
        )
        total_unnetted_trade_notional += (
            unnetted_trade_notional
        )
        total_trading_cost += (
            trading_cost_today
        )
        total_unnetted_cost += (
            unnetted_cost_today
        )
        total_netting_cost_savings += (
            netting_cost_savings
        )

        # 5) End-of-day exposures and daily row.
        gross_exposure = 0.0
        net_exposure = 0.0

        for cohort in active:
            g, n = active_exposure(
                cohort,
                mark_row,
            )

            gross_exposure += g
            net_exposure += n

        daily_return = (
            equity
            / equity_start
            - 1.0
            if equity_start > EPS
            else np.nan
        )

        turnover_netted = (
            netted_trade_notional
            / equity_start
            if equity_start > EPS
            else np.nan
        )

        turnover_unnetted = (
            unnetted_trade_notional
            / equity_start
            if equity_start > EPS
            else np.nan
        )

        daily_rows.append(
            (
                int(
                    dates[t]
                ),
                mode,
                float(
                    equity
                ),
                (
                    float(
                        daily_return
                    )
                    if np.isfinite(
                        daily_return
                    )
                    else None
                ),
                float(
                    mtm_pnl
                ),
                float(
                    trading_cost_today
                ),
                int(
                    len(
                        active
                    )
                ),
                (
                    float(
                        gross_exposure
                        / equity
                    )
                    if equity > EPS
                    else None
                ),
                (
                    float(
                        net_exposure
                        / equity
                    )
                    if equity > EPS
                    else None
                ),
                (
                    float(
                        turnover_netted
                    )
                    if np.isfinite(
                        turnover_netted
                    )
                    else None
                ),
                (
                    float(
                        turnover_unnetted
                    )
                    if np.isfinite(
                        turnover_unnetted
                    )
                    else None
                ),
                float(
                    buy_notional
                ),
                float(
                    sell_notional
                ),
                float(
                    netting_cost_savings
                ),
                int(
                    bool(
                        entered_cases
                    )
                ),
                ";".join(
                    entered_cases
                ),
                (
                    float(
                        entered_qics[0]
                    )
                    if len(
                        entered_qics
                    ) == 1
                    else (
                        ";".join(
                            f"{x:.12g}"
                            for x in entered_qics
                        )
                        if entered_qics
                        else None
                    )
                ),
                ";".join(
                    entered_sources
                ),
                ";".join(
                    entered_decisions
                ),
            )
        )

    totals = {
        "TotalBuyNotional": total_buy_notional,
        "TotalSellNotional": total_sell_notional,
        "TotalNettedTurnoverNotional": total_netted_trade_notional,
        "TotalUnnettedTurnoverNotional": total_unnetted_trade_notional,
        "TotalTradingCost": total_trading_cost,
        "TotalUnnettedCost": total_unnetted_cost,
        "TotalNettingCostSavings": total_netting_cost_savings,
        "ExitMissingRawCount": total_exit_missing_raw,
    }

    return {
        "mode": mode,
        "daily_rows": daily_rows,
        "cohort_rows": cohort_rows,
        "totals": totals,
    }


# =============================================================================
# PERFORMANCE METRICS
# =============================================================================

def max_drawdown(
    equity: np.ndarray,
) -> float:
    peak = np.maximum.accumulate(
        equity
    )

    dd = (
        equity
        / peak
        - 1.0
    )

    return float(
        np.nanmin(
            dd
        )
    )


def result_arrays(
    result,
):
    daily = result[
        "daily_rows"
    ]

    dates = np.asarray(
        [
            int(
                r[0]
            )
            for r in daily
        ],
        dtype=np.int32,
    )

    equity = np.asarray(
        [
            float(
                r[2]
            )
            for r in daily
        ],
        dtype=np.float64,
    )

    returns = np.asarray(
        [
            np.nan
            if r[3] is None
            else float(
                r[3]
            )
            for r in daily
        ],
        dtype=np.float64,
    )

    gross_exposure = np.asarray(
        [
            np.nan
            if r[7] is None
            else float(
                r[7]
            )
            for r in daily
        ],
        dtype=np.float64,
    )

    net_exposure = np.asarray(
        [
            np.nan
            if r[8] is None
            else float(
                r[8]
            )
            for r in daily
        ],
        dtype=np.float64,
    )

    turnover = np.asarray(
        [
            np.nan
            if r[9] is None
            else float(
                r[9]
            )
            for r in daily
        ],
        dtype=np.float64,
    )

    return {
        "dates": dates,
        "equity": equity,
        "returns": returns,
        "gross_exposure": gross_exposure,
        "net_exposure": net_exposure,
        "turnover": turnover,
    }


def performance_metrics(
    result,
    no_cost_result,
):
    a = result_arrays(
        result
    )
    a0 = result_arrays(
        no_cost_result
    )

    valid_ret = a[
        "returns"
    ][
        np.isfinite(
            a[
                "returns"
            ]
        )
    ]

    n = len(
        valid_ret
    )

    final_equity = float(
        a[
            "equity"
        ][
            -1
        ]
    )

    total_return = (
        final_equity
        - 1.0
    )

    annualized_return = (
        final_equity ** (
            252.0
            / max(
                1,
                n,
            )
        )
        - 1.0
        if final_equity > 0
        else np.nan
    )

    annualized_vol = (
        float(
            np.std(
                valid_ret,
                ddof=1,
            )
            * math.sqrt(
                252.0
            )
        )
        if n > 1
        else np.nan
    )

    sharpe = (
        float(
            np.mean(
                valid_ret
            )
            / np.std(
                valid_ret,
                ddof=1,
            )
            * math.sqrt(
                252.0
            )
        )
        if (
            n > 1
            and np.std(
                valid_ret,
                ddof=1,
            ) > EPS
        )
        else np.nan
    )

    positive_ratio = (
        float(
            np.mean(
                valid_ret > 0
            )
        )
        if n
        else np.nan
    )

    cohort_returns = np.asarray(
        [
            np.nan
            if r[12] is None
            else float(
                r[12]
            )
            for r in result[
                "cohort_rows"
            ]
        ],
        dtype=np.float64,
    )

    cohort_returns = cohort_returns[
        np.isfinite(
            cohort_returns
        )
    ]

    totals = result[
        "totals"
    ]

    unnetted = float(
        totals[
            "TotalUnnettedTurnoverNotional"
        ]
    )

    netted = float(
        totals[
            "TotalNettedTurnoverNotional"
        ]
    )

    turnover_reduction = (
        1.0
        - netted
        / unnetted
        if unnetted > EPS
        else np.nan
    )

    no_cost_equity = float(
        a0[
            "equity"
        ][
            -1
        ]
    )

    no_cost_return = (
        no_cost_equity
        - 1.0
    )

    no_cost_cagr = (
        no_cost_equity ** (
            252.0
            / max(
                1,
                len(
                    a0[
                        "returns"
                    ]
                ),
            )
        )
        - 1.0
        if no_cost_equity > 0
        else np.nan
    )

    return {
        "Strategy": result[
            "mode"
        ],
        "Start": int(
            a[
                "dates"
            ][
                0
            ]
        ),
        "End": int(
            a[
                "dates"
            ][
                -1
            ]
        ),
        "CompletedCohorts": len(
            result[
                "cohort_rows"
            ]
        ),
        "TotalReturnNet": float(
            total_return
        ),
        "CAGRNet": (
            float(
                annualized_return
            )
            if np.isfinite(
                annualized_return
            )
            else None
        ),
        "AnnualizedVol": (
            float(
                annualized_vol
            )
            if np.isfinite(
                annualized_vol
            )
            else None
        ),
        "SharpeNet": (
            float(
                sharpe
            )
            if np.isfinite(
                sharpe
            )
            else None
        ),
        "MaxDrawdownNet": max_drawdown(
            a[
                "equity"
            ]
        ),
        "PositiveDailyRatio": (
            float(
                positive_ratio
            )
            if np.isfinite(
                positive_ratio
            )
            else None
        ),
        "TotalReturnNoCost": float(
            no_cost_return
        ),
        "CAGRNoCost": (
            float(
                no_cost_cagr
            )
            if np.isfinite(
                no_cost_cagr
            )
            else None
        ),
        "MaxDrawdownNoCost": max_drawdown(
            a0[
                "equity"
            ]
        ),
        "AverageGrossExposure": float(
            np.nanmean(
                a[
                    "gross_exposure"
                ]
            )
        ),
        "AverageNetExposure": float(
            np.nanmean(
                a[
                    "net_exposure"
                ]
            )
        ),
        "AverageDailyTurnover": float(
            np.nanmean(
                a[
                    "turnover"
                ]
            )
        ),
        "TurnoverReductionFromNetting": (
            float(
                turnover_reduction
            )
            if np.isfinite(
                turnover_reduction
            )
            else None
        ),
        "TotalTradingCost": float(
            totals[
                "TotalTradingCost"
            ]
        ),
        "Mean5DCohortGrossReturn": (
            float(
                np.mean(
                    cohort_returns
                )
            )
            if len(
                cohort_returns
            )
            else None
        ),
        "Positive5DCohortRatio": (
            float(
                np.mean(
                    cohort_returns > 0
                )
            )
            if len(
                cohort_returns
            )
            else None
        ),
        "ExitMissingRawPriceCount": int(
            totals[
                "ExitMissingRawCount"
            ]
        ),
    }


def yearly_rows_for_result(
    result,
    no_cost_result,
):
    a = result_arrays(
        result
    )
    a0 = result_arrays(
        no_cost_result
    )

    rows = []

    years = sorted(
        {
            int(
                d
            )
            // 10000
            for d in a[
                "dates"
            ]
        }
    )

    for year in years:
        idx = np.flatnonzero(
            (
                a[
                    "dates"
                ]
                // 10000
            )
            == year
        )

        if len(
            idx
        ) == 0:
            continue

        first = int(
            idx[
                0
            ]
        )
        last = int(
            idx[
                -1
            ]
        )

        start_eq = (
            1.0
            if first == 0
            else float(
                a[
                    "equity"
                ][
                    first - 1
                ]
            )
        )

        start_eq0 = (
            1.0
            if first == 0
            else float(
                a0[
                    "equity"
                ][
                    first - 1
                ]
            )
        )

        rows.append(
            (
                year,
                result[
                    "mode"
                ],
                float(
                    a[
                        "equity"
                    ][
                        last
                    ]
                    / start_eq
                    - 1.0
                ),
                float(
                    a0[
                        "equity"
                    ][
                        last
                    ]
                    / start_eq0
                    - 1.0
                ),
                float(
                    np.nanmean(
                        a[
                            "turnover"
                        ][
                            idx
                        ]
                    )
                ),
            )
        )

    return rows


# =============================================================================
# CHARTS
# =============================================================================

def rolling_mean(
    x: np.ndarray,
    window: int = 20,
) -> np.ndarray:
    out = np.full_like(
        x,
        np.nan,
        dtype=np.float64,
    )

    for i in range(
        window - 1,
        len(
            x
        ),
    ):
        out[
            i
        ] = np.nanmean(
            x[
                i
                - window
                + 1:
                i
                + 1
            ]
        )

    return out


def graph_dates(
    date_values: np.ndarray,
):
    return [
        datetime.strptime(
            yyyymmdd_text(
                int(
                    d
                )
            ),
            "%Y-%m-%d",
        )
        for d in date_values
    ]


def save_graphs(
    results: dict,
    no_cost_results: dict,
    yearly_rows,
    output_dir: Path,
):
    graph_dir = (
        output_dir
        / "graphs"
    )

    graph_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    saved = []

    base_arrays = result_arrays(
        results[
            "LONG_SHORT"
        ]
    )

    dates = graph_dates(
        base_arrays[
            "dates"
        ]
    )

    # Net equity comparison.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    for mode in STRATEGY_MODES:
        a = result_arrays(
            results[
                mode
            ]
        )

        plt.plot(
            dates,
            a[
                "equity"
            ],
            label=mode,
        )

    plt.axhline(
        1.0,
        linewidth=0.8,
    )
    plt.title(
        "DeltaQIC5(K=5): daily net equity comparison"
    )
    plt.xlabel(
        "Date"
    )
    plt.ylabel(
        "Equity (initial = 1)"
    )
    plt.legend()
    plt.grid(
        True,
        alpha=0.25,
    )
    plt.tight_layout()

    p = graph_dir / "01_net_equity_comparison.png"
    plt.savefig(
        p,
        dpi=180,
    )
    plt.close(
        fig
    )
    saved.append(
        p
    )

    # No-cost comparison.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    for mode in STRATEGY_MODES:
        a = result_arrays(
            no_cost_results[
                mode
            ]
        )

        plt.plot(
            dates,
            a[
                "equity"
            ],
            label=mode,
        )

    plt.axhline(
        1.0,
        linewidth=0.8,
    )
    plt.title(
        "DeltaQIC5(K=5): equity before transaction costs"
    )
    plt.xlabel(
        "Date"
    )
    plt.ylabel(
        "Equity (initial = 1)"
    )
    plt.legend()
    plt.grid(
        True,
        alpha=0.25,
    )
    plt.tight_layout()

    p = graph_dir / "02_no_cost_equity_comparison.png"
    plt.savefig(
        p,
        dpi=180,
    )
    plt.close(
        fig
    )
    saved.append(
        p
    )

    # Drawdown comparison.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    for mode in STRATEGY_MODES:
        a = result_arrays(
            results[
                mode
            ]
        )

        peak = np.maximum.accumulate(
            a[
                "equity"
            ]
        )

        dd = (
            a[
                "equity"
            ]
            / peak
            - 1.0
        )

        plt.plot(
            dates,
            dd * 100.0,
            label=mode,
        )

    plt.axhline(
        0.0,
        linewidth=0.8,
    )
    plt.title(
        "DeltaQIC5(K=5): daily drawdown comparison"
    )
    plt.xlabel(
        "Date"
    )
    plt.ylabel(
        "Drawdown (%)"
    )
    plt.legend()
    plt.grid(
        True,
        alpha=0.25,
    )
    plt.tight_layout()

    p = graph_dir / "03_drawdown_comparison.png"
    plt.savefig(
        p,
        dpi=180,
    )
    plt.close(
        fig
    )
    saved.append(
        p
    )

    # Annual returns.
    years = sorted(
        {
            int(
                r[0]
            )
            for r in yearly_rows
        }
    )

    x = np.arange(
        len(
            years
        )
    )
    width = 0.25

    fig = plt.figure(
        figsize=(
            13,
            5,
        )
    )

    for j, mode in enumerate(
        STRATEGY_MODES
    ):
        lookup = {
            int(
                r[0]
            ): float(
                r[2]
            )
            * 100.0
            for r in yearly_rows
            if r[1] == mode
        }

        vals = [
            lookup.get(
                year,
                np.nan,
            )
            for year in years
        ]

        plt.bar(
            x
            + (
                j - 1
            )
            * width,
            vals,
            width=width,
            label=mode,
        )

    plt.axhline(
        0.0,
        linewidth=0.8,
    )
    plt.xticks(
        x,
        years,
        rotation=45,
    )
    plt.title(
        "DeltaQIC5(K=5): calendar-year net returns"
    )
    plt.xlabel(
        "Year"
    )
    plt.ylabel(
        "Net return (%)"
    )
    plt.legend()
    plt.tight_layout()

    p = graph_dir / "04_annual_net_returns.png"
    plt.savefig(
        p,
        dpi=180,
    )
    plt.close(
        fig
    )
    saved.append(
        p
    )

    # Turnover comparison, rolling 20 days.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    for mode in STRATEGY_MODES:
        a = result_arrays(
            results[
                mode
            ]
        )

        plt.plot(
            dates,
            rolling_mean(
                a[
                    "turnover"
                ],
                20,
            )
            * 100.0,
            label=mode,
        )

    plt.title(
        "DeltaQIC5(K=5): netted daily turnover (20D average)"
    )
    plt.xlabel(
        "Date"
    )
    plt.ylabel(
        "Turnover (% of equity)"
    )
    plt.legend()
    plt.grid(
        True,
        alpha=0.25,
    )
    plt.tight_layout()

    p = graph_dir / "05_turnover_comparison.png"
    plt.savefig(
        p,
        dpi=180,
    )
    plt.close(
        fig
    )
    saved.append(
        p
    )

    return saved


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
                len(header) - 1
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
        + "".join(
            rels
        )
        + "</Relationships>"
    )

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
        help=(
            "필수 Excel 6개가 있는 상위폴더. "
            "생략 시 이 py 폴더 이하 재귀탐색."
        ),
    )

    parser.add_argument(
        "--start",
        default=None,
        help="실제 진입 시작일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--end",
        default=None,
        help="실제 진입 종료일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--buy-cost-bps",
        type=float,
        default=DEFAULT_BUY_COST_BPS,
        help="매수 거래비용 bps. 기본 5",
    )

    parser.add_argument(
        "--sell-cost-bps",
        type=float,
        default=DEFAULT_SELL_COST_BPS,
        help="매도 거래비용 bps. 기본 20",
    )

    parser.add_argument(
        "--allow-nonpositive-delta-qic5",
        action="store_true",
        help=(
            "지정하면 최고 PCA DeltaQIC5가 0 이하인 날도 "
            "신규 cohort 진입"
        ),
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과 폴더",
    )

    args = parser.parse_args()

    if args.buy_cost_bps < 0:
        raise ValueError(
            "--buy-cost-bps는 0 이상이어야 합니다."
        )

    if args.sell_cost_bps < 0:
        raise ValueError(
            "--sell-cost-bps는 0 이상이어야 합니다."
        )

    buy_cost_rate = (
        args.buy_cost_bps
        / 10000.0
    )

    sell_cost_rate = (
        args.sell_cost_bps
        / 10000.0
    )

    require_positive_delta_qic5 = (
        not args.allow_nonpositive_delta_qic5
    )

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
            / "DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST"
        )
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "DELTA_QIC5_K5_PCA_THREE_PORTFOLIO_BACKTEST.xlsx"
    )

    start_date = parse_cli_date(
        args.start
    )

    end_date = parse_cli_date(
        args.end
    )

    file_map = {
        "qic": find_input_file(
            data_dir,
            "QUANTILE_RANK_IC_5BIN_ANALYSIS",
        ),
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
        "price": find_full_history_price_file(
            data_dir
        ),
    }

    print(
        "="
        * 100
    )
    print(
        "DELTA QIC5 K=5 -> SAME CASE/Q1/Q5 -> LONG+SHORT / LONG / SHORT"
    )
    print(
        "="
        * 100
    )

    for key, path in file_map.items():
        print(
            f"{key:<8}: {path}"
        )

    print()

    # 1) Common252.
    print(
        "[1/8] Common252 universe 읽기...",
        flush=True,
    )

    dates, tickers, common = read_wide_matrix(
        file_map[
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
        f"      dates={len(dates):,}, tickers={len(tickers):,}",
        flush=True,
    )

    # 2) Price.
    print(
        "[2/8] 수정주가 읽기 / 정렬...",
        flush=True,
    )

    d_price, t_price, price0 = read_wide_matrix(
        file_map[
            "price"
        ],
        "수정주가",
    )

    raw_prices = align_matrix(
        dates,
        tickers,
        d_price,
        t_price,
        price0,
        "수정주가",
    )

    raw_prices[
        (~np.isfinite(raw_prices))
        | (raw_prices <= 0)
    ] = np.nan

    mark_prices = forward_fill_prices(
        raw_prices
    )

    print(
        f"      finite raw prices={int(np.isfinite(raw_prices).sum()):,}",
        flush=True,
    )

    # 3) Cases + DeltaQIC5.
    print(
        "[3/8] PCA Case 정의 + DAILY_FWD5 DeltaQIC5_vs_Raw 읽기...",
        flush=True,
    )

    cases = read_pca_case_definitions(
        file_map[
            "qic"
        ]
    )

    print(
        f"      PCA candidate cases={len(cases)}",
        flush=True,
    )

    delta_qic_matrix = read_delta_qic5_fwd5_matrix(
        file_map[
            "qic"
        ],
        dates,
        cases,
    )

    # 4) Lagged DeltaQIC5 selection.
    print(
        "[4/8] 5거래일 시차를 둔 DeltaQIC5 최고 Case 선택...",
        flush=True,
    )

    (
        decisions,
        skipped_nonpositive_delta_qic5,
        skipped_no_delta_qic5,
    ) = build_decision_records(
        dates=dates,
        delta_qic_matrix=delta_qic_matrix,
        cases=cases,
        start_date=start_date,
        end_date=end_date,
        require_positive_delta_qic5=require_positive_delta_qic5,
    )

    print(
        f"      valid DeltaQIC5 decisions={len(decisions):,}",
        flush=True,
    )
    print(
        f"      skipped nonpositive DeltaQIC5={skipped_nonpositive_delta_qic5:,}",
        flush=True,
    )

    # 5) Same Q1/Q5 selections for all strategies.
    print(
        "[5/8] 선택 Case의 현재 Signal로 공통 Q5/Q1 생성...",
        flush=True,
    )

    (
        specs,
        skipped_signal,
        skipped_entry,
    ) = build_cohort_specs(
        decisions=decisions,
        cases=cases,
        file_map=file_map,
        dates=dates,
        tickers=tickers,
        common_bool=common_bool,
        raw_prices=raw_prices,
    )

    print(
        f"      common cohort selections={len(specs):,}",
        flush=True,
    )
    print(
        f"      skipped signal formation={skipped_signal:,}",
        flush=True,
    )
    print(
        f"      skipped entry tradability={skipped_entry:,}",
        flush=True,
    )

    # 6) Three net and no-cost simulations.
    print(
        "[6/8] 세 포트폴리오 시뮬레이션...",
        flush=True,
    )

    results = {}
    no_cost_results = {}

    for mode in STRATEGY_MODES:
        print(
            f"      {mode} (net cost)",
            flush=True,
        )

        results[
            mode
        ] = simulate_strategy(
            mode=mode,
            specs=specs,
            dates=dates,
            tickers=tickers,
            raw_prices=raw_prices,
            mark_prices=mark_prices,
            buy_cost_rate=buy_cost_rate,
            sell_cost_rate=sell_cost_rate,
        )

        no_cost_results[
            mode
        ] = simulate_strategy(
            mode=mode,
            specs=specs,
            dates=dates,
            tickers=tickers,
            raw_prices=raw_prices,
            mark_prices=mark_prices,
            buy_cost_rate=0.0,
            sell_cost_rate=0.0,
        )

    # 7) Metrics / graphs.
    print(
        "[7/8] 성과표 + 그래프 생성...",
        flush=True,
    )

    metrics = {
        mode: performance_metrics(
            results[
                mode
            ],
            no_cost_results[
                mode
            ],
        )
        for mode in STRATEGY_MODES
    }

    yearly_rows = []

    for mode in STRATEGY_MODES:
        yearly_rows.extend(
            yearly_rows_for_result(
                results[
                    mode
                ],
                no_cost_results[
                    mode
                ],
            )
        )

    graph_paths = save_graphs(
        results=results,
        no_cost_results=no_cost_results,
        yearly_rows=yearly_rows,
        output_dir=output_dir,
    )

    # 8) Excel.
    print(
        "[8/8] 비교 Excel 저장...",
        flush=True,
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(
            td
        )

        sheet_specs = []

        # PARAMETERS.
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Selector",
                "Lagged DeltaQIC5_vs_Raw from DAILY_FWD5",
            ),
            (
                "Candidate_Cases",
                "PCA only: W60/W120/W252 x PC1/PC12/PC123 x N5/N10/N30/N60",
            ),
            (
                "DeltaQIC5_Availability",
                "DeltaQIC5 at source s becomes known after close of s+5",
            ),
            (
                "Decision",
                "At t=s+5 choose highest known PCA DeltaQIC5_vs_Raw",
            ),
            (
                "Stock_Selection",
                "Use current t Signal; Q5 residual losers and Q1 residual winners",
            ),
            (
                "Execution",
                "Enter at next trading-day close t+1",
            ),
            (
                "Holding_Days",
                HOLD_DAYS,
            ),
            (
                "Daily_New_Cohort_Gross",
                COHORT_GROSS_FRACTION,
            ),
            (
                "LONG_SHORT",
                "Q5 Long 10% + Q1 Short 10%",
            ),
            (
                "LONG_ONLY",
                "Q5 Long 20%",
            ),
            (
                "SHORT_ONLY",
                "Q1 Short 20%",
            ),
            (
                "Buy_Cost_Bps",
                args.buy_cost_bps,
            ),
            (
                "Sell_Cost_Bps",
                args.sell_cost_bps,
            ),
            (
                "Cost_Method",
                "Stock-level daily netting before transaction costs",
            ),
            (
                "Require_Positive_DeltaQIC5",
                "NO"
                if args.allow_nonpositive_delta_qic5
                else "YES",
            ),
            (
                "Short_Borrow_Fee",
                "NOT INCLUDED",
            ),
            (
                "Common_Selections",
                len(
                    specs
                ),
            ),
            (
                "Skipped_NonPositive_DeltaQIC5",
                skipped_nonpositive_delta_qic5,
            ),
            (
                "Skipped_No_DeltaQIC5",
                skipped_no_delta_qic5,
            ),
            (
                "Skipped_Signal_Formation",
                skipped_signal,
            ),
            (
                "Skipped_Entry_Tradability",
                skipped_entry,
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
                105,
            ],
        )

        sheet_specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # COMPARISON SUMMARY.
        p = td / "comparison_summary.xml"

        summary_header = [
            "Strategy",
            "Start",
            "End",
            "CompletedCohorts",
            "TotalReturnNet",
            "CAGRNet",
            "AnnualizedVol",
            "SharpeNet",
            "MaxDrawdownNet",
            "PositiveDailyRatio",
            "TotalReturnNoCost",
            "CAGRNoCost",
            "MaxDrawdownNoCost",
            "AverageGrossExposure",
            "AverageNetExposure",
            "AverageDailyTurnover",
            "TurnoverReductionFromNetting",
            "TotalTradingCost",
            "Mean5DCohortGrossReturn",
            "Positive5DCohortRatio",
            "ExitMissingRawPriceCount",
        ]

        summary_rows = []

        for mode in STRATEGY_MODES:
            m = metrics[
                mode
            ]

            summary_rows.append(
                tuple(
                    m[
                        key
                    ]
                    for key in summary_header
                )
            )

        write_table_sheet(
            p,
            summary_header,
            summary_rows,
            [
                20,
                14,
                14,
                18,
                17,
                15,
                18,
                15,
                18,
                19,
                19,
                16,
                21,
                21,
                20,
                22,
                27,
                18,
                24,
                22,
                23,
            ],
        )

        sheet_specs.append(
            (
                "COMPARISON_SUMMARY",
                p,
            )
        )

        # YEARLY.
        p = td / "yearly.xml"

        write_table_sheet(
            p,
            [
                "Year",
                "Strategy",
                "NetReturn",
                "NoCostReturn",
                "AverageTurnover",
            ],
            sorted(
                yearly_rows,
                key=lambda r: (
                    r[0],
                    r[1],
                ),
            ),
            [
                12,
                20,
                16,
                18,
                20,
            ],
        )

        sheet_specs.append(
            (
                "YEARLY_COMPARISON",
                p,
            )
        )

        # DAILY COMPARISON.
        p = td / "daily_comparison.xml"

        arrays = {
            mode: result_arrays(
                results[
                    mode
                ]
            )
            for mode in STRATEGY_MODES
        }

        arrays0 = {
            mode: result_arrays(
                no_cost_results[
                    mode
                ]
            )
            for mode in STRATEGY_MODES
        }

        base_dates = arrays[
            "LONG_SHORT"
        ][
            "dates"
        ]

        comparison_rows = []

        for i, d in enumerate(
            base_dates
        ):
            comparison_rows.append(
                (
                    yyyymmdd_text(
                        int(
                            d
                        )
                    ),
                    arrays[
                        "LONG_SHORT"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays[
                        "LONG_ONLY"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays[
                        "SHORT_ONLY"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays0[
                        "LONG_SHORT"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays0[
                        "LONG_ONLY"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays0[
                        "SHORT_ONLY"
                    ][
                        "equity"
                    ][
                        i
                    ],
                    arrays[
                        "LONG_SHORT"
                    ][
                        "returns"
                    ][
                        i
                    ],
                    arrays[
                        "LONG_ONLY"
                    ][
                        "returns"
                    ][
                        i
                    ],
                    arrays[
                        "SHORT_ONLY"
                    ][
                        "returns"
                    ][
                        i
                    ],
                )
            )

        write_table_sheet(
            p,
            [
                "Date",
                "Equity_LONG_SHORT",
                "Equity_LONG_ONLY",
                "Equity_SHORT_ONLY",
                "NoCostEquity_LONG_SHORT",
                "NoCostEquity_LONG_ONLY",
                "NoCostEquity_SHORT_ONLY",
                "Return_LONG_SHORT",
                "Return_LONG_ONLY",
                "Return_SHORT_ONLY",
            ],
            comparison_rows,
            [
                14,
                20,
                20,
                20,
                25,
                25,
                25,
                20,
                20,
                20,
            ],
        )

        sheet_specs.append(
            (
                "DAILY_COMPARISON",
                p,
            )
        )

        # Detailed DAILY sheets.
        daily_header = [
            "Date",
            "Strategy",
            "Equity",
            "DailyNetReturn",
            "GrossPnLAmount",
            "TradingCostAmount",
            "ActiveCohorts",
            "GrossExposure",
            "NetExposure",
            "Turnover_NETTED",
            "Turnover_UNNETTED",
            "BuyNotional_NETTED",
            "SellNotional_NETTED",
            "CostSavedByNetting",
            "NewCohortEntered",
            "SelectedCase",
            "LaggedDeltaQIC5_FWD5",
            "DeltaQICSourceDate",
            "DecisionDate",
        ]

        mode_sheet_names = {
            "LONG_SHORT": "DAILY_LONG_SHORT",
            "LONG_ONLY": "DAILY_LONG_ONLY",
            "SHORT_ONLY": "DAILY_SHORT_ONLY",
        }

        for mode in STRATEGY_MODES:
            p = td / f"daily_{mode.lower()}.xml"

            rows = (
                (
                    yyyymmdd_text(
                        int(
                            r[0]
                        )
                    ),
                    *r[
                        1:
                    ],
                )
                for r in results[
                    mode
                ][
                    "daily_rows"
                ]
            )

            write_table_sheet(
                p,
                daily_header,
                rows,
                [
                    14,
                    18,
                    15,
                    18,
                    18,
                    19,
                    15,
                    16,
                    16,
                    18,
                    20,
                    20,
                    20,
                    20,
                    18,
                    26,
                    18,
                    16,
                    16,
                ],
            )

            sheet_specs.append(
                (
                    mode_sheet_names[
                        mode
                    ],
                    p,
                )
            )

        # Common cohort selection.
        p = td / "selection.xml"

        selection_rows = [
            (
                yyyymmdd_text(
                    int(
                        dates[
                            spec.delta_qic_source_idx
                        ]
                    )
                ),
                yyyymmdd_text(
                    int(
                        dates[
                            spec.decision_idx
                        ]
                    )
                ),
                yyyymmdd_text(
                    int(
                        dates[
                            spec.entry_idx
                        ]
                    )
                ),
                yyyymmdd_text(
                    int(
                        dates[
                            spec.exit_idx
                        ]
                    )
                ),
                spec.case_id,
                float(
                    spec.lagged_delta_qic5
                ),
                int(
                    len(
                        spec.long_idx
                    )
                ),
                int(
                    len(
                        spec.short_idx
                    )
                ),
                int(
                    spec.entry_missing_filtered
                ),
                ",".join(
                    tickers[i]
                    for i in spec.long_idx
                ),
                ",".join(
                    tickers[i]
                    for i in spec.short_idx
                ),
            )
            for spec in specs
        ]

        write_table_sheet(
            p,
            [
                "DeltaQICSourceDate",
                "DecisionDate",
                "EntryDate",
                "ExitDate",
                "CaseID",
                "LaggedDeltaQIC5_FWD5",
                "Q5_LongN",
                "Q1_ShortN",
                "EntryMissingFiltered",
                "Q5_Tickers",
                "Q1_Tickers",
            ],
            selection_rows,
            [
                16,
                16,
                16,
                16,
                28,
                18,
                13,
                13,
                21,
                55,
                55,
            ],
        )

        sheet_specs.append(
            (
                "COHORT_SELECTION",
                p,
            )
        )

        # Cohort performance by mode.
        cohort_header = [
            "DeltaQICSourceDate",
            "DecisionDate",
            "EntryDate",
            "ExitDate",
            "Strategy",
            "CaseID",
            "LaggedDeltaQIC5_FWD5",
            "LongN",
            "ShortN",
            "EntryEquity",
            "InitialGrossAllocation",
            "GrossPnL",
            "GrossReturnOnGross",
            "EntryMissingFiltered",
            "ExitMissingRawCount",
            "LongTickers",
            "ShortTickers",
        ]

        cohort_sheet_names = {
            "LONG_SHORT": "COHORTS_LONG_SHORT",
            "LONG_ONLY": "COHORTS_LONG_ONLY",
            "SHORT_ONLY": "COHORTS_SHORT_ONLY",
        }

        for mode in STRATEGY_MODES:
            p = td / f"cohorts_{mode.lower()}.xml"

            rows = (
                (
                    yyyymmdd_text(
                        int(
                            r[0]
                        )
                    ),
                    yyyymmdd_text(
                        int(
                            r[1]
                        )
                    ),
                    yyyymmdd_text(
                        int(
                            r[2]
                        )
                    ),
                    yyyymmdd_text(
                        int(
                            r[3]
                        )
                    ),
                    *r[
                        4:
                    ],
                )
                for r in results[
                    mode
                ][
                    "cohort_rows"
                ]
            )

            write_table_sheet(
                p,
                cohort_header,
                rows,
                [
                    16,
                    16,
                    16,
                    16,
                    18,
                    28,
                    18,
                    11,
                    11,
                    16,
                    23,
                    16,
                    21,
                    22,
                    22,
                    55,
                    55,
                ],
            )

            sheet_specs.append(
                (
                    cohort_sheet_names[
                        mode
                    ],
                    p,
                )
            )

        # Case usage is common across all three.
        p = td / "usage.xml"

        usage = Counter(
            spec.case_id
            for spec in specs
        )

        total_specs = max(
            1,
            len(
                specs
            ),
        )

        usage_rows = [
            (
                case_id,
                count,
                count
                / total_specs,
            )
            for case_id, count in usage.most_common()
        ]

        write_table_sheet(
            p,
            [
                "CaseID",
                "SelectedCohorts",
                "UsageRatio",
            ],
            usage_rows,
            [
                28,
                20,
                16,
            ],
        )

        sheet_specs.append(
            (
                "CASE_USAGE",
                p,
            )
        )

        package_xlsx(
            output_path,
            sheet_specs,
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
        "먼저 확인할 시트:"
    )
    print(
        "  1) COMPARISON_SUMMARY"
    )
    print(
        "  2) DAILY_COMPARISON"
    )
    print(
        "  3) COHORT_SELECTION"
    )
    print(
        "  4) YEARLY_COMPARISON"
    )
    print()
    print(
        "그래프:"
    )

    for p in graph_paths:
        print(
            " ",
            p,
        )


if __name__ == "__main__":
    main()
