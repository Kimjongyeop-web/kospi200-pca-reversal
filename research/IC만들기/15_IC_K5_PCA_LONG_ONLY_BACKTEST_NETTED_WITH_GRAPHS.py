# -*- coding: utf-8 -*-
"""
15_IC_K5_PCA_LONG_ONLY_BACKTEST_NETTED_WITH_GRAPHS.py
=================================

IC(K=5)를 실제 매매 의사결정에 직접 사용하는 PCA 단기반전 LONG-ONLY 백테스트.

핵심 타임라인
-------------
signal date = s

IC_s,5 는:
    s+1 ~ s+5 미래수익률을 포함하므로
    s+5 종가가 끝난 뒤에야 알 수 있다.

따라서 현재 decision date = d = s+5 종가 후:
    1) 이미 확정된 IC_s,5 중 PCA Case 최고값을 선택
    2) 과거 s일 Signal을 사는 것이 아니라,
       "현재 d일"의 선택된 PCA Case Signal로 종목을 다시 정렬
    3) 현재 d일 Signal 상위 20% = Q5 residual loser만 Long
    4) Q1 Short는 사용하지 않음
    5) 다음 거래일 d+1 종가에 신규 cohort 진입
    6) 5거래일 보유 후 종가 청산

즉:
    IC_{t-5,5}
      -> 현재 t의 PCA Case 선택
      -> 현재 t의 Q5 Long 종목 선택
      -> t+1 종가 진입
      -> 5거래일 보유

이 구조는 T일의 미래 5일을 사용한 IC를 T일에 이용하는 look-ahead를 막는다.

포트폴리오 구조
---------------
- 후보 Case: PCA Case만 사용 (Raw 제외)
- IC horizon: FWD5만 사용
- 최고 PCA IC가 0 이하이면 신규 cohort는 Cash (기본값)
- 매일 총 Equity의 20%를 신규 Q5 Long cohort에 배정
    Long = 20%
    Short = 0%
- 신규 cohort는 정확히 5 close-to-close return intervals 보유
- 정상상태에서는 5개의 20% Long cohort가 겹쳐 약 100% Long exposure
- 각 cohort 내부 Q5 종목은 equal-weight
- T 종가 정보로 Signal을 만들기 때문에 T+1 종가 진입
- 거래비용: 매수 5bp / 매도 20bp
- 중요: 거래비용은 cohort별 전량청산/전량진입 금액에 각각 부과하지 않는다.
  매일 만기 cohort 제거 + 신규 cohort 추가를 모두 반영한 뒤,
  종목별 signed share 변화량을 합산(netting)해서 실제 순매수/순매도 금액에만 부과한다.

예:
    기존 삼성전자 +0.30%, 새 목표 +0.35% -> +0.05%만 매수
    기존 삼성전자 +0.30%, 새 목표 +0.20% -> -0.10%만 매도
    기존 삼성전자 +0.30%, 새 목표 -0.20% -> -0.50% 매도
    만기 Long과 신규 Long 수량이 동일 -> 거래 0, 비용 0

예:
    Day 1 cohort: 20% gross
    Day 2 cohort: +20%
    ...
    Day 5: 총 약 100% gross
    Day 6: Day1 cohort 청산 + 새 Day6 cohort 20% 진입

거래정지/가격결측
-----------------
- 신규 진입일 수정주가가 없는 종목은 체결 불가로 제외하고 해당 leg 내부 재가중
- 보유 중 가격이 일시 결측이면 마지막 수정주가로 mark-to-market(그날 수익 0)
- 가격이 다시 생기면 마지막 관측가격 대비 gap이 그날 반영
- 예정 청산일 원시 수정주가가 없으면 마지막 mark 가격으로 청산 가정하고
  COHORTS의 ExitMissingRawCount에 기록
  (5일 horizon + Common252 적격종목이라 보통 드문 케이스지만 반드시 QC 가능)

필요 입력 Excel 6개
-------------------
1) STATE_IC_TRAIN_LIBRARY.xlsx
   - CASE_DEFINITION
   - CASE_DAILY_IC
   - IC_FWD5 사용

2) 00_COMMON252_BASE.xlsx
   - COMMON252_ELIGIBLE

3) 02_PCA_W60_CUMULATIVE.xlsx
4) 04_PCA_W120_CUMULATIVE.xlsx
5) 06_PCA_W252_CUMULATIVE.xlsx
   - 현재 decision date의 PCA reversal Signal 산출

6) KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx
   - 수정주가
   - 실제 다음날 종가 진입 / 일별 mark-to-market / 5일 후 청산

필요하지 않은 파일
------------------
- FUTURE_LOG_RETURNS_5_10_20_60.xlsx
  (실제 백테스트 PnL은 수정주가에서 직접 계산)
- PCA DAILY 파일
- Spread 결과파일

출력
----
IC_K5_PCA_PORTFOLIO_BACKTEST/
    IC_K5_PCA_PORTFOLIO_BACKTEST.xlsx

시트
----
PARAMETERS
SUMMARY
DAILY
COHORTS
CASE_USAGE

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

IC_HORIZON = 5
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


def read_ic_fwd5_matrix(
    path: Path,
    ref_dates: np.ndarray,
    cases: list[CaseDef],
) -> np.ndarray:
    it = iter_sheet_rows(
        path,
        "CASE_DAILY_IC",
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

    required = [
        "Date",
        "CaseID",
        "IC_FWD5",
    ]

    missing = [
        x
        for x in required
        if x not in headers
    ]

    if missing:
        raise KeyError(
            f"CASE_DAILY_IC 필수열 누락: {missing}"
        )

    date_col = headers.index(
        "Date"
    )

    case_col = headers.index(
        "CaseID"
    )

    ic_col = headers.index(
        "IC_FWD5"
    )

    date_map = {
        int(d): i
        for i, d in enumerate(
            ref_dates
        )
    }

    case_map = {
        c.case_id: j
        for j, c in enumerate(
            cases
        )
    }

    out = np.full(
        (
            len(ref_dates),
            len(cases),
        ),
        np.nan,
        dtype=np.float64,
    )

    rows = 0

    for cells in it:
        case_id = str(
            cells.get(
                case_col,
                "",
            )
            or ""
        ).strip()

        cj = case_map.get(
            case_id
        )

        if cj is None:
            continue

        raw_date = cells.get(
            date_col
        )

        raw_ic = cells.get(
            ic_col
        )

        if (
            raw_date in (None, "")
            or raw_ic in (None, "")
        ):
            continue

        try:
            d = parse_date_value(
                raw_date
            )

            ic = float(
                raw_ic
            )
        except Exception:
            continue

        di = date_map.get(
            d
        )

        if (
            di is not None
            and np.isfinite(ic)
        ):
            out[
                di,
                cj,
            ] = ic

        rows += 1

    print(
        f"      CASE_DAILY_IC PCA rows read={rows:,}",
        flush=True,
    )

    return out


# =============================================================================
# SIGNAL SELECTION
# =============================================================================

@dataclass
class DecisionRecord:
    ic_source_idx: int
    decision_idx: int
    entry_idx: int
    exit_idx: int
    case_idx: int
    case_id: str
    lagged_ic: float


@dataclass
class CohortSpec:
    ic_source_idx: int
    decision_idx: int
    entry_idx: int
    exit_idx: int
    case_id: str
    lagged_ic: float
    long_idx: np.ndarray
    short_idx: np.ndarray
    entry_missing_filtered: int


def build_decision_records(
    dates: np.ndarray,
    ic_matrix: np.ndarray,
    cases: list[CaseDef],
    start_date: int | None,
    end_date: int | None,
    require_positive_ic: bool,
):
    records = []
    skipped_nonpositive = 0
    skipped_no_ic = 0

    T = len(dates)

    # entry = decision + 1
    # ic_source = decision - 5
    # exit = entry + 5
    for decision_idx in range(
        IC_HORIZON,
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

        ic_source_idx = (
            decision_idx
            - IC_HORIZON
        )

        row = ic_matrix[
            ic_source_idx
        ]

        finite = np.isfinite(
            row
        )

        if not finite.any():
            skipped_no_ic += 1
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

        best_ic = float(
            row[
                case_idx
            ]
        )

        if (
            require_positive_ic
            and best_ic <= 0
        ):
            skipped_nonpositive += 1
            continue

        records.append(
            DecisionRecord(
                ic_source_idx=ic_source_idx,
                decision_idx=decision_idx,
                entry_idx=entry_idx,
                exit_idx=exit_idx,
                case_idx=case_idx,
                case_id=cases[
                    case_idx
                ].case_id,
                lagged_ic=best_ic,
            )
        )

    return (
        records,
        skipped_nonpositive,
        skipped_no_ic,
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

            long_idx, _unused_q1 = legs

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

            missing_count = int(
                (~long_ok).sum()
            )

            long_idx = long_idx[
                long_ok
            ]

            # LONG-ONLY: Q1 short leg is intentionally unused.
            short_idx = np.asarray(
                [],
                dtype=int,
            )

            if len(long_idx) < MIN_LEG_STOCKS:
                skipped_entry += 1
                continue

            specs.append(
                CohortSpec(
                    ic_source_idx=rec.ic_source_idx,
                    decision_idx=rec.decision_idx,
                    entry_idx=rec.entry_idx,
                    exit_idx=rec.exit_idx,
                    case_id=rec.case_id,
                    lagged_ic=rec.lagged_ic,
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


def simulate_portfolio(
    specs: list[CohortSpec],
    dates: np.ndarray,
    tickers: list[str],
    raw_prices: np.ndarray,
    mark_prices: np.ndarray,
    buy_cost_rate: float,
    sell_cost_rate: float,
):
    """
    LONG-ONLY + stock-level daily NETTING.

    Every entry day:
      - invest 20% of current pre-cost equity in Q5 only
      - equal-weight among Q5 names
      - hold exactly 5 close-to-close intervals

    Normal steady state:
      - five overlapping 20% cohorts
      - about 100% long exposure
      - remaining capital is cash whenever fewer than five cohorts are active

    Cost rule after stock-level netting:
      net share change > 0 -> BUY cost
      net share change < 0 -> SELL cost

    Thus if an expiring cohort sells Samsung while the new cohort buys Samsung
    on the same close, only the difference in shares is actually traded.
    """

    entries_by_day = defaultdict(list)
    for spec in specs:
        entries_by_day[spec.entry_idx].append(spec)

    if not specs:
        raise RuntimeError("진입 가능한 long-only cohort가 0개입니다.")

    first_entry = min(x.entry_idx for x in specs)
    last_exit = max(x.exit_idx for x in specs)

    equity = 1.0
    active: list[ActiveCohort] = []

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

    for t in range(first_entry, last_exit + 1):
        equity_start = equity
        mtm_pnl = 0.0

        # 1) Overnight PnL for cohorts already held.
        if t > 0:
            for c in active:
                prev_px = mark_prices[
                    t - 1,
                    c.spec.long_idx
                ]
                now_px = mark_prices[
                    t,
                    c.spec.long_idx
                ]

                pnl = float(
                    np.nansum(
                        c.long_shares
                        * (now_px - prev_px)
                    )
                )

                c.gross_pnl += pnl
                mtm_pnl += pnl

            equity += mtm_pnl

        equity_before_trade = equity
        mark_row = mark_prices[t]

        # Positive quantity = buy, negative = sell.
        net_trade_qty = np.zeros(
            N,
            dtype=np.float64,
        )

        unnetted_buy_notional = 0.0
        unnetted_sell_notional = 0.0

        # 2) Expiring long cohorts -> sell shares.
        still_active = []

        for c in active:
            if c.spec.exit_idx != t:
                still_active.append(c)
                continue

            net_trade_qty[
                c.spec.long_idx
            ] -= c.long_shares

            exit_notional = float(
                np.nansum(
                    np.abs(
                        c.long_shares
                        * mark_row[
                            c.spec.long_idx
                        ]
                    )
                )
            )

            unnetted_sell_notional += exit_notional

            raw_exit = raw_prices[t]

            missing_exit = int(
                (
                    ~np.isfinite(
                        raw_exit[
                            c.spec.long_idx
                        ]
                    )
                ).sum()
            )

            total_exit_missing_raw += missing_exit

            gross_return_on_gross = (
                c.gross_pnl
                / c.initial_gross
                if c.initial_gross > EPS
                else np.nan
            )

            cohort_rows.append(
                (
                    int(
                        dates[
                            c.spec.ic_source_idx
                        ]
                    ),
                    int(
                        dates[
                            c.spec.decision_idx
                        ]
                    ),
                    int(
                        dates[
                            c.spec.entry_idx
                        ]
                    ),
                    int(
                        dates[
                            c.spec.exit_idx
                        ]
                    ),
                    c.spec.case_id,
                    float(
                        c.spec.lagged_ic
                    ),
                    int(
                        len(
                            c.spec.long_idx
                        )
                    ),
                    float(
                        c.entry_equity
                    ),
                    float(
                        c.initial_gross
                    ),
                    float(
                        c.gross_pnl
                    ),
                    (
                        float(
                            gross_return_on_gross
                        )
                        if np.isfinite(
                            gross_return_on_gross
                        )
                        else None
                    ),
                    int(
                        c.spec.entry_missing_filtered
                    ),
                    missing_exit,
                    ",".join(
                        tickers[i]
                        for i in c.spec.long_idx
                    ),
                )
            )

        active = still_active

        # 3) New 20% Q5 long cohort at today's close.
        entered_case_list = []
        entered_ic_list = []
        entered_ic_source_list = []
        entered_decision_list = []

        todays_specs = entries_by_day.get(
            t,
            [],
        )

        for spec in todays_specs:
            if equity_before_trade <= EPS:
                continue

            long_target = (
                COHORT_GROSS_FRACTION
                * equity_before_trade
            )

            entry_row = raw_prices[t]
            long_price = entry_row[
                spec.long_idx
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
                    spec.long_idx
                )
            )

            long_shares = (
                long_each
                / long_price
            )

            # Long open = BUY.
            net_trade_qty[
                spec.long_idx
            ] += long_shares

            unnetted_buy_notional += long_target

            active.append(
                ActiveCohort(
                    spec=spec,
                    initial_gross=long_target,
                    long_shares=long_shares,
                    short_shares=np.asarray(
                        [],
                        dtype=np.float64,
                    ),
                    gross_pnl=0.0,
                    entry_equity=equity_before_trade,
                )
            )

            entered_case_list.append(
                spec.case_id
            )
            entered_ic_list.append(
                float(
                    spec.lagged_ic
                )
            )
            entered_ic_source_list.append(
                yyyymmdd_text(
                    int(
                        dates[
                            spec.ic_source_idx
                        ]
                    )
                )
            )
            entered_decision_list.append(
                yyyymmdd_text(
                    int(
                        dates[
                            spec.decision_idx
                        ]
                    )
                )
            )

        # 4) Net all stock-level exit and entry orders.
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
        total_netted_trade_notional += netted_trade_notional
        total_unnetted_trade_notional += unnetted_trade_notional
        total_trading_cost += trading_cost_today
        total_unnetted_cost += unnetted_cost_today
        total_netting_cost_savings += netting_cost_savings

        # 5) End-of-day long exposure.
        gross_exposure = 0.0
        net_exposure = 0.0

        for c in active:
            long_val = (
                c.long_shares
                * mark_row[
                    c.spec.long_idx
                ]
            )

            position_value = float(
                np.nansum(
                    long_val
                )
            )

            gross_exposure += abs(
                position_value
            )
            net_exposure += position_value

        daily_return = (
            equity
            / equity_start
            - 1.0
            if equity_start > EPS
            else np.nan
        )

        turnover_ratio = (
            netted_trade_notional
            / equity_start
            if equity_start > EPS
            else np.nan
        )

        unnetted_turnover_ratio = (
            unnetted_trade_notional
            / equity_start
            if equity_start > EPS
            else np.nan
        )

        entered_flag = int(
            len(
                entered_case_list
            ) > 0
        )

        cash_weight = (
            1.0
            - net_exposure / equity
            if equity > EPS
            else np.nan
        )

        daily_rows.append(
            (
                int(
                    dates[t]
                ),
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
                        cash_weight
                    )
                    if np.isfinite(
                        cash_weight
                    )
                    else None
                ),
                (
                    float(
                        turnover_ratio
                    )
                    if np.isfinite(
                        turnover_ratio
                    )
                    else None
                ),
                (
                    float(
                        unnetted_turnover_ratio
                    )
                    if np.isfinite(
                        unnetted_turnover_ratio
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
                    netted_trade_notional
                ),
                float(
                    unnetted_trade_notional
                ),
                float(
                    netting_cost_savings
                ),
                entered_flag,
                ";".join(
                    entered_case_list
                ),
                (
                    float(
                        entered_ic_list[0]
                    )
                    if len(
                        entered_ic_list
                    ) == 1
                    else (
                        ";".join(
                            f"{x:.12g}"
                            for x in entered_ic_list
                        )
                        if entered_ic_list
                        else None
                    )
                ),
                ";".join(
                    entered_ic_source_list
                ),
                ";".join(
                    entered_decision_list
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

    return (
        daily_rows,
        cohort_rows,
        totals,
    )


# =============================================================================
# SUMMARY
# =============================================================================

def max_drawdown(
    equity: np.ndarray,
):
    peak = np.maximum.accumulate(
        equity
    )

    dd = (
        equity
        / peak
        - 1.0
    )

    return float(
        np.min(
            dd
        )
    )


def max_drawdown_series(
    equity: np.ndarray,
) -> np.ndarray:
    peak = np.maximum.accumulate(
        equity
    )

    return (
        equity / peak - 1.0
    )


def performance_metrics_from_daily(
    daily_rows,
):
    dates = [
        r[0]
        for r in daily_rows
    ]

    equity = np.asarray(
        [
            r[1]
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    ret = np.asarray(
        [
            np.nan
            if r[2] is None
            else float(
                r[2]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    valid_ret = ret[
        np.isfinite(
            ret
        )
    ]

    n = len(
        valid_ret
    )

    total_return = (
        float(
            equity[-1] - 1.0
        )
        if len(
            equity
        )
        else np.nan
    )

    annualized_return = (
        float(
            equity[-1] ** (
                252.0
                / max(
                    1,
                    n,
                )
            )
            - 1.0
        )
        if (
            len(
                equity
            )
            and equity[-1] > 0
        )
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

    mdd = (
        max_drawdown(
            equity
        )
        if len(
            equity
        )
        else np.nan
    )

    positive_daily = (
        float(
            np.mean(
                valid_ret > 0
            )
        )
        if n
        else np.nan
    )

    return {
        "dates": dates,
        "equity": equity,
        "ret": ret,
        "total_return": total_return,
        "annualized_return": annualized_return,
        "annualized_vol": annualized_vol,
        "sharpe": sharpe,
        "mdd": mdd,
        "positive_daily": positive_daily,
    }


def build_summary(
    daily_rows,
    cohort_rows,
    totals,
    buy_cost_bps: float,
    sell_cost_bps: float,
    skipped_nonpositive: int,
    skipped_no_ic: int,
    skipped_signal: int,
    skipped_entry: int,
    no_cost_daily_rows=None,
):
    perf = performance_metrics_from_daily(
        daily_rows
    )

    dates = perf["dates"]

    gross_exposure = np.asarray(
        [
            np.nan
            if r[6] is None
            else float(
                r[6]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    net_exposure = np.asarray(
        [
            np.nan
            if r[7] is None
            else float(
                r[7]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    cash_weight = np.asarray(
        [
            np.nan
            if r[8] is None
            else float(
                r[8]
            )
            for r in daily_rows
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
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    unnetted_turnover = np.asarray(
        [
            np.nan
            if r[10] is None
            else float(
                r[10]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    total_netted = float(
        totals[
            "TotalNettedTurnoverNotional"
        ]
    )

    total_unnetted = float(
        totals[
            "TotalUnnettedTurnoverNotional"
        ]
    )

    turnover_reduction = (
        1.0
        - total_netted
        / total_unnetted
        if total_unnetted > EPS
        else np.nan
    )

    total_cost = float(
        totals[
            "TotalTradingCost"
        ]
    )

    total_unnetted_cost = float(
        totals[
            "TotalUnnettedCost"
        ]
    )

    cost_reduction = (
        1.0
        - total_cost
        / total_unnetted_cost
        if total_unnetted_cost > EPS
        else np.nan
    )

    cohort_gross_returns = np.asarray(
        [
            np.nan
            if r[10] is None
            else float(
                r[10]
            )
            for r in cohort_rows
        ],
        dtype=np.float64,
    )

    cohort_gross_returns = cohort_gross_returns[
        np.isfinite(
            cohort_gross_returns
        )
    ]

    summary = [
        (
            "Strategy",
            "IC(K=5) PCA dynamic Q5 LONG ONLY",
        ),
        (
            "Backtest_Start",
            yyyymmdd_text(
                int(
                    dates[0]
                )
            ),
        ),
        (
            "Backtest_End",
            yyyymmdd_text(
                int(
                    dates[-1]
                )
            ),
        ),
        (
            "IC_Horizon",
            5,
        ),
        (
            "Holding_Days",
            HOLD_DAYS,
        ),
        (
            "Daily_New_Long_Cohort",
            COHORT_GROSS_FRACTION,
        ),
        (
            "Target_Long_Exposure_When_Fully_Invested",
            COHORT_GROSS_FRACTION
            * HOLD_DAYS,
        ),
        (
            "Short_Exposure",
            0.0,
        ),
        (
            "Buy_Cost_Bps",
            buy_cost_bps,
        ),
        (
            "Sell_Cost_Bps",
            sell_cost_bps,
        ),
        (
            "Long_RoundTrip_Cost_Bps_if_no_netting",
            buy_cost_bps
            + sell_cost_bps,
        ),
        (
            "Transaction_Cost_Method",
            "STOCK_LEVEL_DAILY_NETTING",
        ),
        (
            "Completed_Cohorts",
            len(
                cohort_rows
            ),
        ),
        (
            "Total_Return_NET",
            perf[
                "total_return"
            ],
        ),
        (
            "Annualized_Return_NET",
            perf[
                "annualized_return"
            ],
        ),
        (
            "Annualized_Volatility_NET",
            perf[
                "annualized_vol"
            ],
        ),
        (
            "Sharpe_RF0_NET",
            perf[
                "sharpe"
            ],
        ),
        (
            "Max_Drawdown_NET",
            perf[
                "mdd"
            ],
        ),
        (
            "Positive_Daily_Return_Ratio_NET",
            perf[
                "positive_daily"
            ],
        ),
        (
            "Average_Gross_Exposure",
            float(
                np.nanmean(
                    gross_exposure
                )
            ),
        ),
        (
            "Average_Net_Long_Exposure",
            float(
                np.nanmean(
                    net_exposure
                )
            ),
        ),
        (
            "Average_Cash_Weight",
            float(
                np.nanmean(
                    cash_weight
                )
            ),
        ),
        (
            "Average_Daily_Turnover_NETTED",
            float(
                np.nanmean(
                    turnover
                )
            ),
        ),
        (
            "Average_Daily_Turnover_UNNETTED",
            float(
                np.nanmean(
                    unnetted_turnover
                )
            ),
        ),
        (
            "Total_Netted_Trade_Notional",
            total_netted,
        ),
        (
            "Total_Unnetted_Trade_Notional",
            total_unnetted,
        ),
        (
            "Turnover_Reduction_From_Netting",
            float(
                turnover_reduction
            )
            if np.isfinite(
                turnover_reduction
            )
            else None,
        ),
        (
            "Total_Trading_Cost_NETTED",
            total_cost,
        ),
        (
            "Total_Trading_Cost_if_UNNETTED",
            total_unnetted_cost,
        ),
        (
            "Total_Cost_Saved_By_Netting",
            float(
                totals[
                    "TotalNettingCostSavings"
                ]
            ),
        ),
        (
            "Cost_Reduction_From_Netting",
            float(
                cost_reduction
            )
            if np.isfinite(
                cost_reduction
            )
            else None,
        ),
        (
            "Total_Buy_Notional",
            float(
                totals[
                    "TotalBuyNotional"
                ]
            ),
        ),
        (
            "Total_Sell_Notional",
            float(
                totals[
                    "TotalSellNotional"
                ]
            ),
        ),
        (
            "Mean_5D_Cohort_Gross_Return",
            float(
                np.mean(
                    cohort_gross_returns
                )
            )
            if len(
                cohort_gross_returns
            )
            else None,
        ),
        (
            "Median_5D_Cohort_Gross_Return",
            float(
                np.median(
                    cohort_gross_returns
                )
            )
            if len(
                cohort_gross_returns
            )
            else None,
        ),
        (
            "Positive_5D_Cohort_Gross_Ratio",
            float(
                np.mean(
                    cohort_gross_returns > 0
                )
            )
            if len(
                cohort_gross_returns
            )
            else None,
        ),
        (
            "Exit_Missing_Raw_Price_Count",
            int(
                totals[
                    "ExitMissingRawCount"
                ]
            ),
        ),
        (
            "Skipped_NonPositive_IC",
            skipped_nonpositive,
        ),
        (
            "Skipped_No_IC",
            skipped_no_ic,
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

    if no_cost_daily_rows is not None:
        p0 = performance_metrics_from_daily(
            no_cost_daily_rows
        )

        summary.extend(
            [
                (
                    "Total_Return_NO_COST",
                    p0[
                        "total_return"
                    ],
                ),
                (
                    "Annualized_Return_NO_COST",
                    p0[
                        "annualized_return"
                    ],
                ),
                (
                    "Sharpe_RF0_NO_COST",
                    p0[
                        "sharpe"
                    ],
                ),
                (
                    "Max_Drawdown_NO_COST",
                    p0[
                        "mdd"
                    ],
                ),
            ]
        )

    return summary


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
# PERFORMANCE CHARTS
# =============================================================================

def save_performance_charts(
    daily_rows,
    no_cost_daily_rows,
    cohort_rows,
    output_dir: Path,
):
    """
    Save daily performance PNGs.

    No explicit colors/styles are forced; matplotlib defaults are used.
    """
    chart_dir = (
        output_dir
        / "graphs"
    )

    chart_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    dates = [
        datetime.strptime(
            yyyymmdd_text(
                int(
                    r[0]
                )
            ),
            "%Y-%m-%d",
        )
        for r in daily_rows
    ]

    equity = np.asarray(
        [
            float(
                r[1]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    daily_ret = np.asarray(
        [
            np.nan
            if r[2] is None
            else float(
                r[2]
            )
            for r in daily_rows
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
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    unnetted_turnover = np.asarray(
        [
            np.nan
            if r[10] is None
            else float(
                r[10]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    gross_exposure = np.asarray(
        [
            np.nan
            if r[6] is None
            else float(
                r[6]
            )
            for r in daily_rows
        ],
        dtype=np.float64,
    )

    no_cost_equity = np.asarray(
        [
            float(
                r[1]
            )
            for r in no_cost_daily_rows
        ],
        dtype=np.float64,
    )

    drawdown = max_drawdown_series(
        equity
    )

    def rolling_mean(
        x,
        window=20,
    ):
        x = np.asarray(
            x,
            dtype=np.float64,
        )

        out = np.full_like(
            x,
            np.nan,
        )

        for i in range(
            window - 1,
            len(
                x
            ),
        ):
            out[i] = np.nanmean(
                x[
                    i
                    - window
                    + 1:
                    i
                    + 1
                ]
            )

        return out

    saved = []

    # 1) Net Equity
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        equity,
        linewidth=1.0,
    )

    plt.axhline(
        1.0,
        linewidth=0.8,
    )

    plt.title(
        "IC(K=5) PCA Q5 long-only: daily net equity"
    )

    plt.xlabel(
        "Date"
    )

    plt.ylabel(
        "Equity (initial = 1)"
    )

    plt.grid(
        True,
        alpha=0.25,
    )

    plt.tight_layout()

    p = (
        chart_dir
        / "01_daily_net_equity.png"
    )

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

    # 2) Net vs no-cost counterfactual.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        no_cost_equity,
        linewidth=1.0,
        label="No transaction cost",
    )

    plt.plot(
        dates,
        equity,
        linewidth=1.0,
        label="Net: buy 5bp / sell 20bp",
    )

    plt.axhline(
        1.0,
        linewidth=0.8,
    )

    plt.title(
        "Q5 long-only: no-cost vs net equity"
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

    p = (
        chart_dir
        / "02_no_cost_vs_net_equity.png"
    )

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

    # 3) Drawdown.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        drawdown * 100.0,
        linewidth=1.0,
    )

    plt.axhline(
        0.0,
        linewidth=0.8,
    )

    plt.title(
        "Q5 long-only: daily drawdown"
    )

    plt.xlabel(
        "Date"
    )

    plt.ylabel(
        "Drawdown (%)"
    )

    plt.grid(
        True,
        alpha=0.25,
    )

    plt.tight_layout()

    p = (
        chart_dir
        / "03_daily_drawdown.png"
    )

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

    # 4) Daily net return.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        daily_ret * 100.0,
        linewidth=0.7,
    )

    plt.axhline(
        0.0,
        linewidth=0.8,
    )

    plt.title(
        "Q5 long-only: daily net return"
    )

    plt.xlabel(
        "Date"
    )

    plt.ylabel(
        "Daily return (%)"
    )

    plt.grid(
        True,
        alpha=0.25,
    )

    plt.tight_layout()

    p = (
        chart_dir
        / "04_daily_net_return.png"
    )

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

    # 5) Turnover: 20-day rolling mean.
    turn20 = rolling_mean(
        turnover,
        20,
    )

    unnet20 = rolling_mean(
        unnetted_turnover,
        20,
    )

    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        unnet20 * 100.0,
        linewidth=1.0,
        label="Unnetted turnover (20D avg)",
    )

    plt.plot(
        dates,
        turn20 * 100.0,
        linewidth=1.0,
        label="Netted turnover (20D avg)",
    )

    plt.title(
        "Q5 long-only: daily turnover and netting effect"
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

    p = (
        chart_dir
        / "05_turnover_netted_vs_unnetted.png"
    )

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

    # 6) Exposure.
    fig = plt.figure(
        figsize=(
            12,
            5,
        )
    )

    plt.plot(
        dates,
        gross_exposure * 100.0,
        linewidth=0.8,
    )

    plt.axhline(
        100.0,
        linewidth=0.8,
    )

    plt.title(
        "Q5 long-only: daily long exposure"
    )

    plt.xlabel(
        "Date"
    )

    plt.ylabel(
        "Long exposure (% of equity)"
    )

    plt.grid(
        True,
        alpha=0.25,
    )

    plt.tight_layout()

    p = (
        chart_dir
        / "06_daily_long_exposure.png"
    )

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

    # 7) Mean 5-day cohort gross return by lagged-IC quintile.
    ic_values = np.asarray(
        [
            float(
                r[5]
            )
            for r in cohort_rows
        ],
        dtype=np.float64,
    )

    cohort_returns = np.asarray(
        [
            np.nan
            if r[10] is None
            else float(
                r[10]
            )
            for r in cohort_rows
        ],
        dtype=np.float64,
    )

    valid = (
        np.isfinite(
            ic_values
        )
        & np.isfinite(
            cohort_returns
        )
    )

    if valid.sum() >= 25:
        iv = ic_values[
            valid
        ]
        cr = cohort_returns[
            valid
        ]

        ranks = np.argsort(
            np.argsort(
                iv,
                kind="mergesort",
            ),
            kind="mergesort",
        )

        q = np.minimum(
            4,
            (
                ranks
                * 5
                // len(
                    ranks
                )
            ),
        )

        qmeans = [
            float(
                np.mean(
                    cr[
                        q == k
                    ]
                )
            )
            * 100.0
            for k in range(
                5
            )
        ]

        labels = [
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "Q5",
        ]

        fig = plt.figure(
            figsize=(
                9,
                5,
            )
        )

        plt.bar(
            labels,
            qmeans,
        )

        plt.axhline(
            0.0,
            linewidth=0.8,
        )

        plt.title(
            "Lagged IC quintile vs subsequent Q5 long 5-day gross return"
        )

        plt.xlabel(
            "Lagged realized IC quintile (Q5 = highest)"
        )

        plt.ylabel(
            "Mean 5-day gross return (%)"
        )

        plt.tight_layout()

        p = (
            chart_dir
            / "07_lagged_ic_quintile_vs_long_return.png"
        )

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
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--data-dir",
        default=None,
        help="6개 입력 Excel이 있는 상위폴더. 생략 시 이 py 폴더 이하 재귀탐색.",
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
        "--allow-negative-ic",
        action="store_true",
        help="지정하면 최고 PCA IC가 0 이하인 날도 신규 reversal cohort 진입",
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

    require_positive_ic = (
        not args.allow_negative_ic
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
        else script_dir
        / "IC_K5_PCA_LONG_ONLY_BACKTEST"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "IC_K5_PCA_LONG_ONLY_BACKTEST.xlsx"
    )

    start_date = parse_cli_date(
        args.start
    )

    end_date = parse_cli_date(
        args.end
    )

    # ---------------------------------------------------------------------
    # Files
    # ---------------------------------------------------------------------
    file_map = {
        "ic": find_input_file(
            data_dir,
            "STATE_IC_TRAIN_LIBRARY",
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

    print("=" * 100)
    print("IC K=5 -> PCA CASE -> Q5 LONG ONLY -> DAILY 20% COHORT -> 5-DAY HOLD | NETTED COST")
    print("=" * 100)

    for k, p in file_map.items():
        print(
            f"{k:<8}: {p}"
        )

    print()

    # ---------------------------------------------------------------------
    # Base Common252
    # ---------------------------------------------------------------------
    print(
        "[1/7] Common252 universe 읽기...",
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

    print(
        f"      dates={len(dates):,}, tickers={len(tickers):,}",
        flush=True,
    )

    # ---------------------------------------------------------------------
    # Price
    # ---------------------------------------------------------------------
    print(
        "[2/7] Full History 수정주가 읽기 / 정렬...",
        flush=True,
    )

    d_price, t_price, price_raw0 = read_wide_matrix(
        file_map["price"],
        "수정주가",
    )

    raw_prices = align_matrix(
        dates,
        tickers,
        d_price,
        t_price,
        price_raw0,
        "수정주가",
    )

    # nonpositive invalid
    raw_prices[
        ~np.isfinite(
            raw_prices
        )
        | (
            raw_prices <= 0
        )
    ] = np.nan

    mark_prices = forward_fill_prices(
        raw_prices
    )

    print(
        f"      finite raw prices={int(np.isfinite(raw_prices).sum()):,}",
        flush=True,
    )

    # ---------------------------------------------------------------------
    # IC
    # ---------------------------------------------------------------------
    print(
        "[3/7] PCA Case 정의 + IC_FWD5 읽기...",
        flush=True,
    )

    cases = read_pca_case_definitions(
        file_map["ic"]
    )

    print(
        f"      PCA candidate cases={len(cases)}",
        flush=True,
    )

    ic_matrix = read_ic_fwd5_matrix(
        file_map["ic"],
        dates,
        cases,
    )

    (
        decisions,
        skipped_nonpositive,
        skipped_no_ic,
    ) = build_decision_records(
        dates=dates,
        ic_matrix=ic_matrix,
        cases=cases,
        start_date=start_date,
        end_date=end_date,
        require_positive_ic=require_positive_ic,
    )

    print(
        f"      valid IC decisions={len(decisions):,}",
        flush=True,
    )

    print(
        f"      skipped nonpositive IC={skipped_nonpositive:,}",
        flush=True,
    )

    # ---------------------------------------------------------------------
    # Current signals -> current Q5 Long
    # ---------------------------------------------------------------------
    print(
        "[4/7] 선택 Case별 현재 decision-date Signal로 Q5 Long 종목 생성...",
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
        f"      tradable cohort specs={len(specs):,}",
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

    # ---------------------------------------------------------------------
    # Sim
    # ---------------------------------------------------------------------
    print(
        "[5/7] Daily overlapping cohort portfolio 시뮬레이션...",
        flush=True,
    )

    (
        daily_rows,
        cohort_rows,
        totals,
    ) = simulate_portfolio(
        specs=specs,
        dates=dates,
        tickers=tickers,
        raw_prices=raw_prices,
        mark_prices=mark_prices,
        buy_cost_rate=buy_cost_rate,
        sell_cost_rate=sell_cost_rate,
    )

    # Same signals / same long-only construction, but zero transaction cost.
    (
        no_cost_daily_rows,
        _no_cost_cohort_rows,
        _no_cost_totals,
    ) = simulate_portfolio(
        specs=specs,
        dates=dates,
        tickers=tickers,
        raw_prices=raw_prices,
        mark_prices=mark_prices,
        buy_cost_rate=0.0,
        sell_cost_rate=0.0,
    )

    # ---------------------------------------------------------------------
    # Summary / usage
    # ---------------------------------------------------------------------
    print(
        "[6/7] 성과 요약...",
        flush=True,
    )

    summary_rows = build_summary(
        daily_rows=daily_rows,
        cohort_rows=cohort_rows,
        totals=totals,
        buy_cost_bps=args.buy_cost_bps,
        sell_cost_bps=args.sell_cost_bps,
        skipped_nonpositive=skipped_nonpositive,
        skipped_no_ic=skipped_no_ic,
        skipped_signal=skipped_signal,
        skipped_entry=skipped_entry,
        no_cost_daily_rows=no_cost_daily_rows,
    )

    usage = Counter(
        r[4]
        for r in cohort_rows
    )

    total_cohorts = max(
        1,
        len(
            cohort_rows
        ),
    )

    usage_rows = [
        (
            case_id,
            count,
            count
            / total_cohorts,
        )
        for case_id, count in usage.most_common()
    ]

    # ---------------------------------------------------------------------
    # Output
    # ---------------------------------------------------------------------
    print(
        "[7/7] Excel + 성과 그래프 저장...",
        flush=True,
    )

    chart_paths = save_performance_charts(
        daily_rows=daily_rows,
        no_cost_daily_rows=no_cost_daily_rows,
        cohort_rows=cohort_rows,
        output_dir=output_dir,
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs_xml = []

        # PARAMETERS ------------------------------------------------------
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Strategy",
                "Lagged realized IC(K=5) selects PCA Case; current Q5 Long only",
            ),
            (
                "IC_Availability",
                "IC at signal date s becomes known after close of s+5",
            ),
            (
                "Decision",
                "At current close t=s+5 choose highest already-known PCA IC_FWD5",
            ),
            (
                "StockSelection",
                "Use current date t signal of selected PCA Case, not old signal date s",
            ),
            (
                "Execution",
                "Enter at next trading-day close t+1",
            ),
            (
                "Long",
                "Current Signal top 20% (Q5 residual losers)",
            ),
            (
                "Short",
                "NOT USED",
            ),
            (
                "Daily_New_Cohort_Gross",
                COHORT_GROSS_FRACTION,
            ),
            (
                "Long_Gross_Per_Cohort",
                COHORT_GROSS_FRACTION,
            ),
            (
                "Short_Gross_Per_Cohort",
                0.0,
            ),
            (
                "Holding_Days",
                HOLD_DAYS,
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
                "Transaction_Cost_Method",
                "Stock-level daily NETTING of expiring + new cohort trades",
            ),
            (
                "Netting_Rule",
                "Charge cost only on final signed share change per ticker after all daily exit/entry orders are offset",
            ),
            (
                "Long_RoundTrip_Bps",
                args.buy_cost_bps + args.sell_cost_bps,
            ),

            (
                "Require_Positive_IC",
                "NO"
                if args.allow_negative_ic
                else "YES",
            ),
            (
                "Candidate_Cases",
                "PCA only: W60/W120/W252 x PC1/PC12/PC123 x N5/N10/N30/N60",
            ),
        ]

        write_table_sheet(
            p,
            ["Item", "Value"],
            parameter_rows,
            [28, 100],
        )

        specs_xml.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # SUMMARY ---------------------------------------------------------
        p = td / "summary.xml"

        write_table_sheet(
            p,
            ["Metric", "Value"],
            summary_rows,
            [34, 28],
        )

        specs_xml.append(
            (
                "SUMMARY",
                p,
            )
        )

        # DAILY -----------------------------------------------------------
        p = td / "daily.xml"

        daily_header = [
            "Date",
            "Equity",
            "DailyNetReturn",
            "GrossPnLAmount",
            "TradingCostAmount",
            "ActiveCohorts",
            "GrossExposure",
            "NetLongExposure",
            "CashWeight",
            "Turnover_NETTED",
            "Turnover_UNNETTED",
            "BuyNotional_NETTED",
            "SellNotional_NETTED",
            "TradeNotional_NETTED",
            "TradeNotional_UNNETTED",
            "CostSavedByNetting",
            "NewCohortEntered",
            "SelectedCase",
            "LaggedIC_FWD5",
            "ICSourceDate",
            "DecisionDate",
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
                15,
                17,
                18,
                19,
                15,
                16,
                18,
                15,
                18,
                20,
                20,
                20,
                21,
                23,
                20,
                18,
                26,
                17,
                15,
                15,
            ],
        )

        specs_xml.append(
            (
                "DAILY",
                p,
            )
        )

        # COHORTS ---------------------------------------------------------
        p = td / "cohorts.xml"

        cohort_header = [
            "ICSourceDate",
            "DecisionDate",
            "EntryDate",
            "ExitDate",
            "CaseID",
            "LaggedIC_FWD5",
            "LongN",
            "EntryEquity",
            "InitialLongAllocation",
            "GrossPnL",
            "GrossReturnOnGross",
            "EntryMissingFiltered",
            "ExitMissingRawCount",
            "LongTickers",
        ]

        cohort_out = (
            (
                yyyymmdd_text(
                    r[0]
                ),
                yyyymmdd_text(
                    r[1]
                ),
                yyyymmdd_text(
                    r[2]
                ),
                yyyymmdd_text(
                    r[3]
                ),
                *r[4:],
            )
            for r in cohort_rows
        )

        write_table_sheet(
            p,
            cohort_header,
            cohort_out,
            [
                15,
                15,
                15,
                15,
                26,
                17,
                11,
                16,
                22,
                16,
                20,
                20,
                20,
                45,
            ],
        )

        specs_xml.append(
            (
                "COHORTS",
                p,
            )
        )

        # DAILY_NO_COST -----------------------------------------------------
        p = td / "daily_no_cost.xml"

        no_cost_daily_out = (
            (
                yyyymmdd_text(
                    r[0]
                ),
                *r[1:],
            )
            for r in no_cost_daily_rows
        )

        write_table_sheet(
            p,
            daily_header,
            no_cost_daily_out,
            [
                13,
                15,
                17,
                18,
                19,
                15,
                16,
                18,
                15,
                18,
                20,
                20,
                20,
                21,
                23,
                20,
                18,
                26,
                17,
                15,
                15,
            ],
        )

        specs_xml.append(
            (
                "DAILY_NO_COST",
                p,
            )
        )

        # CASE USAGE ------------------------------------------------------
        p = td / "usage.xml"

        write_table_sheet(
            p,
            [
                "CaseID",
                "CompletedCohorts",
                "UsageRatio",
            ],
            usage_rows,
            [
                28,
                20,
                16,
            ],
        )

        specs_xml.append(
            (
                "CASE_USAGE",
                p,
            )
        )

        package_xlsx(
            output_path,
            specs_xml,
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
    print("=" * 100)
    print("DONE")
    print("=" * 100)
    print("OUTPUT:", output_path)
    print()
    print("먼저 확인할 시트:")
    print("  1) SUMMARY")
    print("  2) DAILY")
    print("  3) COHORTS")
    print("  4) CASE_USAGE")
    print("  5) DAILY_NO_COST")
    print()
    print("그래프:")
    for p in chart_paths:
        print(" ", p)


if __name__ == "__main__":
    main()
