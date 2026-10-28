# -*- coding: utf-8 -*-
"""
06_build_state_ic_train_library_STRUCTURE_FIXED.py
==================================================

목적
----
아래 8개 Excel 결과파일을 이용해,

1) 날짜별 PCA 직접출력 기반 시장상태 X_T (18차원)
2) 날짜별 40개 Case × 미래 5/10/20/60일 Spearman Rank IC
3) X_T + Case + IC를 합친 Train State-Case Library

를 하나의 Excel로 만든다.

[STRUCTURE FIX]
- XLSX [Content_Types].xml에서 workbook.xml / styles.xml Override 중복 등록을 제거했다.
- X, Case, Signal, IC, Train leakage 계산 로직은 이전 버전과 동일하다.

이 단계에서는:
- Long/Short 종목을 고르지 않는다.
- Portfolio PnL을 계산하지 않는다.
- 현재 날짜의 추천 Case도 아직 선택하지 않는다.
- 과거 상황 라이브러리만 구축한다.

================================================================================
필요한 입력파일 8개
================================================================================

같은 폴더(또는 하위폴더)에 아래 파일을 둔다.

1. 00_COMMON252_BASE.xlsx
2. 01_PCA_W60_DAILY.xlsx
3. 02_PCA_W60_CUMULATIVE.xlsx
4. 03_PCA_W120_DAILY.xlsx
5. 04_PCA_W120_CUMULATIVE.xlsx
6. 05_PCA_W252_DAILY.xlsx
7. 06_PCA_W252_CUMULATIVE.xlsx
8. FUTURE_LOG_RETURNS_5_10_20_60.xlsx

파일명 뒤에 (1), (2) 등이 붙어도 자동 탐색한다.

================================================================================
시장상태 X_T : PCA에서 직접 나온 값만 사용
================================================================================

각 PCA window W = 60 / 120 / 252에 대해:

    EVR1_W = PC1 Explained Ratio
    EVR2_W = PC2 Explained Ratio
    EVR3_W = PC3 Explained Ratio

그리고 Factor Score 자체는 PC마다 분산 크기가 다르므로

    G1_W = Factor1_W / sqrt(lambda1_W)
    G2_W = Factor2_W / sqrt(lambda2_W)
    G3_W = Factor3_W / sqrt(lambda3_W)

로 PC 자체의 표준편차 단위 shock으로 바꾼다.

따라서 총:

    6개 × 3 windows = 18개 상태변수

X_T = [
    EVR1_60, EVR2_60, EVR3_60, G1_60, G2_60, G3_60,
    EVR1_120, EVR2_120, EVR3_120, G1_120, G2_120, G3_120,
    EVR1_252, EVR2_252, EVR3_252, G1_252, G2_252, G3_252
]

평균상관계수, 시장수익률, VIX 등은 이번 첫 버전 X에 넣지 않는다.

================================================================================
Case 정의 : 실질적 40개
================================================================================

[Raw]
Raw는 PCA window와 무관하므로 중복하지 않는다.

    RAW_N5
    RAW_N10
    RAW_N30
    RAW_N60

총 4개.

[PCA-adjusted]
PCA Window 3개:
    W60, W120, W252

PC 제거 3개:
    PC1
    PC12   (= PC1+PC2)
    PC123  (= PC1+PC2+PC3)

과거 Formation 4개:
    N5, N10, N30, N60

3 × 3 × 4 = 36개.

총:
    4 + 36 = 40 cases

================================================================================
Signal 정의
================================================================================

각 case의 cumulative log return 자체는
"낮을수록 과거 loser" 이므로 reversal 방향을 양(+)으로 맞추기 위해:

    Signal(i,T) = - CumulativeReturn(i,T)

예:
    W60_PC12_N10:
    Signal(i,T) = - PC12_CUM10(i,T)

Raw:
    RAW_N10:
    Signal(i,T) = - RAW_CUM_LOG_10(i,T)

================================================================================
IC 정의
================================================================================

미래 horizon k = 5 / 10 / 20 / 60.

날짜 T의 Common252 종목 횡단면에서:

    IC_T(case, k)
        = Spearman_i(
              Signal_case(i,T),
              FWD_LOG_k(i,T)
          )

중요:
- Signal과 Future Return은 종목별 값.
- IC는 "날짜 T × case × 미래 horizon"마다 숫자 하나.
- IC > 0 : 과거 loser일수록 미래수익률 순위가 높은 reversal 방향.
- IC < 0 : momentum 방향.

================================================================================
Train/Test leakage 방지
================================================================================

--train-end 를 지정하면,
미래 k일 수익률이 Train 종료일을 넘어가는 날짜의 IC_k는 Train label로 사용하지 않는다.

예:
Train 종료 index가 E이고 FWD20이면,

    T + 20 <= E

인 날짜에 대해서만 IC20을 저장한다.

따라서 Test 미래정보가 Train IC에 섞이지 않는다.

권장 실행:
    python 06_build_state_ic_train_library.py ^
        --train-start 2009-01-08 ^
        --train-end 2018-12-31

Train 날짜는 연구 설계에 맞게 추후 확정한다.

train-start/end를 생략하면 전체 가용기간을 library로 만들지만,
최종 연구에서는 반드시 Train/Test 기간을 사전에 고정하는 것을 권장한다.

================================================================================
출력
================================================================================

STATE_IC_LIBRARY/
    STATE_IC_TRAIN_LIBRARY.xlsx

시트:
    README
    CASE_DEFINITION
    MARKET_STATE_X
    CASE_DAILY_IC
    TRAIN_STATE_CASE_LIBRARY

CASE_DAILY_IC:
    Date
    CaseID
    PCA_Window
    PC_Remove
    Formation_N
    IC_FWD5 / N_FWD5
    IC_FWD10 / N_FWD10
    IC_FWD20 / N_FWD20
    IC_FWD60 / N_FWD60

TRAIN_STATE_CASE_LIBRARY:
    위 Case/IC 값 + 그 날짜의 X 18개가 같은 행에 붙는다.

이 파일이 다음 단계에서:
"현재 X_T와 유사한 Train 날짜를 찾고,
그 유사상황에서 어느 Case의 Local IC가 좋았는가"
를 계산하기 위한 역사적 라이브러리가 된다.

필수 패키지:
    pip install numpy

pandas / openpyxl / h5py는 필요하지 않는다.
"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import numpy as np


# =============================================================================
# 0. CONSTANTS
# =============================================================================

PCA_WINDOWS = (60, 120, 252)
FORMATION_HORIZONS = (5, 10, 30, 60)
FUTURE_HORIZONS = (5, 10, 20, 60)

MIN_IC_STOCKS = 30
EPS = 1e-15

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# =============================================================================
# 1. BASIC HELPERS
# =============================================================================

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
    """파일명의 (1), 공백 등을 무시한 탐색 보조."""
    s = s.lower()
    s = re.sub(r"\(\d+\)", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def find_input_file(base: Path, key: str) -> Path:
    """
    base 및 하위폴더를 재귀 탐색한다.
    key는 파일명에 반드시 포함되어야 하는 핵심 문자열.
    여러 개면 최근 수정 파일 선택.
    """
    key_norm = normalize_filename_token(key)

    candidates = []

    for p in base.rglob("*.xlsx"):
        name_norm = normalize_filename_token(p.name)

        if key_norm in name_norm:
            # 출력 파일을 다시 입력으로 잡는 것 방지
            if "state_ic_train_library" in name_norm:
                continue

            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            f"\n필수 파일을 찾지 못했습니다: {key}\n"
            f"검색 폴더: {base}\n"
        )

    return max(
        candidates,
        key=lambda p: p.stat().st_mtime,
    )


def col_to_idx(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref)

    if not m:
        raise ValueError(cell_ref)

    n = 0

    for ch in m.group(1):
        n = n * 26 + (ord(ch) - 64)

    return n - 1


def idx_to_col(idx_zero_based: int) -> str:
    n = idx_zero_based + 1
    s = ""

    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s

    return s


def excel_serial_to_yyyymmdd(value) -> int:
    d = datetime(1899, 12, 30) + timedelta(days=float(value))
    return int(d.strftime("%Y%m%d"))


def parse_date_value(value) -> int:
    """
    Excel serial / YYYYMMDD / YYYY-MM-DD / datetime 유사 값을 YYYYMMDD int로.
    """
    if value is None or value == "":
        raise ValueError("empty date")

    if isinstance(value, (int, float, np.integer, np.floating)):
        x = float(value)

        # Excel serial
        if 20000 <= x <= 80000:
            return excel_serial_to_yyyymmdd(x)

        # YYYYMMDD
        if 19000101 <= x <= 21001231:
            return int(x)

    s = str(value).strip()

    if re.fullmatch(r"\d{8}", s):
        return int(s)

    if re.fullmatch(r"\d+(\.\d+)?", s):
        x = float(s)

        if 20000 <= x <= 80000:
            return excel_serial_to_yyyymmdd(x)

    # ISO/date text
    digits = re.findall(r"\d+", s)

    if len(digits) >= 3:
        y, m, d = map(int, digits[:3])
        return y * 10000 + m * 100 + d

    raise ValueError(f"날짜 파싱 실패: {value!r}")


def parse_cli_date(s: str | None) -> int | None:
    if not s:
        return None

    return parse_date_value(s)


# =============================================================================
# 2. XLSX STREAM READER
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

    result = []

    for si in root:
        result.append(
            "".join(
                (t.text or "")
                for t in si.iter(
                    f"{{{NS_MAIN}}}t"
                )
            )
        )

    return result


def cell_value(
    cell: ET.Element,
    shared: list[str],
):
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
            for event, elem in ET.iterparse(
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

                    ci = col_to_idx(ref)

                    cells[ci] = cell_value(
                        c,
                        shared,
                    )

                yield cells

                elem.clear()


def read_wide_matrix(
    path: Path,
    sheet_name: str,
) -> tuple[np.ndarray, list[str], np.ndarray]:
    """
    Date × Ticker wide matrix.
    반환:
        dates: T
        tickers: N
        arr: T x N float64
    """
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

    tickers = []

    for c in range(
        1,
        max_col + 1,
    ):
        tickers.append(
            normalize_ticker(
                header.get(c, "")
            )
        )

    dates = []
    rows = []

    for cells in it:
        if 0 not in cells:
            continue

        try:
            date = parse_date_value(
                cells[0]
            )
        except Exception:
            continue

        row = np.full(
            len(tickers),
            np.nan,
            dtype=np.float64,
        )

        for c, value in cells.items():
            if c == 0:
                continue

            j = c - 1

            if (
                j < 0
                or j >= len(tickers)
                or value in (None, "")
            ):
                continue

            try:
                row[j] = float(value)
            except Exception:
                pass

        dates.append(
            date
        )

        rows.append(
            row
        )

    return (
        np.asarray(
            dates,
            dtype=np.int32,
        ),
        tickers,
        np.vstack(rows)
        if rows
        else np.empty(
            (0, len(tickers))
        ),
    )


def read_summary(
    path: Path,
    sheet_name: str = "DAILY_SUMMARY",
) -> tuple[np.ndarray, dict[str, np.ndarray]]:
    """
    DAILY_SUMMARY -> Date + numeric columns.
    """
    it = iter_sheet_rows(
        path,
        sheet_name,
    )

    header_cells = next(it)

    max_col = max(
        header_cells.keys()
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

    colmap = {
        h: i
        for i, h in enumerate(
            headers
        )
        if h
    }

    required = [
        "Date",
        "PC1_Eigenvalue",
        "PC2_Eigenvalue",
        "PC3_Eigenvalue",
        "PC1_ExplainedRatio",
        "PC2_ExplainedRatio",
        "PC3_ExplainedRatio",
        "PC1_FactorScore",
        "PC2_FactorScore",
        "PC3_FactorScore",
    ]

    missing = [
        c
        for c in required
        if c not in colmap
    ]

    if missing:
        raise KeyError(
            f"{path.name}: DAILY_SUMMARY 필수열 누락: {missing}\n"
            f"실제열={headers}"
        )

    dates = []

    values = {
        c: []
        for c in required
        if c != "Date"
    }

    for cells in it:
        date_val = cells.get(
            colmap["Date"]
        )

        if date_val in (
            None,
            "",
        ):
            continue

        try:
            d = parse_date_value(
                date_val
            )
        except Exception:
            continue

        dates.append(d)

        for name in values:
            raw = cells.get(
                colmap[name]
            )

            try:
                values[name].append(
                    float(raw)
                    if raw not in (
                        None,
                        "",
                    )
                    else np.nan
                )
            except Exception:
                values[name].append(
                    np.nan
                )

    return (
        np.asarray(
            dates,
            dtype=np.int32,
        ),
        {
            k: np.asarray(
                v,
                dtype=np.float64,
            )
            for k, v in values.items()
        },
    )


# =============================================================================
# 3. ALIGNMENT / SPEARMAN
# =============================================================================

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


def rank_average_ties(
    x: np.ndarray,
) -> np.ndarray:
    """
    scipy 없이 average-tie rank.
    Rank = 1 ... n.
    """
    n = len(x)

    order = np.argsort(
        x,
        kind="mergesort",
    )

    sorted_x = x[order]

    ranks_sorted = np.empty(
        n,
        dtype=np.float64,
    )

    i = 0

    while i < n:
        j = i + 1

        while (
            j < n
            and sorted_x[j]
            == sorted_x[i]
        ):
            j += 1

        # i..j-1 의 1-based 평균 rank
        rank = (
            (i + 1)
            + j
        ) / 2.0

        ranks_sorted[
            i:j
        ] = rank

        i = j

    ranks = np.empty(
        n,
        dtype=np.float64,
    )

    ranks[
        order
    ] = ranks_sorted

    return ranks


def spearman_ic(
    signal: np.ndarray,
    future: np.ndarray,
    eligible: np.ndarray,
) -> tuple[float, int]:
    valid = (
        eligible
        & np.isfinite(signal)
        & np.isfinite(future)
    )

    n = int(
        valid.sum()
    )

    if n < MIN_IC_STOCKS:
        return np.nan, n

    x = signal[
        valid
    ]

    y = future[
        valid
    ]

    rx = rank_average_ties(
        x
    )

    ry = rank_average_ties(
        y
    )

    sx = float(
        np.std(
            rx,
            ddof=1,
        )
    )

    sy = float(
        np.std(
            ry,
            ddof=1,
        )
    )

    if (
        sx <= EPS
        or sy <= EPS
    ):
        return np.nan, n

    ic = float(
        np.corrcoef(
            rx,
            ry,
        )[0, 1]
    )

    return ic, n


# =============================================================================
# 4. CASE DEFINITIONS
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    pca_window: str
    pc_remove: str
    formation_n: int
    source_file_key: str
    source_sheet: str


def make_cases() -> list[CaseDef]:
    cases = []

    # Raw: 4 unique cases
    for n in FORMATION_HORIZONS:
        cases.append(
            CaseDef(
                case_id=f"RAW_N{n}",
                pca_window="RAW",
                pc_remove="RAW",
                formation_n=n,
                source_file_key="base",
                source_sheet=f"RAW_CUM_LOG_{n}",
            )
        )

    # PCA-adjusted: 36 cases
    pc_map = {
        "PC1": "PC1",
        "PC12": "PC12",
        "PC123": "PC123",
    }

    for W in PCA_WINDOWS:
        for pc_name, sheet_prefix in pc_map.items():
            for n in FORMATION_HORIZONS:
                cases.append(
                    CaseDef(
                        case_id=f"W{W}_{pc_name}_N{n}",
                        pca_window=str(W),
                        pc_remove=pc_name,
                        formation_n=n,
                        source_file_key=f"cum{W}",
                        source_sheet=f"{sheet_prefix}_CUM{n}",
                    )
                )

    if len(cases) != 40:
        raise RuntimeError(
            f"case count != 40: {len(cases)}"
        )

    return cases


# =============================================================================
# 5. BUILD X
# =============================================================================

def build_market_state_x(
    daily_files: dict[int, Path],
) -> tuple[
    np.ndarray,
    list[str],
    np.ndarray,
]:
    """
    return:
        dates
        feature_names (18)
        X : T x 18
    """
    dates_ref = None

    all_blocks = []

    feature_names = []

    for W in PCA_WINDOWS:
        dates, s = read_summary(
            daily_files[W]
        )

        if dates_ref is None:
            dates_ref = dates
        else:
            require_same_dates(
                dates_ref,
                dates,
                f"W{W} DAILY_SUMMARY",
            )

        ev1 = s[
            "PC1_Eigenvalue"
        ]
        ev2 = s[
            "PC2_Eigenvalue"
        ]
        ev3 = s[
            "PC3_Eigenvalue"
        ]

        ratio1 = s[
            "PC1_ExplainedRatio"
        ]
        ratio2 = s[
            "PC2_ExplainedRatio"
        ]
        ratio3 = s[
            "PC3_ExplainedRatio"
        ]

        f1 = s[
            "PC1_FactorScore"
        ]
        f2 = s[
            "PC2_FactorScore"
        ]
        f3 = s[
            "PC3_FactorScore"
        ]

        g1 = np.full_like(
            f1,
            np.nan,
        )
        g2 = np.full_like(
            f2,
            np.nan,
        )
        g3 = np.full_like(
            f3,
            np.nan,
        )

        valid1 = (
            np.isfinite(f1)
            & np.isfinite(ev1)
            & (ev1 > EPS)
        )

        valid2 = (
            np.isfinite(f2)
            & np.isfinite(ev2)
            & (ev2 > EPS)
        )

        valid3 = (
            np.isfinite(f3)
            & np.isfinite(ev3)
            & (ev3 > EPS)
        )

        g1[
            valid1
        ] = (
            f1[valid1]
            / np.sqrt(
                ev1[valid1]
            )
        )

        g2[
            valid2
        ] = (
            f2[valid2]
            / np.sqrt(
                ev2[valid2]
            )
        )

        g3[
            valid3
        ] = (
            f3[valid3]
            / np.sqrt(
                ev3[valid3]
            )
        )

        block = np.column_stack(
            [
                ratio1,
                ratio2,
                ratio3,
                g1,
                g2,
                g3,
            ]
        )

        all_blocks.append(
            block
        )

        feature_names.extend(
            [
                f"EVR1_{W}",
                f"EVR2_{W}",
                f"EVR3_{W}",
                f"G1_{W}",
                f"G2_{W}",
                f"G3_{W}",
            ]
        )

    X = np.column_stack(
        all_blocks
    )

    return (
        dates_ref,
        feature_names,
        X,
    )


# =============================================================================
# 6. TRAIN DATE RULES
# =============================================================================

def make_train_mask(
    dates: np.ndarray,
    train_start: int | None,
    train_end: int | None,
) -> tuple[np.ndarray, int]:
    mask = np.ones(
        len(dates),
        dtype=bool,
    )

    if train_start is not None:
        mask &= (
            dates >= train_start
        )

    if train_end is not None:
        mask &= (
            dates <= train_end
        )

    valid_idx = np.flatnonzero(
        mask
    )

    if len(valid_idx) == 0:
        raise RuntimeError(
            "Train 날짜가 0개입니다."
        )

    train_end_index = int(
        valid_idx[-1]
    )

    return mask, train_end_index


# =============================================================================
# 7. OUTPUT XLSX WRITER
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
    style = ' s="1"' if header else ""

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
    row_iter,
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
            row_iter,
            start=2,
        ):
            out.write(
                f'<row r="{last_row}">'
            )

            for c, value in enumerate(
                row,
                start=1,
            ):
                if (
                    value is None
                    or (
                        isinstance(
                            value,
                            float,
                        )
                        and not math.isfinite(
                            value
                        )
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

    # workbook.xml / styles.xml are already declared once in
    # CONTENT_TYPES_TEMPLATE above.
    # Only worksheet overrides are added here.
    # (Previous version duplicated workbook/styles Override entries and
    # could trigger Excel's "We found a problem with some content" repair warning.)
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

        for i, (
            _,
            path,
        ) in enumerate(
            sheet_specs,
            start=1,
        ):
            z.write(
                path,
                f"xl/worksheets/sheet{i}.xml",
            )


# =============================================================================
# 8. MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--data-dir",
        default=None,
        help=(
            "8개 입력 Excel이 있는 상위폴더. "
            "생략하면 이 .py가 있는 폴더를 재귀 탐색."
        ),
    )

    parser.add_argument(
        "--train-start",
        default=None,
        help="Train 시작일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--train-end",
        default=None,
        help="Train 종료일 YYYY-MM-DD",
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
        / "STATE_IC_LIBRARY"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    train_start = parse_cli_date(
        args.train_start
    )

    train_end = parse_cli_date(
        args.train_end
    )

    print("=" * 100)
    print("BUILD PCA MARKET STATE X + CASE DAILY IC + TRAIN LIBRARY")
    print("=" * 100)
    print("DATA DIR:", data_dir)
    print("TRAIN START:", train_start)
    print("TRAIN END  :", train_end)
    print()

    # -------------------------------------------------------------------------
    # Find 8 files
    # -------------------------------------------------------------------------
    file_map = {
        "base": find_input_file(
            data_dir,
            "00_COMMON252_BASE",
        ),
        "daily60": find_input_file(
            data_dir,
            "01_PCA_W60_DAILY",
        ),
        "cum60": find_input_file(
            data_dir,
            "02_PCA_W60_CUMULATIVE",
        ),
        "daily120": find_input_file(
            data_dir,
            "03_PCA_W120_DAILY",
        ),
        "cum120": find_input_file(
            data_dir,
            "04_PCA_W120_CUMULATIVE",
        ),
        "daily252": find_input_file(
            data_dir,
            "05_PCA_W252_DAILY",
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

    print("INPUT FILES")

    for k, p in file_map.items():
        print(
            f"  {k:<10}: {p}"
        )

    # -------------------------------------------------------------------------
    # Read Common252 universe
    # -------------------------------------------------------------------------
    print(
        "\n[1/5] Common252 universe 읽기...",
        flush=True,
    )

    (
        dates,
        tickers,
        common252,
    ) = read_wide_matrix(
        file_map["base"],
        "COMMON252_ELIGIBLE",
    )

    common252_bool = (
        np.nan_to_num(
            common252,
            nan=0.0,
        )
        == 1
    )

    T, N = common252.shape

    print(
        f"      T={T:,} N={N:,}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Future matrices
    # -------------------------------------------------------------------------
    print(
        "[2/5] 미래 FWD 5/10/20/60 읽기...",
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

        print(
            f"      FWD{h} finite={int(np.isfinite(arr).sum()):,}",
            flush=True,
        )

    # -------------------------------------------------------------------------
    # Market State X
    # -------------------------------------------------------------------------
    print(
        "[3/5] PCA 직접출력 -> X 18개 구성...",
        flush=True,
    )

    daily_files = {
        60: file_map[
            "daily60"
        ],
        120: file_map[
            "daily120"
        ],
        252: file_map[
            "daily252"
        ],
    }

    (
        x_dates,
        feature_names,
        X,
    ) = build_market_state_x(
        daily_files
    )

    require_same_dates(
        dates,
        x_dates,
        "MARKET STATE X",
    )

    if X.shape != (
        T,
        18,
    ):
        raise RuntimeError(
            f"X shape error: {X.shape}"
        )

    print(
        f"      X shape={X.shape}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Train mask
    # -------------------------------------------------------------------------
    train_mask, train_end_index = make_train_mask(
        dates,
        train_start,
        train_end,
    )

    train_indices = np.flatnonzero(
        train_mask
    )

    print(
        "      Train dates="
        f"{len(train_indices):,} | "
        f"{dates[train_indices[0]]} ~ "
        f"{dates[train_indices[-1]]}",
        flush=True,
    )

    if train_end is None:
        print(
            "\n[주의] --train-end 를 지정하지 않았습니다. "
            "전체 가용기간으로 library를 만듭니다.\n"
            "최종 연구에서는 Train/Test를 먼저 고정하는 것을 권장합니다.",
            flush=True,
        )

    # -------------------------------------------------------------------------
    # Cases and IC
    # -------------------------------------------------------------------------
    print(
        "\n[4/5] 40개 Case × 미래 4개 horizon 일별 Spearman IC 계산...",
        flush=True,
    )

    cases = make_cases()

    # 결과:
    # case_id -> IC T x 4, N T x 4
    ic_results = {}

    # cumulative 파일 cache는 case group별 한 sheet만 읽고 버림.
    for case_no, case in enumerate(
        cases,
        start=1,
    ):
        source_path = file_map[
            case.source_file_key
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

        # Reversal Score = - cumulative return
        signal = -cumulative

        ic_arr = np.full(
            (
                T,
                len(
                    FUTURE_HORIZONS
                ),
            ),
            np.nan,
            dtype=np.float64,
        )

        n_arr = np.zeros(
            (
                T,
                len(
                    FUTURE_HORIZONS
                ),
            ),
            dtype=np.int16,
        )

        for t_idx in train_indices:
            elig = common252_bool[
                t_idx
            ]

            for h_idx, h in enumerate(
                FUTURE_HORIZONS
            ):
                # Train end leakage 방지:
                # 미래 h 거래일이 Train 밖으로 넘어가면 label 금지.
                if (
                    train_end is not None
                    and t_idx + h
                    > train_end_index
                ):
                    continue

                ic, nobs = spearman_ic(
                    signal[
                        t_idx
                    ],
                    future[h][
                        t_idx
                    ],
                    elig,
                )

                ic_arr[
                    t_idx,
                    h_idx,
                ] = ic

                n_arr[
                    t_idx,
                    h_idx,
                ] = nobs

        ic_results[
            case.case_id
        ] = (
            ic_arr,
            n_arr,
        )

        if (
            case_no % 5 == 0
            or case_no == 1
            or case_no == len(cases)
        ):
            valid_counts = [
                int(
                    np.isfinite(
                        ic_arr[
                            :,
                            j,
                        ]
                    ).sum()
                )
                for j in range(
                    len(
                        FUTURE_HORIZONS
                    )
                )
            ]

            print(
                f"      [{case_no:02d}/40] "
                f"{case.case_id:<20} "
                f"IC valid={valid_counts}",
                flush=True,
            )

        # large matrix 해제
        del cumulative
        del signal

    # -------------------------------------------------------------------------
    # Output XLSX
    # -------------------------------------------------------------------------
    print(
        "\n[5/5] Excel library 저장...",
        flush=True,
    )

    output_path = (
        output_dir
        / "STATE_IC_TRAIN_LIBRARY.xlsx"
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        sheet_specs = []

        # README --------------------------------------------------------------
        readme_path = td / "sheet_readme.xml"

        readme_header = [
            "항목",
            "내용",
        ]

        readme_rows = [
            [
                "목적",
                "PCA 직접출력 기반 시장상태 X와 과거 Case별 일별 IC를 합친 Train 상황 라이브러리",
            ],
            [
                "시장상태 X",
                "PCA60/120/252 각각 EVR1~3 + G1~3, 총 18개",
            ],
            [
                "G 정의",
                "G_k,W = FactorScore_k,W / sqrt(Eigenvalue_k,W)",
            ],
            [
                "Case 수",
                "Raw 4 + PCA-adjusted 36 = 총 40",
            ],
            [
                "Signal",
                "Signal = - cumulative log return; 클수록 과거 loser",
            ],
            [
                "IC",
                "날짜 T 횡단면 Spearman(Signal, 실제 FWD 로그수익률)",
            ],
            [
                "Future",
                "FWD 5/10/20/60, PCA 제거 및 표준화 없음",
            ],
            [
                "Train 시작",
                str(
                    train_start
                    if train_start is not None
                    else "AUTO"
                ),
            ],
            [
                "Train 종료",
                str(
                    train_end
                    if train_end is not None
                    else "AUTO/전체"
                ),
            ],
            [
                "Train leakage",
                "train-end 지정 시 T+h가 Train 종료를 넘는 IC_h는 NaN 처리",
            ],
            [
                "다음 단계",
                "현재 X_T와 유사한 Train 날짜를 검색하고 그 주변의 Local IC로 적합한 Case 추천",
            ],
        ]

        write_table_sheet(
            readme_path,
            readme_header,
            iter(
                readme_rows
            ),
            widths=[
                25,
                95,
            ],
        )

        sheet_specs.append(
            (
                "README",
                readme_path,
            )
        )

        # CASE DEFINITION -----------------------------------------------------
        case_path = td / "sheet_cases.xml"

        case_header = [
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Source_Sheet",
        ]

        def case_rows():
            for c in cases:
                yield [
                    c.case_id,
                    c.pca_window,
                    c.pc_remove,
                    c.formation_n,
                    c.source_sheet,
                ]

        write_table_sheet(
            case_path,
            case_header,
            case_rows(),
            widths=[
                24,
                14,
                14,
                14,
                22,
            ],
        )

        sheet_specs.append(
            (
                "CASE_DEFINITION",
                case_path,
            )
        )

        # MARKET STATE X ------------------------------------------------------
        x_path = td / "sheet_x.xml"

        x_header = [
            "Date",
            *feature_names,
        ]

        def x_rows():
            for i in train_indices:
                row = [
                    str(
                        int(
                            dates[i]
                        )
                    )
                ]

                row.extend(
                    [
                        (
                            None
                            if not np.isfinite(v)
                            else float(v)
                        )
                        for v in X[i]
                    ]
                )

                yield row

        write_table_sheet(
            x_path,
            x_header,
            x_rows(),
            widths=[
                12,
                *(
                    [14]
                    * len(
                        feature_names
                    )
                ),
            ],
        )

        sheet_specs.append(
            (
                "MARKET_STATE_X",
                x_path,
            )
        )

        # CASE DAILY IC -------------------------------------------------------
        ic_path = td / "sheet_ic.xml"

        ic_header = [
            "Date",
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
        ]

        for h in FUTURE_HORIZONS:
            ic_header.extend(
                [
                    f"IC_FWD{h}",
                    f"N_FWD{h}",
                ]
            )

        def daily_ic_rows():
            for i in train_indices:
                for case in cases:
                    ic_arr, n_arr = (
                        ic_results[
                            case.case_id
                        ]
                    )

                    row = [
                        str(
                            int(
                                dates[i]
                            )
                        ),
                        case.case_id,
                        case.pca_window,
                        case.pc_remove,
                        case.formation_n,
                    ]

                    for j, h in enumerate(
                        FUTURE_HORIZONS
                    ):
                        ic = ic_arr[
                            i,
                            j,
                        ]

                        nobs = int(
                            n_arr[
                                i,
                                j,
                            ]
                        )

                        row.extend(
                            [
                                (
                                    None
                                    if not np.isfinite(
                                        ic
                                    )
                                    else float(ic)
                                ),
                                nobs,
                            ]
                        )

                    yield row

        write_table_sheet(
            ic_path,
            ic_header,
            daily_ic_rows(),
            widths=[
                12,
                24,
                13,
                13,
                13,
                13,
                11,
                13,
                11,
                13,
                11,
                13,
                11,
            ],
        )

        sheet_specs.append(
            (
                "CASE_DAILY_IC",
                ic_path,
            )
        )

        # TRAIN STATE-CASE LIBRARY -------------------------------------------
        library_path = (
            td
            / "sheet_library.xml"
        )

        library_header = [
            "Date",
            *feature_names,
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
        ]

        for h in FUTURE_HORIZONS:
            library_header.extend(
                [
                    f"IC_FWD{h}",
                    f"N_FWD{h}",
                ]
            )

        def library_rows():
            for i in train_indices:
                xrow = [
                    (
                        None
                        if not np.isfinite(v)
                        else float(v)
                    )
                    for v in X[i]
                ]

                for case in cases:
                    ic_arr, n_arr = (
                        ic_results[
                            case.case_id
                        ]
                    )

                    row = [
                        str(
                            int(
                                dates[i]
                            )
                        ),
                        *xrow,
                        case.case_id,
                        case.pca_window,
                        case.pc_remove,
                        case.formation_n,
                    ]

                    for j, h in enumerate(
                        FUTURE_HORIZONS
                    ):
                        ic = ic_arr[
                            i,
                            j,
                        ]

                        row.extend(
                            [
                                (
                                    None
                                    if not np.isfinite(
                                        ic
                                    )
                                    else float(ic)
                                ),
                                int(
                                    n_arr[
                                        i,
                                        j,
                                    ]
                                ),
                            ]
                        )

                    yield row

        library_widths = (
            [12]
            + [13] * 18
            + [
                24,
                13,
                13,
                13,
            ]
            + [13, 11] * 4
        )

        write_table_sheet(
            library_path,
            library_header,
            library_rows(),
            widths=library_widths,
        )

        sheet_specs.append(
            (
                "TRAIN_STATE_CASE_LIBRARY",
                library_path,
            )
        )

        package_xlsx(
            output_path,
            sheet_specs,
        )

    # -------------------------------------------------------------------------
    # Final validation
    # -------------------------------------------------------------------------
    with ZipFile(
        output_path
    ) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성 xlsx ZIP 오류: {bad}"
            )

    print()
    print("=" * 100)
    print("DONE")
    print("=" * 100)
    print("OUTPUT:", output_path)
    print("Cases :", len(cases))
    print("X dim :", len(feature_names))
    print("Train dates:", len(train_indices))
    print(
        "Library rows:",
        f"{len(train_indices) * len(cases):,}",
    )
    print()
    print(
        "다음 단계에서는 이 파일의 MARKET_STATE_X와 "
        "TRAIN_STATE_CASE_LIBRARY를 사용해 현재 X_T와 유사한 과거 상황을 찾고 "
        "Local IC 기준으로 추천 Case를 계산하면 됩니다."
    )


if __name__ == "__main__":
    main()
