# -*- coding: utf-8 -*-
"""
09_ridge_ic_annual_expanding_walkforward_MINIMAL.py
===================================================

연도별 Expanding Walk-Forward Ridge IC Case Selection.

기본 구조
---------
- 2025년 Test:
    2024년 말까지 당시 관측 완료된 과거 IC로 학습
- 2026년 Test:
    2025년까지 새로 관측 완료된 과거 IC를 포함해
    alpha / X 표준화 / Ridge를 다시 계산

즉:
    과거 -> 다음 연도 OOS Test
    Test가 끝나면 그 결과는 다음 연도에서 과거가 됨.

연도 안에서도 매일:
    현재 시점까지 maturity가 끝난 과거 IC를 추가해
    Ridge 계수를 expanding update 한다.

Look-ahead 방지
---------------
현재 날짜 T, horizon h에서 학습 가능한 signal date s는:

    s + h <= T - 1

만 허용한다.

연초 alpha 재선택
-----------------
각 Test 연도 시작 전에 그 시점까지 관측 가능한 과거 label만 이용하고,
그 과거 안에서 다시 앞부분 Fit / 뒷부분 Validation으로 나눠
alpha를 horizon별로 고른다.

출력은 연구에 필요한 3개 시트만 생성:
- PARAMETERS
- YEARLY_SUMMARY
- DAILY

기본 first test year = 2025.
필수 외부 패키지: numpy
"""

from __future__ import annotations

import argparse
import math
import re
import tempfile
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import numpy as np


# =============================================================================
# CONSTANTS
# =============================================================================

FUTURE_HORIZONS = (5, 10, 20, 60)

X_COLS = [
    "EVR1_60", "EVR2_60", "EVR3_60", "G1_60", "G2_60", "G3_60",
    "EVR1_120", "EVR2_120", "EVR3_120", "G1_120", "G2_120", "G3_120",
    "EVR1_252", "EVR2_252", "EVR3_252", "G1_252", "G2_252", "G3_252",
]

DEFAULT_ALPHAS = (0.01, 0.1, 1.0, 10.0, 100.0)

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


def normalize_filename_token(s: str) -> str:
    s = s.lower()
    s = re.sub(r"\(\d+\)", "", s)
    s = re.sub(r"\s+", "", s)
    return s


def find_input_file(base: Path) -> Path:
    candidates = []

    for p in base.rglob("*.xlsx"):
        name = normalize_filename_token(p.name)

        if "state_ic_train_library" in name:
            if "ridge_ic_case_selection" in name:
                continue

            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            "STATE_IC_TRAIN_LIBRARY*.xlsx 를 찾지 못했습니다.\n"
            f"검색폴더: {base}"
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

    digits = re.findall(r"\d+", s)

    if len(digits) >= 3:
        y, m, d = map(int, digits[:3])
        return y * 10000 + m * 100 + d

    raise ValueError(f"날짜 파싱 실패: {value!r}")


def parse_cli_date(s: str | None) -> int | None:
    if not s:
        return None

    return parse_date_value(s)


def yyyymmdd_text(x: int) -> str:
    s = str(int(x))
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}"


def parse_alpha_grid(s: str | None) -> tuple[float, ...]:
    if not s:
        return DEFAULT_ALPHAS

    vals = []

    for part in s.split(","):
        x = float(part.strip())

        if x < 0:
            raise ValueError("alpha는 0 이상이어야 합니다.")

        vals.append(x)

    if not vals:
        raise ValueError("alpha 후보가 없습니다.")

    return tuple(vals)


# =============================================================================
# XLSX STREAM READER
# =============================================================================

def workbook_sheet_targets(path: Path) -> dict[str, str]:
    with ZipFile(path) as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))

        relmap = {
            r.attrib["Id"]: r.attrib["Target"]
            for r in rels
        }

        out = {}

        sheets = wb.find(f"{{{NS_MAIN}}}sheets")

        for s in sheets:
            rid = s.attrib[f"{{{NS_REL}}}id"]
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

    root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    result = []

    for si in root:
        result.append(
            "".join(
                (t.text or "")
                for t in si.iter(f"{{{NS_MAIN}}}t")
            )
        )

    return result


def cell_value(cell: ET.Element, shared: list[str]):
    typ = cell.attrib.get("t")

    if typ == "inlineStr":
        t = cell.find(f".//{{{NS_MAIN}}}t")
        return t.text if t is not None else ""

    v = cell.find(f"{{{NS_MAIN}}}v")

    if v is None:
        return None

    if typ == "s":
        return shared[int(v.text)]

    if typ == "b":
        return int(v.text)

    return v.text


def iter_sheet_rows(xlsx_path: Path, sheet_name: str):
    targets = workbook_sheet_targets(xlsx_path)

    if sheet_name not in targets:
        raise KeyError(
            f"{xlsx_path.name}: 시트 없음: {sheet_name}\n"
            f"실제 시트={list(targets)}"
        )

    target = targets[sheet_name]

    with ZipFile(xlsx_path) as z:
        shared = load_shared_strings(z)

        with z.open(target) as f:
            for _, elem in ET.iterparse(f, events=("end",)):
                if elem.tag != f"{{{NS_MAIN}}}row":
                    continue

                cells = {}

                for c in elem:
                    if c.tag != f"{{{NS_MAIN}}}c":
                        continue

                    ref = c.attrib.get("r")

                    if not ref:
                        continue

                    cells[col_to_idx(ref)] = cell_value(c, shared)

                yield cells
                elem.clear()


def read_market_state_x(path: Path):
    it = iter_sheet_rows(path, "MARKET_STATE_X")
    header_cells = next(it)

    max_col = max(header_cells)

    headers = [
        str(header_cells.get(c, "") or "").strip()
        for c in range(max_col + 1)
    ]

    colmap = {
        h: i
        for i, h in enumerate(headers)
        if h
    }

    required = ["Date", *X_COLS]

    missing = [
        c for c in required
        if c not in colmap
    ]

    if missing:
        raise KeyError(
            f"MARKET_STATE_X 필수열 누락: {missing}"
        )

    dates = []
    rows = []

    for cells in it:
        raw_date = cells.get(colmap["Date"])

        if raw_date in (None, ""):
            continue

        try:
            d = parse_date_value(raw_date)
        except Exception:
            continue

        xrow = np.full(
            len(X_COLS),
            np.nan,
            dtype=np.float64,
        )

        for j, name in enumerate(X_COLS):
            raw = cells.get(colmap[name])

            if raw in (None, ""):
                continue

            try:
                xrow[j] = float(raw)
            except Exception:
                pass

        dates.append(d)
        rows.append(xrow)

    if not rows:
        raise RuntimeError(
            "MARKET_STATE_X 유효 행 없음"
        )

    return (
        np.asarray(dates, dtype=np.int32),
        np.vstack(rows),
    )


def read_case_definition(path: Path):
    it = iter_sheet_rows(path, "CASE_DEFINITION")
    header_cells = next(it)

    max_col = max(header_cells)

    headers = [
        str(header_cells.get(c, "") or "").strip()
        for c in range(max_col + 1)
    ]

    colmap = {
        h: i
        for i, h in enumerate(headers)
        if h
    }

    if "CaseID" not in colmap:
        raise KeyError("CASE_DEFINITION: CaseID 열 없음")

    cases = []
    seen = set()

    for cells in it:
        case_id = str(
            cells.get(
                colmap["CaseID"],
                "",
            )
            or ""
        ).strip()

        if not case_id or case_id in seen:
            continue

        seen.add(case_id)
        cases.append(case_id)

    if not cases:
        raise RuntimeError(
            "CASE_DEFINITION 유효 Case 없음"
        )

    return cases


def read_case_daily_ic(
    path: Path,
    all_dates: np.ndarray,
    cases: list[str],
):
    it = iter_sheet_rows(path, "CASE_DAILY_IC")
    header_cells = next(it)

    max_col = max(header_cells)

    headers = [
        str(header_cells.get(c, "") or "").strip()
        for c in range(max_col + 1)
    ]

    colmap = {
        h: i
        for i, h in enumerate(headers)
        if h
    }

    required = [
        "Date",
        "CaseID",
        *[f"IC_FWD{h}" for h in FUTURE_HORIZONS],
    ]

    missing = [
        c for c in required
        if c not in colmap
    ]

    if missing:
        raise KeyError(
            f"CASE_DAILY_IC 필수열 누락: {missing}"
        )

    date_to_idx = {
        int(d): i
        for i, d in enumerate(all_dates)
    }

    case_to_idx = {
        c: i
        for i, c in enumerate(cases)
    }

    T = len(all_dates)
    C = len(cases)

    matrices = {
        h: np.full(
            (T, C),
            np.nan,
            dtype=np.float64,
        )
        for h in FUTURE_HORIZONS
    }

    rows_read = 0

    for cells in it:
        raw_date = cells.get(colmap["Date"])
        case_id = str(
            cells.get(
                colmap["CaseID"],
                "",
            )
            or ""
        ).strip()

        if (
            raw_date in (None, "")
            or case_id not in case_to_idx
        ):
            continue

        try:
            d = parse_date_value(raw_date)
        except Exception:
            continue

        ti = date_to_idx.get(d)

        if ti is None:
            continue

        ci = case_to_idx[case_id]

        for h in FUTURE_HORIZONS:
            raw = cells.get(
                colmap[f"IC_FWD{h}"]
            )

            if raw in (None, ""):
                continue

            try:
                matrices[h][ti, ci] = float(raw)
            except Exception:
                pass

        rows_read += 1

        if rows_read % 50000 == 0:
            print(
                f"      CASE_DAILY_IC {rows_read:,} rows...",
                flush=True,
            )

    return matrices, rows_read


# =============================================================================
# STATS
# =============================================================================

def rank_average_ties(x: np.ndarray) -> np.ndarray:
    x = np.asarray(x, dtype=np.float64)

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
            and sorted_x[j] == sorted_x[i]
        ):
            j += 1

        rank = (
            (i + 1)
            + j
        ) / 2.0

        ranks_sorted[i:j] = rank
        i = j

    ranks = np.empty(
        n,
        dtype=np.float64,
    )

    ranks[order] = ranks_sorted

    return ranks


def safe_spearman(
    a: np.ndarray,
    b: np.ndarray,
) -> float:
    mask = (
        np.isfinite(a)
        & np.isfinite(b)
    )

    if int(mask.sum()) < 5:
        return np.nan

    ra = rank_average_ties(
        a[mask]
    )

    rb = rank_average_ties(
        b[mask]
    )

    sa = float(
        np.std(
            ra,
            ddof=1,
        )
    )

    sb = float(
        np.std(
            rb,
            ddof=1,
        )
    )

    if sa <= EPS or sb <= EPS:
        return np.nan

    return float(
        np.corrcoef(
            ra,
            rb,
        )[0, 1]
    )


def hac_mean_test(
    values: np.ndarray,
    lag: int,
):
    """
    Newey-West HAC t-test for H0: mean = 0.
    p-value는 asymptotic normal two-sided.
    """
    x = np.asarray(
        values,
        dtype=np.float64,
    )

    x = x[np.isfinite(x)]

    T = len(x)

    if T < 5:
        return np.nan, np.nan, np.nan

    mean = float(
        np.mean(x)
    )

    u = x - mean

    gamma0 = float(
        np.dot(u, u) / T
    )

    lrv = gamma0

    L = min(
        int(lag),
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
        return mean, np.nan, np.nan

    se = math.sqrt(
        lrv / T
    )

    t = mean / se

    p = math.erfc(
        abs(t) / math.sqrt(2.0)
    )

    return mean, float(t), float(p)



# =============================================================================
# ANNUAL WALK-FORWARD
# =============================================================================

def complete_x_indices(
    X: np.ndarray,
) -> np.ndarray:
    mask = np.all(
        np.isfinite(X),
        axis=1,
    )

    idx = np.flatnonzero(mask)

    if len(idx) < 100:
        raise RuntimeError(
            "X 18개가 완전한 날짜가 너무 적습니다."
        )

    return idx


def date_year(
    yyyymmdd: int,
) -> int:
    return int(yyyymmdd) // 10000


def make_scaler(
    X: np.ndarray,
    idx: np.ndarray,
):
    xx = X[idx]

    mu = np.mean(
        xx,
        axis=0,
    )

    sd = np.std(
        xx,
        axis=0,
        ddof=1,
    )

    bad = (
        ~np.isfinite(sd)
        | (sd <= EPS)
    )

    if bad.any():
        bad_cols = [
            X_COLS[i]
            for i in np.flatnonzero(bad)
        ]

        raise RuntimeError(
            f"X 표준편차 0/NaN: {bad_cols}"
        )

    return mu, sd


def z_transform(
    X: np.ndarray,
    mu: np.ndarray,
    sd: np.ndarray,
):
    return (
        X - mu
    ) / sd


def design_matrix(
    Z: np.ndarray,
):
    return np.column_stack(
        (
            np.ones(
                len(Z),
                dtype=np.float64,
            ),
            Z,
        )
    )


def design_row(
    zrow: np.ndarray,
):
    return np.concatenate(
        (
            np.array(
                [1.0],
                dtype=np.float64,
            ),
            zrow.astype(
                np.float64,
                copy=False,
            ),
        )
    )


def ridge_fit_batch(
    Z: np.ndarray,
    Y: np.ndarray,
    alpha: float,
):
    valid = (
        np.all(
            np.isfinite(Z),
            axis=1,
        )
        & np.all(
            np.isfinite(Y),
            axis=1,
        )
    )

    Z = Z[valid]
    Y = Y[valid]

    if len(Z) < 50:
        raise RuntimeError(
            f"Ridge 학습표본 부족: {len(Z)}"
        )

    D = design_matrix(Z)

    penalty = np.eye(
        D.shape[1],
        dtype=np.float64,
    ) * float(alpha)

    # intercept는 penalty 제외
    penalty[0, 0] = 0.0

    beta = np.linalg.solve(
        D.T @ D + penalty,
        D.T @ Y,
    )

    return beta


def select_alpha_at_year_start(
    X: np.ndarray,
    label_idx: np.ndarray,
    Y_all: np.ndarray,
    alphas: tuple[float, ...],
    inner_fit_ratio: float,
):
    """
    label_idx는 Test 연도 시작 전에 이미 maturity가 끝난 날짜만 들어온다.

    Train 내부 chronological Fit/Validation:
    각 alpha에 대해 Validation 날짜마다 40개 predicted IC 중 최고 Case를 고르고,
    그 Case의 실제 IC 평균이 가장 높은 alpha를 선택.
    """
    # 모든 40 Case label이 있는 날짜만 사용
    full_label = np.all(
        np.isfinite(
            Y_all[label_idx]
        ),
        axis=1,
    )

    idx = label_idx[
        full_label
    ]

    if len(idx) < 200:
        raise RuntimeError(
            f"alpha 선택용 label 부족: {len(idx)}"
        )

    cut = int(
        math.floor(
            len(idx)
            * inner_fit_ratio
        )
    )

    cut = max(
        100,
        min(
            cut,
            len(idx) - 50,
        ),
    )

    fit_idx = idx[:cut]
    val_idx = idx[cut:]

    mu, sd = make_scaler(
        X,
        fit_idx,
    )

    Z_fit = z_transform(
        X[fit_idx],
        mu,
        sd,
    )

    Z_val = z_transform(
        X[val_idx],
        mu,
        sd,
    )

    Y_fit = Y_all[
        fit_idx
    ]

    Y_val = Y_all[
        val_idx
    ]

    best_alpha = None
    best_score = -np.inf

    for alpha in alphas:
        beta = ridge_fit_batch(
            Z_fit,
            Y_fit,
            alpha,
        )

        pred = (
            design_matrix(Z_val)
            @ beta
        )

        selected_actual = []

        for r in range(
            len(val_idx)
        ):
            pvec = pred[r]
            avec = Y_val[r]

            if (
                not np.isfinite(pvec).any()
                or not np.isfinite(avec).any()
            ):
                continue

            ci = int(
                np.nanargmax(pvec)
            )

            actual = avec[ci]

            if np.isfinite(actual):
                selected_actual.append(
                    float(actual)
                )

        if not selected_actual:
            continue

        score = float(
            np.mean(
                selected_actual
            )
        )

        if score > best_score:
            best_score = score
            best_alpha = float(alpha)

    if best_alpha is None:
        raise RuntimeError(
            "alpha 선택 실패"
        )

    return best_alpha, best_score


def build_year_start_state(
    X: np.ndarray,
    Y_all: np.ndarray,
    label_idx: np.ndarray,
    alpha: float,
):
    """
    Test 연도 시작 시점의 최종 scaler + Ridge sufficient statistics.
    """
    full_label = np.all(
        np.isfinite(
            Y_all[label_idx]
        ),
        axis=1,
    )

    idx = label_idx[
        full_label
    ]

    if len(idx) < 100:
        raise RuntimeError(
            f"연도 시작 학습표본 부족: {len(idx)}"
        )

    mu, sd = make_scaler(
        X,
        idx,
    )

    Z = z_transform(
        X[idx],
        mu,
        sd,
    )

    D = design_matrix(Z)
    Y = Y_all[idx]

    p = D.shape[1]
    C = Y.shape[1]

    xtx = D.T @ D
    xty = D.T @ Y

    added = np.zeros(
        len(X),
        dtype=bool,
    )

    added[idx] = True

    return {
        "mu": mu,
        "sd": sd,
        "alpha": float(alpha),
        "xtx": xtx,
        "xty": xty,
        "added": added,
        "n": int(len(idx)),
        "p": p,
        "C": C,
    }


def solve_state_beta(
    state,
):
    penalty = np.eye(
        state["p"],
        dtype=np.float64,
    ) * state["alpha"]

    penalty[0, 0] = 0.0

    return np.linalg.solve(
        state["xtx"]
        + penalty,
        state["xty"],
    )


def add_new_matured_label(
    state,
    X: np.ndarray,
    Y_all: np.ndarray,
    s: int,
):
    if (
        s < 0
        or s >= len(X)
        or state["added"][s]
    ):
        return

    if not np.all(
        np.isfinite(
            X[s]
        )
    ):
        return

    y = Y_all[s]

    if not np.all(
        np.isfinite(y)
    ):
        return

    z = (
        X[s]
        - state["mu"]
    ) / state["sd"]

    d = design_row(z)

    state["xtx"] += np.outer(
        d,
        d,
    )

    state["xty"] += np.outer(
        d,
        y,
    )

    state["added"][s] = True
    state["n"] += 1


def global_case_at_year_start(
    Y_all: np.ndarray,
    label_idx: np.ndarray,
    cases: list[str],
):
    Y = Y_all[
        label_idx
    ]

    mean_ic = np.nanmean(
        Y,
        axis=0,
    )

    if not np.isfinite(
        mean_ic
    ).any():
        raise RuntimeError(
            "Global Case 계산 실패"
        )

    ci = int(
        np.nanargmax(
            mean_ic
        )
    )

    return {
        "case_idx": ci,
        "case_id": cases[ci],
        "train_mean_ic": float(
            mean_ic[ci]
        ),
    }


def annual_expanding_walkforward(
    all_dates: np.ndarray,
    X: np.ndarray,
    valid_idx: np.ndarray,
    ic_matrices: dict[int, np.ndarray],
    cases: list[str],
    first_test_year: int,
    last_test_year: int,
    alphas: tuple[float, ...],
    inner_fit_ratio: float,
):
    daily_rows = []
    yearly_model_rows = []

    available_years = sorted(
        {
            date_year(
                int(all_dates[i])
            )
            for i in valid_idx
        }
    )

    test_years = [
        y
        for y in available_years
        if first_test_year
        <= y
        <= last_test_year
    ]

    if not test_years:
        raise RuntimeError(
            "요청한 Test 연도가 데이터에 없습니다."
        )

    print(
        "[4/6] 연도별 expanding walk-forward 시작...",
        flush=True,
    )

    for test_year in test_years:
        year_test_idx = np.asarray(
            [
                i
                for i in valid_idx
                if date_year(
                    int(
                        all_dates[i]
                    )
                )
                == test_year
            ],
            dtype=int,
        )

        if len(
            year_test_idx
        ) == 0:
            continue

        first_t = int(
            year_test_idx[0]
        )

        train_through_idx = (
            first_t - 1
        )

        print()
        print(
            f"  Test {test_year}: "
            f"{yyyymmdd_text(all_dates[year_test_idx[0]])} ~ "
            f"{yyyymmdd_text(all_dates[year_test_idx[-1]])}",
            flush=True,
        )

        horizon_states = {}
        global_info = {}

        # -------------------------------------------------------------
        # 연초: horizon별 alpha/scaler/Ridge/Global 재계산
        # -------------------------------------------------------------
        for h in FUTURE_HORIZONS:
            # 연초 시점에 실제로 label이 완전히 알려진 signal date만
            label_idx = valid_idx[
                (valid_idx + h)
                <= train_through_idx
            ]

            if len(
                label_idx
            ) < 200:
                raise RuntimeError(
                    f"{test_year} FWD{h}: 연초 학습 label 부족"
                )

            alpha, val_score = select_alpha_at_year_start(
                X=X,
                label_idx=label_idx,
                Y_all=ic_matrices[h],
                alphas=alphas,
                inner_fit_ratio=inner_fit_ratio,
            )

            state = build_year_start_state(
                X=X,
                Y_all=ic_matrices[h],
                label_idx=label_idx,
                alpha=alpha,
            )

            ginfo = global_case_at_year_start(
                Y_all=ic_matrices[h],
                label_idx=label_idx,
                cases=cases,
            )

            horizon_states[h] = state
            global_info[h] = ginfo

            yearly_model_rows.append(
                (
                    test_year,
                    h,
                    yyyymmdd_text(
                        int(
                            all_dates[
                                train_through_idx
                            ]
                        )
                    ),
                    int(
                        state["n"]
                    ),
                    float(alpha),
                    float(val_score),
                    ginfo["case_id"],
                    float(
                        ginfo[
                            "train_mean_ic"
                        ]
                    ),
                )
            )

            print(
                f"      FWD{h}: alpha={alpha:g} | "
                f"Train labels={state['n']:,} | "
                f"Global={ginfo['case_id']}",
                flush=True,
            )

        # -------------------------------------------------------------
        # 해당 Test 연도 안에서는 매일 matured label 추가
        # -------------------------------------------------------------
        for pos, t in enumerate(
            year_test_idx
        ):
            t = int(t)

            for h in FUTURE_HORIZONS:
                state = horizon_states[h]

                # 현재 T를 예측하기 직전까지 확정된 가장 최근 label
                newest_s = (
                    t - 1 - h
                )

                add_new_matured_label(
                    state=state,
                    X=X,
                    Y_all=ic_matrices[h],
                    s=newest_s,
                )

                beta = solve_state_beta(
                    state
                )

                z_t = (
                    X[t]
                    - state["mu"]
                ) / state["sd"]

                pred = (
                    design_row(
                        z_t
                    )
                    @ beta
                )

                if not np.isfinite(
                    pred
                ).any():
                    continue

                selected_ci = int(
                    np.nanargmax(
                        pred
                    )
                )

                recommended = cases[
                    selected_ci
                ]

                predicted_ic = float(
                    pred[
                        selected_ci
                    ]
                )

                actual_vec = ic_matrices[h][
                    t
                ]

                actual_ic = (
                    float(
                        actual_vec[
                            selected_ci
                        ]
                    )
                    if np.isfinite(
                        actual_vec[
                            selected_ci
                        ]
                    )
                    else None
                )

                rank_spear = safe_spearman(
                    pred,
                    actual_vec,
                )

                gb = global_info[h]
                gb_ci = int(
                    gb["case_idx"]
                )

                global_actual = (
                    float(
                        actual_vec[
                            gb_ci
                        ]
                    )
                    if np.isfinite(
                        actual_vec[
                            gb_ci
                        ]
                    )
                    else None
                )

                daily_rows.append(
                    (
                        int(
                            all_dates[t]
                        ),
                        test_year,
                        h,
                        recommended,
                        predicted_ic,
                        actual_ic,
                        gb["case_id"],
                        global_actual,
                        (
                            float(
                                rank_spear
                            )
                            if np.isfinite(
                                rank_spear
                            )
                            else None
                        ),
                    )
                )

            if (
                (pos + 1) % 100 == 0
                or (pos + 1)
                == len(
                    year_test_idx
                )
            ):
                print(
                    f"        {pos + 1:,}/{len(year_test_idx):,} days",
                    flush=True,
                )

    return (
        daily_rows,
        yearly_model_rows,
    )


# =============================================================================
# YEARLY SUMMARY
# =============================================================================

def build_yearly_summary(
    daily_rows,
    yearly_model_rows,
):
    model_lookup = {
        (
            int(r[0]),
            int(r[1]),
        ): r
        for r in yearly_model_rows
    }

    years = sorted(
        {
            int(r[1])
            for r in daily_rows
        }
    )

    rows = []

    for year in years:
        for h in FUTURE_HORIZONS:
            subset = [
                r
                for r in daily_rows
                if r[1] == year
                and r[2] == h
            ]

            if not subset:
                continue

            ridge = np.asarray(
                [
                    np.nan
                    if r[5] is None
                    else float(r[5])
                    for r in subset
                ],
                dtype=np.float64,
            )

            glob = np.asarray(
                [
                    np.nan
                    if r[7] is None
                    else float(r[7])
                    for r in subset
                ],
                dtype=np.float64,
            )

            spear = np.asarray(
                [
                    np.nan
                    if r[8] is None
                    else float(r[8])
                    for r in subset
                ],
                dtype=np.float64,
            )

            pair = (
                np.isfinite(ridge)
                & np.isfinite(glob)
            )

            diff = (
                ridge[pair]
                - glob[pair]
            )

            ridge_mean, ridge_t, ridge_p = hac_mean_test(
                ridge,
                lag=max(
                    1,
                    h - 1,
                ),
            )

            diff_mean, diff_t, diff_p = hac_mean_test(
                diff,
                lag=max(
                    1,
                    h - 1,
                ),
            )

            global_mean = (
                float(
                    np.nanmean(
                        glob
                    )
                )
                if np.isfinite(
                    glob
                ).any()
                else np.nan
            )

            mean_spear = (
                float(
                    np.nanmean(
                        spear
                    )
                )
                if np.isfinite(
                    spear
                ).any()
                else np.nan
            )

            mr = model_lookup[
                (year, h)
            ]

            rows.append(
                (
                    year,
                    h,
                    mr[2],               # TrainThrough
                    mr[3],               # TrainLabelN
                    mr[4],               # Alpha
                    mr[6],               # GlobalCase
                    int(
                        np.isfinite(
                            ridge
                        ).sum()
                    ),
                    (
                        float(ridge_mean)
                        if np.isfinite(
                            ridge_mean
                        )
                        else None
                    ),
                    (
                        float(ridge_t)
                        if np.isfinite(
                            ridge_t
                        )
                        else None
                    ),
                    (
                        float(ridge_p)
                        if np.isfinite(
                            ridge_p
                        )
                        else None
                    ),
                    (
                        float(global_mean)
                        if np.isfinite(
                            global_mean
                        )
                        else None
                    ),
                    (
                        float(diff_mean)
                        if np.isfinite(
                            diff_mean
                        )
                        else None
                    ),
                    (
                        float(diff_t)
                        if np.isfinite(
                            diff_t
                        )
                        else None
                    ),
                    (
                        float(diff_p)
                        if np.isfinite(
                            diff_p
                        )
                        else None
                    ),
                    (
                        float(mean_spear)
                        if np.isfinite(
                            mean_spear
                        )
                        else None
                    ),
                )
            )

    return rows

# =============================================================================
# XLSX STREAM WRITER
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
    sheet_specs: list[tuple[str, Path]],
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
# OUTPUT
# =============================================================================

def build_output_annual(
    output_path: Path,
    first_test_year: int,
    last_test_year: int,
    alphas: tuple[float, ...],
    inner_fit_ratio: float,
    yearly_summary_rows,
    daily_rows,
):
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs = []

        # PARAMETERS
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Model",
                "Annual Expanding Walk-Forward Ridge",
            ),
            (
                "First_Test_Year",
                first_test_year,
            ),
            (
                "Last_Test_Year",
                last_test_year,
            ),
            (
                "Yearly_Refresh",
                "alpha + X scaler + Ridge + Global Case",
            ),
            (
                "Within_Year_Update",
                "newly matured past IC added daily",
            ),
            (
                "X_Dimension",
                len(X_COLS),
            ),
            (
                "Alpha_Candidates",
                ",".join(
                    f"{a:g}"
                    for a in alphas
                ),
            ),
            (
                "Inner_Fit_Ratio",
                inner_fit_ratio,
            ),
        ]

        write_table_sheet(
            p,
            ["Item", "Value"],
            parameter_rows,
            [28, 50],
        )

        specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # YEARLY SUMMARY
        p = td / "yearly_summary.xml"

        summary_header = [
            "TestYear",
            "Horizon",
            "TrainThrough",
            "TrainLabelN",
            "Alpha",
            "GlobalCase",
            "ValidDates",
            "Ridge_MeanActualIC",
            "Ridge_HAC_t",
            "Ridge_HAC_p",
            "Global_MeanActualIC",
            "DeltaIC_RidgeMinusGlobal",
            "Delta_HAC_t",
            "Delta_HAC_p",
            "Mean_CaseRankSpearman",
        ]

        write_table_sheet(
            p,
            summary_header,
            yearly_summary_rows,
            [17] * len(
                summary_header
            ),
        )

        specs.append(
            (
                "YEARLY_SUMMARY",
                p,
            )
        )

        # DAILY
        p = td / "daily.xml"

        daily_header = [
            "Date",
            "TestYear",
            "Horizon",
            "RecommendedCase",
            "PredictedIC",
            "ActualIC",
            "GlobalCase",
            "GlobalActualIC",
            "CaseRankSpearman",
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
                12,
                11,
                26,
                16,
                16,
                26,
                18,
                20,
            ],
        )

        specs.append(
            (
                "DAILY",
                p,
            )
        )

        package_xlsx(
            output_path,
            specs,
        )

    with ZipFile(
        output_path
    ) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성 XLSX ZIP 오류: {bad}"
            )


# =============================================================================
# MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--input",
        default=None,
        help="STATE_IC_TRAIN_LIBRARY.xlsx 경로",
    )

    parser.add_argument(
        "--data-dir",
        default=None,
        help="입력파일 자동탐색 상위폴더",
    )

    parser.add_argument(
        "--first-test-year",
        type=int,
        default=2025,
        help="첫 OOS Test 연도. 기본 2025",
    )

    parser.add_argument(
        "--last-test-year",
        type=int,
        default=None,
        help="마지막 Test 연도. 생략 시 데이터 마지막 연도",
    )

    parser.add_argument(
        "--alphas",
        default=None,
        help="alpha 후보. 예: 0.01,0.1,1,10,100",
    )

    parser.add_argument(
        "--inner-fit-ratio",
        type=float,
        default=0.70,
        help="연초 alpha 선택용 과거 Fit 비율. 기본 0.70",
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과 폴더",
    )

    args = parser.parse_args()

    if not (
        0.50
        <= args.inner_fit_ratio
        <= 0.90
    ):
        raise ValueError(
            "--inner-fit-ratio는 0.50~0.90 범위"
        )

    alphas = parse_alpha_grid(
        args.alphas
    )

    try:
        script_dir = Path(
            __file__
        ).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    if args.input:
        input_path = Path(
            args.input
        ).resolve()
    else:
        search_dir = (
            Path(
                args.data_dir
            ).resolve()
            if args.data_dir
            else script_dir
        )

        input_path = find_input_file(
            search_dir
        )

    if not input_path.exists():
        raise FileNotFoundError(
            input_path
        )

    output_dir = (
        Path(
            args.output_dir
        ).resolve()
        if args.output_dir
        else script_dir
        / "RIDGE_ANNUAL_WALKFORWARD"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "RIDGE_ANNUAL_WALKFORWARD_MINIMAL.xlsx"
    )

    print("=" * 100)
    print("ANNUAL EXPANDING WALK-FORWARD RIDGE IC CASE SELECTION")
    print("=" * 100)
    print("INPUT:", input_path)
    print(
        "FIRST TEST YEAR:",
        args.first_test_year,
    )
    print()

    print(
        "[1/6] MARKET_STATE_X 읽기...",
        flush=True,
    )

    all_dates, X = read_market_state_x(
        input_path
    )

    print(
        f"      dates={len(all_dates):,} | "
        f"{yyyymmdd_text(all_dates[0])} ~ "
        f"{yyyymmdd_text(all_dates[-1])}",
        flush=True,
    )

    print(
        "[2/6] CASE_DEFINITION 읽기...",
        flush=True,
    )

    cases = read_case_definition(
        input_path
    )

    print(
        f"      cases={len(cases)}",
        flush=True,
    )

    print(
        "[3/6] CASE_DAILY_IC 읽기...",
        flush=True,
    )

    ic_matrices, rows_read = read_case_daily_ic(
        input_path,
        all_dates,
        cases,
    )

    print(
        f"      rows={rows_read:,}",
        flush=True,
    )

    valid_idx = complete_x_indices(
        X
    )

    data_last_year = max(
        date_year(
            int(all_dates[i])
        )
        for i in valid_idx
    )

    last_test_year = (
        args.last_test_year
        if args.last_test_year is not None
        else data_last_year
    )

    if (
        args.first_test_year
        > last_test_year
    ):
        raise ValueError(
            "first-test-year가 last-test-year보다 큽니다."
        )

    daily_rows, yearly_model_rows = annual_expanding_walkforward(
        all_dates=all_dates,
        X=X,
        valid_idx=valid_idx,
        ic_matrices=ic_matrices,
        cases=cases,
        first_test_year=args.first_test_year,
        last_test_year=last_test_year,
        alphas=alphas,
        inner_fit_ratio=args.inner_fit_ratio,
    )

    print(
        "[5/6] 연도별 결과 요약...",
        flush=True,
    )

    yearly_summary_rows = build_yearly_summary(
        daily_rows,
        yearly_model_rows,
    )

    print(
        "[6/6] 최소 결과 Excel 저장...",
        flush=True,
    )

    build_output_annual(
        output_path=output_path,
        first_test_year=args.first_test_year,
        last_test_year=last_test_year,
        alphas=alphas,
        inner_fit_ratio=args.inner_fit_ratio,
        yearly_summary_rows=yearly_summary_rows,
        daily_rows=daily_rows,
    )

    print()
    print("=" * 100)
    print("DONE")
    print("=" * 100)
    print("OUTPUT:", output_path)
    print()
    print("기본 구조:")
    print("  2025 Test <- 2024년까지 과거")
    print("  2026 Test <- 2025년까지 새로 포함")
    print()
    print("연도 시작마다:")
    print("  alpha / X 표준화 / Ridge / Global Case 전부 재계산")
    print()
    print("연도 안에서는:")
    print("  새로 maturity가 끝난 과거 Daily IC를 매일 추가")
    print()
    print("먼저 YEARLY_SUMMARY만 보면 됩니다.")


if __name__ == "__main__":
    main()
