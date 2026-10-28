# -*- coding: utf-8 -*-
"""
08_ridge_ic_case_selection_MINIMAL.py
=====================================

목적
----
STATE_IC_TRAIN_LIBRARY.xlsx를 이용해

    PCA 시장상태 X 18개
        -> 40개 Case의 미래 IC 예측
        -> 예측 IC가 가장 높은 Case 추천
        -> 실제 Test IC로 사후 검증

을 수행한다.

이번 버전은 연구에 필요한 결과만 남긴다.

출력 Excel 시트
---------------
1) PARAMETERS
   Train/Test 구간과 모델 설정

2) SUMMARY
   Horizon별 Ridge 성과와 Global 고정 Case 비교

3) DAILY
   날짜별 추천 Case / 예측 IC / 실제 IC

사용하지 않는 값
-----------------
- Oracle Case
- 실제 40개 중 순위
- Percentile
- Nearest date
- KNN / LocalMeanIC
- 불필요한 중간 계산표

Ridge 구조
----------
각 Horizon h에서 날짜 T에 40개 Case의 IC를 동시에 예측한다.

    X_T (18개 PCA 상태)
        -> Predicted IC of 40 Cases

추천:
    RecommendedCase_T,h
        = argmax_case PredictedIC_T,case,h

Test에서는 expanding 방식으로,
현재 날짜 전에 이미 완전히 실현된 과거 IC만 계속 추가해 Ridge를 업데이트한다.

Look-ahead 방지:
    FWD h의 과거 IC label은
        signal_date + h <= current_date - 1
    인 경우만 현재 예측에 사용한다.

Alpha 선택:
    Train 내부를 시간순으로 Fit / Validation으로 나누고
    alpha 후보 중 Validation에서
    "예측 IC 최고 Case를 골랐을 때 실제 IC 평균"이 가장 높은 alpha를
    Horizon별로 하나 선택한다.

기본 alpha 후보:
    0.01, 0.1, 1, 10, 100

필수 외부 패키지:
    numpy
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
# SPLIT
# =============================================================================

def build_split(
    all_dates: np.ndarray,
    X: np.ndarray,
    train_start: int | None,
    train_end: int | None,
    test_start: int | None,
    test_end: int | None,
):
    complete_x = np.all(
        np.isfinite(X),
        axis=1,
    )

    valid_idx = np.flatnonzero(
        complete_x
    )

    if len(valid_idx) < 100:
        raise RuntimeError(
            "X 18개가 완전한 날짜가 너무 적습니다."
        )

    valid_dates = all_dates[
        valid_idx
    ]

    tr_start = (
        int(train_start)
        if train_start is not None
        else int(valid_dates[0])
    )

    if train_end is None:
        cut = max(
            1,
            int(
                math.floor(
                    len(valid_idx)
                    * 0.70
                )
            ),
        ) - 1

        tr_end = int(
            valid_dates[cut]
        )

        print(
            "[주의] --train-end 미지정: 완전한 X 날짜 앞 70% Train / 뒤 30% Test",
            flush=True,
        )
    else:
        tr_end = int(train_end)

    train_idx = valid_idx[
        (valid_dates >= tr_start)
        & (valid_dates <= tr_end)
    ]

    if len(train_idx) < 100:
        raise RuntimeError(
            "Train 날짜가 너무 적습니다."
        )

    if test_start is None:
        later = valid_idx[
            valid_dates
            > all_dates[
                train_idx[-1]
            ]
        ]

        if len(later) == 0:
            raise RuntimeError(
                "Train 뒤 Test 날짜가 없습니다."
            )

        te_start = int(
            all_dates[
                later[0]
            ]
        )
    else:
        te_start = int(test_start)

    te_end = (
        int(test_end)
        if test_end is not None
        else int(valid_dates[-1])
    )

    test_idx = valid_idx[
        (valid_dates >= te_start)
        & (valid_dates <= te_end)
    ]

    if len(test_idx) == 0:
        raise RuntimeError(
            "Test 날짜가 없습니다."
        )

    if (
        all_dates[train_idx[-1]]
        >= all_dates[test_idx[0]]
    ):
        raise RuntimeError(
            "Train/Test 기간이 겹칩니다."
        )

    return valid_idx, train_idx, test_idx


# =============================================================================
# RIDGE
# =============================================================================

def make_fixed_train_scaler(
    X: np.ndarray,
    train_idx: np.ndarray,
):
    train_x = X[
        train_idx
    ]

    mu = np.mean(
        train_x,
        axis=0,
    )

    sd = np.std(
        train_x,
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

    Z = (
        X - mu
    ) / sd

    return Z, mu, sd


def design_row(
    zrow: np.ndarray,
) -> np.ndarray:
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
    """
    Z: n x 18, already standardized
    Y: n x C

    intercept 포함 design D=[1,Z].
    intercept에는 penalty를 주지 않는다.
    """
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

    if len(Z) < 30:
        raise RuntimeError(
            f"Ridge 학습표본 부족: {len(Z)}"
        )

    D = np.column_stack(
        (
            np.ones(
                len(Z),
                dtype=np.float64,
            ),
            Z,
        )
    )

    xtx = D.T @ D
    xty = D.T @ Y

    penalty = np.eye(
        D.shape[1],
        dtype=np.float64,
    ) * float(alpha)

    penalty[0, 0] = 0.0

    beta = np.linalg.solve(
        xtx + penalty,
        xty,
    )

    return beta


def ridge_predict_batch(
    beta: np.ndarray,
    zrow: np.ndarray,
):
    d = design_row(
        zrow
    )

    return d @ beta


def select_alpha_per_horizon(
    all_dates: np.ndarray,
    Z: np.ndarray,
    train_idx: np.ndarray,
    ic_matrices: dict[int, np.ndarray],
    alphas: tuple[float, ...],
    inner_fit_ratio: float,
):
    """
    Outer Train 내부 chronological holdout.

    alpha 평가 기준:
      Validation 각 날짜에서
      Predicted IC가 가장 높은 Case를 선택하고,
      그 Case의 실제 IC 평균을 최대화.
    """
    n_train = len(train_idx)

    cut = int(
        math.floor(
            n_train
            * inner_fit_ratio
        )
    )

    cut = max(
        100,
        min(
            cut,
            n_train - 50,
        ),
    )

    inner_fit_idx = train_idx[
        :cut
    ]

    inner_val_idx = train_idx[
        cut:
    ]

    inner_fit_end = int(
        inner_fit_idx[-1]
    )

    outer_train_end = int(
        train_idx[-1]
    )

    selected = {}

    print(
        "[4/7] Horizon별 Ridge alpha 선택...",
        flush=True,
    )

    for h in FUTURE_HORIZONS:
        fit_idx = inner_fit_idx[
            (inner_fit_idx + h)
            <= inner_fit_end
        ]

        val_idx = inner_val_idx[
            (inner_val_idx + h)
            <= outer_train_end
        ]

        if (
            len(fit_idx) < 100
            or len(val_idx) < 20
        ):
            raise RuntimeError(
                f"FWD{h}: inner Train/Validation 표본 부족"
            )

        Y_fit = ic_matrices[h][
            fit_idx
        ]

        best_alpha = None
        best_score = -np.inf

        for alpha in alphas:
            beta = ridge_fit_batch(
                Z[fit_idx],
                Y_fit,
                alpha,
            )

            D_val = np.column_stack(
                (
                    np.ones(
                        len(val_idx),
                        dtype=np.float64,
                    ),
                    Z[val_idx],
                )
            )

            pred = D_val @ beta
            actual = ic_matrices[h][
                val_idx
            ]

            selected_actual = []

            for r in range(
                len(val_idx)
            ):
                pvec = pred[r]
                avec = actual[r]

                if (
                    not np.isfinite(pvec).any()
                    or not np.isfinite(avec).any()
                ):
                    continue

                ci = int(
                    np.nanargmax(pvec)
                )

                a = avec[ci]

                if np.isfinite(a):
                    selected_actual.append(
                        float(a)
                    )

            score = (
                float(
                    np.mean(
                        selected_actual
                    )
                )
                if selected_actual
                else -np.inf
            )

            if score > best_score:
                best_score = score
                best_alpha = float(alpha)

        if best_alpha is None:
            raise RuntimeError(
                f"FWD{h}: alpha 선택 실패"
            )

        selected[h] = best_alpha

        print(
            f"      FWD{h}: alpha={best_alpha:g} | validation selected mean IC={best_score:.6f}",
            flush=True,
        )

    return selected


def compute_global_baseline(
    train_idx: np.ndarray,
    ic_matrices: dict[int, np.ndarray],
    cases: list[str],
):
    train_end = int(
        train_idx[-1]
    )

    result = {}

    for h in FUTURE_HORIZONS:
        idx = train_idx[
            (train_idx + h)
            <= train_end
        ]

        Y = ic_matrices[h][
            idx
        ]

        mean_ic = np.nanmean(
            Y,
            axis=0,
        )

        if not np.isfinite(mean_ic).any():
            raise RuntimeError(
                f"FWD{h}: Global mean IC 계산 실패"
            )

        ci = int(
            np.nanargmax(
                mean_ic
            )
        )

        result[h] = {
            "case_idx": ci,
            "case_id": cases[ci],
            "train_mean_ic": float(
                mean_ic[ci]
            ),
        }

    return result


def run_expanding_ridge(
    all_dates: np.ndarray,
    Z: np.ndarray,
    valid_idx: np.ndarray,
    train_idx: np.ndarray,
    test_idx: np.ndarray,
    ic_matrices: dict[int, np.ndarray],
    cases: list[str],
    selected_alpha: dict[int, float],
    global_info: dict,
):
    """
    Test 날짜마다 expanding Ridge.

    현재 T를 예측할 때:
        s + h <= T - 1
    인 과거 Daily IC만 사용.
    """
    p = 1 + len(X_COLS)
    C = len(cases)

    daily_rows = []

    # horizon별 sufficient statistics를 초기화한다.
    states = {}

    first_test_idx = int(
        test_idx[0]
    )

    for h in FUTURE_HORIZONS:
        alpha = float(
            selected_alpha[h]
        )

        xtx = np.zeros(
            (p, p),
            dtype=np.float64,
        )

        xty = np.zeros(
            (p, C),
            dtype=np.float64,
        )

        added = np.zeros(
            len(all_dates),
            dtype=bool,
        )

        # 첫 Test 날짜 예측 직전까지 maturity가 끝난 데이터만 초기 포함
        eligible = valid_idx[
            (valid_idx + h)
            <= (first_test_idx - 1)
        ]

        n_added = 0

        for s in eligible:
            y = ic_matrices[h][s]

            if not np.all(
                np.isfinite(y)
            ):
                continue

            d = design_row(
                Z[s]
            )

            xtx += np.outer(
                d,
                d,
            )

            xty += np.outer(
                d,
                y,
            )

            added[s] = True
            n_added += 1

        if n_added < 100:
            raise RuntimeError(
                f"FWD{h}: 초기 Ridge 표본 부족 {n_added}"
            )

        states[h] = {
            "alpha": alpha,
            "xtx": xtx,
            "xty": xty,
            "added": added,
            "n": n_added,
        }

    print(
        "[6/7] Test expanding Ridge 예측...",
        flush=True,
    )

    for pos, t in enumerate(
        test_idx
    ):
        t = int(t)

        for h in FUTURE_HORIZONS:
            state = states[h]

            # 현재 T 예측 전에 새로 maturity가 끝난 과거 label 추가
            # 조건: s+h <= T-1
            newest_s = t - 1 - h

            if newest_s >= 0:
                # 완전한 X 날짜이며 아직 미추가인 경우
                if (
                    newest_s < len(all_dates)
                    and np.all(
                        np.isfinite(
                            Z[newest_s]
                        )
                    )
                    and not state["added"][
                        newest_s
                    ]
                ):
                    y = ic_matrices[h][
                        newest_s
                    ]

                    if np.all(
                        np.isfinite(y)
                    ):
                        d = design_row(
                            Z[newest_s]
                        )

                        state["xtx"] += np.outer(
                            d,
                            d,
                        )

                        state["xty"] += np.outer(
                            d,
                            y,
                        )

                        state["added"][
                            newest_s
                        ] = True

                        state["n"] += 1

            penalty = np.eye(
                p,
                dtype=np.float64,
            ) * state["alpha"]

            penalty[0, 0] = 0.0

            beta = np.linalg.solve(
                state["xtx"]
                + penalty,
                state["xty"],
            )

            pred = ridge_predict_batch(
                beta,
                Z[t],
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

            recommended_case = cases[
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

            case_rank_spear = safe_spearman(
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
                    h,
                    recommended_case,
                    predicted_ic,
                    actual_ic,
                    gb["case_id"],
                    global_actual,
                    (
                        float(
                            case_rank_spear
                        )
                        if np.isfinite(
                            case_rank_spear
                        )
                        else None
                    ),
                )
            )

        if (
            (pos + 1) % 250 == 0
            or (pos + 1) == len(test_idx)
        ):
            print(
                f"      {pos + 1:,}/{len(test_idx):,} Test dates",
                flush=True,
            )

    return daily_rows


# =============================================================================
# SUMMARY
# =============================================================================

def build_summary(
    daily_rows,
    selected_alpha,
    global_info,
):
    rows = []

    for h in FUTURE_HORIZONS:
        subset = [
            r for r in daily_rows
            if r[1] == h
        ]

        ridge = np.asarray(
            [
                np.nan
                if r[4] is None
                else float(r[4])
                for r in subset
            ],
            dtype=np.float64,
        )

        glob = np.asarray(
            [
                np.nan
                if r[6] is None
                else float(r[6])
                for r in subset
            ],
            dtype=np.float64,
        )

        spear = np.asarray(
            [
                np.nan
                if r[7] is None
                else float(r[7])
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
            lag=max(1, h - 1),
        )

        diff_mean, diff_t, diff_p = hac_mean_test(
            diff,
            lag=max(1, h - 1),
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

        rows.append(
            (
                h,
                float(
                    selected_alpha[h]
                ),
                global_info[h]["case_id"],
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

def build_output(
    output_path: Path,
    all_dates: np.ndarray,
    train_idx: np.ndarray,
    test_idx: np.ndarray,
    cases: list[str],
    alphas: tuple[float, ...],
    inner_fit_ratio: float,
    summary_rows,
    daily_rows,
):
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs = []

        # PARAMETERS ---------------------------------------------------------
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Train_Start",
                yyyymmdd_text(
                    all_dates[
                        train_idx[0]
                    ]
                ),
            ),
            (
                "Train_End",
                yyyymmdd_text(
                    all_dates[
                        train_idx[-1]
                    ]
                ),
            ),
            (
                "Test_Start",
                yyyymmdd_text(
                    all_dates[
                        test_idx[0]
                    ]
                ),
            ),
            (
                "Test_End",
                yyyymmdd_text(
                    all_dates[
                        test_idx[-1]
                    ]
                ),
            ),
            (
                "Model",
                "Expanding Ridge IC Prediction",
            ),
            (
                "X_Dimension",
                len(X_COLS),
            ),
            (
                "Case_Count",
                len(cases),
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
            [24, 42],
        )

        specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # SUMMARY ------------------------------------------------------------
        p = td / "summary.xml"

        summary_header = [
            "Horizon",
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
            summary_rows,
            [16] * len(
                summary_header
            ),
        )

        specs.append(
            (
                "SUMMARY",
                p,
            )
        )

        # DAILY --------------------------------------------------------------
        p = td / "daily.xml"

        daily_header = [
            "Date",
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
        "--train-start",
        default=None,
        help="Train 시작일 YYYY-MM-DD",
    )

    parser.add_argument(
        "--train-end",
        default=None,
        help="Train 종료일 YYYY-MM-DD. 생략 시 완전한 X 날짜 앞 70%% Train.",
    )

    parser.add_argument(
        "--test-start",
        default=None,
        help="Test 시작일 YYYY-MM-DD. 생략 시 Train 다음 거래일.",
    )

    parser.add_argument(
        "--test-end",
        default=None,
        help="Test 종료일 YYYY-MM-DD",
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
        help="Outer Train 내부 alpha 선택용 Fit 비율. 기본 0.70",
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
            "--inner-fit-ratio는 0.50~0.90 권장 범위"
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
        / "RIDGE_IC_CASE_SELECTION"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "RIDGE_IC_CASE_SELECTION_MINIMAL.xlsx"
    )

    print("=" * 100)
    print("RIDGE: PCA X -> CASE IC PREDICTION -> CASE SELECTION")
    print("=" * 100)
    print("INPUT :", input_path)
    print(
        "ALPHAS:",
        ", ".join(
            f"{a:g}"
            for a in alphas
        ),
    )
    print()

    print(
        "[1/7] MARKET_STATE_X 읽기...",
        flush=True,
    )

    all_dates, X = read_market_state_x(
        input_path
    )

    print(
        f"      dates={len(all_dates):,} | "
        f"{yyyymmdd_text(all_dates[0])} ~ {yyyymmdd_text(all_dates[-1])}",
        flush=True,
    )

    print(
        "[2/7] CASE_DEFINITION 읽기...",
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
        "[3/7] CASE_DAILY_IC 읽기...",
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

    train_start = parse_cli_date(
        args.train_start
    )

    train_end = parse_cli_date(
        args.train_end
    )

    test_start = parse_cli_date(
        args.test_start
    )

    test_end = parse_cli_date(
        args.test_end
    )

    valid_idx, train_idx, test_idx = build_split(
        all_dates,
        X,
        train_start,
        train_end,
        test_start,
        test_end,
    )

    print(
        f"      Train: {yyyymmdd_text(all_dates[train_idx[0]])} ~ "
        f"{yyyymmdd_text(all_dates[train_idx[-1]])} "
        f"({len(train_idx):,} days)",
        flush=True,
    )

    print(
        f"      Test : {yyyymmdd_text(all_dates[test_idx[0]])} ~ "
        f"{yyyymmdd_text(all_dates[test_idx[-1]])} "
        f"({len(test_idx):,} days)",
        flush=True,
    )

    # Test 전체에서는 Outer Train 기준 X scaler 고정
    Z, _, _ = make_fixed_train_scaler(
        X,
        train_idx,
    )

    selected_alpha = select_alpha_per_horizon(
        all_dates=all_dates,
        Z=Z,
        train_idx=train_idx,
        ic_matrices=ic_matrices,
        alphas=alphas,
        inner_fit_ratio=args.inner_fit_ratio,
    )

    print(
        "[5/7] Global fixed baseline 계산...",
        flush=True,
    )

    global_info = compute_global_baseline(
        train_idx,
        ic_matrices,
        cases,
    )

    daily_rows = run_expanding_ridge(
        all_dates=all_dates,
        Z=Z,
        valid_idx=valid_idx,
        train_idx=train_idx,
        test_idx=test_idx,
        ic_matrices=ic_matrices,
        cases=cases,
        selected_alpha=selected_alpha,
        global_info=global_info,
    )

    summary_rows = build_summary(
        daily_rows,
        selected_alpha,
        global_info,
    )

    print(
        "[7/7] 최소 결과 Excel 저장...",
        flush=True,
    )

    build_output(
        output_path=output_path,
        all_dates=all_dates,
        train_idx=train_idx,
        test_idx=test_idx,
        cases=cases,
        alphas=alphas,
        inner_fit_ratio=args.inner_fit_ratio,
        summary_rows=summary_rows,
        daily_rows=daily_rows,
    )

    print()
    print("=" * 100)
    print("DONE")
    print("=" * 100)
    print("OUTPUT:", output_path)
    print()
    print("시트는 3개만 생성:")
    print("  PARAMETERS")
    print("  SUMMARY")
    print("  DAILY")
    print()
    print("가장 먼저 SUMMARY에서 볼 것:")
    print("  Ridge_MeanActualIC")
    print("  DeltaIC_RidgeMinusGlobal")
    print("  Delta_HAC_p")
    print("  Mean_CaseRankSpearman")


if __name__ == "__main__":
    main()
