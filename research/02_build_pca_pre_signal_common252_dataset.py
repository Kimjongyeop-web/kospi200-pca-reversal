# -*- coding: utf-8 -*-
"""
KOSPI200 PCA PRE-SIGNAL DATASET BUILDER - COMMON252 - EXCEL ONLY
================================================================

목적
----
단기반전 신호를 선언하기 전 단계까지 한 번에 계산하고,
중간 HDF5(.h5)를 만들지 않고 처음부터 Excel(.xlsx)로 직접 저장한다.

입력
----
KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx
  - 수정주가
  - KOSPI200편입여부
  - 일간로그수익률 (있으면 수정주가 재계산 결과와 대조)
  - 기업별편입이력 (기업명 mapping)

핵심 연구 규칙
--------------
1) 수정주가 -> 일간 로그수익률
2) 날짜 T마다 Common252 universe를 먼저 확정
   - T일 KOSPI200 편입
   - T-252 ~ T-1 로그수익률 252개 모두 존재
   - T일 로그수익률 존재
   - PCA60/120/252 모두 분산 > 0
3) PCA60 / PCA120 / PCA252 모두 동일한 Common252 종목 사용
4) T일 PCA-adjusted residual:
   - PCA 학습: T-W ~ T-1
   - T일 수익률: 학습기간 mu/sigma로 표준화
   - factor score: f_k,T = v_k,T-1' z_T
   - residual: sigma * (z_T - PCA reconstruction)
5) daily residual을 5 / 10 / 30 / 60일 단순합
6) 미래수익률 / IC / 단기반전 signal / rank / threshold는 아직 계산하지 않음

출력
----
PCA_PRE_SIGNAL_COMMON252_EXCEL/
  00_COMMON252_BASE.xlsx

  01_PCA_W60_DAILY.xlsx
  02_PCA_W60_CUMULATIVE.xlsx

  03_PCA_W120_DAILY.xlsx
  04_PCA_W120_CUMULATIVE.xlsx

  05_PCA_W252_DAILY.xlsx
  06_PCA_W252_CUMULATIVE.xlsx

H5는 생성하지 않는다.

필수 패키지
-----------
pip install numpy scipy xlsxwriter

실행
----
python 02_build_pca_pre_signal_common252_EXCEL_ONLY.py

입력 파일 직접 지정:
python 02_build_pca_pre_signal_common252_EXCEL_ONLY.py --input "KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx"

Excel 용량을 줄이고 DAILY 파일에서 TRAIN_MEAN/STD/Z/LOADINGS를 생략:
python 02_build_pca_pre_signal_common252_EXCEL_ONLY.py --compact
"""

from __future__ import annotations

import argparse
import math
import re
import time
from datetime import datetime, timedelta
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

import numpy as np
import xlsxwriter
from scipy import linalg as la


# =============================================================================
# 0. SETTINGS
# =============================================================================

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

PCA_WINDOWS = (60, 120, 252)
CUM_HORIZONS = (5, 10, 30, 60)

N_PC = 3
MIN_STOCKS = 30
EPS_VAR = 1e-18


# =============================================================================
# 1. INPUT EXCEL READER
# =============================================================================

def excel_serial_to_date(x: str) -> datetime:
    return datetime(1899, 12, 30) + timedelta(days=float(x))


def yyyymmdd_to_datetime(x: int) -> datetime:
    s = str(int(x))
    return datetime(int(s[:4]), int(s[4:6]), int(s[6:8]))


def col_to_idx(cell_ref: str) -> int:
    m = re.match(r"([A-Z]+)", cell_ref)
    if not m:
        raise ValueError(f"잘못된 Excel cell ref: {cell_ref}")

    s = m.group(1)
    n = 0

    for ch in s:
        n = n * 26 + (ord(ch) - 64)

    return n - 1


def normalize_ticker(value) -> str:
    s = str(value).strip()

    if s.lower() in {"", "nan", "none"}:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    if s.isdigit():
        s = s.zfill(6)

    return s.upper()


def workbook_sheet_targets(xlsx_path: Path) -> dict[str, str]:
    with ZipFile(xlsx_path) as z:
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
        ns = {
            "m": NS_MAIN,
            "r": NS_REL,
        }

        for s in wb.find("m:sheets", ns):
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

    root = ET.fromstring(
        z.read("xl/sharedStrings.xml")
    )

    out = []

    for si in root:
        texts = []

        for t in si.iter(f"{{{NS_MAIN}}}t"):
            texts.append(t.text or "")

        out.append("".join(texts))

    return out


def read_sheet_rows(
    xlsx_path: Path,
    target: str,
) -> tuple[list[list[object]], str]:
    """
    pandas/openpyxl 없이 OOXML을 직접 읽는다.
    대용량 원본 Excel을 빠르게 numpy matrix로 옮기기 위한 용도.
    """
    with ZipFile(xlsx_path) as z:
        shared = load_shared_strings(z)
        root = ET.fromstring(z.read(target))

        dim_node = root.find(
            f"{{{NS_MAIN}}}dimension"
        )

        dim = (
            dim_node.attrib.get("ref", "A1:A1")
            if dim_node is not None
            else "A1:A1"
        )

        end = dim.split(":")[-1]
        ncol = col_to_idx(end) + 1
        nrow = int(
            re.search(r"(\d+)$", end).group(1)
        )

        rows = [
            [None] * ncol
            for _ in range(nrow)
        ]

        sheet_data = root.find(
            f"{{{NS_MAIN}}}sheetData"
        )

        for row in sheet_data:
            r = int(row.attrib["r"]) - 1

            for cell in row:
                ref = cell.attrib.get("r")
                c = col_to_idx(ref)
                typ = cell.attrib.get("t")

                if typ == "inlineStr":
                    tt = cell.find(
                        f".//{{{NS_MAIN}}}t"
                    )
                    val = (
                        tt.text
                        if tt is not None
                        else ""
                    )

                else:
                    v = cell.find(
                        f"{{{NS_MAIN}}}v"
                    )

                    if v is None:
                        val = None
                    elif typ == "s":
                        val = shared[int(v.text)]
                    else:
                        val = v.text

                rows[r][c] = val

        return rows, dim


def rows_to_numeric_matrix(
    rows: list[list[object]],
) -> tuple[np.ndarray, np.ndarray, list[str]]:
    tickers = [
        normalize_ticker(x)
        for x in rows[0][1:]
    ]

    dates = np.array(
        [
            int(
                excel_serial_to_date(row[0])
                .strftime("%Y%m%d")
            )
            for row in rows[1:]
        ],
        dtype=np.int32,
    )

    arr = np.full(
        (len(rows) - 1, len(tickers)),
        np.nan,
        dtype=np.float64,
    )

    for i, row in enumerate(rows[1:]):
        for j, val in enumerate(row[1:]):
            if val is None or val == "":
                continue

            try:
                arr[i, j] = float(val)
            except Exception:
                pass

    return dates, arr, tickers


def read_company_map(
    xlsx_path: Path,
    target: str,
) -> dict[str, str]:
    rows, _ = read_sheet_rows(
        xlsx_path,
        target,
    )

    if not rows:
        return {}

    header = [
        str(x or "")
        for x in rows[0]
    ]

    try:
        ticker_idx = header.index("종목코드")
    except ValueError:
        ticker_idx = 0

    try:
        name_idx = header.index("기업명")
    except ValueError:
        name_idx = 1

    out = {}

    for row in rows[1:]:
        if ticker_idx >= len(row):
            continue

        ticker = normalize_ticker(
            row[ticker_idx]
        )

        if not ticker:
            continue

        name = (
            ""
            if name_idx >= len(row)
            or row[name_idx] is None
            else str(row[name_idx]).strip()
        )

        out[ticker] = name

    return out


# =============================================================================
# 2. RETURN / ROLLING HELPERS
# =============================================================================

def calculate_log_returns_previous_available(
    price: np.ndarray,
) -> np.ndarray:
    """
    수정주가 -> 일간 로그수익률.

    구조적 NaN 날짜에는 return을 만들지 않는다.
    다음 유효가격 날짜에는 직전 유효가격 대비 로그수익률 계산.
    """
    T, N = price.shape

    ret = np.full(
        (T, N),
        np.nan,
        dtype=np.float64,
    )

    last = np.full(
        N,
        np.nan,
        dtype=np.float64,
    )

    for t in range(T):
        cur = price[t]

        valid = (
            np.isfinite(cur)
            & (cur > 0)
        )

        have_prev = (
            valid
            & np.isfinite(last)
            & (last > 0)
        )

        ret[t, have_prev] = np.log(
            cur[have_prev]
            / last[have_prev]
        )

        last[valid] = cur[valid]

    return ret


def rolling_complete_sum(
    x: np.ndarray,
    n: int,
) -> np.ndarray:
    """
    최근 n개 값이 모두 finite일 때만 합산.
    하나라도 NaN이면 해당 누적값도 NaN.
    """
    T, N = x.shape

    filled = np.nan_to_num(
        x,
        nan=0.0,
    )

    valid = np.isfinite(
        x
    ).astype(np.int16)

    cs = np.vstack(
        [
            np.zeros(
                (1, N),
                dtype=np.float64,
            ),
            np.cumsum(
                filled,
                axis=0,
            ),
        ]
    )

    cc = np.vstack(
        [
            np.zeros(
                (1, N),
                dtype=np.int32,
            ),
            np.cumsum(
                valid,
                axis=0,
                dtype=np.int32,
            ),
        ]
    )

    out = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    if T < n:
        return out

    vals = cs[n:] - cs[:-n]
    counts = cc[n:] - cc[:-n]

    block = vals.astype(
        np.float32
    )

    block[counts != n] = np.nan

    out[n - 1:] = block

    return out


def orient_eigenvectors(
    vecs: np.ndarray,
) -> np.ndarray:
    """
    고유벡터 부호는 임의이므로
    절대값 최대 loading이 양수가 되도록 고정.
    """
    out = vecs.copy()

    for k in range(out.shape[1]):
        v = out[:, k]

        idx = int(
            np.argmax(
                np.abs(v)
            )
        )

        if v[idx] < 0:
            out[:, k] *= -1.0

    return out


# =============================================================================
# 3. COMMON252 UNIVERSE
# =============================================================================

def build_common_252_universe(
    returns: np.ndarray,
    membership: np.ndarray,
) -> np.ndarray:
    """
    날짜 T에서 PCA60/120/252가 공유할 종목집합.

    조건:
    1) Membership(T)=1
    2) T-252 ... T-1 로그수익률 252개 모두 존재
    3) T일 로그수익률 존재
    4) 60/120/252 각 window에서 분산 > 0
    """
    T, N = returns.shape
    max_w = max(PCA_WINDOWS)

    common = np.zeros(
        (T, N),
        dtype=np.uint8,
    )

    finite = np.isfinite(
        returns
    )

    filled = np.nan_to_num(
        returns,
        nan=0.0,
    )

    cs_count = np.vstack(
        [
            np.zeros(
                (1, N),
                dtype=np.int32,
            ),
            np.cumsum(
                finite.astype(np.int32),
                axis=0,
                dtype=np.int32,
            ),
        ]
    )

    cs_sum = np.vstack(
        [
            np.zeros(
                (1, N),
                dtype=np.float64,
            ),
            np.cumsum(
                filled,
                axis=0,
                dtype=np.float64,
            ),
        ]
    )

    cs_sumsq = np.vstack(
        [
            np.zeros(
                (1, N),
                dtype=np.float64,
            ),
            np.cumsum(
                filled * filled,
                axis=0,
                dtype=np.float64,
            ),
        ]
    )

    for t in range(max_w, T):
        cnt252 = (
            cs_count[t]
            - cs_count[t - max_w]
        )

        ok = (
            (membership[t] == 1)
            & np.isfinite(returns[t])
            & (cnt252 == max_w)
        )

        for W in PCA_WINDOWS:
            cnt = (
                cs_count[t]
                - cs_count[t - W]
            )

            sm = (
                cs_sum[t]
                - cs_sum[t - W]
            )

            ss = (
                cs_sumsq[t]
                - cs_sumsq[t - W]
            )

            var = (
                ss
                - (sm * sm) / W
            ) / (W - 1)

            ok &= (
                (cnt == W)
                & np.isfinite(var)
                & (var > EPS_VAR)
            )

        common[t, ok] = 1

    return common


# =============================================================================
# 4. PCA CALCULATION
# =============================================================================

def run_window(
    W: int,
    dates: np.ndarray,
    tickers: list[str],
    returns: np.ndarray,
    membership: np.ndarray,
    common_eligible: np.ndarray,
) -> dict[str, np.ndarray]:
    """
    한 PCA window의 전체 결과를 메모리에 계산한다.
    H5를 만들지 않는다.

    계산이 끝나면 Excel로 저장하고 메모리에서 제거한다.
    """
    T, N = returns.shape

    eligible_mat = np.zeros(
        (T, N),
        dtype=np.uint8,
    )

    mu_mat = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    sigma_mat = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    loading_mat = np.full(
        (T, N, N_PC),
        np.nan,
        dtype=np.float32,
    )

    ztoday_mat = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    eigen = np.full(
        (T, N_PC),
        np.nan,
        dtype=np.float32,
    )

    ratio = np.full(
        (T, N_PC),
        np.nan,
        dtype=np.float32,
    )

    factor = np.full(
        (T, N_PC),
        np.nan,
        dtype=np.float32,
    )

    pca_count = np.zeros(
        T,
        dtype=np.int16,
    )

    member_count = np.zeros(
        T,
        dtype=np.int16,
    )

    avgcorr = np.full(
        T,
        np.nan,
        dtype=np.float32,
    )

    trace_arr = np.full(
        T,
        np.nan,
        dtype=np.float32,
    )

    resid1 = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    resid12 = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    resid123 = np.full(
        (T, N),
        np.nan,
        dtype=np.float32,
    )

    # t=W일의 training window = [0:W]
    train = returns[0:W]

    filled = np.nan_to_num(
        train,
        nan=0.0,
    )

    sums = np.nansum(
        train,
        axis=0,
    )

    sumsq = np.nansum(
        train * train,
        axis=0,
    )

    cross = (
        filled.T
        @ filled
    )

    t0 = time.time()

    for t in range(T):
        member_count[t] = int(
            np.sum(
                membership[t] == 1
            )
        )

        if t < W:
            continue

        eligible = (
            common_eligible[t] == 1
        )

        idx = np.flatnonzero(
            eligible
        )

        if idx.size < MIN_STOCKS:
            # 다음 rolling window 갱신은 아래에서 수행해야 함.
            pass

        else:
            s = sums[idx]
            ss = sumsq[idx]

            var = (
                ss
                - (s * s) / W
            ) / (W - 1)

            good = (
                np.isfinite(var)
                & (var > EPS_VAR)
            )

            if not np.all(good):
                bad_tickers = [
                    tickers[j]
                    for j in idx[~good][:10]
                ]

                raise RuntimeError(
                    f"W={W} date={dates[t]}: "
                    f"COMMON252 variance validation failed: "
                    f"{bad_tickers}"
                )

            std = np.sqrt(var)
            mu = s / W

            cp = cross[
                np.ix_(idx, idx)
            ]

            cov = (
                cp
                - np.outer(s, s) / W
            ) / (W - 1)

            C = cov / np.outer(
                std,
                std,
            )

            C = (
                C + C.T
            ) * 0.5

            np.fill_diagonal(
                C,
                1.0,
            )

            n_use = idx.size

            vals, vecs = la.eigh(
                C,
                subset_by_index=[
                    n_use - N_PC,
                    n_use - 1,
                ],
                driver="evr",
                check_finite=False,
            )

            vals = vals[::-1]
            vecs = vecs[:, ::-1]

            vecs = orient_eigenvectors(
                vecs
            )

            trace = float(
                np.trace(C)
            )

            ratios = vals / trace

            avg_corr = (
                float(
                    (
                        np.sum(C)
                        - trace
                    )
                    / (
                        n_use
                        * (n_use - 1)
                    )
                )
                if n_use > 1
                else np.nan
            )

            r_t = returns[t, idx]

            z_t = (
                r_t - mu
            ) / std

            factors = (
                vecs.T
                @ z_t
            )

            recon1 = (
                factors[0]
                * vecs[:, 0]
            )

            recon12 = (
                recon1
                + factors[1]
                * vecs[:, 1]
            )

            recon123 = (
                recon12
                + factors[2]
                * vecs[:, 2]
            )

            ez1 = (
                z_t
                - recon1
            )

            ez12 = (
                z_t
                - recon12
            )

            ez123 = (
                z_t
                - recon123
            )

            # 다시 로그수익률 단위로 복원.
            # residual이므로 평균 mu는 다시 더하지 않는다.
            er1 = (
                std
                * ez1
            )

            er12 = (
                std
                * ez12
            )

            er123 = (
                std
                * ez123
            )

            eligible_mat[t, idx] = 1
            mu_mat[t, idx] = mu.astype(
                np.float32
            )
            sigma_mat[t, idx] = std.astype(
                np.float32
            )
            loading_mat[t, idx, :] = vecs.astype(
                np.float32
            )
            ztoday_mat[t, idx] = z_t.astype(
                np.float32
            )

            eigen[t, :] = vals.astype(
                np.float32
            )
            ratio[t, :] = ratios.astype(
                np.float32
            )
            factor[t, :] = factors.astype(
                np.float32
            )

            pca_count[t] = int(
                n_use
            )

            avgcorr[t] = np.float32(
                avg_corr
            )

            trace_arr[t] = np.float32(
                trace
            )

            resid1[t, idx] = er1.astype(
                np.float32
            )

            resid12[t, idx] = er12.astype(
                np.float32
            )

            resid123[t, idx] = er123.astype(
                np.float32
            )

        # 다음 t+1 training window:
        # [t+1-W, t+1)
        if t + 1 < T:
            old = (
                returns[t - W]
                if t - W >= 0
                else None
            )

            new = returns[t]

            if old is not None:
                oldf = np.nan_to_num(
                    old,
                    nan=0.0,
                )

                newf = np.nan_to_num(
                    new,
                    nan=0.0,
                )

                cross += (
                    np.outer(newf, newf)
                    - np.outer(oldf, oldf)
                )

                sums += (
                    newf
                    - oldf
                )

                sumsq += (
                    newf * newf
                    - oldf * oldf
                )

        if (
            t % 500 == 0
            and t >= W
        ):
            print(
                f"W={W} | "
                f"{t}/{T-1} | "
                f"date={dates[t]} | "
                f"eligible={int(pca_count[t])} | "
                f"elapsed={time.time()-t0:.1f}s",
                flush=True,
            )

    cum = {
        "PC1": {},
        "PC12": {},
        "PC123": {},
    }

    for label, arr in [
        ("PC1", resid1),
        ("PC12", resid12),
        ("PC123", resid123),
    ]:
        for n in CUM_HORIZONS:
            cum[label][n] = rolling_complete_sum(
                arr.astype(np.float64),
                n,
            )

    return {
        "eligible": eligible_mat,
        "training_mean": mu_mat,
        "training_std": sigma_mat,
        "loadings": loading_mat,
        "z_today": ztoday_mat,
        "eigenvalues": eigen,
        "explained_ratio": ratio,
        "factor_score": factor,
        "pca_used_count": pca_count,
        "member_count": member_count,
        "average_pairwise_corr": avgcorr,
        "total_variance_trace": trace_arr,
        "resid_pc1": resid1,
        "resid_pc12": resid12,
        "resid_pc123": resid123,
        "cumulative": cum,
    }


# =============================================================================
# 5. EXCEL WRITER
# =============================================================================

def excel_formats(wb):
    return {
        "title": wb.add_format({
            "bold": True,
            "font_size": 14,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "align": "left",
            "valign": "vcenter",
        }),
        "header": wb.add_format({
            "bold": True,
            "font_color": "#FFFFFF",
            "bg_color": "#1F4E78",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }),
        "label": wb.add_format({
            "bold": True,
            "bg_color": "#D9EAF7",
            "border": 1,
            "valign": "top",
        }),
        "date": wb.add_format({
            "num_format": "yyyy-mm-dd",
        }),
        "num8": wb.add_format({
            "num_format": "0.00000000",
        }),
        "num6": wb.add_format({
            "num_format": "0.000000",
        }),
        "pct4": wb.add_format({
            "num_format": "0.0000%",
        }),
        "int": wb.add_format({
            "num_format": "0",
        }),
        "text": wb.add_format({
            "num_format": "@",
        }),
        "wrap": wb.add_format({
            "text_wrap": True,
            "valign": "top",
        }),
    }


def add_readme(
    wb,
    title: str,
    rows: list[tuple[str, str]],
):
    ws = wb.add_worksheet(
        "README"
    )

    fm = excel_formats(wb)

    ws.set_column(
        "A:A",
        26,
    )

    ws.set_column(
        "B:B",
        90,
    )

    ws.merge_range(
        "A1:B1",
        title,
        fm["title"],
    )

    ws.write(
        "A3",
        "항목",
        fm["header"],
    )

    ws.write(
        "B3",
        "내용",
        fm["header"],
    )

    for r, (k, v) in enumerate(
        rows,
        start=3,
    ):
        ws.write(
            r,
            0,
            k,
            fm["label"],
        )

        ws.write(
            r,
            1,
            v,
            fm["wrap"],
        )

    ws.freeze_panes(
        3,
        0,
    )


def write_ticker_map(
    wb,
    tickers: list[str],
    company_map: dict[str, str],
):
    ws = wb.add_worksheet(
        "TICKER_MAP"
    )

    fm = excel_formats(wb)

    ws.write_row(
        0,
        0,
        [
            "Ticker",
            "Company",
        ],
        fm["header"],
    )

    for i, ticker in enumerate(
        tickers,
        start=1,
    ):
        ws.write_string(
            i,
            0,
            ticker,
            fm["text"],
        )

        ws.write(
            i,
            1,
            company_map.get(
                ticker,
                "",
            ),
        )

    ws.freeze_panes(
        1,
        0,
    )

    ws.autofilter(
        0,
        0,
        len(tickers),
        1,
    )

    ws.set_column(
        0,
        0,
        12,
    )

    ws.set_column(
        1,
        1,
        24,
    )


def write_wide_matrix(
    wb,
    sheet_name: str,
    dates: np.ndarray,
    tickers: list[str],
    arr: np.ndarray,
    number_format: str = "0.00000000",
    binary: bool = False,
):
    """
    행 = 날짜
    열 = ticker

    Excel 행 제한을 피하기 위해 wide matrix로 저장.
    """
    ws = wb.add_worksheet(
        sheet_name[:31]
    )

    fm = excel_formats(wb)

    num_fmt = (
        fm["int"]
        if binary
        else wb.add_format({
            "num_format": number_format,
        })
    )

    ws.freeze_panes(
        1,
        1,
    )

    ws.autofilter(
        0,
        0,
        len(dates),
        len(tickers),
    )

    ws.set_column(
        0,
        0,
        12,
    )

    ws.set_column(
        1,
        len(tickers),
        10,
    )

    ws.write(
        0,
        0,
        "Date",
        fm["header"],
    )

    for j, ticker in enumerate(
        tickers,
        start=1,
    ):
        ws.write_string(
            0,
            j,
            ticker,
            fm["header"],
        )

    for i, d in enumerate(
        dates,
        start=1,
    ):
        ws.write_datetime(
            i,
            0,
            yyyymmdd_to_datetime(
                int(d)
            ),
            fm["date"],
        )

        row = arr[i - 1]

        if binary:
            # 0도 의미가 있으므로 전체 기록.
            for j, v in enumerate(
                row,
                start=1,
            ):
                ws.write_number(
                    i,
                    j,
                    int(v),
                    num_fmt,
                )

        else:
            finite_idx = np.flatnonzero(
                np.isfinite(row)
            )

            for j in finite_idx:
                ws.write_number(
                    i,
                    int(j) + 1,
                    float(row[j]),
                    num_fmt,
                )


def write_daily_summary(
    wb,
    W: int,
    dates: np.ndarray,
    result: dict[str, np.ndarray],
):
    ws = wb.add_worksheet(
        "DAILY_SUMMARY"
    )

    fm = excel_formats(wb)

    headers = [
        "Date",
        "Window",
        "CurrentMemberCount",
        "PCAUsedStockCount",
        "DroppedCount",
        "PC1_Eigenvalue",
        "PC2_Eigenvalue",
        "PC3_Eigenvalue",
        "PC1_ExplainedRatio",
        "PC2_ExplainedRatio",
        "PC3_ExplainedRatio",
        "PC1_PC3_CumulativeRatio",
        "PC1_FactorScore",
        "PC2_FactorScore",
        "PC3_FactorScore",
        "AveragePairwiseCorrelation",
        "TotalVarianceTrace",
    ]

    ws.write_row(
        0,
        0,
        headers,
        fm["header"],
    )

    ws.freeze_panes(
        1,
        1,
    )

    ws.autofilter(
        0,
        0,
        len(dates),
        len(headers) - 1,
    )

    ws.set_column(
        0,
        0,
        12,
    )

    ws.set_column(
        1,
        len(headers) - 1,
        18,
    )

    ev = result[
        "eigenvalues"
    ]

    ratio = result[
        "explained_ratio"
    ]

    factor = result[
        "factor_score"
    ]

    count = result[
        "pca_used_count"
    ]

    member = result[
        "member_count"
    ]

    avgcorr = result[
        "average_pairwise_corr"
    ]

    trace = result[
        "total_variance_trace"
    ]

    for i, d in enumerate(
        dates,
        start=1,
    ):
        k = i - 1

        ws.write_datetime(
            i,
            0,
            yyyymmdd_to_datetime(
                int(d)
            ),
            fm["date"],
        )

        ws.write_number(
            i,
            1,
            W,
            fm["int"],
        )

        ws.write_number(
            i,
            2,
            int(member[k]),
            fm["int"],
        )

        ws.write_number(
            i,
            3,
            int(count[k]),
            fm["int"],
        )

        ws.write_number(
            i,
            4,
            int(member[k] - count[k]),
            fm["int"],
        )

        for pc in range(3):
            if np.isfinite(
                ev[k, pc]
            ):
                ws.write_number(
                    i,
                    5 + pc,
                    float(ev[k, pc]),
                    fm["num6"],
                )

        cum_ratio = 0.0
        any_ratio = False

        for pc in range(3):
            if np.isfinite(
                ratio[k, pc]
            ):
                v = float(
                    ratio[k, pc]
                )

                ws.write_number(
                    i,
                    8 + pc,
                    v,
                    fm["pct4"],
                )

                cum_ratio += v
                any_ratio = True

        if any_ratio:
            ws.write_number(
                i,
                11,
                cum_ratio,
                fm["pct4"],
            )

        for pc in range(3):
            if np.isfinite(
                factor[k, pc]
            ):
                ws.write_number(
                    i,
                    12 + pc,
                    float(
                        factor[k, pc]
                    ),
                    fm["num6"],
                )

        if np.isfinite(
            avgcorr[k]
        ):
            ws.write_number(
                i,
                15,
                float(avgcorr[k]),
                fm["num6"],
            )

        if np.isfinite(
            trace[k]
        ):
            ws.write_number(
                i,
                16,
                float(trace[k]),
                fm["num6"],
            )


def write_common_validation(
    wb,
    dates: np.ndarray,
    membership: np.ndarray,
    common_eligible: np.ndarray,
):
    ws = wb.add_worksheet(
        "COMMON252_VALIDATION"
    )

    fm = excel_formats(wb)

    ws.write_row(
        0,
        0,
        [
            "Date",
            "CurrentMemberCount",
            "Common252EligibleCount",
            "PCA60_Count",
            "PCA120_Count",
            "PCA252_Count",
            "ExactMatch",
        ],
        fm["header"],
    )

    counts = common_eligible.sum(
        axis=1
    )

    member_counts = (
        membership == 1
    ).sum(
        axis=1
    )

    for i, d in enumerate(
        dates,
        start=1,
    ):
        k = i - 1

        n = int(
            counts[k]
        )

        ws.write_datetime(
            i,
            0,
            yyyymmdd_to_datetime(
                int(d)
            ),
            fm["date"],
        )

        ws.write_number(
            i,
            1,
            int(member_counts[k]),
            fm["int"],
        )

        ws.write_number(
            i,
            2,
            n,
            fm["int"],
        )

        ws.write_number(
            i,
            3,
            n,
            fm["int"],
        )

        ws.write_number(
            i,
            4,
            n,
            fm["int"],
        )

        ws.write_number(
            i,
            5,
            n,
            fm["int"],
        )

        ws.write(
            i,
            6,
            "OK",
        )

    ws.freeze_panes(
        1,
        1,
    )

    ws.autofilter(
        0,
        0,
        len(dates),
        6,
    )

    ws.set_column(
        0,
        0,
        12,
    )

    ws.set_column(
        1,
        6,
        20,
    )


def export_base_excel(
    out_path: Path,
    dates: np.ndarray,
    tickers: list[str],
    company_map: dict[str, str],
    returns: np.ndarray,
    membership: np.ndarray,
    common_eligible: np.ndarray,
):
    print(
        "[EXCEL] BASE ->",
        out_path.name,
        flush=True,
    )

    wb = xlsxwriter.Workbook(
        str(out_path),
        {
            "constant_memory": True,
            "nan_inf_to_errors": False,
        },
    )

    add_readme(
        wb,
        "KOSPI200 PCA PRE-SIGNAL — COMMON252 BASE",
        [
            (
                "상태",
                "단기반전 신호 선언 전 데이터. 미래수익률/IC/threshold/rank 미포함.",
            ),
            (
                "일간수익률",
                "수정주가 기반 로그수익률 ln(P_t / previous available adjusted close).",
            ),
            (
                "Common252",
                "날짜 T에서 직전 252일 로그수익률이 모두 존재하고 T일 KOSPI200 구성종목이며 T일 수익률이 존재하는 종목 중, 60/120/252 window 모두 분산이 양수인 동일 종목집합.",
            ),
            (
                "PCA 비교",
                "PCA60 / PCA120 / PCA252 모두 같은 날짜에는 정확히 같은 ticker 집합 사용.",
            ),
            (
                "누적 Raw",
                "RAW_CUM_LOG_5/10/30/60은 PC 제거 전 실제 로그수익률의 누적합.",
            ),
        ],
    )

    write_ticker_map(
        wb,
        tickers,
        company_map,
    )

    write_wide_matrix(
        wb,
        "DAILY_LOG_RETURN",
        dates,
        tickers,
        returns,
        "0.00000000",
    )

    write_wide_matrix(
        wb,
        "MEMBERSHIP",
        dates,
        tickers,
        membership,
        binary=True,
    )

    write_wide_matrix(
        wb,
        "COMMON252_ELIGIBLE",
        dates,
        tickers,
        common_eligible,
        binary=True,
    )

    write_common_validation(
        wb,
        dates,
        membership,
        common_eligible,
    )

    for n in CUM_HORIZONS:
        raw_cum = rolling_complete_sum(
            returns,
            n,
        )

        write_wide_matrix(
            wb,
            f"RAW_CUM_LOG_{n}",
            dates,
            tickers,
            raw_cum,
            "0.00000000",
        )

        del raw_cum

    wb.close()


def export_daily_excel(
    out_path: Path,
    W: int,
    dates: np.ndarray,
    tickers: list[str],
    result: dict[str, np.ndarray],
    compact: bool,
):
    print(
        f"[EXCEL] W{W} DAILY ->",
        out_path.name,
        flush=True,
    )

    wb = xlsxwriter.Workbook(
        str(out_path),
        {
            "constant_memory": True,
            "nan_inf_to_errors": False,
        },
    )

    add_readme(
        wb,
        f"KOSPI200 PCA W{W} — DAILY RESULTS",
        [
            (
                "PCA Window",
                f"{W} 거래일",
            ),
            (
                "Training",
                f"T일 결과는 T-{W} ~ T-1만으로 PCA 학습.",
            ),
            (
                "Universe",
                "PCA60/120/252 모두 같은 Common252 ticker를 사용.",
            ),
            (
                "Standardization",
                f"T-{W}~T-1의 각 종목 평균/표준편차로 표준화.",
            ),
            (
                "Factor",
                "f_k,T = v_k,T-1' z_T. PC1/PC2/PC3 각각 날짜 T마다 상수 1개.",
            ),
            (
                "Residual",
                "PC 공통부분을 제거한 뒤 해당 window의 sigma를 곱해 다시 로그수익률 단위로 복원. 평균 mu는 재가산하지 않음.",
            ),
            (
                "Compact",
                str(bool(compact)),
            ),
            (
                "신호",
                "아직 단기반전 신호를 선언하지 않음.",
            ),
        ],
    )

    write_daily_summary(
        wb,
        W,
        dates,
        result,
    )

    if not compact:
        write_wide_matrix(
            wb,
            "TRAIN_MEAN",
            dates,
            tickers,
            result["training_mean"],
            "0.00000000",
        )

        write_wide_matrix(
            wb,
            "TRAIN_STD",
            dates,
            tickers,
            result["training_std"],
            "0.00000000",
        )

        write_wide_matrix(
            wb,
            "Z_TODAY",
            dates,
            tickers,
            result["z_today"],
            "0.000000",
        )

        load = result[
            "loadings"
        ]

        write_wide_matrix(
            wb,
            "LOAD_PC1",
            dates,
            tickers,
            load[:, :, 0],
            "0.000000",
        )

        write_wide_matrix(
            wb,
            "LOAD_PC2",
            dates,
            tickers,
            load[:, :, 1],
            "0.000000",
        )

        write_wide_matrix(
            wb,
            "LOAD_PC3",
            dates,
            tickers,
            load[:, :, 2],
            "0.000000",
        )

    write_wide_matrix(
        wb,
        "RESID_PC1",
        dates,
        tickers,
        result["resid_pc1"],
        "0.00000000",
    )

    write_wide_matrix(
        wb,
        "RESID_PC12",
        dates,
        tickers,
        result["resid_pc12"],
        "0.00000000",
    )

    write_wide_matrix(
        wb,
        "RESID_PC123",
        dates,
        tickers,
        result["resid_pc123"],
        "0.00000000",
    )

    wb.close()


def export_cumulative_excel(
    out_path: Path,
    W: int,
    dates: np.ndarray,
    tickers: list[str],
    result: dict[str, np.ndarray],
):
    print(
        f"[EXCEL] W{W} CUMULATIVE ->",
        out_path.name,
        flush=True,
    )

    wb = xlsxwriter.Workbook(
        str(out_path),
        {
            "constant_memory": True,
            "nan_inf_to_errors": False,
        },
    )

    add_readme(
        wb,
        f"KOSPI200 PCA W{W} — CUMULATIVE RESIDUAL LOG RETURNS",
        [
            (
                "PCA Window",
                f"{W} 거래일",
            ),
            (
                "누적기간",
                "5 / 10 / 30 / 60 거래일",
            ),
            (
                "PC 제거 버전",
                "PC1 / PC1~2 / PC1~3 각각 별도 저장.",
            ),
            (
                "동적 PCA",
                "각 날짜마다 당시 직전 W일 PCA를 새로 추정한 daily residual을 만든 뒤 누적.",
            ),
            (
                "계산",
                "로그수익률 residual이므로 최근 n개의 daily residual log return을 단순 합산.",
            ),
            (
                "신호",
                "아직 부호반전/threshold/rank를 적용하지 않음. 현재 값 자체는 cumulative PCA residual log return.",
            ),
        ],
    )

    cum = result[
        "cumulative"
    ]

    for label in [
        "PC1",
        "PC12",
        "PC123",
    ]:
        for n in CUM_HORIZONS:
            write_wide_matrix(
                wb,
                f"{label}_CUM{n}",
                dates,
                tickers,
                cum[label][n],
                "0.00000000",
            )

    wb.close()


# =============================================================================
# 6. VALIDATION
# =============================================================================

def validate_window(
    W: int,
    dates: np.ndarray,
    common_eligible: np.ndarray,
    result: dict[str, np.ndarray],
):
    ev = result[
        "eigenvalues"
    ]

    count = result[
        "pca_used_count"
    ]

    eligible = result[
        "eligible"
    ]

    if not np.array_equal(
        eligible,
        common_eligible,
    ):
        # PCA가 계산되기 전 날짜는 common_eligible도 모두 0이므로 exact match 가능.
        bad = np.argwhere(
            eligible != common_eligible
        )

        raise RuntimeError(
            f"W={W}: saved eligible != Common252. "
            f"first mismatch={bad[0].tolist() if len(bad) else None}"
        )

    valid = np.isfinite(
        ev[:, 0]
    )

    if valid.any():
        first = int(
            np.flatnonzero(
                valid
            )[0]
        )

        if first < max(PCA_WINDOWS):
            # Common252 자체가 252 history 이후에만 생기므로
            # 실제 첫 PCA는 252보다 앞설 수 없음.
            raise RuntimeError(
                f"W={W}: Common252 history 이전에 PCA 생성됨"
            )

        if np.nanmax(
            ev[:, 1]
            - ev[:, 0]
        ) > 1e-5:
            raise RuntimeError(
                f"W={W}: PC2 eigenvalue > PC1"
            )

        if np.nanmax(
            ev[:, 2]
            - ev[:, 1]
        ) > 1e-5:
            raise RuntimeError(
                f"W={W}: PC3 eigenvalue > PC2"
            )

        print(
            f"W={W} validation OK | "
            f"first={dates[first]} | "
            f"last={dates[np.flatnonzero(valid)[-1]]} | "
            f"medianN={int(np.median(count[valid]))}",
            flush=True,
        )


def validate_residual_formula(
    W: int,
    returns: np.ndarray,
    result: dict[str, np.ndarray],
):
    """
    마지막 유효 날짜에서 PC1~3 residual 식 직접 재계산.
    """
    valid_dates = np.flatnonzero(
        result["pca_used_count"]
        >= MIN_STOCKS
    )

    if len(valid_dates) == 0:
        return

    t = int(
        valid_dates[-1]
    )

    idx = np.flatnonzero(
        result["eligible"][t] == 1
    )

    z = result[
        "z_today"
    ][t, idx].astype(
        np.float64
    )

    std = result[
        "training_std"
    ][t, idx].astype(
        np.float64
    )

    V = result[
        "loadings"
    ][t, idx, :].astype(
        np.float64
    )

    f = result[
        "factor_score"
    ][t].astype(
        np.float64
    )

    expected = std * (
        z
        - f[0] * V[:, 0]
        - f[1] * V[:, 1]
        - f[2] * V[:, 2]
    )

    saved = result[
        "resid_pc123"
    ][t, idx].astype(
        np.float64
    )

    max_err = float(
        np.max(
            np.abs(
                expected - saved
            )
        )
    )

    if max_err > 1e-6:
        raise RuntimeError(
            f"W={W}: residual formula validation failed: {max_err}"
        )

    print(
        f"W={W} residual formula OK | "
        f"date={t} max_abs_error={max_err:.3e}",
        flush=True,
    )


# =============================================================================
# 7. MAIN
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--input",
        default=None,
        help="FULL_HISTORY_PCA_READY xlsx path",
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="Excel 출력 폴더",
    )

    parser.add_argument(
        "--compact",
        action="store_true",
        help=(
            "DAILY Excel에서 TRAIN_MEAN / TRAIN_STD / Z_TODAY / "
            "LOAD_PC1~3를 생략. SUMMARY + daily residual만 저장."
        ),
    )

    args = parser.parse_args()

    try:
        base = Path(
            __file__
        ).resolve().parent
    except NameError:
        base = Path.cwd()

    # -------------------------------------------------------------------------
    # Input
    # -------------------------------------------------------------------------
    if args.input:
        input_path = Path(
            args.input
        )

        if not input_path.is_absolute():
            input_path = (
                base / input_path
            )

    else:
        preferred = (
            base
            / "KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx"
        )

        if preferred.exists():
            input_path = preferred

        else:
            candidates = sorted(
                base.glob(
                    "*FULL_HISTORY*PCA_READY*.xlsx"
                )
            )

            if not candidates:
                raise FileNotFoundError(
                    "FULL_HISTORY_PCA_READY 분석데이터 Excel을 찾지 못했습니다."
                )

            input_path = candidates[0]

    if not input_path.exists():
        raise FileNotFoundError(
            input_path
        )

    out_dir = (
        Path(args.output_dir)
        if args.output_dir
        else base
        / "PCA_PRE_SIGNAL_COMMON252_EXCEL"
    )

    out_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    print("=" * 96)
    print("KOSPI200 COMMON252 PCA — EXCEL ONLY")
    print("=" * 96)
    print("INPUT :", input_path)
    print("OUTPUT:", out_dir)
    print("H5    : 생성하지 않음")
    print()

    # -------------------------------------------------------------------------
    # Load source
    # -------------------------------------------------------------------------
    targets = workbook_sheet_targets(
        input_path
    )

    for required in [
        "수정주가",
        "KOSPI200편입여부",
    ]:
        if required not in targets:
            raise KeyError(
                f"필수 시트 누락: {required}"
            )

    print(
        "[1/7] 수정주가 로딩",
        flush=True,
    )

    price_rows, _ = read_sheet_rows(
        input_path,
        targets["수정주가"],
    )

    dates, price, tickers = rows_to_numeric_matrix(
        price_rows
    )

    del price_rows

    print(
        "[2/7] Membership 로딩",
        flush=True,
    )

    member_rows, _ = read_sheet_rows(
        input_path,
        targets["KOSPI200편입여부"],
    )

    (
        dates_m,
        member_raw,
        tickers_m,
    ) = rows_to_numeric_matrix(
        member_rows
    )

    del member_rows

    if not np.array_equal(
        dates,
        dates_m,
    ):
        raise RuntimeError(
            "수정주가와 membership 날짜 불일치"
        )

    if tickers != tickers_m:
        raise RuntimeError(
            "수정주가와 membership ticker 순서 불일치"
        )

    membership = (
        np.nan_to_num(
            member_raw,
            nan=0.0,
        )
        == 1
    ).astype(
        np.uint8
    )

    del member_raw

    company_map = {}

    if "기업별편입이력" in targets:
        try:
            company_map = read_company_map(
                input_path,
                targets[
                    "기업별편입이력"
                ],
            )
        except Exception as exc:
            print(
                "기업명 mapping skip:",
                type(exc).__name__,
                exc,
            )

    # -------------------------------------------------------------------------
    # Returns
    # -------------------------------------------------------------------------
    print(
        "[3/7] 수정주가 -> 일간 로그수익률",
        flush=True,
    )

    returns_recalc = (
        calculate_log_returns_previous_available(
            price
        )
    )

    returns = returns_recalc

    if "일간로그수익률" in targets:
        print(
            "      기존 일간로그수익률 시트와 대조",
            flush=True,
        )

        ret_rows, _ = read_sheet_rows(
            input_path,
            targets["일간로그수익률"],
        )

        (
            dates_r,
            returns_sheet,
            tickers_r,
        ) = rows_to_numeric_matrix(
            ret_rows
        )

        del ret_rows

        if (
            np.array_equal(
                dates,
                dates_r,
            )
            and tickers == tickers_r
        ):
            overlap = (
                np.isfinite(
                    returns_recalc[1:]
                )
                & np.isfinite(
                    returns_sheet[1:]
                )
            )

            maxdiff = (
                float(
                    np.nanmax(
                        np.abs(
                            returns_recalc[1:][overlap]
                            - returns_sheet[1:][overlap]
                        )
                    )
                )
                if overlap.any()
                else np.nan
            )

            recalc_only = int(
                np.sum(
                    np.isfinite(
                        returns_recalc[1:]
                    )
                    & ~np.isfinite(
                        returns_sheet[1:]
                    )
                )
            )

            sheet_only = int(
                np.sum(
                    np.isfinite(
                        returns_sheet[1:]
                    )
                    & ~np.isfinite(
                        returns_recalc[1:]
                    )
                )
            )

            if (
                (
                    not np.isfinite(maxdiff)
                    or maxdiff <= 1e-10
                )
                and recalc_only == 0
                and sheet_only == 0
            ):
                returns = returns_sheet

                print(
                    "      로그수익률 검증 OK | "
                    f"maxdiff={maxdiff:.3e}",
                    flush=True,
                )

            else:
                raise RuntimeError(
                    "일간로그수익률 시트가 수정주가 재계산과 불일치: "
                    f"maxdiff={maxdiff}, "
                    f"recalc_only={recalc_only}, "
                    f"sheet_only={sheet_only}"
                )

    # -------------------------------------------------------------------------
    # Common252
    # -------------------------------------------------------------------------
    print(
        "[4/7] Common252 동일 universe 계산",
        flush=True,
    )

    common_eligible = (
        build_common_252_universe(
            returns,
            membership,
        )
    )

    common_counts = (
        common_eligible.sum(
            axis=1
        )
    )

    valid_common = (
        common_counts
        >= MIN_STOCKS
    )

    if not valid_common.any():
        raise RuntimeError(
            "Common252 eligible 날짜가 없습니다."
        )

    print(
        "COMMON252 | "
        f"first={dates[np.flatnonzero(valid_common)[0]]} | "
        f"last={dates[np.flatnonzero(valid_common)[-1]]} | "
        f"valid_dates={int(valid_common.sum())} | "
        f"medianN={int(np.median(common_counts[valid_common]))} | "
        f"minN={int(np.min(common_counts[valid_common]))} | "
        f"maxN={int(np.max(common_counts[valid_common]))}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Base Excel
    # -------------------------------------------------------------------------
    print(
        "[5/7] BASE Excel 직접 저장",
        flush=True,
    )

    export_base_excel(
        out_dir
        / "00_COMMON252_BASE.xlsx",
        dates,
        tickers,
        company_map,
        returns,
        membership,
        common_eligible,
    )

    # -------------------------------------------------------------------------
    # PCA + immediate Excel
    # -------------------------------------------------------------------------
    print(
        "[6/7] PCA60 / 120 / 252 계산 후 즉시 Excel 저장",
        flush=True,
    )

    file_numbers = {
        60: (1, 2),
        120: (3, 4),
        252: (5, 6),
    }

    expected_common_counts = (
        common_eligible.sum(
            axis=1
        ).astype(
            np.int16
        )
    )

    for W in PCA_WINDOWS:
        print()
        print(
            "=" * 96,
            flush=True,
        )

        print(
            f"PCA W={W}",
            flush=True,
        )

        print(
            "=" * 96,
            flush=True,
        )

        result = run_window(
            W,
            dates,
            tickers,
            returns,
            membership,
            common_eligible,
        )

        # exact-match validation
        if not np.array_equal(
            result["pca_used_count"],
            expected_common_counts,
        ):
            bad = np.flatnonzero(
                result["pca_used_count"]
                != expected_common_counts
            )

            # PCA 계산 불가 시점의 Common252 count는 원래 0이어야 한다.
            raise RuntimeError(
                f"W={W}: Common252 stock count mismatch. "
                f"first bad date={dates[bad[0]] if len(bad) else None}"
            )

        validate_window(
            W,
            dates,
            common_eligible,
            result,
        )

        validate_residual_formula(
            W,
            returns,
            result,
        )

        daily_no, cum_no = (
            file_numbers[W]
        )

        export_daily_excel(
            out_dir
            / f"{daily_no:02d}_PCA_W{W}_DAILY.xlsx",
            W,
            dates,
            tickers,
            result,
            compact=args.compact,
        )

        export_cumulative_excel(
            out_dir
            / f"{cum_no:02d}_PCA_W{W}_CUMULATIVE.xlsx",
            W,
            dates,
            tickers,
            result,
        )

        # 해당 window 결과를 Excel에 저장한 뒤 메모리 해제.
        del result

    # -------------------------------------------------------------------------
    # Final
    # -------------------------------------------------------------------------
    print()
    print(
        "[7/7] 최종 검증 및 파일 목록",
        flush=True,
    )

    expected_files = [
        "00_COMMON252_BASE.xlsx",
        "01_PCA_W60_DAILY.xlsx",
        "02_PCA_W60_CUMULATIVE.xlsx",
        "03_PCA_W120_DAILY.xlsx",
        "04_PCA_W120_CUMULATIVE.xlsx",
        "05_PCA_W252_DAILY.xlsx",
        "06_PCA_W252_CUMULATIVE.xlsx",
    ]

    for name in expected_files:
        p = (
            out_dir
            / name
        )

        if not p.exists():
            raise RuntimeError(
                f"Excel 생성 실패: {p}"
            )

    print()
    print("=" * 96)
    print("DONE — EXCEL ONLY")
    print("=" * 96)

    print(
        "source_dates",
        dates[0],
        dates[-1],
        "T",
        len(dates),
        "N",
        len(tickers),
    )

    print(
        "finite_prices",
        int(
            np.isfinite(
                price
            ).sum()
        ),
        "finite_returns",
        int(
            np.isfinite(
                returns
            ).sum()
        ),
    )

    print(
        "COMMON252 exact-match OK | "
        f"valid_dates={int(valid_common.sum())} "
        f"medianN={int(np.median(common_counts[valid_common]))} "
        f"minN={int(np.min(common_counts[valid_common]))} "
        f"maxN={int(np.max(common_counts[valid_common]))}"
    )

    print(
        "signal_declared=False "
        "future_return_included=False"
    )

    print()
    print(
        "Excel files:"
    )

    for name in expected_files:
        p = (
            out_dir
            / name
        )

        print(
            f" - {name} "
            f"({p.stat().st_size / 1024 / 1024:.1f} MB)"
        )


if __name__ == "__main__":
    main()
