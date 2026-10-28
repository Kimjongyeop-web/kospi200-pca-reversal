# -*- coding: utf-8 -*-
"""
10_check_top_tail_reversal_signal.py
====================================

목적
----
전체 횡단면 Rank IC가 낮더라도,
Signal 상위 일부 종목(가장 강한 loser)만 놓고 보면
미래수익률을 잘 잡고 있는지 확인한다.

이번 코드는 40개 Case에 대해:

    Signal Top 5%
    Signal Top 10%
    Signal Top 20%

를 각각 뽑고, 미래 5/10/20/60일 수익률에서 다음만 계산한다.

1) MeanTopFwdLog
   Signal 상위 종목들의 평균 미래 로그수익률

2) MeanUniverseFwdLog
   같은 날짜 분석가능 종목 전체의 평균 미래 로그수익률

3) MeanExcessLog
   Top 그룹 평균 - Universe 평균
   -> 시장 전체 상승/하락을 제외하고 상위 Signal 종목이 더 좋았는지

4) Excess_HAC_t / Excess_HAC_p
   MeanExcessLog가 0보다 유의하게 큰지 Newey-West HAC 검정

5) PositiveStockRate
   선택된 Top 종목들 중 미래수익률 > 0 인 종목 비율

6) PrecisionAtPct
   Signal Top q% 종목 중 실제 미래수익률도 Top q%였던 비율

7) PrecisionLift
   PrecisionAtPct / q
   랜덤 선택이면 기대값 약 1.
   예: Top10% Precision=0.25 이면 Lift=2.5

8) TopBeatsUniverseDateRatio
   날짜별로 Top 그룹 평균 미래수익률이
   전체 Universe 평균보다 높았던 날짜 비율

핵심 해석
---------
IC는 200개 전체 순서를 평가하지만,
이 파일은 "우리가 실제로 관심 있는 상위 tail 종목을 잘 골랐는가?"를 본다.

예:
    TopPct = 10
    PrecisionAtPct = 0.25
    PrecisionLift = 2.5

이면 Signal 상위 10% 종목 중 25%가
실제 미래수익률 상위 10%에도 들어갔다는 뜻이며,
랜덤 기대치 10%보다 2.5배 높다.

------------------------------------------------------------------------------
필요 입력 Excel 5개
------------------------------------------------------------------------------

1. 00_COMMON252_BASE.xlsx
   - COMMON252_ELIGIBLE
   - RAW_CUM_LOG_5
   - RAW_CUM_LOG_10
   - RAW_CUM_LOG_30
   - RAW_CUM_LOG_60

2. 02_PCA_W60_CUMULATIVE.xlsx

3. 04_PCA_W120_CUMULATIVE.xlsx

4. 06_PCA_W252_CUMULATIVE.xlsx

5. FUTURE_LOG_RETURNS_5_10_20_60.xlsx
   - FWD_LOG_5
   - FWD_LOG_10
   - FWD_LOG_20
   - FWD_LOG_60

DAILY PCA 파일이나 STATE_IC_TRAIN_LIBRARY.xlsx는 필요 없다.

------------------------------------------------------------------------------
Case
------------------------------------------------------------------------------

Raw:
    RAW_N5 / RAW_N10 / RAW_N30 / RAW_N60 = 4개

PCA:
    W60/W120/W252
    x PC1/PC12/PC123
    x N5/N10/N30/N60
    = 36개

총 40개.

Signal:
    Signal = - cumulative return
즉 Signal이 클수록 PCA-adjusted 과거 loser.

------------------------------------------------------------------------------
출력
------------------------------------------------------------------------------

TAIL_REVERSAL_CHECK/
    TOP_TAIL_REVERSAL_CHECK.xlsx

시트 3개만 생성:

PARAMETERS
    분석 설정

SUMMARY
    전체기간 Case x Horizon x TopPct 결과

YEARLY
    연도별 동일 결과 (안정성 확인)

------------------------------------------------------------------------------
실행
------------------------------------------------------------------------------

기본:
    python 10_check_top_tail_reversal_signal.py

기간 제한:
    python 10_check_top_tail_reversal_signal.py --start 2010-01-01 --end 2026-07-31

Top 비율 변경:
    python 10_check_top_tail_reversal_signal.py --top-pcts 5,10,20

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


# =============================================================================
# BASIC
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
            if "top_tail_reversal_check" in name:
                continue
            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            f"\n필수 입력파일을 찾지 못했습니다: {key}\n"
            f"검색폴더: {base}\n"
        )

    return max(candidates, key=lambda p: p.stat().st_mtime)


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

    raise ValueError(f"date parse failed: {value!r}")


def parse_cli_date(s: str | None) -> int | None:
    if not s:
        return None
    return parse_date_value(s)


def yyyymmdd_text(x: int) -> str:
    s = str(int(x))
    return f"{s[:4]}-{s[4:6]}-{s[6:8]}"


def parse_top_pcts(s: str) -> tuple[float, ...]:
    out = []

    for x in s.split(","):
        p = float(x.strip())

        if not (0 < p < 100):
            raise ValueError("top-pcts는 0~100 사이여야 합니다.")

        out.append(p / 100.0)

    if not out:
        raise ValueError("top-pcts가 비어 있습니다.")

    return tuple(out)


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

    return [
        "".join(
            (t.text or "")
            for t in si.iter(f"{{{NS_MAIN}}}t")
        )
        for si in root
    ]


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


def read_wide_matrix(
    path: Path,
    sheet_name: str,
) -> tuple[np.ndarray, list[str], np.ndarray]:
    it = iter_sheet_rows(path, sheet_name)

    try:
        header = next(it)
    except StopIteration:
        raise RuntimeError(f"{path.name}/{sheet_name}: 빈 시트")

    max_col = max(header)

    tickers = [
        normalize_ticker(header.get(c, ""))
        for c in range(1, max_col + 1)
    ]

    dates = []
    rows = []

    for cells in it:
        if 0 not in cells:
            continue

        try:
            d = parse_date_value(cells[0])
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
        raise RuntimeError(f"{path.name}/{sheet_name}: 유효 데이터 없음")

    return (
        np.asarray(dates, dtype=np.int32),
        tickers,
        np.vstack(rows),
    )


def require_same_dates(a, b, label):
    if not np.array_equal(a, b):
        raise RuntimeError(f"날짜 불일치: {label}")


def require_same_tickers(a, b, label):
    if a != b:
        raise RuntimeError(f"ticker 불일치: {label}")


# =============================================================================
# CASES
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    source_key: str
    sheet_name: str


def make_cases() -> list[CaseDef]:
    cases = []

    for n in FORMATION_HORIZONS:
        cases.append(
            CaseDef(
                case_id=f"RAW_N{n}",
                source_key="base",
                sheet_name=f"RAW_CUM_LOG_{n}",
            )
        )

    for W in PCA_WINDOWS:
        for pc in ("PC1", "PC12", "PC123"):
            for n in FORMATION_HORIZONS:
                cases.append(
                    CaseDef(
                        case_id=f"W{W}_{pc}_N{n}",
                        source_key=f"cum{W}",
                        sheet_name=f"{pc}_CUM{n}",
                    )
                )

    if len(cases) != 40:
        raise RuntimeError(f"Case 수 오류: {len(cases)}")

    return cases


# =============================================================================
# METRICS
# =============================================================================

def hac_mean_test(
    values: np.ndarray,
    lag: int,
):
    x = np.asarray(values, dtype=np.float64)
    x = x[np.isfinite(x)]

    T = len(x)

    if T < 5:
        return np.nan, np.nan, np.nan

    mean = float(np.mean(x))
    u = x - mean

    gamma0 = float(np.dot(u, u) / T)
    lrv = gamma0

    L = min(int(lag), T - 1)

    for ell in range(1, L + 1):
        gamma = float(
            np.dot(
                u[ell:],
                u[:-ell],
            )
            / T
        )

        weight = 1.0 - ell / (L + 1.0)
        lrv += 2.0 * weight * gamma

    if lrv <= EPS:
        return mean, np.nan, np.nan

    se = math.sqrt(lrv / T)
    t = mean / se

    # asymptotic two-sided normal p-value
    p = math.erfc(abs(t) / math.sqrt(2.0))

    return mean, float(t), float(p)


def evaluate_one_date(
    signal: np.ndarray,
    future: np.ndarray,
    eligible: np.ndarray,
    top_pct: float,
):
    valid = (
        eligible
        & np.isfinite(signal)
        & np.isfinite(future)
    )

    idx = np.flatnonzero(valid)
    n = len(idx)

    if n < 30:
        return None

    k = max(
        1,
        int(
            math.ceil(
                top_pct * n
            )
        ),
    )

    sig = signal[idx]
    fwd = future[idx]

    # Signal 큰 순서 = 더 강한 loser
    sig_order = np.argsort(-sig, kind="mergesort")
    top_local = sig_order[:k]

    # 실제 미래수익률 높은 순서
    fwd_order = np.argsort(-fwd, kind="mergesort")
    actual_top_local = fwd_order[:k]

    top_fwd = fwd[top_local]

    top_mean = float(np.mean(top_fwd))
    universe_mean = float(np.mean(fwd))
    excess = top_mean - universe_mean

    positive_stock_n = int(np.sum(top_fwd > 0))
    selected_n = len(top_fwd)

    intersection = int(
        len(
            np.intersect1d(
                top_local,
                actual_top_local,
                assume_unique=False,
            )
        )
    )

    precision = intersection / k

    return {
        "n_universe": n,
        "selected_n": selected_n,
        "top_mean": top_mean,
        "universe_mean": universe_mean,
        "excess": excess,
        "positive_stock_n": positive_stock_n,
        "precision": float(precision),
        "beats_universe": int(top_mean > universe_mean),
    }


def aggregate_metrics(
    daily_records: list[dict],
    top_pct: float,
    horizon: int,
):
    if not daily_records:
        return None

    top_mean = np.asarray(
        [r["top_mean"] for r in daily_records],
        dtype=np.float64,
    )

    universe_mean = np.asarray(
        [r["universe_mean"] for r in daily_records],
        dtype=np.float64,
    )

    excess = np.asarray(
        [r["excess"] for r in daily_records],
        dtype=np.float64,
    )

    precision = np.asarray(
        [r["precision"] for r in daily_records],
        dtype=np.float64,
    )

    total_positive = sum(
        r["positive_stock_n"]
        for r in daily_records
    )

    total_selected = sum(
        r["selected_n"]
        for r in daily_records
    )

    beats_ratio = float(
        np.mean(
            [
                r["beats_universe"]
                for r in daily_records
            ]
        )
    )

    _, hac_t, hac_p = hac_mean_test(
        excess,
        lag=max(1, horizon - 1),
    )

    mean_precision = float(np.mean(precision))
    lift = mean_precision / top_pct

    return {
        "DateCount": len(daily_records),
        "MeanTopFwdLog": float(np.mean(top_mean)),
        "MeanUniverseFwdLog": float(np.mean(universe_mean)),
        "MeanExcessLog": float(np.mean(excess)),
        "Excess_HAC_t": hac_t,
        "Excess_HAC_p": hac_p,
        "PositiveStockRate": (
            total_positive / total_selected
            if total_selected
            else np.nan
        ),
        "PrecisionAtPct": mean_precision,
        "PrecisionLift": float(lift),
        "TopBeatsUniverseDateRatio": beats_ratio,
    }


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


def xml_inline_cell(ref: str, value, header: bool = False) -> str:
    style = ' s="1"' if header else ""

    return (
        f'<c r="{ref}" t="inlineStr"{style}>'
        f'<is><t>{escape(str(value))}</t></is>'
        f'</c>'
    )


def xml_number_cell(ref: str, value) -> str:
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
    with path.open("w", encoding="utf-8", newline="") as out:
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

            for i, w in enumerate(widths, start=1):
                out.write(
                    f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
                )

            out.write("</cols>")

        out.write("<sheetData>")
        out.write('<row r="1">')

        for c, name in enumerate(header, start=1):
            out.write(
                xml_inline_cell(
                    f"{idx_to_col(c-1)}1",
                    name,
                    header=True,
                )
            )

        out.write("</row>")

        last_row = 1

        for last_row, row in enumerate(rows, start=2):
            out.write(f'<row r="{last_row}">')

            for c, value in enumerate(row, start=1):
                if value is None:
                    continue

                if isinstance(value, float) and not math.isfinite(value):
                    continue

                ref = f"{idx_to_col(c-1)}{last_row}"

                if isinstance(
                    value,
                    (int, float, np.integer, np.floating),
                ):
                    out.write(xml_number_cell(ref, value))
                else:
                    out.write(xml_inline_cell(ref, value))

            out.write("</row>")

        out.write("</sheetData>")

        if header:
            last_col = idx_to_col(len(header) - 1)
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
        for i, (name, _) in enumerate(sheet_specs, start=1)
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

    for i, _ in enumerate(sheet_specs, start=1):
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

    for i in range(1, len(sheet_specs) + 1):
        overrides.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )

    content_types = CONTENT_TYPES_TEMPLATE.format(
        sheet_overrides="".join(overrides)
    )

    with ZipFile(
        output_path,
        "w",
        compression=ZIP_DEFLATED,
        compresslevel=6,
    ) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", ROOT_RELS)
        z.writestr("xl/workbook.xml", workbook_xml)
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        z.writestr("xl/styles.xml", STYLES)

        for i, (_, p) in enumerate(sheet_specs, start=1):
            z.write(p, f"xl/worksheets/sheet{i}.xml")


# =============================================================================
# MAIN ANALYSIS
# =============================================================================

def main():
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--data-dir",
        default=None,
        help="5개 입력 Excel이 있는 상위폴더. 생략 시 이 py 폴더 재귀탐색.",
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
        "--top-pcts",
        default="5,10,20",
        help="확인할 Signal 상위 비율. 예: 5,10,20",
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과폴더",
    )

    args = parser.parse_args()

    try:
        script_dir = Path(__file__).resolve().parent
    except NameError:
        script_dir = Path.cwd()

    data_dir = (
        Path(args.data_dir).resolve()
        if args.data_dir
        else script_dir
    )

    output_dir = (
        Path(args.output_dir).resolve()
        if args.output_dir
        else script_dir / "TAIL_REVERSAL_CHECK"
    )

    output_dir.mkdir(parents=True, exist_ok=True)

    output_path = (
        output_dir
        / "TOP_TAIL_REVERSAL_CHECK.xlsx"
    )

    start_date = parse_cli_date(args.start)
    end_date = parse_cli_date(args.end)
    top_pcts = parse_top_pcts(args.top_pcts)

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
    print("TOP-TAIL REVERSAL SIGNAL CHECK")
    print("=" * 100)

    for k, p in file_map.items():
        print(f"{k:<8}: {p}")

    print()

    # -------------------------------------------------------------------------
    # Base universe
    # -------------------------------------------------------------------------
    print("[1/5] Common252 universe 읽기...", flush=True)

    dates, tickers, common = read_wide_matrix(
        file_map["base"],
        "COMMON252_ELIGIBLE",
    )

    common_bool = (
        np.nan_to_num(common, nan=0.0)
        == 1
    )

    T, N = common.shape

    date_mask = np.ones(T, dtype=bool)

    if start_date is not None:
        date_mask &= dates >= start_date

    if end_date is not None:
        date_mask &= dates <= end_date

    analysis_idx = np.flatnonzero(date_mask)

    if len(analysis_idx) == 0:
        raise RuntimeError("분석 날짜가 0개입니다.")

    print(
        f"      dates={len(analysis_idx):,} | "
        f"{yyyymmdd_text(dates[analysis_idx[0]])} ~ "
        f"{yyyymmdd_text(dates[analysis_idx[-1]])}",
        flush=True,
    )

    # -------------------------------------------------------------------------
    # Futures
    # -------------------------------------------------------------------------
    print("[2/5] Future returns 읽기...", flush=True)

    future = {}

    for h in FUTURE_HORIZONS:
        d, t, arr = read_wide_matrix(
            file_map["future"],
            f"FWD_LOG_{h}",
        )

        require_same_dates(dates, d, f"FWD{h}")
        require_same_tickers(tickers, t, f"FWD{h}")

        future[h] = arr

    # -------------------------------------------------------------------------
    # Cases
    # -------------------------------------------------------------------------
    cases = make_cases()

    print(
        f"[3/5] 40개 Case × {len(top_pcts)}개 Top 비율 × 4 horizon 분석...",
        flush=True,
    )

    summary_rows = []
    yearly_rows = []

    years = sorted(
        {
            int(dates[i]) // 10000
            for i in analysis_idx
        }
    )

    for case_no, case in enumerate(cases, start=1):
        source_path = file_map[case.source_key]

        d, t, cumulative = read_wide_matrix(
            source_path,
            case.sheet_name,
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

        # reversal signal: larger = bigger past loser
        signal = -cumulative

        for h in FUTURE_HORIZONS:
            fwd = future[h]

            for pct in top_pcts:
                daily_records = []
                yearly_records = {
                    y: []
                    for y in years
                }

                for ti in analysis_idx:
                    rec = evaluate_one_date(
                        signal=signal[ti],
                        future=fwd[ti],
                        eligible=common_bool[ti],
                        top_pct=pct,
                    )

                    if rec is None:
                        continue

                    daily_records.append(rec)

                    y = int(dates[ti]) // 10000
                    yearly_records[y].append(rec)

                agg = aggregate_metrics(
                    daily_records,
                    pct,
                    h,
                )

                if agg is not None:
                    summary_rows.append(
                        (
                            case.case_id,
                            h,
                            int(round(pct * 100)),
                            agg["DateCount"],
                            agg["MeanTopFwdLog"],
                            agg["MeanUniverseFwdLog"],
                            agg["MeanExcessLog"],
                            agg["Excess_HAC_t"],
                            agg["Excess_HAC_p"],
                            agg["PositiveStockRate"],
                            agg["PrecisionAtPct"],
                            agg["PrecisionLift"],
                            agg["TopBeatsUniverseDateRatio"],
                        )
                    )

                for y in years:
                    yagg = aggregate_metrics(
                        yearly_records[y],
                        pct,
                        h,
                    )

                    if yagg is None:
                        continue

                    yearly_rows.append(
                        (
                            y,
                            case.case_id,
                            h,
                            int(round(pct * 100)),
                            yagg["DateCount"],
                            yagg["MeanTopFwdLog"],
                            yagg["MeanExcessLog"],
                            yagg["PositiveStockRate"],
                            yagg["PrecisionAtPct"],
                            yagg["PrecisionLift"],
                            yagg["TopBeatsUniverseDateRatio"],
                        )
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
    # Output
    # -------------------------------------------------------------------------
    print("[4/5] 결과 Excel 저장...", flush=True)

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs = []

        # PARAMETERS
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Purpose",
                "Check whether reversal Signal top tail predicts future top-return stocks even when full Rank IC is weak",
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
                "TopPct",
                ",".join(
                    str(
                        int(
                            round(
                                p * 100
                            )
                        )
                    )
                    for p in top_pcts
                ),
            ),
            (
                "Cases",
                40,
            ),
            (
                "FutureHorizons",
                "5,10,20,60",
            ),
            (
                "Signal",
                "-CumulativeReturn (higher = stronger past loser)",
            ),
            (
                "PrecisionLift",
                "PrecisionAtPct / TopPct; random expectation = 1",
            ),
        ]

        write_table_sheet(
            p,
            ["Item", "Value"],
            parameter_rows,
            [24, 95],
        )

        specs.append(("PARAMETERS", p))

        # SUMMARY
        p = td / "summary.xml"

        summary_header = [
            "CaseID",
            "Horizon",
            "TopPct",
            "DateCount",
            "MeanTopFwdLog",
            "MeanUniverseFwdLog",
            "MeanExcessLog",
            "Excess_HAC_t",
            "Excess_HAC_p",
            "PositiveStockRate",
            "PrecisionAtPct",
            "PrecisionLift",
            "TopBeatsUniverseDateRatio",
        ]

        write_table_sheet(
            p,
            summary_header,
            summary_rows,
            [24, 12, 12, 14, 18, 20, 18, 16, 16, 18, 18, 16, 24],
        )

        specs.append(("SUMMARY", p))

        # YEARLY
        p = td / "yearly.xml"

        yearly_header = [
            "Year",
            "CaseID",
            "Horizon",
            "TopPct",
            "DateCount",
            "MeanTopFwdLog",
            "MeanExcessLog",
            "PositiveStockRate",
            "PrecisionAtPct",
            "PrecisionLift",
            "TopBeatsUniverseDateRatio",
        ]

        write_table_sheet(
            p,
            yearly_header,
            yearly_rows,
            [12, 24, 12, 12, 14, 18, 18, 18, 18, 16, 24],
        )

        specs.append(("YEARLY", p))

        package_xlsx(
            output_path,
            specs,
        )

    with ZipFile(output_path) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성 XLSX ZIP 오류: {bad}"
            )

    print("[5/5] 완료", flush=True)
    print()
    print("=" * 100)
    print("OUTPUT:", output_path)
    print("=" * 100)
    print()
    print("SUMMARY에서 먼저 볼 열:")
    print("  MeanExcessLog")
    print("  Excess_HAC_p")
    print("  PrecisionAtPct")
    print("  PrecisionLift")
    print("  TopBeatsUniverseDateRatio")


if __name__ == "__main__":
    main()
