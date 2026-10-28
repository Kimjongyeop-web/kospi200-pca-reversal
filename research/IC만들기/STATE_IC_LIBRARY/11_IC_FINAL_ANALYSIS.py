# -*- coding: utf-8 -*-
"""
11_IC_FINAL_ANALYSIS.py
=======================

연구 본체용 최종 IC 통계 정리.

입력
----
STATE_IC_TRAIN_LIBRARY.xlsx

사용 시트
---------
- CASE_DEFINITION
- CASE_DAILY_IC

출력
----
IC_FINAL_ANALYSIS/
    IC_FINAL_ANALYSIS.xlsx

시트는 3개만 생성:
1) PARAMETERS
2) IC_SUMMARY
3) BEST_COMPARISON

핵심 질문
---------
1. Raw / PC1 / PC12 / PC123 중 어떤 신호의 Mean Rank IC가 높은가?
2. 과거 formation n = 5/10/30/60 중 어디가 강한가?
3. 미래 horizon k = 5/10/20/60 중 어디까지 효과가 유지되는가?
4. PCA 제거가 같은 n의 Raw 신호보다 IC를 유의하게 개선하는가?

IC_SUMMARY 열
-------------
CaseID
PCA_Window
PC_Remove
Formation_N
Horizon
ValidDates
MeanIC
MedianIC
PositiveICRatio
HAC_t
HAC_p
DeltaIC_vs_Raw
Delta_HAC_t
Delta_HAC_p

BEST_COMPARISON 열
------------------
Formation_N
Horizon
Raw_MeanIC
Best_PCA_Case
Best_PCA_MeanIC
DeltaIC_vs_Raw
Delta_HAC_t
Delta_HAC_p

통계
----
Daily IC는 formation/future return overlap 때문에 시계열 자기상관이 생길 수 있으므로
Newey-West HAC standard error를 사용한다.

보수적 lag:
    lag = max(Formation_N, Horizon) - 1

Delta IC:
    DeltaIC_t = IC_PCA_t - IC_RAW_same_n_t

즉 PC 제거가 같은 과거 formation n의 Raw reversal보다
실제로 예측력을 개선했는지를 날짜별 차이로 검정한다.

p-value는 양측검정(two-sided), 유의수준 자체를 코드에서 강제하지 않는다.

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


FUTURE_HORIZONS = (5, 10, 20, 60)
FORMATION_HORIZONS = (5, 10, 30, 60)

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
            if "ic_final_analysis" in name:
                continue
            candidates.append(p)

    if not candidates:
        raise FileNotFoundError(
            "STATE_IC_TRAIN_LIBRARY*.xlsx 를 찾지 못했습니다.\n"
            f"검색폴더: {base}"
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


# =============================================================================
# CASE DEFINITION
# =============================================================================

@dataclass(frozen=True)
class CaseDef:
    case_id: str
    pca_window: str
    pc_remove: str
    formation_n: int


def read_case_definition(path: Path) -> list[CaseDef]:
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

    required = [
        "CaseID",
        "PCA_Window",
        "PC_Remove",
        "Formation_N",
    ]

    missing = [
        c for c in required
        if c not in colmap
    ]

    if missing:
        raise KeyError(
            f"CASE_DEFINITION 필수열 누락: {missing}"
        )

    out = []
    seen = set()

    for cells in it:
        case_id = str(
            cells.get(colmap["CaseID"], "") or ""
        ).strip()

        if not case_id or case_id in seen:
            continue

        seen.add(case_id)

        pca_window = str(
            cells.get(colmap["PCA_Window"], "") or ""
        ).strip()

        pc_remove = str(
            cells.get(colmap["PC_Remove"], "") or ""
        ).strip()

        try:
            formation_n = int(
                float(
                    cells.get(colmap["Formation_N"])
                )
            )
        except Exception:
            raise ValueError(
                f"{case_id}: Formation_N 파싱 실패"
            )

        out.append(
            CaseDef(
                case_id=case_id,
                pca_window=pca_window,
                pc_remove=pc_remove,
                formation_n=formation_n,
            )
        )

    if not out:
        raise RuntimeError("CASE_DEFINITION 유효 Case 없음")

    return out


# =============================================================================
# DAILY IC LOAD
# =============================================================================

def read_daily_ic(
    path: Path,
    cases: list[CaseDef],
    start_date: int | None,
    end_date: int | None,
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

    case_ids = [c.case_id for c in cases]
    case_to_idx = {
        c: i
        for i, c in enumerate(case_ids)
    }

    # case -> horizon -> list[(date, ic)]
    records = {
        c: {
            h: []
            for h in FUTURE_HORIZONS
        }
        for c in case_ids
    }

    rows_read = 0
    min_date = None
    max_date = None

    for cells in it:
        raw_date = cells.get(colmap["Date"])
        case_id = str(
            cells.get(colmap["CaseID"], "") or ""
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

        if start_date is not None and d < start_date:
            continue

        if end_date is not None and d > end_date:
            continue

        min_date = d if min_date is None else min(min_date, d)
        max_date = d if max_date is None else max(max_date, d)

        for h in FUTURE_HORIZONS:
            raw = cells.get(
                colmap[f"IC_FWD{h}"]
            )

            if raw in (None, ""):
                continue

            try:
                ic = float(raw)
            except Exception:
                continue

            if np.isfinite(ic):
                records[case_id][h].append(
                    (d, ic)
                )

        rows_read += 1

        if rows_read % 50000 == 0:
            print(
                f"      CASE_DAILY_IC {rows_read:,} rows...",
                flush=True,
            )

    if min_date is None:
        raise RuntimeError("선택 기간에 유효 IC가 없습니다.")

    return records, rows_read, min_date, max_date


# =============================================================================
# HAC
# =============================================================================

def hac_mean_test(
    values: np.ndarray,
    lag: int,
):
    """
    H0: E[x] = 0
    Newey-West HAC, Bartlett weights.
    Asymptotic two-sided normal p-value.
    """
    x = np.asarray(values, dtype=np.float64)
    x = x[np.isfinite(x)]

    T = len(x)

    if T < 5:
        return np.nan, np.nan, np.nan

    mean = float(np.mean(x))
    u = x - mean

    gamma0 = float(
        np.dot(u, u) / T
    )

    lrv = gamma0

    L = min(
        max(0, int(lag)),
        T - 1,
    )

    for ell in range(1, L + 1):
        gamma = float(
            np.dot(
                u[ell:],
                u[:-ell],
            ) / T
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

    se = math.sqrt(lrv / T)
    t = mean / se

    p = math.erfc(
        abs(t) / math.sqrt(2.0)
    )

    return mean, float(t), float(p)


def series_from_records(
    rows: list[tuple[int, float]],
):
    if not rows:
        return (
            np.empty(0, dtype=np.int32),
            np.empty(0, dtype=np.float64),
        )

    rows = sorted(rows, key=lambda x: x[0])

    dates = np.asarray(
        [r[0] for r in rows],
        dtype=np.int32,
    )

    values = np.asarray(
        [r[1] for r in rows],
        dtype=np.float64,
    )

    return dates, values


def aligned_delta(
    pca_rows: list[tuple[int, float]],
    raw_rows: list[tuple[int, float]],
):
    p_dates, p_vals = series_from_records(pca_rows)
    r_dates, r_vals = series_from_records(raw_rows)

    if len(p_dates) == 0 or len(r_dates) == 0:
        return np.empty(0, dtype=np.float64)

    raw_map = {
        int(d): float(v)
        for d, v in zip(r_dates, r_vals)
    }

    delta = []

    for d, v in zip(p_dates, p_vals):
        rv = raw_map.get(int(d))

        if rv is None:
            continue

        if np.isfinite(v) and np.isfinite(rv):
            delta.append(
                float(v - rv)
            )

    return np.asarray(
        delta,
        dtype=np.float64,
    )


# =============================================================================
# ANALYSIS
# =============================================================================

def build_analysis(
    cases: list[CaseDef],
    records,
):
    case_map = {
        c.case_id: c
        for c in cases
    }

    raw_by_n = {
        c.formation_n: c.case_id
        for c in cases
        if c.pc_remove.upper() == "RAW"
    }

    missing_raw = [
        n
        for n in FORMATION_HORIZONS
        if n not in raw_by_n
    ]

    if missing_raw:
        raise RuntimeError(
            f"Raw Case 누락: n={missing_raw}"
        )

    summary_rows = []
    detail_map = {}

    for case in cases:
        for h in FUTURE_HORIZONS:
            _, ic = series_from_records(
                records[
                    case.case_id
                ][h]
            )

            if len(ic) == 0:
                continue

            valid_n = int(
                np.isfinite(ic).sum()
            )

            mean_ic = float(
                np.mean(ic)
            )

            median_ic = float(
                np.median(ic)
            )

            positive_ratio = float(
                np.mean(ic > 0)
            )

            lag = max(
                case.formation_n,
                h,
            ) - 1

            _, hac_t, hac_p = hac_mean_test(
                ic,
                lag=lag,
            )

            delta_mean = None
            delta_t = None
            delta_p = None

            if case.pc_remove.upper() != "RAW":
                raw_case_id = raw_by_n[
                    case.formation_n
                ]

                delta = aligned_delta(
                    records[
                        case.case_id
                    ][h],
                    records[
                        raw_case_id
                    ][h],
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

            row = (
                case.case_id,
                case.pca_window,
                case.pc_remove,
                case.formation_n,
                h,
                valid_n,
                mean_ic,
                median_ic,
                positive_ratio,
                (
                    float(hac_t)
                    if np.isfinite(hac_t)
                    else None
                ),
                (
                    float(hac_p)
                    if np.isfinite(hac_p)
                    else None
                ),
                delta_mean,
                delta_t,
                delta_p,
            )

            summary_rows.append(row)

            detail_map[
                (
                    case.case_id,
                    h,
                )
            ] = {
                "MeanIC": mean_ic,
                "HAC_t": hac_t,
                "HAC_p": hac_p,
                "DeltaIC": delta_mean,
                "Delta_HAC_t": delta_t,
                "Delta_HAC_p": delta_p,
            }

    # -------------------------------------------------------------------------
    # Same n, same h: Raw vs best PCA case
    # -------------------------------------------------------------------------
    best_rows = []

    for n in FORMATION_HORIZONS:
        raw_case = raw_by_n[n]

        for h in FUTURE_HORIZONS:
            raw_info = detail_map.get(
                (
                    raw_case,
                    h,
                )
            )

            if raw_info is None:
                continue

            pca_candidates = [
                c
                for c in cases
                if (
                    c.pc_remove.upper() != "RAW"
                    and c.formation_n == n
                    and (
                        c.case_id,
                        h,
                    )
                    in detail_map
                )
            ]

            if not pca_candidates:
                continue

            best = max(
                pca_candidates,
                key=lambda c: detail_map[
                    (
                        c.case_id,
                        h,
                    )
                ]["MeanIC"],
            )

            binfo = detail_map[
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
                    float(
                        raw_info[
                            "MeanIC"
                        ]
                    ),
                    best.case_id,
                    float(
                        binfo[
                            "MeanIC"
                        ]
                    ),
                    binfo[
                        "DeltaIC"
                    ],
                    binfo[
                        "Delta_HAC_t"
                    ],
                    binfo[
                        "Delta_HAC_p"
                    ],
                )
            )

    summary_rows.sort(
        key=lambda r: (
            r[4],   # horizon
            r[3],   # n
            -r[6],  # mean ic desc
        )
    )

    best_rows.sort(
        key=lambda r: (
            r[1],  # horizon
            r[0],  # n
        )
    )

    return summary_rows, best_rows


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
        / "IC_FINAL_ANALYSIS"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "IC_FINAL_ANALYSIS.xlsx"
    )

    start_date = parse_cli_date(
        args.start
    )

    end_date = parse_cli_date(
        args.end
    )

    print("=" * 100)
    print("FINAL RANK IC ANALYSIS")
    print("=" * 100)
    print("INPUT:", input_path)
    print()

    print(
        "[1/4] CASE_DEFINITION 읽기...",
        flush=True,
    )

    cases = read_case_definition(
        input_path
    )

    print(
        f"      Cases={len(cases)}",
        flush=True,
    )

    print(
        "[2/4] CASE_DAILY_IC 읽기...",
        flush=True,
    )

    (
        records,
        rows_read,
        actual_start,
        actual_end,
    ) = read_daily_ic(
        input_path,
        cases,
        start_date,
        end_date,
    )

    print(
        f"      rows={rows_read:,}",
        flush=True,
    )

    print(
        f"      period={yyyymmdd_text(actual_start)} ~ {yyyymmdd_text(actual_end)}",
        flush=True,
    )

    print(
        "[3/4] Mean IC / HAC / Delta IC vs Raw 계산...",
        flush=True,
    )

    summary_rows, best_rows = build_analysis(
        cases,
        records,
    )

    print(
        "[4/4] Excel 저장...",
        flush=True,
    )

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        specs = []

        # PARAMETERS
        p = td / "parameters.xml"

        parameter_rows = [
            (
                "Input",
                input_path.name,
            ),
            (
                "Start",
                yyyymmdd_text(actual_start),
            ),
            (
                "End",
                yyyymmdd_text(actual_end),
            ),
            (
                "Cases",
                len(cases),
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
                "IC",
                "Daily cross-sectional Spearman Rank IC",
            ),
            (
                "DeltaIC",
                "PCA Case IC - same Formation_N Raw IC",
            ),
            (
                "HAC_Lag",
                "max(Formation_N, Horizon)-1",
            ),
            (
                "p_value",
                "Newey-West HAC, two-sided asymptotic normal",
            ),
        ]

        write_table_sheet(
            p,
            ["Item", "Value"],
            parameter_rows,
            [26, 74],
        )

        specs.append(
            (
                "PARAMETERS",
                p,
            )
        )

        # IC SUMMARY
        p = td / "ic_summary.xml"

        summary_header = [
            "CaseID",
            "PCA_Window",
            "PC_Remove",
            "Formation_N",
            "Horizon",
            "ValidDates",
            "MeanIC",
            "MedianIC",
            "PositiveICRatio",
            "HAC_t",
            "HAC_p",
            "DeltaIC_vs_Raw",
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
                15,
                15,
                18,
                14,
                14,
                18,
                16,
                16,
            ],
        )

        specs.append(
            (
                "IC_SUMMARY",
                p,
            )
        )

        # BEST COMPARISON
        p = td / "best.xml"

        best_header = [
            "Formation_N",
            "Horizon",
            "Raw_Case",
            "Raw_MeanIC",
            "Best_PCA_Case",
            "Best_PCA_MeanIC",
            "DeltaIC_vs_Raw",
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
                16,
                26,
                18,
                18,
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
    print("  1) BEST_COMPARISON")
    print("  2) IC_SUMMARY")
    print()
    print("핵심 열:")
    print("  MeanIC")
    print("  HAC_p")
    print("  DeltaIC_vs_Raw")
    print("  Delta_HAC_p")


if __name__ == "__main__":
    main()
