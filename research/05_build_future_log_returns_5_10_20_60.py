# -*- coding: utf-8 -*-
"""
KOSPI200 미래 실제 누적 로그수익률 5/10/20/60 생성
=================================================

[입력 기준]
반드시 아래 Full History + membership 보정 완료본을 사용한다.

    KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx

파일명 뒤에 (1), (2) 등이 붙어 있어도 자동 탐색한다.
QC 파일은 계산 입력으로 사용하지 않는다.

[미래수익률 정의]
기준일 T에서:

    FWD_LOG_5(T)  = ln(P[T+5]  / P[T])
    FWD_LOG_10(T) = ln(P[T+10] / P[T])
    FWD_LOG_20(T) = ln(P[T+20] / P[T])
    FWD_LOG_60(T) = ln(P[T+60] / P[T])

이는 각각 T+1부터 T+h까지의 일간 로그수익률 합과 같다.

[중요]
- PCA 제거하지 않음
- Z-score 표준화하지 않음
- 실제 수정주가 기반 로그수익률
- T일의 Common252 eligible 종목만 결과에 남김
- T일 수익률 자체는 미래수익률에 포함하지 않음
- T+h의 실제 수정주가가 없으면 NaN(빈칸)
- 아직 단기반전 신호 / Rank / IC는 계산하지 않음

[Common252 기준]
날짜 T에서:
1) T일 KOSPI200 편입 = 1
2) T일 로그수익률 존재
3) T-252 ~ T-1의 로그수익률 252개 모두 존재
4) PCA60 / PCA120 / PCA252 세 window에서 모두 분산 > 0

=> 기존 PCA60/120/252 비교에서 사용한 동일 universe와 같은 원칙.

[출력]
FUTURE_LOG_RETURNS_5_10_20_60/
    FUTURE_LOG_RETURNS_5_10_20_60.xlsx

시트:
    README
    VALIDATION
    TICKER_MAP
    COMMON252_ELIGIBLE
    FWD_LOG_5
    FWD_LOG_10
    FWD_LOG_20
    FWD_LOG_60

대형 시트는:
    행 = 기준일 T
    열 = ticker

[필수 패키지]
    pip install numpy

[실행]
    python 05_build_future_log_returns_5_10_20_60.py

입력 직접 지정:
    python 05_build_future_log_returns_5_10_20_60.py --input "KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026(2).xlsx"
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
# SETTINGS
# =============================================================================

PCA_WINDOWS = (60, 120, 252)
FWD_HORIZONS = (5, 10, 20, 60)

EPS_VAR = 1e-18
MIN_STOCKS = 30

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# =============================================================================
# INPUT EXCEL READER
# =============================================================================

def normalize_ticker(value) -> str:
    s = str(value).strip()

    if s.lower() in {"", "nan", "none"}:
        return ""

    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]

    if s.isdigit():
        s = s.zfill(6)

    return s.upper()


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


def excel_serial_to_datetime(x) -> datetime:
    return datetime(1899, 12, 30) + timedelta(days=float(x))


def find_input_file(base_dir: Path, arg_input: str | None) -> Path:
    if arg_input:
        p = Path(arg_input)

        if not p.is_absolute():
            p = base_dir / p

        if not p.exists():
            raise FileNotFoundError(f"입력 파일을 찾지 못했습니다: {p}")

        return p

    exact = base_dir / "KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx"

    if exact.exists():
        return exact

    candidates = [
        p for p in base_dir.glob("*FULL_HISTORY*PCA_READY*.xlsx")
        if "QC" not in p.name.upper()
        and "FUTURE_LOG_RETURNS" not in p.name.upper()
    ]

    if not candidates:
        raise FileNotFoundError(
            "FULL_HISTORY_PCA_READY 분석데이터를 찾지 못했습니다.\n"
            "예: KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026(2).xlsx"
        )

    # 같은 파일의 (1), (2) 등 여러 버전이 있으면 최근 수정본 우선
    return max(candidates, key=lambda p: p.stat().st_mtime)


def workbook_sheet_targets(xlsx_path: Path) -> dict[str, str]:
    with ZipFile(xlsx_path) as z:
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
        text = "".join(
            (t.text or "")
            for t in si.iter(f"{{{NS_MAIN}}}t")
        )
        result.append(text)

    return result


def read_sheet_rows(xlsx_path: Path, target: str) -> list[list[object]]:
    with ZipFile(xlsx_path) as z:
        shared = load_shared_strings(z)
        root = ET.fromstring(z.read(target))

        dim_node = root.find(f"{{{NS_MAIN}}}dimension")

        if dim_node is not None:
            dim = dim_node.attrib.get("ref", "A1:A1")
            end = dim.split(":")[-1]
            ncol = col_to_idx(end) + 1
            nrow = int(re.search(r"(\d+)$", end).group(1))
        else:
            ncol = 1
            nrow = 1

        rows = [
            [None] * ncol
            for _ in range(nrow)
        ]

        sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")

        if sheet_data is None:
            return rows

        for row in sheet_data:
            ri = int(row.attrib["r"]) - 1

            for cell in row:
                ci = col_to_idx(cell.attrib["r"])
                typ = cell.attrib.get("t")

                if typ == "inlineStr":
                    t = cell.find(f".//{{{NS_MAIN}}}t")
                    val = t.text if t is not None else ""

                else:
                    v = cell.find(f"{{{NS_MAIN}}}v")

                    if v is None:
                        val = None
                    elif typ == "s":
                        val = shared[int(v.text)]
                    else:
                        val = v.text

                rows[ri][ci] = val

        return rows


def rows_to_numeric_matrix(
    rows: list[list[object]],
) -> tuple[np.ndarray, np.ndarray, list[str]]:
    tickers = [
        normalize_ticker(x)
        for x in rows[0][1:]
    ]

    dates = []

    for row in rows[1:]:
        x = row[0]

        if isinstance(x, (int, float)):
            d = excel_serial_to_datetime(x)
        else:
            s = str(x).strip()

            if re.fullmatch(r"\d+(\.\d+)?", s):
                d = excel_serial_to_datetime(float(s))
            else:
                d = datetime.fromisoformat(s[:10])

        dates.append(int(d.strftime("%Y%m%d")))

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

    return (
        np.asarray(dates, dtype=np.int32),
        arr,
        tickers,
    )


def read_company_map(
    xlsx_path: Path,
    target: str,
) -> dict[str, str]:
    rows = read_sheet_rows(xlsx_path, target)

    if not rows:
        return {}

    header = [
        str(x or "")
        for x in rows[0]
    ]

    ticker_idx = 0
    name_idx = 1

    for i, h in enumerate(header):
        key = re.sub(r"[\s_]+", "", h).lower()

        if "종목코드" in key or "ticker" in key:
            ticker_idx = i

        if "기업명" in key or "종목명" in key or key == "name":
            name_idx = i

    result = {}

    for row in rows[1:]:
        if ticker_idx >= len(row) or row[ticker_idx] in (None, ""):
            continue

        ticker = normalize_ticker(row[ticker_idx])

        name = (
            ""
            if name_idx >= len(row) or row[name_idx] is None
            else str(row[name_idx]).strip()
        )

        result[ticker] = name

    return result


# =============================================================================
# DATA VALIDATION / COMMON252
# =============================================================================

def validate_full_history(
    price: np.ndarray,
    membership: np.ndarray,
):
    finite = np.isfinite(price)

    total = int(finite.sum())

    nonmember = int(
        (
            finite
            & (membership == 0)
        ).sum()
    )

    member_missing = int(
        (
            ~finite
            & (membership == 1)
        ).sum()
    )

    if nonmember == 0:
        raise RuntimeError(
            "\n[STOP] 이 입력은 Full History가 아닌 것으로 보입니다.\n"
            "membership=0 기간의 가격이 0개입니다.\n\n"
            "반드시 다음 계열의 파일을 사용하세요:\n"
            "KOSPI200_분석데이터_FULL_HISTORY_PCA_READY_2008_2026.xlsx\n"
        )

    if member_missing != 0:
        raise RuntimeError(
            f"\n[STOP] 편입 중 수정주가 누락이 {member_missing}개 있습니다.\n"
            "PCA_READY membership 보정본인지 확인하세요."
        )

    return total, nonmember, member_missing


def build_common252(
    returns: np.ndarray,
    membership: np.ndarray,
) -> np.ndarray:
    """
    기존 Common252 PCA와 같은 universe 원칙.
    """
    T, N = returns.shape

    common = np.zeros(
        (T, N),
        dtype=np.uint8,
    )

    finite = np.isfinite(returns)

    filled = np.nan_to_num(
        returns,
        nan=0.0,
    )

    cs_count = np.vstack([
        np.zeros((1, N), dtype=np.int32),
        np.cumsum(
            finite.astype(np.int32),
            axis=0,
            dtype=np.int32,
        ),
    ])

    cs_sum = np.vstack([
        np.zeros((1, N), dtype=np.float64),
        np.cumsum(filled, axis=0),
    ])

    cs_sq = np.vstack([
        np.zeros((1, N), dtype=np.float64),
        np.cumsum(filled * filled, axis=0),
    ])

    for t in range(252, T):
        cnt252 = (
            cs_count[t]
            - cs_count[t - 252]
        )

        ok = (
            (membership[t] == 1)
            & np.isfinite(returns[t])
            & (cnt252 == 252)
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
                cs_sq[t]
                - cs_sq[t - W]
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
# FUTURE RETURNS
# =============================================================================

def calculate_future_returns(
    price: np.ndarray,
    common252: np.ndarray,
) -> dict[int, np.ndarray]:
    """
    실제 미래 누적 로그수익률.

    FWD_LOG_h(T) = ln(P[T+h] / P[T])

    주의:
    - PCA 제거 X
    - 표준화 X
    - 기준일 T Common252 종목만 남김
    """
    T, N = price.shape

    result = {}

    for h in FWD_HORIZONS:
        out = np.full(
            (T, N),
            np.nan,
            dtype=np.float64,
        )

        p_t = price[:-h]
        p_h = price[h:]

        valid = (
            np.isfinite(p_t)
            & np.isfinite(p_h)
            & (p_t > 0)
            & (p_h > 0)
        )

        block = np.full_like(
            p_t,
            np.nan,
        )

        block[valid] = np.log(
            p_h[valid] / p_t[valid]
        )

        out[:-h] = block

        # 신호와 같은 T일 universe에서 IC를 계산하기 위해
        # 기준일 T Common252 밖의 종목은 제거.
        out[common252 == 0] = np.nan

        result[h] = out

    return result


# =============================================================================
# MINIMAL XLSX WRITER (stdlib only)
# =============================================================================

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
<borders count="1">
<border><left/><right/><top/><bottom/><diagonal/></border>
</borders>
<cellStyleXfs count="1">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
</cellStyleXfs>
<cellXfs count="3">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
</cellXfs>
<cellStyles count="1">
<cellStyle name="Normal" xfId="0" builtinId="0"/>
</cellStyles>
</styleSheet>"""


def inline_cell(ref: str, text: str, style: int | None = None) -> str:
    s = f' s="{style}"' if style is not None else ""
    return (
        f'<c r="{ref}" t="inlineStr"{s}>'
        f'<is><t>{escape(str(text))}</t></is>'
        f'</c>'
    )


def number_cell(ref: str, value: float, style: int = 2) -> str:
    return (
        f'<c r="{ref}" s="{style}">'
        f'<v>{float(value):.12g}</v>'
        f'</c>'
    )


def write_matrix_sheet_xml(
    temp_path: Path,
    dates: np.ndarray,
    tickers: list[str],
    matrix: np.ndarray,
    binary: bool = False,
):
    with temp_path.open(
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
            '<pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/>'
            '</sheetView></sheetViews>'
        )

        out.write(
            '<cols>'
            '<col min="1" max="1" width="12" customWidth="1"/>'
            f'<col min="2" max="{len(tickers)+1}" width="10" customWidth="1"/>'
            '</cols>'
        )

        out.write("<sheetData>")

        out.write('<row r="1">')
        out.write(inline_cell("A1", "Date", 1))

        for j, ticker in enumerate(tickers, start=2):
            out.write(
                inline_cell(
                    f"{idx_to_col(j-1)}1",
                    ticker,
                    1,
                )
            )

        out.write("</row>")

        for i, d in enumerate(dates, start=2):
            out.write(f'<row r="{i}">')

            ds = str(int(d))
            date_text = (
                f"{ds[:4]}-{ds[4:6]}-{ds[6:8]}"
            )

            out.write(
                inline_cell(
                    f"A{i}",
                    date_text,
                )
            )

            row = matrix[i - 2]

            if binary:
                # 0도 정보이므로 모든 종목 기록
                for j, v in enumerate(row, start=2):
                    out.write(
                        number_cell(
                            f"{idx_to_col(j-1)}{i}",
                            int(v),
                            0,
                        )
                    )
            else:
                finite_idx = np.flatnonzero(
                    np.isfinite(row)
                )

                for j0 in finite_idx:
                    out.write(
                        number_cell(
                            f"{idx_to_col(int(j0)+1)}{i}",
                            row[j0],
                            2,
                        )
                    )

            out.write("</row>")

        out.write("</sheetData>")

        last_col = idx_to_col(len(tickers))

        out.write(
            f'<autoFilter ref="A1:{last_col}{len(dates)+1}"/>'
        )

        out.write("</worksheet>")


def write_simple_table_xml(
    temp_path: Path,
    rows: list[list[object]],
):
    with temp_path.open(
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

        out.write("<sheetData>")

        for r, row in enumerate(rows, start=1):
            out.write(f'<row r="{r}">')

            for c, value in enumerate(row, start=1):
                ref = f"{idx_to_col(c-1)}{r}"

                if r == 1:
                    out.write(
                        inline_cell(
                            ref,
                            value,
                            1,
                        )
                    )
                elif isinstance(value, (int, float, np.integer, np.floating)):
                    out.write(
                        number_cell(
                            ref,
                            float(value),
                            0,
                        )
                    )
                else:
                    out.write(
                        inline_cell(
                            ref,
                            value,
                        )
                    )

            out.write("</row>")

        out.write("</sheetData>")
        out.write("</worksheet>")


def write_output_xlsx(
    output_path: Path,
    dates: np.ndarray,
    tickers: list[str],
    company_map: dict[str, str],
    membership: np.ndarray,
    common252: np.ndarray,
    future: dict[int, np.ndarray],
    source_name: str,
    validation_stats: dict,
):
    sheet_names = [
        "README",
        "VALIDATION",
        "TICKER_MAP",
        "COMMON252_ELIGIBLE",
        "FWD_LOG_5",
        "FWD_LOG_10",
        "FWD_LOG_20",
        "FWD_LOG_60",
    ]

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        sheet_files = []

        # README
        readme_rows = [
            ["항목", "내용"],
            ["입력파일", source_name],
            ["미래기간", "5 / 10 / 20 / 60 거래일"],
            ["PCA 제거", "하지 않음"],
            ["표준화", "하지 않음"],
            ["수익률", "실제 수정주가 기반 누적 로그수익률"],
            [
                "정의",
                "FWD_LOG_h(T)=ln(P[T+h]/P[T])=T+1~T+h 일간 로그수익률 합",
            ],
            [
                "Universe",
                "기준일 T의 Common252 eligible 종목만 값 유지",
            ],
            [
                "편출 이후",
                "Full History 가격이 계속 있으면 미래수익률도 계속 계산",
            ],
            [
                "신호/IC",
                "아직 단기반전 신호·Rank·IC는 계산하지 않음",
            ],
        ]

        p = td / "sheet1.xml"
        write_simple_table_xml(p, readme_rows)
        sheet_files.append(p)

        # VALIDATION
        common_counts = common252.sum(axis=1)

        val_rows = [
            [
                "항목",
                "값",
            ],
            [
                "source_dates_start",
                int(dates[0]),
            ],
            [
                "source_dates_end",
                int(dates[-1]),
            ],
            [
                "T",
                len(dates),
            ],
            [
                "N",
                len(tickers),
            ],
            [
                "finite_prices",
                validation_stats["finite_prices"],
            ],
            [
                "nonmember_finite_prices",
                validation_stats["nonmember_finite_prices"],
            ],
            [
                "member_missing_prices",
                validation_stats["member_missing_prices"],
            ],
            [
                "common252_valid_dates",
                int(
                    np.sum(
                        common_counts >= MIN_STOCKS
                    )
                ),
            ],
            [
                "common252_medianN",
                int(
                    np.median(
                        common_counts[
                            common_counts >= MIN_STOCKS
                        ]
                    )
                ),
            ],
            [
                "common252_minN",
                int(
                    np.min(
                        common_counts[
                            common_counts >= MIN_STOCKS
                        ]
                    )
                ),
            ],
            [
                "common252_maxN",
                int(
                    np.max(
                        common_counts[
                            common_counts >= MIN_STOCKS
                        ]
                    )
                ),
            ],
        ]

        for h in FWD_HORIZONS:
            val_rows.append([
                f"finite_FWD_LOG_{h}",
                int(np.isfinite(future[h]).sum()),
            ])

        p = td / "sheet2.xml"
        write_simple_table_xml(p, val_rows)
        sheet_files.append(p)

        # TICKER MAP
        ticker_rows = [
            ["Ticker", "Company"]
        ]

        ticker_rows.extend(
            [
                [
                    t,
                    company_map.get(t, ""),
                ]
                for t in tickers
            ]
        )

        p = td / "sheet3.xml"
        write_simple_table_xml(p, ticker_rows)
        sheet_files.append(p)

        # Common252
        p = td / "sheet4.xml"
        write_matrix_sheet_xml(
            p,
            dates,
            tickers,
            common252,
            binary=True,
        )
        sheet_files.append(p)

        # FWD
        for offset, h in enumerate(
            FWD_HORIZONS,
            start=5,
        ):
            p = td / f"sheet{offset}.xml"

            write_matrix_sheet_xml(
                p,
                dates,
                tickers,
                future[h],
                binary=False,
            )

            sheet_files.append(p)

        # Workbook
        sheets_xml = "".join(
            (
                f'<sheet name="{escape(name)}" '
                f'sheetId="{i}" r:id="rId{i}"/>'
            )
            for i, name in enumerate(
                sheet_names,
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

        rel_items = []

        for i in range(
            1,
            len(sheet_names) + 1,
        ):
            rel_items.append(
                f'<Relationship Id="rId{i}" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                f'Target="worksheets/sheet{i}.xml"/>'
            )

        rel_items.append(
            f'<Relationship Id="rId{len(sheet_names)+1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
            'Target="styles.xml"/>'
        )

        workbook_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(rel_items)
            + '</Relationships>'
        )

        overrides = [
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
            '<Override PartName="/xl/styles.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
        ]

        for i in range(
            1,
            len(sheet_names) + 1,
        ):
            overrides.append(
                f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            )

        content_types = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" '
            'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides)
            + '</Types>'
        )

        root_rels = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships '
            'xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>'
            '</Relationships>'
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
                root_rels,
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

            for i, p in enumerate(
                sheet_files,
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
        help="FULL_HISTORY_PCA_READY xlsx 파일 경로",
    )

    parser.add_argument(
        "--output-dir",
        default=None,
        help="결과 폴더",
    )

    args = parser.parse_args()

    try:
        base_dir = Path(__file__).resolve().parent
    except NameError:
        base_dir = Path.cwd()

    input_path = find_input_file(
        base_dir,
        args.input,
    )

    output_dir = (
        Path(args.output_dir)
        if args.output_dir
        else base_dir
        / "FUTURE_LOG_RETURNS_5_10_20_60"
    )

    output_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    output_path = (
        output_dir
        / "FUTURE_LOG_RETURNS_5_10_20_60.xlsx"
    )

    print("=" * 96)
    print("KOSPI200 FUTURE LOG RETURNS 5 / 10 / 20 / 60")
    print("=" * 96)
    print("INPUT :", input_path)
    print("OUTPUT:", output_path)
    print()
    print("PCA removed  = False")
    print("Standardized = False")
    print("Signal       = False")
    print()

    targets = workbook_sheet_targets(
        input_path
    )

    required = [
        "수정주가",
        "일간로그수익률",
        "KOSPI200편입여부",
    ]

    for sheet in required:
        if sheet not in targets:
            raise KeyError(
                f"필수 시트 없음: {sheet}"
            )

    print("[1/5] 수정주가 읽기...", flush=True)

    price_rows = read_sheet_rows(
        input_path,
        targets["수정주가"],
    )

    dates, price, tickers = (
        rows_to_numeric_matrix(
            price_rows
        )
    )

    del price_rows

    print("[2/5] 일간로그수익률 읽기...", flush=True)

    return_rows = read_sheet_rows(
        input_path,
        targets["일간로그수익률"],
    )

    dates_r, daily_log, tickers_r = (
        rows_to_numeric_matrix(
            return_rows
        )
    )

    del return_rows

    print("[3/5] Membership 읽기...", flush=True)

    member_rows = read_sheet_rows(
        input_path,
        targets["KOSPI200편입여부"],
    )

    dates_m, member_raw, tickers_m = (
        rows_to_numeric_matrix(
            member_rows
        )
    )

    del member_rows

    if not (
        np.array_equal(dates, dates_r)
        and np.array_equal(dates, dates_m)
    ):
        raise RuntimeError(
            "수정주가/로그수익률/membership 날짜 불일치"
        )

    if not (
        tickers == tickers_r
        and tickers == tickers_m
    ):
        raise RuntimeError(
            "수정주가/로그수익률/membership ticker 불일치"
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
        company_map = read_company_map(
            input_path,
            targets["기업별편입이력"],
        )

    finite_price = np.isfinite(price)

    stats = {
        "finite_prices": int(
            finite_price.sum()
        ),
        "nonmember_finite_prices": int(
            (
                finite_price
                & (membership == 0)
            ).sum()
        ),
        "member_missing_prices": int(
            (
                ~finite_price
                & (membership == 1)
            ).sum()
        ),
    }

    validate_full_history(
        price,
        membership,
    )

    print(
        "Full History OK | "
        f"finite_prices={stats['finite_prices']:,} | "
        f"nonmember_prices={stats['nonmember_finite_prices']:,} | "
        f"member_missing={stats['member_missing_prices']:,}",
        flush=True,
    )

    print("[4/5] Common252 + 미래수익률 계산...", flush=True)

    common252 = build_common252(
        daily_log,
        membership,
    )

    counts = common252.sum(axis=1)

    valid_dates = counts >= MIN_STOCKS

    if not valid_dates.any():
        raise RuntimeError(
            "Common252 유효 날짜가 없습니다."
        )

    print(
        "COMMON252 exact-match basis | "
        f"valid_dates={int(valid_dates.sum())} "
        f"medianN={int(np.median(counts[valid_dates]))} "
        f"minN={int(np.min(counts[valid_dates]))} "
        f"maxN={int(np.max(counts[valid_dates]))}",
        flush=True,
    )

    future = calculate_future_returns(
        price,
        common252,
    )

    for h in FWD_HORIZONS:
        print(
            f"FWD_LOG_{h} finite="
            f"{int(np.isfinite(future[h]).sum()):,}",
            flush=True,
        )

    print("[5/5] Excel 저장...", flush=True)

    write_output_xlsx(
        output_path,
        dates,
        tickers,
        company_map,
        membership,
        common252,
        future,
        input_path.name,
        stats,
    )

    # 기본 ZIP integrity check
    with ZipFile(output_path) as z:
        bad = z.testzip()

        if bad is not None:
            raise RuntimeError(
                f"생성된 xlsx ZIP 오류: {bad}"
            )

    print()
    print("=" * 96)
    print("DONE")
    print("=" * 96)
    print("File:", output_path)
    print()
    print("FWD_LOG_h(T) = ln(P[T+h] / P[T])")
    print("PCA removed   = False")
    print("Standardized  = False")
    print("Signal        = not declared")


if __name__ == "__main__":
    main()
