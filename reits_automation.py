#!/usr/bin/env python3
"""대한민국 상장 리츠 정보를 자동 수집해 엑셀(xlsx)로 저장합니다.

핵심 목표:
- pykrx가 없거나 네트워크가 제한돼도 최소 결과를 생성
- KRX KIND 크롤링(대체 경로) + 정적 백업 목록으로 빈 결과 방지
- 외부 라이브러리 없이 xlsx 생성 가능
"""

from __future__ import annotations

import argparse
import html
import io
import json
import os
import re
import urllib.parse
import urllib.request
import zipfile
from dataclasses import dataclass
from datetime import datetime
from typing import Optional
from xml.etree import ElementTree as ET

DART_BASE = "https://opendart.fss.or.kr/api"
KIND_DOWNLOAD_URL = "https://kind.krx.co.kr/corpgeneral/corpList.do?method=download&searchType=13"
DEFAULT_REPORT_BEGIN_DATE = "20200101"

# 네트워크가 막힌 환경에서도 최소 결과를 보장하기 위한 백업 목록
STATIC_REITS = [
    ("088980", "맥쿼리인프라"),
    ("330590", "롯데리츠"),
    ("348950", "제이알글로벌리츠"),
    ("357120", "코람코에너지리츠"),
    ("365550", "ESR켄달스퀘어리츠"),
    ("377190", "디앤디플랫폼리츠"),
    ("395400", "SK리츠"),
    ("417310", "코람코더원리츠"),
    ("432320", "KB스타리츠"),
    ("448730", "삼성FN리츠"),
]

ASSET_TYPE_PATTERNS = {
    "오피스": ["오피스", "office"],
    "리테일": ["리테일", "쇼핑", "백화점", "마트", "상업시설", "retail"],
    "물류": ["물류", "logistics", "센터"],
    "호텔/숙박": ["호텔", "숙박", "레지던스"],
    "주거": ["주거", "임대주택", "아파트"],
    "데이터센터": ["데이터센터", "data center"],
    "복합": ["복합", "mixed-use"],
}

LEASE_STRUCTURE_PATTERNS = [
    "마스터리스",
    "책임임대차",
    "고정임대료",
    "변동임대료",
    "매출연동",
    "NNN",
    "트리플넷",
]

VACANCY_REGEXES = [
    re.compile(r"공실률\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%"),
    re.compile(r"vacancy\s*rate\s*[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%", re.I),
]


@dataclass
class ReitRecord:
    ticker: str
    name: str
    asset_type: Optional[str] = None
    lease_structure: Optional[str] = None
    vacancy_rate: Optional[float] = None
    dividend_yield: Optional[float] = None
    source: str = ""
    note: str = ""


def load_env_file(path: str = ".env") -> None:
    if not os.path.exists(path):
        return
    with open(path, "r", encoding="utf-8") as fp:
        for raw in fp:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip("\"'")
            if key and key not in os.environ:
                os.environ[key] = value


def try_import_pykrx_stock():
    try:
        from pykrx import stock  # type: ignore

        return stock
    except Exception:
        return None


def http_get(url: str, params: Optional[dict[str, str]] = None, timeout: int = 12) -> bytes:
    if params:
        query = urllib.parse.urlencode(params)
        sep = "&" if "?" in url else "?"
        url = f"{url}{sep}{query}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def get_listed_reits_from_pykrx(base_date: str) -> list[tuple[str, str]]:
    stock = try_import_pykrx_stock()
    if stock is None:
        return []
    tickers = stock.get_market_ticker_list(base_date, market="ALL")
    rows = []
    for ticker in tickers:
        name = stock.get_market_ticker_name(ticker)
        if "리츠" in name or "인프라" in name:
            rows.append((ticker, name))
    return sorted(rows, key=lambda x: x[1])


def get_listed_reits_from_kind() -> list[tuple[str, str]]:
    raw = http_get(KIND_DOWNLOAD_URL)
    text = None
    for enc in ("euc-kr", "cp949", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        return []

    rows: list[tuple[str, str]] = []
    tr_matches = re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.I | re.S)
    for tr in tr_matches:
        tds = re.findall(r"<td[^>]*>(.*?)</td>", tr, flags=re.I | re.S)
        if len(tds) < 2:
            continue
        clean = [re.sub(r"<.*?>", "", c).strip() for c in tds]
        name = html.unescape(clean[0])
        code_candidates = re.findall(r"\d{6}", " ".join(clean))
        if not code_candidates:
            continue
        code = code_candidates[0]
        if "리츠" in name or "인프라" in name:
            rows.append((code, name))

    uniq = sorted({ticker: name for ticker, name in rows}.items(), key=lambda x: x[1])
    return list(uniq)


def get_listed_reits(base_date: str) -> tuple[list[tuple[str, str]], str]:
    by_pykrx = get_listed_reits_from_pykrx(base_date)
    if by_pykrx:
        return by_pykrx, "pykrx"

    try:
        by_kind = get_listed_reits_from_kind()
        if by_kind:
            return by_kind, "KIND"
    except Exception:
        pass

    return STATIC_REITS[:], "static"


def get_dividend_yields(base_date: str) -> dict[str, float]:
    stock = try_import_pykrx_stock()
    if stock is None:
        return {}
    try:
        fundamentals = stock.get_market_fundamental_by_ticker(base_date, market="ALL")
    except Exception:
        return {}

    if "DIV" not in fundamentals.columns:
        return {}

    result: dict[str, float] = {}
    for ticker, row in fundamentals.iterrows():
        try:
            div = row.get("DIV")
            if div is not None and str(div) != "nan":
                result[str(ticker).zfill(6)] = float(div)
        except Exception:
            continue
    return result


def _xml_first_text(node: ET.Element, tag: str) -> str:
    found = node.find(tag)
    return (found.text or "").strip() if found is not None else ""


def get_corp_code_mapping(api_key: str) -> dict[str, str]:
    raw = http_get(f"{DART_BASE}/corpCode.xml", {"crtfc_key": api_key})
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        with zf.open("CORPCODE.xml") as fp:
            root = ET.fromstring(fp.read())

    mapping: dict[str, str] = {}
    for item in root.findall("list"):
        stock_code = _xml_first_text(item, "stock_code")
        corp_code = _xml_first_text(item, "corp_code")
        if stock_code and corp_code:
            mapping[stock_code] = corp_code
    return mapping


def find_latest_report_receipt(api_key: str, corp_code: str, begin_date: str) -> Optional[str]:
    raw = http_get(
        f"{DART_BASE}/list.json",
        {
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bgn_de": begin_date,
            "pblntf_ty": "A",
            "last_reprt_at": "Y",
            "page_count": "10",
        },
    )
    data = json.loads(raw.decode("utf-8"))
    if data.get("status") != "000" or not data.get("list"):
        return None
    reports = sorted(data["list"], key=lambda x: x.get("rcept_dt", ""), reverse=True)
    return reports[0].get("rcept_no")


def download_report_text(api_key: str, rcept_no: str) -> str:
    raw = http_get(f"{DART_BASE}/document.xml", {"crtfc_key": api_key, "rcept_no": rcept_no})
    with zipfile.ZipFile(io.BytesIO(raw)) as zf:
        xml_names = [name for name in zf.namelist() if name.lower().endswith(".xml")]
        texts = []
        for name in xml_names:
            with zf.open(name) as fp:
                try:
                    root = ET.fromstring(fp.read())
                except ET.ParseError:
                    continue
                text = " ".join(root.itertext()).strip()
                if text:
                    texts.append(text)
    return "\n".join(texts)


def infer_asset_type(text: str) -> Optional[str]:
    lower = text.lower()
    found = [cat for cat, pats in ASSET_TYPE_PATTERNS.items() if any(p.lower() in lower for p in pats)]
    return ", ".join(sorted(set(found))) if found else None


def infer_lease_structure(text: str) -> Optional[str]:
    found = [kw for kw in LEASE_STRUCTURE_PATTERNS if kw.lower() in text.lower()]
    return ", ".join(sorted(set(found))) if found else None


def infer_vacancy_rate(text: str) -> Optional[float]:
    values = []
    for regex in VACANCY_REGEXES:
        for match in regex.finditer(text):
            try:
                val = float(match.group(1))
                if 0 <= val <= 100:
                    values.append(val)
            except Exception:
                continue
    return min(values) if values else None


def build_dataset(base_date: str, dart_api_key: Optional[str]) -> tuple[list[ReitRecord], str]:
    reits, source_name = get_listed_reits(base_date)
    dividends = get_dividend_yields(base_date)

    corp_map: dict[str, str] = {}
    if dart_api_key:
        try:
            corp_map = get_corp_code_mapping(dart_api_key)
        except Exception as exc:
            print(f"[경고] DART 기업코드 수집 실패: {exc}")

    records: list[ReitRecord] = []
    for ticker, name in reits:
        rec = ReitRecord(ticker=ticker, name=name, dividend_yield=dividends.get(ticker), source=source_name)

        if dart_api_key and ticker in corp_map:
            try:
                rcept_no = find_latest_report_receipt(dart_api_key, corp_map[ticker], DEFAULT_REPORT_BEGIN_DATE)
                if rcept_no:
                    text = download_report_text(dart_api_key, rcept_no)
                    rec.asset_type = infer_asset_type(text)
                    rec.lease_structure = infer_lease_structure(text)
                    rec.vacancy_rate = infer_vacancy_rate(text)
                    rec.source = f"{source_name}+DART:{rcept_no}"
                else:
                    rec.note = "DART 보고서 없음"
            except Exception as exc:
                rec.note = f"DART 수집 실패: {exc}"
        elif dart_api_key:
            rec.note = "DART 종목 매핑 없음"

        records.append(rec)

    return records, source_name


def _col_to_letter(idx: int) -> str:
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


def _xml_escape(v: str) -> str:
    return (
        v.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def write_xlsx(records: list[ReitRecord], output_path: str) -> None:
    columns = [
        "종목코드",
        "리츠명",
        "자산 유형(추정)",
        "임차 구조(추정)",
        "공실률(%)",
        "배당 수익률(%)",
        "데이터 출처",
        "비고",
    ]

    rows = [columns]
    for r in records:
        rows.append(
            [
                r.ticker,
                r.name,
                r.asset_type or "",
                r.lease_structure or "",
                "" if r.vacancy_rate is None else f"{r.vacancy_rate}",
                "" if r.dividend_yield is None else f"{r.dividend_yield}",
                r.source,
                r.note,
            ]
        )

    sheet_rows = []
    for r_idx, row in enumerate(rows, start=1):
        cells = []
        for c_idx, value in enumerate(row, start=1):
            ref = f"{_col_to_letter(c_idx)}{r_idx}"
            v = _xml_escape(str(value))
            cells.append(f'<c r="{ref}" t="inlineStr"><is><t>{v}</t></is></c>')
        sheet_rows.append(f"<row r=\"{r_idx}\">{''.join(cells)}</row>")

    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(sheet_rows)}</sheetData>"
        "</worksheet>"
    )

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""

    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="REITs" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>"""

    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""

    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
</styleSheet>"""

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
        zf.writestr("xl/styles.xml", styles)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="대한민국 상장 리츠 자동 수집기")
    parser.add_argument("--date", default=datetime.today().strftime("%Y%m%d"), help="기준일(YYYYMMDD)")
    parser.add_argument("--output", default="korea_listed_reits.xlsx", help="출력 엑셀 파일 경로")
    parser.add_argument("--dart-api-key", default=os.getenv("OPEN_DART_API_KEY"), help="OpenDART API Key")
    return parser.parse_args()


def main() -> int:
    load_env_file()
    args = parse_args()

    if not args.dart_api_key:
        print("[안내] OPEN_DART_API_KEY가 없어 DART 기반 컬럼은 비어 있을 수 있습니다.")

    records, src = build_dataset(args.date, args.dart_api_key)
    write_xlsx(records, args.output)

    print(f"총 {len(records)}개 리츠를 {args.output}로 저장했습니다. (종목 소스: {src})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
