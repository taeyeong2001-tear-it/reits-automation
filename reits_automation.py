#!/usr/bin/env python3
"""대한민국 상장 리츠 정보를 자동 수집해 엑셀로 저장합니다."""

from __future__ import annotations

import argparse
import io
import os
import re
import sys
import zipfile
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, Optional
from xml.etree import ElementTree as ET


DART_BASE = "https://opendart.fss.or.kr/api"
DEFAULT_REPORT_BEGIN_DATE = "20200101"


def load_env_file(path: str = ".env") -> None:
    """간단한 .env 로더 (KEY=VALUE). 이미 설정된 환경변수는 덮어쓰지 않습니다."""
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


def require_module(module_name: str):
    try:
        return __import__(module_name)
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            f"필수 패키지 '{module_name}'가 설치되어 있지 않습니다. "
            "`pip install -r requirements.txt`를 먼저 실행하세요."
        ) from exc


def create_session():
    requests = require_module("requests")
    session = requests.Session()
    adapter = requests.adapters.HTTPAdapter(max_retries=3)
    session.mount("http://", adapter)
    session.mount("https://", adapter)
    return session


def get_listed_reits(base_date: str) -> list[tuple[str, str]]:
    stock = require_module("pykrx").stock
    tickers = stock.get_market_ticker_list(base_date, market="ALL")
    reits = []
    for ticker in tickers:
        name = stock.get_market_ticker_name(ticker)
        if "리츠" in name:
            reits.append((ticker, name))
    return sorted(reits, key=lambda x: x[1])


def get_dividend_yields(base_date: str) -> dict[str, float]:
    pd = require_module("pandas")
    stock = require_module("pykrx").stock
    fundamentals = stock.get_market_fundamental_by_ticker(base_date, market="ALL")
    if "DIV" not in fundamentals.columns:
        return {}

    result: dict[str, float] = {}
    for ticker, row in fundamentals.iterrows():
        div = row.get("DIV")
        if pd.notna(div):
            result[str(ticker).zfill(6)] = float(div)
    return result


def _xml_first_text(node: ET.Element, tag: str) -> str:
    found = node.find(tag)
    return (found.text or "").strip() if found is not None else ""


def get_corp_code_mapping(api_key: str, session) -> dict[str, str]:
    resp = session.get(f"{DART_BASE}/corpCode.xml", params={"crtfc_key": api_key}, timeout=30)
    resp.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        with zf.open("CORPCODE.xml") as fp:
            root = ET.fromstring(fp.read())

    mapping: dict[str, str] = {}
    for item in root.findall("list"):
        stock_code = _xml_first_text(item, "stock_code")
        corp_code = _xml_first_text(item, "corp_code")
        if stock_code and corp_code:
            mapping[stock_code] = corp_code
    return mapping


def find_latest_report_receipt(
    api_key: str,
    corp_code: str,
    begin_date: str,
    session,
) -> Optional[str]:
    params = {
        "crtfc_key": api_key,
        "corp_code": corp_code,
        "bgn_de": begin_date,
        "pblntf_ty": "A",
        "last_reprt_at": "Y",
        "page_count": 10,
    }
    resp = session.get(f"{DART_BASE}/list.json", params=params, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    if data.get("status") != "000" or not data.get("list"):
        return None

    reports = sorted(data["list"], key=lambda x: x.get("rcept_dt", ""), reverse=True)
    return reports[0].get("rcept_no")


def download_report_text(api_key: str, rcept_no: str, session) -> str:
    resp = session.get(
        f"{DART_BASE}/document.xml",
        params={"crtfc_key": api_key, "rcept_no": rcept_no},
        timeout=60,
    )
    resp.raise_for_status()

    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        xml_names = [name for name in zf.namelist() if name.lower().endswith(".xml")]
        texts = []
        for name in xml_names:
            with zf.open(name) as fp:
                try:
                    root = ET.fromstring(fp.read())
                except ET.ParseError:
                    continue
                text = " ".join(root.itertext())
                if text.strip():
                    texts.append(text.strip())
        return "\n".join(texts)


def infer_asset_type(text: str) -> Optional[str]:
    lower = text.lower()
    found = []
    for category, patterns in ASSET_TYPE_PATTERNS.items():
        if any(pattern.lower() in lower for pattern in patterns):
            found.append(category)

    if not found:
        return None
    return ", ".join(sorted(set(found)))


def infer_lease_structure(text: str) -> Optional[str]:
    found = [kw for kw in LEASE_STRUCTURE_PATTERNS if kw.lower() in text.lower()]
    if not found:
        return None
    return ", ".join(sorted(set(found)))


def infer_vacancy_rate(text: str) -> Optional[float]:
    values = []
    for regex in VACANCY_REGEXES:
        for match in regex.finditer(text):
            try:
                value = float(match.group(1))
            except ValueError:
                continue
            if 0 <= value <= 100:
                values.append(value)
    if not values:
        return None
    return min(values)


def build_dataset(base_date: str, dart_api_key: Optional[str]) -> Any:
    pd = require_module("pandas")
    reits = get_listed_reits(base_date)
    dividends = get_dividend_yields(base_date)

    records: list[ReitRecord] = []

    corp_map: dict[str, str] = {}
    session = create_session()
    if dart_api_key:
        corp_map = get_corp_code_mapping(dart_api_key, session)

    for ticker, name in reits:
        record = ReitRecord(ticker=ticker, name=name, dividend_yield=dividends.get(ticker))

        if dart_api_key and ticker in corp_map:
            try:
                rcept_no = find_latest_report_receipt(
                    dart_api_key,
                    corp_map[ticker],
                    DEFAULT_REPORT_BEGIN_DATE,
                    session,
                )
                if rcept_no:
                    text = download_report_text(dart_api_key, rcept_no, session)
                    record.asset_type = infer_asset_type(text)
                    record.lease_structure = infer_lease_structure(text)
                    record.vacancy_rate = infer_vacancy_rate(text)
                    record.source = f"DART:{rcept_no}"
                    if not any([record.asset_type, record.lease_structure, record.vacancy_rate]):
                        record.note = "보고서 파싱 성공, 주요 지표 미추출"
                else:
                    record.note = "DART 보고서 없음"
            except Exception as exc:  # noqa: BLE001
                record.note = f"DART 수집 실패: {exc}"
        else:
            if dart_api_key:
                record.note = "DART 종목 매핑 없음"

        records.append(record)

    ordered_columns = [
        "종목코드",
        "리츠명",
        "자산 유형(추정)",
        "임차 구조(추정)",
        "공실률(%)",
        "배당 수익률(%)",
        "데이터 출처",
        "비고",
    ]
    df = pd.DataFrame([asdict(record) for record in records]).rename(
        columns={
            "ticker": "종목코드",
            "name": "리츠명",
            "asset_type": "자산 유형(추정)",
            "lease_structure": "임차 구조(추정)",
            "vacancy_rate": "공실률(%)",
            "dividend_yield": "배당 수익률(%)",
            "source": "데이터 출처",
            "note": "비고",
        }
    )
    return df.reindex(columns=ordered_columns)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="대한민국 상장 리츠 자동 수집기")
    parser.add_argument("--date", default=datetime.today().strftime("%Y%m%d"), help="기준일(YYYYMMDD)")
    parser.add_argument("--output", default="korea_listed_reits.xlsx", help="출력 엑셀 파일 경로")
    parser.add_argument(
        "--dart-api-key",
        default=os.getenv("OPEN_DART_API_KEY"),
        help="OpenDART API Key (미지정시 배당 수익률/종목정보 위주로 생성)",
    )
    return parser.parse_args()


def main() -> int:
    load_env_file()
    args = parse_args()

    try:
        if not args.dart_api_key:
            print("[안내] OPEN_DART_API_KEY가 없어 DART 기반 컬럼(자산 유형/임차 구조/공실률)은 비어 있을 수 있습니다.")
        df = build_dataset(args.date, args.dart_api_key)
        df.to_excel(args.output, index=False)
    except RuntimeError as exc:
        print(f"[실패] {exc}", file=sys.stderr)
        return 2

    print(f"총 {len(df)}개 리츠를 {args.output}로 저장했습니다.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
