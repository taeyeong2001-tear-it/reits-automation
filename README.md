# reits-automation

대한민국 상장 리츠(REITs)를 대상으로 아래 항목을 자동 수집해 엑셀로 저장하는 스크립트입니다.

- 자산 유형(추정)
- 임차 구조(추정)
- 공실률(가능 시)
- 배당 수익률

## 설치

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 실행

기본 실행(오늘 날짜 기준, `korea_listed_reits.xlsx` 생성):

```bash
python reits_automation.py
```

기준일/출력 파일 지정:

```bash
python reits_automation.py --date 20260218 --output reits_20260218.xlsx
```

OpenDART API 키를 설정하면 사업보고서 원문 기반으로 `자산 유형/임차 구조/공실률`을 추가 추정합니다.

```bash
export OPEN_DART_API_KEY="YOUR_API_KEY"
python reits_automation.py
```

또는 옵션으로 직접 전달:

```bash
python reits_automation.py --dart-api-key YOUR_API_KEY
```

`.env` 파일로도 설정할 수 있습니다:

```bash
cat > .env <<EOF
OPEN_DART_API_KEY=YOUR_API_KEY
EOF
python reits_automation.py
```

## 데이터 소스

- **KRX (pykrx)**: 상장 리츠 목록(종목명 `리츠` 포함), 배당 수익률
- **OpenDART API(선택)**: 최근 사업보고서 텍스트를 파싱해 자산 유형/임차 구조/공실률 추정

## 개선된 동작

- `--help`는 의존성 설치 전에도 정상 동작합니다.
- 필수 패키지가 누락되면 `pip install -r requirements.txt` 안내 메시지를 출력하고 종료합니다.
- OpenDART XML 파싱을 표준 라이브러리로 처리해 불필요한 의존성을 줄였습니다.

## 주의사항

- `자산 유형/임차 구조/공실률`은 보고서 내 텍스트 키워드 기반 자동 추정치입니다.
- 공시 문서 형식이 회사별로 달라 일부 리츠는 값이 비어 있을 수 있습니다.
- OpenDART API 키가 없으면 DART 기반 컬럼은 비어 있을 수 있으며, 실행 시 안내 메시지가 출력됩니다.
