# reits-automation

대한민국 상장 리츠(REITs)의 정보를 자동 수집해 엑셀(`.xlsx`)로 저장합니다.

수집 항목:
- 자산 유형(추정)
- 임차 구조(추정)
- 공실률(가능 시)
- 배당 수익률(가능 시)

## 특징 (요청 반영)

- **기존 방식 외 대체 크롤링 추가**
  - 1순위: `pykrx`
  - 2순위: **KIND 상장법인 다운로드 페이지 크롤링**
  - 3순위: 정적 백업 REIT 목록
- 위 순서로 동작하므로, 특정 소스 실패 시에도 결과가 아예 비는 문제를 줄였습니다.
- 외부 라이브러리 없이도 `.xlsx` 파일 생성이 가능합니다.

## 실행

```bash
python reits_automation.py
```

출력 파일 지정:

```bash
python reits_automation.py --output reits_result.xlsx
```

OpenDART API 키를 사용하면 공시 기반 추정 컬럼이 보강됩니다.

```bash
export OPEN_DART_API_KEY="YOUR_API_KEY"
python reits_automation.py
```

`.env`도 지원합니다.

```bash
cat > .env <<EOF
OPEN_DART_API_KEY=YOUR_API_KEY
EOF
python reits_automation.py
```

## 참고

- 네트워크 제한 환경에서는 일부 크롤링/공시 조회가 실패할 수 있으며, 이 경우 가능한 대체 소스로 계속 진행합니다.
