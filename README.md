# 입출고·불량 데이터 조회 앱

## 기능

### 입고
- 공장별 필터
- SKU 검색
- SKU별 입고수량 집계
- Excel 다운로드

### 출고
- 파트별 필터
- 매장명별 필터
- SKU 검색
- SKU별 출고수량 집계
- 매장·파트·SKU별 Excel 다운로드

### 불량
- 공장별 필터
- 불량타입별 필터
- SKU 검색
- 공장·불량타입·SKU별 불량수량 집계
- Excel 다운로드

## 분류 규칙

### 출고 전표명
- 일본 → 일본
- 미국 → 미국
- 시딩 / 협찬 / 본부장 → 협찬
- 해외파트 → B2B
- 오프라인반품 / 오프라인 매장 반품 → 오프라인 반품
- 출고요청 → `_` 사이의 이름을 매장명으로 유지
- 그 외 전표명은 원문 유지

### 불량타입
- 전체 → 전체
- 렌즈 → 렌즈
- 테 → 테
- 그 외 → 기타

### 공장
- C2-S → C2-S
- C2 / c2 → C2
- C5 → C5
- 그 외 / 공란 → 미상

## 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

브라우저에서 표시되는 주소로 접속하면 됩니다.

## Streamlit Community Cloud

GitHub 저장소에 아래 3개 파일을 올립니다.

- app.py
- requirements.txt
- 분석raw.xlsx

그 다음 Streamlit Community Cloud에서 GitHub 저장소를 연결하고 `app.py`를 Main file로 선택하여 배포합니다.
