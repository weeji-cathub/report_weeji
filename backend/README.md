# Excel ↔ Web Bridge (Backend)

직업병 사례 회의서식 웹앱이 호출하는 로컬 FastAPI 서버.
xlwings로 **이미 열려있는 Excel 인스턴스에 attach**해서 `대상` / `대상(2025)` 시트를 직접 read/write 합니다.

## 사전 조건

- Windows + Microsoft Excel
- 마스터 `.xlsm` 워크북을 사용자가 직접 Excel에서 열어둘 것
  (파일명이 `app.py`의 `WORKBOOK` 상수와 정확히 일치해야 함)

## 설치

```powershell
pip install -r requirements.txt
```

## 실행

```powershell
python app.py
```

- Server: <http://127.0.0.1:8765>
- Swagger UI: <http://127.0.0.1:8765/docs>

## 엔드포인트

| Method | Path                                       | 용도                                                |
| ------ | ------------------------------------------ | --------------------------------------------------- |
| GET    | `/case/{연번}`                              | 사례 1건 조회 (47열 헤더 → 값 dict)                |
| POST   | `/review`                                   | 동일 연번 행의 지정 헤더 업데이트 (whitelist 3개)  |
| GET    | `/search?field=&keyword=&limit=`            | 대상 + 대상(2025) 검색, 보고일 desc 정렬            |
| POST   | `/write`                                    | (legacy) 1셀 write 검증용                           |

`/review` 허용 키: `질환 및 발생형태`, `업무관련성`, `유해인자분류`

## 동시성

Excel COM은 STA(single-threaded apartment) 모델이라 동시 호출 시 RPC 충돌이 날 수 있습니다.
모든 엔드포인트는 `EXCEL_LOCK`(threading.Lock)으로 직렬화 — 단일 사용자 / 단일 Excel 인스턴스 환경 전제.

## CORS

로컬 단일 사용자 환경 가정으로 `allow_origins=["*"]`. `file://`에서 직접 호출하는 frontend도 허용됩니다.

## 미확정 / 후속

- `DATE_HDR = "보고일"`은 추정값입니다. 실제 시트 헤더명과 다르면 `app.py` 상단 한 줄만 변경하세요.
