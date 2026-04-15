# 스톡옵션 행사 자동화 시스템

S2W 주식업무 스톡옵션 행사 프로세스 자동화 로컬 웹 애플리케이션

## 실행 방법

```bash
# 패키지 설치
pip install -r requirements.txt

# 서버 실행
run.bat

# 브라우저에서 접속
http://localhost:5000
```

## 프로젝트 구조

```
stock-option-app/
├── app.py                  # Flask 메인 애플리케이션
├── database.py             # SQLite DB 헬퍼 함수
├── run.bat                 # 서버 실행 스크립트
├── requirements.txt        # Python 패키지 의존성
│
├── data/                   # 런타임 데이터 (Git 제외)
│   ├── stockops.db        # SQLite 데이터베이스
│   ├── uploads/           # 업로드된 서류
│   └── outputs/           # 생성된 결과물
│
├── processors/            # 비즈니스 로직 프로세서
│   ├── pdf_merger.py     # PDF 합본
│   ├── pdf_name_extractor.py  # PDF 이름 추출 (OCR)
│   ├── ocr_reader.py     # 주민번호 추출 (EasyOCR)
│   ├── excel_writer.py   # Excel 생성
│   ├── docx_writer.py    # Word 문서 생성
│   ├── hwpx_writer.py    # 한글 HWPX 생성
│   ├── step04_generator.py  # Step04 서류 생성
│   └── step05_generator.py  # Step05 서류 생성
│
├── templates/             # Jinja2 HTML 템플릿
│   ├── base.html         # 베이스 레이아웃
│   ├── index.html        # 홈
│   ├── step01.html       # 신청서 접수
│   ├── step03.html       # 납입금 납입
│   ├── step033.html      # 의무보유
│   ├── step04.html       # 등기신청
│   └── step05.html       # 예탁원 신주발행의뢰
│
├── templates_hwp/         # 한글/워드 원본 템플릿
├── templates_step04/      # Step04 전용 템플릿 및 서류
├── templates_step05/      # Step05 전용 템플릿 및 서류
│
├── static/
│   ├── css/style.css     # 스타일시트
│   └── js/               # JavaScript 파일
│       ├── app.js        # 공통 함수
│       └── step01.js     # Step01 전용
│
└── migrations/            # DB 마이그레이션 스크립트
    ├── add_attachment8.py
    ├── migrate_add_ocr_columns.py
    └── migrate_step05_update.py
```

## 주요 기능

### Step 01 - 신청서 접수
- 신청자 명단 관리 (추가/편집/삭제/순서변경/엑셀 import)
- 서류 업로드 (신청서, 신분증, 계좌사본)
- PDF 파일명 또는 내용에서 자동 이름 추출
- 서류별 PDF 합본 생성
- 직원 개인 제출 링크 자동 생성

### Step 03 - 행사대금 납입
- 행사내역 Excel 생성
- HWPX 자동 생성: 수납의뢰서, 영수증, 보관증명서

### Step 03-3 - 의무보유
- 의무보유 대상자 관리
- Word 문서 자동 생성: 의무보유확약서, 계속보유신청 공문

### Step 04 - 등기신청
- 주주총회의사록 및 조정내역서 통합 PDF
- 정관, 등기위임장 자동 복사
- 주식매수선택권 부여계약서 신청자별 매칭 후 합본
- 법인인감 날인 필요사항 안내

### Step 05 - 예탁원 신주발행의뢰
- 발행가액별 서류 ZIP 자동 생성
- OCR 캐싱 (주민번호, 계좌번호)
- 증권사 코드 자동 매칭
- 신분증 앞면 필터링

## 기술 스택

- **Backend**: Python 3.x, Flask
- **Database**: SQLite3
- **문서 생성**: python-docx, openpyxl, hwpx-writer, docx2pdf
- **PDF 처리**: PyPDF, PyMuPDF
- **OCR**: EasyOCR (로컬 처리, 외부 API 사용 안 함)
- **Frontend**: Vanilla JavaScript, Jinja2

## 보안

⚠️ **모든 개인정보 처리는 로컬에서만 수행됩니다.**
- 주민등록번호, 신분증, 계좌번호 등 민감 정보는 외부 전송 없음
- OCR은 로컬 라이브러리(EasyOCR) 사용
- Flask 서버는 localhost만 바인딩

## 라이선스

내부 사용 전용
