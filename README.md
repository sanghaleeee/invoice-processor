# Invoice Processor

L'Oréal Travel Retail 인보이스 PDF → 엑셀 변환 도구.
배송지 주소에서 리테일러 자동 감지 (Lotte/Shilla/Shinsegae/Hyundai/JDC),
EAN 기반 SKU 마스터 조회로 리테일러 코드 매핑.

## 기능

- **PDF 드래그 앤 드롭** → 자동 파싱 → 엑셀 출력
- **리테일러 자동 감지** — 배송지 주소에서 LOTTE/SHILLA/SHINSEGAE/HYUNDAI/JEJU FREE 키워드识别
- **SKU 마스터 등록** — 엑셀 파일 드롭으로 최신 마스터 등록/업데이트
- **다중 PDF 처리** — 여러 인보이스를 한번에 처리, 합계 결과 표시
- **macOS 앱** (SwiftUI) + **Windows 앱** (tkinter) + **CLI** 모두 지원

## 다운로드

| 플랫폼 | 설치 방법 |
|--------|----------|
| **macOS** | `Invoice Processor.app` (Desktop) — SwiftUI 네이티브 앱 |
| **Windows** | `python invoice_app.py` 또는 PyInstaller로 .exe 빌드 |
| **CLI** | `python3 process_invoice.py invoice.pdf` |

## Quick Start

### macOS
`Invoice Processor.app`에 PDF나 엑셀 파일을 드래그 앤 드롭.

### Windows
```cmd
pip install PyMuPDF openpyxl
python invoice_app.py
```

### CLI (모든 플랫폼)
```bash
pip install PyMuPDF openpyxl
python3 process_invoice.py path/to/invoice.pdf
```

## SKU Master 등록

1. SKU master Excel 파일을 앱에 드래그 (또는 `--sku-master` 인자)
2. `~/hermes work/po-process/sku-master/`에 저장, `config.json`에 경로 기록
3. 우선순위: `--sku-master` 인자 > `config.json` > `~/Downloads/` 최신 파일

### SKU Master 필수 컬럼
- `EAN Code` — 13자리 EAN
- `Lotte Code` — 롯데면세점 코드
- `Shilla Code` — 신라면세점 코드
- `SSG Code` — SSG 코드
- `Hyundai Code` — 현대면세점 코드
- `JDC Code` — JDC 코드

## 출력 엑셀 (`~/Desktop/{InvoiceNo}_processed.xlsx`)

| EAN Number | Retailer Code | Description | Quantity | Unit Price (USD) | Total Price (USD) |
|------------|--------------|-------------|----------|-----------------|-------------------|

마지막 행: Quantity / Unit Price / Total Price SUM 수식

## 리테일러 감지 로직

PDF 1페이지 오른쪽 DELIVERY ADDRESS 영역에서 키워드 탐지:

| 주소 키워드 | 출력 코드 |
|---|---|
| `LOTTE` | Lotte Code |
| `SHILLA` / `신라` | Shilla Code |
| `SHINSEGAE` / `SSG` | SSG Code |
| `HYUNDAI` / `현대` | Hyundai Code |
| `JEJU FREE` / `JDC` | JDC Code |

## 프로젝트 구조

```
po-process/
├── process_invoice.py       # 핵심 엔진 (CLI)
├── invoice_app.py           # tkinter GUI (크로스 플랫폼)
├── Invoice Processor/       # SwiftUI macOS 앱 소스
│   └── Invoice Processor/
│       ├── InvoiceProcessorApp.swift
│       └── ContentView.swift
├── install_deps.bat         # Windows 의존성 설치
├── build_exe.bat            # Windows .exe 빌드
├── sku-master/              # 등록된 SKU master 파일
└── config.json              # 설정 (로컬 경로)
```

## 의존성

- Python 3.9+
- [PyMuPDF](https://pypi.org/project/PyMuPDF/) — PDF 파싱
- [openpyxl](https://pypi.org/project/openpyxl/) — Excel 생성
- (선택) [PyInstaller](https://pypi.org/project/pyinstaller/) — Windows .exe 빌드
- (macOS only) Xcode Command Line Tools — SwiftUI 빌드

## 라이선스

MIT
