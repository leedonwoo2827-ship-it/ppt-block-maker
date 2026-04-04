# PPT Block Maker

원본 PPTX를 분석하여 **세로형 제안서 작성용 프로젝트 폴더**를 자동 생성하는 도구입니다.  
[pptx-vertical-writer](https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer) 플러그인과 함께 사용합니다.

## 주요 기능

| 단계 | 설명 |
|---|---|
| **슬라이드 분석** | Shape 역할 분류 → `slide_index.json` 생성 |
| **텍스트 블록처리** | 텍스트를 `████`로 마스킹하여 레이아웃 전용 템플릿 생성 |
| **슬라이드 분할** | 1장짜리 PPTX로 분할 (`S2001.pptx`, `S2002.pptx`, ...) |
| **MD 생성** | 템플릿별 가이드 문서 생성 (`T0~T9.md` + `GUIDE.md`) |

## 출력 구조

```
출력폴더/
├── docs/              # 템플릿 가이드 MD 파일
│   ├── GUIDE.md       # 워크플로우 가이드
│   ├── T0.md          # 구분페이지
│   ├── T1.md          # 카드형 다중
│   ├── T2.md          # 카드+다이어그램
│   └── ...            # T3~T9
└── templates/
    └── slides/        # 1장짜리 블록처리된 PPTX
        ├── S2001.pptx
        ├── S2002.pptx
        └── ...
```

## 템플릿 타입

| 타입 | 이름 | 용도 |
|---|---|---|
| T0 | 구분페이지 | 섹션/챕터 구분 |
| T1 | 카드형 다중 | 현황/문제점/개선, 비교 분석 |
| T2 | 카드+다이어그램 | 사업 목적, 전략 개요 |
| T3 | 범위/개요 | 사업 범위, 비전 |
| T4 | 다중 데이터테이블 | 복수 테이블, 일정표 |
| T5 | 테이블+다이어그램 | 테이블 + 설명 |
| T6 | 순수 데이터테이블 | 큰 데이터 표, 인력표 |
| T7 | 프로세스+테이블 | 프로세스 흐름 + 데이터 |
| T8 | 이미지중심 | 조직도, 구성도 (이미지 수작업) |
| T9 | 핵심메시지/다이어그램 | CSF, 핵심 포인트 |

## 설치

```bash
pip install -r requirements.txt
```

### 요구 사항

- Python 3.9+
- `mcp`, `python-pptx`, `comtypes`

## 사용법

### CLI

```bash
# 기본 사용
python run.py <원본.pptx> <출력_프로젝트_폴더> --vol <볼륨번호>

# 예시: 기술부문 (vol 2)
python run.py input/기술부문.pptx ./output/project --vol 2

# 예시: 사업관리부문 추가 (vol 3)
python run.py input/사업관리부문.pptx ./output/project --vol 3

# 자동 분류 모드
python run.py input/제안서.pptx ./output/project --vol 0
```

### MCP Server

```bash
python server.py
```

MCP 도구로 두 가지 기능을 제공합니다:

- **`prepare_project`** — 원본 PPTX를 분석하여 프로젝트 폴더를 준비
- **`add_volume`** — 기존 프로젝트에 새 볼륨의 슬라이드를 추가

## 볼륨 번호 체계

슬라이드 번호는 `볼륨 x 1000 + 순번`으로 생성됩니다:

- vol 0: S0001, S0002, ... (자동 분류)
- vol 2: S2001, S2002, ... (II권)
- vol 3: S3001, S3002, ... (III권)

## License

MIT
