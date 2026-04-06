# PPT Block Maker

원본 PPTX를 분석하여 **제안서 작성용 프로젝트 폴더**를 자동 생성하는 CLI 도구입니다.

```
                      원본 PPTX (vol2.pptx, vol3.pptx)
                              │
                              ▼
                ┌─────────────────────────────┐
                │  1단계: ppt-block-maker     │
                │  분석 → 블록처리 → 분할 → MD │
                └─────────────────────────────┘
                              │
              ┌───────────────┼───────────────┐
              ▼               ▼               ▼
    docs/slides/        templates/slides/   templates/
    S????.md            S????.pptx          slide_index.json
    (원본 텍스트)       (████ 블록처리)     (메타데이터)
              │                                │
              ▼                                │
    ┌─────────────────────────────┐            │
    │  2단계: pptx-vertical-writer │            │
    │  S????.md 참고하여           │            │
    │  proposal-body.md 작성       │            │
    └─────────────────────────────┘            │
              │                                │
              ▼                                ▼
    ┌──────────────────────────────────────────┐
    │  3단계: md2pptx                          │
    │  proposal-body.md + 블록 PPTX            │
    │  → 최종 제안서 PPTX 생성                  │
    └──────────────────────────────────────────┘
              │
              ▼
        최종 제안서.pptx
```

---

## 설치

### 1. 저장소 클론

```bash
git clone https://github.com/leedonwoo2827-ship-it/ppt-block-maker.git
cd ppt-block-maker
```

### 2. 의존성 설치

```bash
pip install -r requirements.txt
```

### 요구사항

- Windows + Microsoft PowerPoint (COM 자동화 사용)
- Python 3.9 이상
- 패키지: `mcp`, `python-pptx`, `comtypes`

---

## 프로젝트 폴더 준비

실행 전, 아래와 같은 구조로 프로젝트 폴더를 만들어 두세요.

```
내프로젝트/
├── input/                ← 원본 PPTX 파일을 여기에 넣기
│   ├── vol2.pptx         예) 제안서 Part II
│   └── vol3.pptx         예) 제안서 Part III
├── docs/                 ← (자동 생성됨)
├── templates/            ← (자동 생성됨)
├── references/           ← 참고 PDF 등
├── rfp/                  ← RFP, 제안요청서
├── rawdata/              ← 통계, 보고서
└── output/               ← 최종 산출물
```

> `docs/`, `templates/`는 비워두면 됩니다. 실행 시 자동으로 채워집니다.

---

## 사용법

### 기본 명령어

```powershell
cd D:\00work\260404-PPTblokmaker

python run.py "C:\Users\사용자이름\Documents\pro2ppt\내프로젝트\input\vol2.pptx" "C:\Users\사용자이름\Documents\pro2ppt\내프로젝트" --vol 2
```

### 명령어 형식

```
python run.py <원본PPTX경로> <프로젝트폴더경로> --vol <볼륨번호>
```

| 인자 | 설명 | 기본값 |
|---|---|---|
| `원본PPTX경로` | 분석할 PPTX 파일 | (필수) |
| `프로젝트폴더경로` | 결과를 저장할 폴더 | (필수) |
| `--vol` | 볼륨 번호 (1~9, 0=자동분류) | 0 |
| `--merge-to` | 외부 slide_index.json 경로 (추가 머지) | (선택) |

### 실행 예시

```powershell
# Part II 분석 (vol 2)
python run.py "C:\내프로젝트\input\vol2.pptx" "C:\내프로젝트" --vol 2

# Part III 추가 (같은 프로젝트 폴더에 vol 3 병합)
python run.py "C:\내프로젝트\input\vol3.pptx" "C:\내프로젝트" --vol 3

# 자동 분류 모드
python run.py "C:\내프로젝트\input\제안서.pptx" "C:\내프로젝트"
```

---

## 실행 결과

### 터미널 출력 예시

```
============================================================
PPT Block Maker
  입력: C:\내프로젝트\input\vol2.pptx
  출력: C:\내프로젝트
  볼륨: 2
============================================================

[Step 1/5] 슬라이드 분석...
  Processed 20/128...
  Processed 40/128...
  ...
  slide_index.json 생성 완료

[Step 2/5] 텍스트 블록처리...
  Slides: 128/128 done
  블록처리 완료

[Step 3/5] 슬라이드 분할...
  128/128...
  분할 완료: 128개 파일

[Step 4/5] 개별 슬라이드 MD 생성...
  개별 MD 생성 완료: 128개 파일

[Step 5/5] 그룹 MD 생성...
  그룹 MD 생성 완료

============================================================
완료!
  출력: C:\내프로젝트
  docs/         : 11개 MD (GUIDE + T0~T9)
  docs/slides/  : 128개 MD (개별 슬라이드)
  templates/slides/ : 128개 PPTX (블록처리)
  볼륨: vol2 (S2001~)
============================================================
```

### 생성되는 폴더 구조

```
내프로젝트/
├── docs/
│   ├── GUIDE.md                 # 워크플로우 가이드
│   ├── T0.md ~ T9.md            # 템플릿별 슬라이드 목록 (참고용)
│   └── slides/                  # ★ 핵심 산출물
│       ├── S2001.md             # 원본 텍스트가 @필드에 채워진 스니펫
│       ├── S2002.md
│       ├── S2003.md
│       └── ... (슬라이드 수만큼)
├── templates/
│   ├── slide_index.json         # 슬라이드 메타데이터
│   └── slides/                  # 블록처리된 1장짜리 PPTX
│       ├── S2001.pptx           # 텍스트가 ████로 마스킹됨
│       ├── S2002.pptx
│       └── ... (슬라이드 수만큼)
├── input/
│   └── vol2.pptx                # 원본 (그대로 유지)
└── ...
```

### 개별 슬라이드 MD 예시 (docs/slides/S2003.md)

```markdown
---slide
# [S2003] 추진 전략 및 방법론
template: T1
ref_slide: 2003
---
@governing_message: 본 사업의 추진 배경과 전략적 접근 방향을 제시합니다.
@breadcrumb: II. 본문 > 1. 전략 및 방법론 > 1-1. 개요
@content_1: 대상 지역의 현황 분석을 기반으로 단계적 실행 전략을 수립하였습니다.
@카드1_제목: 추진 배경
@카드1_내용: 기관은 2020년부터 대상 지역의 인프라 개선을 위해...
@카드2_제목: 현지 여건 분석
@카드2_내용: 대상 지역은 전체 인구의 약 60%가 서비스 사각지대에...
```

> 이 MD 파일의 @필드 구조를 참고하여 2단계에서 제안서 본문(proposal-body.md)을 작성합니다.

---

## 실행 파이프라인

```
원본 PPTX
  │
  ├─ Step 1. 슬라이드 분석
  │   → slide_index.json (shape 역할 분류, 원본 텍스트 보존)
  │
  ├─ Step 2. 텍스트 블록처리
  │   → sanitized.pptx (모든 텍스트 → ████ 마스킹)
  │
  ├─ Step 3. 슬라이드 분할
  │   → templates/slides/S2001.pptx, S2002.pptx, ... (1장짜리)
  │
  ├─ Step 4. 개별 슬라이드 MD 생성 ★
  │   → docs/slides/S2001.md, S2002.md, ... (원본 텍스트 포함)
  │
  └─ Step 5. 그룹 MD 생성
      → docs/T0~T9.md + GUIDE.md (템플릿별 참고 목록)
```

**핵심 포인트:**
- **MD 파일** (docs/slides/) → 원본 텍스트가 살아있음 → 2단계 글쓰기 참고용
- **PPTX 파일** (templates/slides/) → 텍스트가 ████로 마스킹됨 → 3단계 변환용

---

## 3단계 연동 워크플로우

| 단계 | 도구 | 입력 | 출력 |
|---|---|---|---|
| **1단계** | ppt-block-maker | 원본 PPTX | docs/slides/, templates/slides/ |
| **2단계** | pptx-vertical-writer | docs/slides/S????.md (참고) | proposal-body.md |
| **3단계** | md2pptx | proposal-body.md + templates/ | 최종 PPTX |

---

## Claude Code Desktop MCP 설정 (선택)

CLI 외에 Claude Code Desktop에서 MCP 도구로도 사용할 수 있습니다.

### 설정 파일 위치

| OS | 경로 |
|---|---|
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |

### JSON 설정

```json
{
  "mcpServers": {
    "pptx-blockmaker": {
      "command": "python",
      "args": [
        "D:/00work/260404-PPTblokmaker/server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

> `args` 경로를 실제 설치 위치로 변경해주세요.

### MCP 도구

| 도구 | 설명 |
|---|---|
| `prepare_project` | 원본 PPTX를 분석하여 프로젝트 폴더 생성 |
| `add_volume` | 기존 프로젝트에 새 볼륨 추가 |

---

## 템플릿 타입

슬라이드는 레이아웃 구조에 따라 자동 분류됩니다.

| 타입 | 이름 | 용도 |
|---|---|---|
| T0 | 구분페이지 | 섹션/챕터 구분 |
| T1 | 카드형 다중 | 현황 분석, 비교, 개선안 |
| T2 | 카드+다이어그램 | 목적, 전략 개요 |
| T3 | 범위/개요 | 범위, 비전 |
| T4 | 다중 데이터테이블 | 복수 테이블, 일정표 |
| T5 | 테이블+다이어그램 | 테이블 + 설명 |
| T6 | 순수 데이터테이블 | 대형 표, 인력표 |
| T7 | 프로세스+테이블 | 프로세스 흐름 + 데이터 |
| T8 | 이미지중심 | 조직도, 구성도 |
| T9 | 핵심메시지 | 핵심 포인트, 다이어그램 |

---

## 볼륨 번호 체계

슬라이드 번호 = `볼륨 x 1000 + 순번`

| 볼륨 | 번호 범위 | 용도 |
|---|---|---|
| vol 0 | S0001 ~ S0999 | 자동 분류 |
| vol 1 | S1001 ~ S1999 | Part I |
| vol 2 | S2001 ~ S2999 | Part II |
| vol 3 | S3001 ~ S3999 | Part III |

같은 폴더에 여러 볼륨을 순차 실행하면 slide_index.json이 자동 병합됩니다.

---

## 트러블슈팅

| 증상 | 해결 |
|---|---|
| PermissionError: sanitized.pptx | `_temp` 폴더 삭제 후 재실행, 또는 PowerPoint 프로세스 종료 |
| PowerPoint COM 오류 | PowerPoint가 설치되어 있는지 확인, 다른 PPTX 파일 닫기 |
| 한글 인코딩 오류 | `PYTHONIOENCODING=utf-8` 환경변수 설정 |
| MCP 서버 미인식 | JSON 문법 확인, server.py 절대경로 확인, 앱 재시작 |

---

## License

MIT
