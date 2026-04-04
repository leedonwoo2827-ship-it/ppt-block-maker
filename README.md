# PPT Block Maker

원본 PPTX를 분석하여 **세로형 제안서 작성용 프로젝트 폴더**를 자동 생성하는 MCP 서버입니다.  
[pptx-vertical-writer](https://github.com/leedonwoo2827-ship-it/pptx-vertical-writer) 플러그인과 함께 사용합니다.

---

## 빠른 시작: Claude Code Desktop에서 MCP 설치

### 1단계: 저장소 클론

```bash
git clone https://github.com/leedonwoo2827-ship-it/ppt-block-maker.git
cd ppt-block-maker
```

### 2단계: 의존성 설치

```bash
pip install -r requirements.txt
```

> 필요 패키지: `mcp`, `python-pptx`, `comtypes`  
> Python 3.9 이상 필요

### 3단계: Claude Code Desktop MCP 설정

Claude Code Desktop 앱의 설정 파일에 아래 내용을 추가합니다.

#### 설정 파일 위치

| OS | 경로 |
|---|---|
| **Windows** | `%APPDATA%\Claude\claude_desktop_config.json` |
| **macOS** | `~/Library/Application Support/Claude/claude_desktop_config.json` |

#### 설정 파일 편집

`claude_desktop_config.json` 파일을 열고, `mcpServers` 항목에 아래 내용을 추가합니다.

**Windows 사용자:**

```json
{
  "mcpServers": {
    "pptx-blockmaker": {
      "command": "python",
      "args": [
        "C:/Users/사용자이름/ppt-block-maker/server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

**macOS 사용자:**

```json
{
  "mcpServers": {
    "pptx-blockmaker": {
      "command": "python3",
      "args": [
        "/Users/사용자이름/ppt-block-maker/server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

> **주의**: `args`의 경로를 실제 클론한 위치로 변경해주세요.

#### 기존에 다른 MCP 서버가 있는 경우

이미 `mcpServers`에 다른 서버가 등록되어 있다면, 기존 항목은 유지하고 `pptx-blockmaker`만 추가합니다:

```json
{
  "mcpServers": {
    "기존-서버": {
      "command": "...",
      "args": ["..."]
    },
    "pptx-blockmaker": {
      "command": "python",
      "args": [
        "C:/Users/사용자이름/ppt-block-maker/server.py"
      ],
      "env": {
        "PYTHONIOENCODING": "utf-8"
      }
    }
  }
}
```

### 4단계: Claude Code Desktop 재시작

설정 파일을 저장한 뒤, Claude Code Desktop 앱을 **완전히 종료**하고 다시 실행합니다.

### 5단계: 연결 확인

Claude Code Desktop에서 아래와 같이 물어보면 MCP 도구가 응답합니다:

```
"기술부문.pptx를 분석해서 프로젝트 폴더를 만들어줘"
```

정상 연결 시 `prepare_project` 도구가 호출되어 작동합니다.

---

## MCP 도구 상세

이 서버는 두 가지 MCP 도구를 제공합니다.

### `prepare_project` — 프로젝트 폴더 생성

원본 PPTX를 분석하여 프로젝트 폴더를 한 번에 준비합니다.

| 파라미터 | 타입 | 필수 | 설명 |
|---|---|---|---|
| `pptx_path` | string | O | 원본 PPTX 파일 경로 |
| `output_dir` | string | O | 출력 프로젝트 폴더 경로 |
| `vol` | int | X | 볼륨 번호 (기본값: 0=자동분류) |

**실행 파이프라인:**

```
원본 PPTX
  │
  ├─ Step 1. 슬라이드 분석 → slide_index.json (Shape 역할 분류)
  │
  ├─ Step 2. 텍스트 블록처리 → sanitized.pptx (텍스트 → ████ 마스킹)
  │
  ├─ Step 3. 슬라이드 분할 → S2001.pptx, S2002.pptx, ... (1장짜리)
  │
  └─ Step 4. MD 생성 → T0~T9.md + GUIDE.md (템플릿 가이드)
```

### `add_volume` — 볼륨 추가

기존 프로젝트에 새 PPTX 볼륨을 추가 분석하여 병합합니다.

| 파라미터 | 타입 | 필수 | 설명 |
|---|---|---|---|
| `pptx_path` | string | O | 추가할 원본 PPTX 파일 경로 |
| `output_dir` | string | O | 기존 프로젝트 폴더 경로 |
| `vol` | int | O | 볼륨 번호 (1~9) |

---

## 출력 구조

```
출력폴더/
├── docs/                    # 템플릿 가이드 MD 파일
│   ├── GUIDE.md             # 전체 워크플로우 가이드
│   ├── T0.md                # 구분페이지 스니펫
│   ├── T1.md                # 카드형 다중 스니펫
│   ├── T2.md                # 카드+다이어그램 스니펫
│   ├── T3.md                # 범위/개요 스니펫
│   ├── T4.md                # 다중 데이터테이블 스니펫
│   ├── T5.md                # 테이블+다이어그램 스니펫
│   ├── T6.md                # 순수 데이터테이블 스니펫
│   ├── T7.md                # 프로세스+테이블 스니펫
│   ├── T8.md                # 이미지중심 스니펫
│   └── T9.md                # 핵심메시지/다이어그램 스니펫
└── templates/
    └── slides/              # 1장짜리 블록처리된 PPTX 템플릿
        ├── S2001.pptx
        ├── S2002.pptx
        ├── S2003.pptx
        └── ...
```

---

## 템플릿 타입 안내

각 슬라이드는 레이아웃 구조에 따라 T0~T9 중 하나로 자동 분류됩니다.

| 타입 | 이름 | 용도 | 주요 요소 |
|---|---|---|---|
| **T0** | 구분페이지 | 섹션/챕터 구분 | 섹션 제목 |
| **T1** | 카드형 다중 | 현황/문제점/개선, 비교 분석 | 카드 N개 (제목+본문) |
| **T2** | 카드+다이어그램 | 사업 목적, 전략 개요 | 카드 + 다이어그램 영역 |
| **T3** | 범위/개요 | 사업 범위, 비전 | 핵심 문구 + 콘텐츠 영역 |
| **T4** | 다중 데이터테이블 | 복수 테이블, 일정표 | 데이터 테이블 N개 |
| **T5** | 테이블+다이어그램 | 테이블 + 설명 | 테이블 + 콘텐츠 텍스트 |
| **T6** | 순수 데이터테이블 | 큰 데이터 표, 인력표 | 대형 단일 테이블 |
| **T7** | 프로세스+테이블 | 프로세스 흐름 + 데이터 | 단계 헤딩 + 테이블 |
| **T8** | 이미지중심 | 조직도, 구성도 | 이미지 (수작업 필요) |
| **T9** | 핵심메시지/다이어그램 | CSF, 핵심 포인트 | 핵심 문구 N개 |

---

## 볼륨 번호 체계

슬라이드 번호는 `볼륨 x 1000 + 순번`으로 생성됩니다.

| 볼륨 | 번호 범위 | 예시 |
|---|---|---|
| vol 0 | S0001 ~ S0999 | 자동 분류 모드 |
| vol 1 | S1001 ~ S1999 | I권 |
| vol 2 | S2001 ~ S2999 | II권 (기술부문 등) |
| vol 3 | S3001 ~ S3999 | III권 (사업관리부문 등) |

같은 프로젝트에 여러 볼륨을 순차적으로 추가하면, 기존 슬라이드와 자동 병합됩니다.

---

## CLI 사용법 (선택)

MCP 서버 외에 CLI로도 직접 실행할 수 있습니다.

```bash
# 기본 사용
python run.py <원본.pptx> <출력_프로젝트_폴더> --vol <볼륨번호>

# 예시: 기술부문 (vol 2)
python run.py input/기술부문.pptx ./output/project --vol 2

# 예시: 사업관리부문 추가 (vol 3)
python run.py input/사업관리부문.pptx ./output/project --vol 3

# 자동 분류 모드 (vol 생략 시 0)
python run.py input/제안서.pptx ./output/project
```

---

## 워크플로우 (pptx-vertical-writer 연동)

```
1. 원본 PPTX 준비
       │
2. ppt-block-maker로 프로젝트 폴더 생성
   (prepare_project 호출)
       │
3. docs/ 폴더의 T?.md에서 ref_slide 확인
       │
4. proposal-body-extended.md 작성
   (스니펫 복사 → @필드 채우기)
       │
5. pptx-vertical-writer로 최종 PPTX 생성
   (create_pptx 호출)
```

---

## 트러블슈팅

### MCP 서버가 인식되지 않는 경우

1. `claude_desktop_config.json`의 JSON 문법 오류 확인 (쉼표, 괄호)
2. `server.py` 경로가 정확한지 확인 — 절대 경로 사용 권장
3. `python --version`으로 Python 3.9+ 확인
4. `pip install -r requirements.txt`로 의존성 재설치
5. Claude Code Desktop 앱을 완전히 종료 후 재시작

### `PYTHONIOENCODING` 설정 이유

한글 파일명/경로 처리 시 인코딩 오류를 방지하기 위해 `utf-8`을 명시합니다.

---

## License

MIT
