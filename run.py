# -*- coding: utf-8 -*-
"""
PPT Block Maker — 원본 PPTX → 프로젝트 ready 변환 도구 (CLI)

사용법:
    python run.py <원본.pptx> <출력_프로젝트_폴더> [--vol 2] [--merge-to <경로>]

예시:
    python run.py input/기술부문.pptx C:/Users/.../pro2ppt/260405-2 --vol 2
    python run.py input/사업관리부문.pptx C:/Users/.../pro2ppt/260405-2 --vol 3

결과:
    출력폴더/
    ├── docs/
    │   ├── GUIDE.md + T0~T9.md
    │   └── slides/        ← S2001.md, S2002.md, ... (원본 텍스트 포함)
    └── templates/
        ├── slide_index.json
        └── slides/        ← S2001.pptx, S2002.pptx, ... (블록처리)
"""
import argparse
import json
import os
import shutil
import sys
from pathlib import Path
from collections import defaultdict

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))


# =========================================================================
# Step 1: 분석 (slide_index.json 생성)
# =========================================================================
def step_analyze(pptx_path, vol_num, output_dir):
    """PPTX를 분석하여 slide_index.json 생성"""
    from template_extractor import extract_slide_index, TEMPLATE_MAP, TEMPLATE_MAP_VOL3
    from template_matcher import analyze_and_match

    pptx_path = os.path.abspath(pptx_path)
    slide_offset = vol_num * 1000

    if vol_num == 2:
        template_map = TEMPLATE_MAP
    elif vol_num == 3:
        template_map = TEMPLATE_MAP_VOL3
    else:
        print(f"  Vol {vol_num}: 자동 분류 모드")
        from pptx import Presentation
        prs = Presentation(pptx_path)
        total = len(prs.slides)
        template_map = {}
        for i in range(1, total + 1):
            try:
                result = analyze_and_match(pptx_path, i)
                best = result.get('best_match', {})
                template_map[i] = best.get('template', 'T14')
            except Exception:
                template_map[i] = 'T14'
            if i % 10 == 0:
                print(f"    분류 중: {i}/{total}...")
        print(f"    분류 완료: {total}장")

    idx_path = os.path.join(output_dir, "slide_index.json")
    extract_slide_index(
        pptx_path, idx_path,
        template_map=template_map,
        slide_offset=slide_offset,
        source_label=f"vol{vol_num}"
    )

    print(f"  slide_index.json 생성 완료: {idx_path}")
    return idx_path


# =========================================================================
# Step 2: 블록처리 (PPTX 텍스트 → ████)
# =========================================================================
def step_sanitize(pptx_path, output_pptx):
    """PPTX 텍스트를 블록처리"""
    from template_sanitizer import sanitize_pptx_aggressive
    sanitize_pptx_aggressive(pptx_path, output_pptx)
    print(f"  블록처리 완료: {output_pptx}")
    return output_pptx


# =========================================================================
# Step 3: 분할 (1장짜리 PPTX로 쪼개기)
# =========================================================================
def step_split(pptx_path, slides_dir, vol_num):
    """블록처리된 PPTX를 1장짜리로 분할"""
    from template_splitter import split_placeholder
    slide_offset = vol_num * 1000
    split_placeholder(pptx_path, slides_dir, slide_offset)
    count = len([f for f in os.listdir(slides_dir) if f.endswith('.pptx')])
    print(f"  분할 완료: {count}개 파일 → {slides_dir}")
    return count


# =========================================================================
# Step 4: 개별 슬라이드 MD 생성 (S2001.md, S2002.md, ...)
# =========================================================================
def step_generate_slide_md(idx_path, slides_md_dir):
    """각 슬라이드별 개별 MD 파일 생성 (원본 텍스트 포함)"""
    os.makedirs(slides_md_dir, exist_ok=True)

    with open(idx_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    count = 0
    for s in data['slides']:
        sn = s['slide_number']
        tmpl = s['template']
        rm = s.get('role_map', {})
        shapes = s.get('shapes', [])

        lines = []
        lines.append('---slide')

        # 제목: breadcrumb 또는 section_title에서 추출
        title = _extract_title(shapes, rm)
        lines.append(f'# [S{sn:04d}] {title}')
        lines.append(f'template: {tmpl}')
        lines.append(f'ref_slide: {sn}')
        if sn >= 3000:
            lines.append('reference_pptx: templates/placeholder_vol3.pptx')
        lines.append('---')

        # @필드 생성 (원본 텍스트 포함)
        if tmpl != 'T0':
            gm_text = _get_role_text(shapes, rm, 'governing_message')
            bc_text = _get_role_text(shapes, rm, 'breadcrumb')
            if gm_text:
                lines.append(f'@governing_message: {gm_text}')
            if bc_text:
                lines.append(f'@breadcrumb: {bc_text}')

        if tmpl == 'T0':
            st_text = _get_role_text(shapes, rm, 'section_title')
            lines.append(f'@content_1: {st_text or "(섹션 제목)"}')

        elif tmpl in ('T1', 'T2'):
            # content_box/content_shape → content
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
            if content_texts:
                lines.append(f'@content_1: {content_texts[0]}')

            # card_table → 카드
            card_indices = rm.get('card_table', [])
            for ci, idx in enumerate(card_indices, 1):
                shape = shapes[idx] if idx < len(shapes) else {}
                preview = shape.get('table_preview', [])
                card_title = preview[0] if len(preview) > 0 else ''
                card_body = preview[1] if len(preview) > 1 else ''
                lines.append(f'@카드{ci}_제목: {card_title}')
                lines.append(f'@카드{ci}_내용: {card_body}')

        elif tmpl == 'T3':
            heading_texts = _get_role_texts(shapes, rm, ['heading_box'])
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
            if heading_texts:
                lines.append(f'@heading_1: {heading_texts[0]}')
            for i, ct in enumerate(content_texts, 1):
                lines.append(f'@content_{i}: {ct}')

        elif tmpl in ('T4', 'T5', 'T6'):
            # 데이터 테이블
            dtbl_indices = rm.get('data_table', [])
            for di, idx in enumerate(dtbl_indices):
                shape = shapes[idx] if idx < len(shapes) else {}
                table_md = _table_to_markdown(shape)
                if table_md:
                    lines.append('')
                    lines.append(table_md)

            if tmpl == 'T5':
                content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
                for i, ct in enumerate(content_texts, 1):
                    lines.append(f'@content_{i}: {ct}')

        elif tmpl == 'T7':
            heading_texts = _get_role_texts(shapes, rm, ['heading_box'])
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
            for i, ht in enumerate(heading_texts, 1):
                lines.append(f'@heading_{i}: {ht}')
            for i, ct in enumerate(content_texts, 1):
                lines.append(f'@content_{i}: {ct}')
            dtbl_indices = rm.get('data_table', [])
            for idx in dtbl_indices:
                shape = shapes[idx] if idx < len(shapes) else {}
                table_md = _table_to_markdown(shape)
                if table_md:
                    lines.append('')
                    lines.append(table_md)

        elif tmpl == 'T8':
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
            if content_texts:
                lines.append(f'@content_1: {content_texts[0]}')
            else:
                lines.append('@content_1: (이미지 설명 - 수작업)')

        elif tmpl == 'T9':
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape'])
            label_texts = _get_role_texts(shapes, rm, ['label_box', 'label_shape'])
            all_texts = content_texts or label_texts
            for i, ct in enumerate(all_texts, 1):
                lines.append(f'@content_{i}: {ct}')

        else:
            # T14 등 기타
            content_texts = _get_role_texts(shapes, rm, ['content_box', 'content_shape', 'text_content'])
            for i, ct in enumerate(content_texts, 1):
                lines.append(f'@content_{i}: {ct}')

        md_path = os.path.join(slides_md_dir, f'S{sn:04d}.md')
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
        count += 1

    print(f"  개별 MD 생성 완료: {count}개 파일 → {slides_md_dir}")
    return count


def _extract_title(shapes, role_map):
    """슬라이드 제목 추출 (breadcrumb → section_title → 첫 번째 텍스트)"""
    for role in ['breadcrumb', 'section_title', 'heading_box']:
        indices = role_map.get(role, [])
        if indices:
            shape = shapes[indices[0]] if indices[0] < len(shapes) else {}
            text = shape.get('text', '').strip()
            if text:
                return text.split('\n')[0][:60]
    return '(제목 없음)'


def _get_role_text(shapes, role_map, role):
    """특정 역할의 첫 번째 shape 텍스트 반환"""
    indices = role_map.get(role, [])
    if indices:
        shape = shapes[indices[0]] if indices[0] < len(shapes) else {}
        return shape.get('text', '').strip()
    return ''


def _get_role_texts(shapes, role_map, roles):
    """여러 역할의 모든 shape 텍스트를 리스트로 반환"""
    texts = []
    for role in roles:
        for idx in role_map.get(role, []):
            if idx < len(shapes):
                text = shapes[idx].get('text', '').strip()
                if text:
                    texts.append(text)
    return texts


def _table_to_markdown(shape):
    """shape의 table_preview를 마크다운 표로 변환"""
    preview = shape.get('table_preview', [])
    size_str = shape.get('table_size', '')
    if not preview or not size_str:
        return ''

    try:
        nrows, ncols = map(int, size_str.split('x'))
    except ValueError:
        return ''

    # table_preview는 최대 3x3 셀 (row-major)
    display_rows = min(nrows, 3)
    display_cols = min(ncols, 3)

    rows = []
    pi = 0
    for r in range(display_rows):
        row = []
        for c in range(display_cols):
            if pi < len(preview):
                row.append(preview[pi])
                pi += 1
            else:
                row.append('')
        rows.append(row)

    if not rows:
        return ''

    lines = []
    # 헤더
    lines.append('| ' + ' | '.join(rows[0]) + ' |')
    lines.append('|' + '|'.join(['---'] * display_cols) + '|')
    # 데이터 행
    for row in rows[1:]:
        lines.append('| ' + ' | '.join(row) + ' |')
    if nrows > 3:
        lines.append(f'| ... ({nrows}행 x {ncols}열) | ' + ' | '.join(['...'] * (display_cols - 1)) + ' |')

    return '\n'.join(lines)


# =========================================================================
# Step 5: 그룹 MD 생성 (T0~T9.md + GUIDE.md) — 참고용
# =========================================================================
def step_generate_md(idx_path, docs_dir):
    """slide_index.json으로부터 T0~T9.md 및 GUIDE.md 생성"""
    os.makedirs(docs_dir, exist_ok=True)

    with open(idx_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    by_tmpl = defaultdict(list)
    for s in data['slides']:
        if s['template'] != 'T14':
            by_tmpl[s['template']].append(s)

    NAMES = {
        'T0': '구분페이지', 'T1': '카드형 다중', 'T2': '카드+다이어그램',
        'T3': '범위/개요', 'T4': '다중 데이터테이블', 'T5': '테이블+다이어그램',
        'T6': '순수 데이터테이블', 'T7': '프로세스+테이블', 'T8': '이미지중심',
        'T9': '핵심메시지/다이어그램'
    }
    WHEN = {
        'T0': '섹션/챕터 구분', 'T1': '현황/문제점/개선, 비교 분석',
        'T2': '사업 목적, 전략 개요', 'T3': '사업 범위, 비전',
        'T4': '복수 테이블, 일정표', 'T5': '테이블 + 설명',
        'T6': '큰 데이터 표, 인력표', 'T7': '프로세스 흐름 + 데이터',
        'T8': '조직도, 구성도 (이미지 수작업)', 'T9': 'CSF, 핵심 포인트'
    }

    total_snippets = 0
    for tmpl in ['T0', 'T1', 'T2', 'T3', 'T4', 'T5', 'T6', 'T7', 'T8', 'T9']:
        slides = by_tmpl.get(tmpl, [])
        if not slides:
            continue
        L = [f'# {tmpl} — {NAMES.get(tmpl, tmpl)}', '',
             f'**용도**: {WHEN.get(tmpl, "")}', f'**슬라이드 수**: {len(slides)}장', '',
             '## 슬라이드 목록', '',
             '| ref_slide | shapes | 카드 | 테이블 | 콘텐츠 | 라벨 | 이미지 |',
             '|---|---|---|---|---|---|---|']
        for s in slides:
            rm = s.get('role_map', {})
            cards = len(rm.get('card_table', []))
            dtbl = len(rm.get('data_table', []))
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            labels = len(rm.get('label_box', []) + rm.get('label_shape', []))
            imgs = len(rm.get('image', []))
            L.append(f'| {s["slide_number"]} | {s["shape_count"]} | {cards} | {dtbl} | {content} | {labels} | {imgs} |')
        L.extend(['', '---', '', f'개별 MD 파일: `docs/slides/S????.md` 참조'])
        total_snippets += len(slides)

        with open(os.path.join(docs_dir, f'{tmpl}.md'), 'w', encoding='utf-8') as f:
            f.write('\n'.join(L))

    # GUIDE.md
    vol_nums = sorted(set(s['slide_number'] // 1000 for s in data['slides']))
    ref_lines = []
    for v in vol_nums:
        slides_in_vol = [s for s in data['slides'] if s['slide_number'] // 1000 == v]
        mn = min(s['slide_number'] for s in slides_in_vol)
        mx = max(s['slide_number'] for s in slides_in_vol)
        ref_lines.append(f'- ref_slide {mn}~{mx}: vol{v}')

    guide = f"""# PPTX 세로형 제안서 가이드

## 워크플로우
1. RFP 분석 → 템플릿 타입 결정 (T0~T9)
2. docs/slides/S????.md 에서 원본 텍스트 확인
3. proposal-body.md 작성 (S????.md 구조를 따라 @필드 채우기)
4. create_pptx 또는 md2pptx로 최종 PPTX 생성

## 템플릿 MD 파일
| 파일 | 용도 |
|---|---|
{chr(10).join(f'| {t}.md | {NAMES.get(t, t)} — {WHEN.get(t, "")} |' for t in sorted(by_tmpl.keys()))}

## 개별 슬라이드 MD
`docs/slides/S????.md` — 각 슬라이드별 원본 텍스트 포함 스니펫

## 글쓰기 규칙
- 거버닝 메시지: 200자, 카드 제목: 15자, 카드 본문: 300자, 핵심 문구: 50~100자

## 출처 주석
`<!-- [rawdata] 파일명, p.페이지 -->` / `<!-- [ref] 파일명, p.페이지 -->` / `<!-- [AI] 설명 -->`

## 참조 PPTX 번호 체계
{chr(10).join(ref_lines)}
"""
    with open(os.path.join(docs_dir, 'GUIDE.md'), 'w', encoding='utf-8') as f:
        f.write(guide)

    print(f"  그룹 MD 생성 완료: {total_snippets}개 슬라이드, {len(by_tmpl)}개 템플릿 파일")


# =========================================================================
# Step 6: 스타트 프롬프트 자동 생성
# =========================================================================
def step_generate_start_prompt(idx_path, output_dir):
    """프로젝트 폴더를 스캔하여 start_prompt.md 자동 생성"""

    with open(idx_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 볼륨 범위 파악
    slides = data.get('slides', [])
    vol_info = {}
    for s in slides:
        v = s['slide_number'] // 1000
        if v not in vol_info:
            vol_info[v] = {'min': s['slide_number'], 'max': s['slide_number'], 'count': 0}
        vol_info[v]['min'] = min(vol_info[v]['min'], s['slide_number'])
        vol_info[v]['max'] = max(vol_info[v]['max'], s['slide_number'])
        vol_info[v]['count'] += 1

    # 템플릿 분포
    from collections import Counter
    tmpl_counts = Counter(s['template'] for s in slides if s['template'] != 'T14')

    # 폴더 스캔
    def scan_folder(folder_path):
        if not os.path.isdir(folder_path):
            return []
        files = []
        for f in sorted(os.listdir(folder_path)):
            fp = os.path.join(folder_path, f)
            if os.path.isfile(fp) and not f.startswith('.'):
                size_kb = os.path.getsize(fp) // 1024
                files.append((f, size_kb))
        return files

    rfp_files = scan_folder(os.path.join(output_dir, 'rfp'))
    ref_files = scan_folder(os.path.join(output_dir, 'references'))
    raw_files = scan_folder(os.path.join(output_dir, 'rawdata'))

    # 파트 정보 생성
    parts_lines = []
    for v in sorted(vol_info.keys()):
        info = vol_info[v]
        parts_lines.append(
            f"## 파트 {v}: vol{v} ({info['count']}장, ref_slide {info['min']}~{info['max']})\n"
            f"- backbone: `./proposal-backbone-part{v}.md`\n"
            f"- 본문: `./proposal-body-part{v}.md`"
        )

    # 파일 목록 포맷
    def format_file_list(files, folder):
        if not files:
            return f"- `{folder}/` — (파일 없음)\n"
        lines = []
        for fname, size in files:
            lines.append(f"- `{folder}/{fname}` ({size}KB)")
        return '\n'.join(lines) + '\n'

    # 스타트 프롬프트 조립
    prompt = f'''"{ output_dir }" 폴더를 기반으로 제안서 본문을 작성합니다.
이 프롬프트는 **backbone + body MD 작성까지만** 진행합니다. PPTX 변환은 별도 CLI 도구(md2pptx)로 수행합니다.

# 1. 반드시 읽을 파일

| 파일 | 용도 |
|---|---|
| `rfp/사업개요.md` | 사업 기본정보, 산출물 목록 |
| `rfp/목차.md` | 전체 목차, 배점, 작성지침 |
| `docs/GUIDE.md` | 워크플로우, 글쓰기 규칙, 출처 주석 |
| `docs/slides/S????.md` | ★ 각 슬라이드별 원본 텍스트 + @필드 구조. 이 구조를 그대로 따를 것 |
| `docs/T0.md` ~ `T9.md` | 템플릿별 슬라이드 목록 (참고용) |

## 프로젝트 내 참고자료

### rfp/
{format_file_list(rfp_files, 'rfp')}
### references/
{format_file_list(ref_files, 'references')}
### rawdata/
{format_file_list(raw_files, 'rawdata')}
# 2. 슬라이드 현황

| 볼륨 | 슬라이드 수 | 번호 범위 |
|---|---|---|
'''

    for v in sorted(vol_info.keys()):
        info = vol_info[v]
        prompt += f'| vol{v} | {info["count"]}장 | S{info["min"]}~S{info["max"]} |\n'

    prompt += f'''
템플릿 분포: {', '.join(f'{t}:{c}장' for t, c in tmpl_counts.most_common())}

# 3. 작업 대상 파트

사용자가 지정한 파트를 작성합니다. 작업 시작 전 어떤 파트를 진행할지 사용자에게 확인하세요.

{chr(10).join(parts_lines)}

# 4. 작업 순서

## Step 1. backbone → `./proposal-backbone-partN.md`
- rfp/목차.md를 읽고 전체 슬라이드 계획표 작성 (슬라이드 번호, 제목, 템플릿, ref_slide, 내용 요약)
- 사용자 확인 후 Step 2 진행

## Step 2. 본문 → `./proposal-body-partN.md`
- `docs/slides/S????.md`에서 @필드 구조 확인 → references/ PDF 참조하여 내용 채우기
- 대단원 완료 시마다 파일 저장 + 사용자에게 진행상황 알림
- 전체 완료되면 사용자에게 최종 확인 요청

**PPTX 변환은 이 프롬프트 범위 밖입니다.** 본문 작성 완료 후 별도 CLI로 변환:
```
python -m md2pptx proposal-body-partN.md -t templates/slides -o output/partN.pptx --continue-on-error -v
```

# 5. 글쓰기 규칙

- 거버닝 메시지: 200자 / 카드 제목: 15자 / 카드 본문: 300자 / 핵심 문구: 50~100자
- **빈 슬라이드 금지** — 모든 @필드에 실질적 텍스트 필수
- **T0(구분페이지) 이외** 모든 슬라이드에 governing_message + 본문 필수
- 이미지 작업 불필요, 색상 마커 사용 안 함, 순수 확장 MD 문법
- 출처 주석은 **본문(@필드)에 넣지 말 것** → `@note` 필드에 내용 요약 + 출처 함께 기재

## 출처 주석 형식
- `<!-- [rawdata] 파일명, p.페이지 -->` — 원문 인용
- `<!-- [ref] 파일명, p.페이지 -->` — 참고
- `<!-- [AI] 설명 -->` — AI 생성 내용

## 분량 기준 (목표의 2배)
- 목차의 p수는 최종 목표이며, **작성 시에는 2배**로 슬라이드 작성
- 1p → 2장 / 3~5p → 6~10장 / 6p+ → 12장 이상

## 템플릿 배정 가이드

| 내용 유형 | 템플릿 | 글 형태 |
|---|---|---|
| 섹션 구분 | T0 | 제목 1줄 |
| 현황/문제점/개선 | T1 | 카드 2~6개 |
| 목적/전략 | T2 | 카드 + 세분화 |
| 범위/비전 | T3 | 거버닝메시지 + 영역 |
| 복수 테이블 | T4 | 마크다운 테이블 x2+ |
| 테이블+설명 | T5 | 테이블 + 다이어그램 텍스트 |
| 큰 데이터 표 | T6 | 마크다운 테이블 |
| 프로세스+테이블 | T7 | 단계별 설명 + 테이블 |
| 이미지 중심 | T8 | 이미지 + 캡션 |
| 핵심 포인트 | T9 | 핵심 문구 3~6개 |
'''

    prompt_path = os.path.join(output_dir, 'start_prompt.md')
    with open(prompt_path, 'w', encoding='utf-8') as f:
        f.write(prompt)

    print(f"  스타트 프롬프트 생성 완료: {prompt_path}")
    return prompt_path


# =========================================================================
# slide_index.json 머지 헬퍼
# =========================================================================
def merge_slide_index(target_path, new_data_path, vol_num):
    """기존 slide_index.json에 새 볼륨 데이터를 머지"""
    with open(target_path, 'r', encoding='utf-8') as f:
        existing = json.load(f)
    with open(new_data_path, 'r', encoding='utf-8') as f:
        new_data = json.load(f)

    vol_min = vol_num * 1000
    vol_max = (vol_num + 1) * 1000
    existing['slides'] = [s for s in existing['slides']
                          if not (vol_min <= s['slide_number'] < vol_max)]
    existing['slides'].extend(new_data['slides'])
    existing['slides'].sort(key=lambda s: s['slide_number'])

    with open(target_path, 'w', encoding='utf-8') as f:
        json.dump(existing, f, ensure_ascii=False, indent=2)

    return target_path


# =========================================================================
# 파이프라인 (server.py에서도 호출)
# =========================================================================
def run_pipeline(pptx_path, output_dir, vol_num, merge_to=None):
    """전체 파이프라인 실행. 결과 요약 문자열 반환."""
    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)

    if not os.path.exists(pptx_path):
        return f"오류: 파일 없음: {pptx_path}"

    docs_dir = os.path.join(output_dir, "docs")
    slides_md_dir = os.path.join(output_dir, "docs", "slides")
    slides_dir = os.path.join(output_dir, "templates", "slides")
    templates_dir = os.path.join(output_dir, "templates")
    temp_dir = os.path.join(output_dir, "_temp")

    for d in [docs_dir, slides_md_dir, slides_dir, templates_dir, temp_dir]:
        os.makedirs(d, exist_ok=True)

    try:
        # Step 1: 분석 (원본 텍스트 포함 slide_index.json)
        print("[Step 1/6] 슬라이드 분석...")
        temp_idx = step_analyze(pptx_path, vol_num, temp_dir)

        # 원본 텍스트 slide_index 보존 (MD 생성용)
        raw_idx = os.path.join(temp_dir, "slide_index_raw.json")
        shutil.copy2(temp_idx, raw_idx)

        # slide_index.json을 templates/에 배치 (블록처리된 버전)
        target_idx = os.path.join(templates_dir, "slide_index.json")
        if os.path.exists(target_idx):
            merge_slide_index(target_idx, temp_idx, vol_num)
            print(f"  slide_index.json 머지 완료: {target_idx}")
        else:
            shutil.copy2(temp_idx, target_idx)
            print(f"  slide_index.json 복사 완료: {target_idx}")

        # Step 2: 블록처리
        print("\n[Step 2/6] 텍스트 블록처리...")
        sanitized_pptx = os.path.join(temp_dir, "sanitized.pptx")
        step_sanitize(pptx_path, sanitized_pptx)

        # Step 3: 분할
        print("\n[Step 3/6] 슬라이드 분할...")
        step_split(sanitized_pptx, slides_dir, vol_num)

        # Step 4: 개별 슬라이드 MD 생성 (원본 텍스트 사용!)
        print("\n[Step 4/6] 개별 슬라이드 MD 생성...")
        step_generate_slide_md(raw_idx, slides_md_dir)

        # Step 5: 그룹 MD 생성 (T0~T9 + GUIDE)
        print("\n[Step 5/6] 그룹 MD 생성...")
        step_generate_md(target_idx, docs_dir)

        # Step 6: 스타트 프롬프트 생성
        print("\n[Step 6/6] 스타트 프롬프트 생성...")
        step_generate_start_prompt(target_idx, output_dir)

        # --merge-to 옵션
        if merge_to:
            merge_to = os.path.abspath(merge_to)
            if os.path.exists(merge_to):
                merge_slide_index(merge_to, temp_idx, vol_num)
                print(f"\n  외부 slide_index.json 머지: {merge_to}")
            else:
                os.makedirs(os.path.dirname(merge_to), exist_ok=True)
                shutil.copy2(temp_idx, merge_to)
                print(f"\n  외부 slide_index.json 복사: {merge_to}")

        # 정리
        shutil.rmtree(temp_dir, ignore_errors=True)

        docs_count = len([f for f in os.listdir(docs_dir) if f.endswith('.md')])
        slides_md_count = len([f for f in os.listdir(slides_md_dir) if f.endswith('.md')])
        slides_count = len([f for f in os.listdir(slides_dir) if f.endswith('.pptx')])

        summary = (
            f"\n{'='*60}\n"
            f"완료!\n"
            f"  출력: {output_dir}\n"
            f"  docs/         : {docs_count}개 MD (GUIDE + T0~T9)\n"
            f"  docs/slides/  : {slides_md_count}개 MD (개별 슬라이드)\n"
            f"  templates/slides/ : {slides_count}개 PPTX (블록처리)\n"
            f"  start_prompt.md   : 2단계 글쓰기용 스타트 프롬프트\n"
            f"  볼륨: vol{vol_num} (S{vol_num}001~)\n"
            f"{'='*60}"
        )
        print(summary)
        return summary

    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        error_msg = f"오류: {type(e).__name__}: {e}"
        print(error_msg)
        return error_msg


# =========================================================================
# 메인
# =========================================================================
def main():
    parser = argparse.ArgumentParser(description='PPT Block Maker — 원본 PPTX → 프로젝트 ready 변환')
    parser.add_argument('pptx', help='원본 PPTX 파일 경로')
    parser.add_argument('output', help='출력 프로젝트 폴더 (docs/, templates/ 생성)')
    parser.add_argument('--vol', type=int, default=0, help='볼륨 번호 (2=II권, 3=III권, 0=자동분류)')
    parser.add_argument('--merge-to', help='외부 slide_index.json 경로 (추가 머지)')

    args = parser.parse_args()

    print(f"\n{'='*60}")
    print(f"PPT Block Maker")
    print(f"  입력: {os.path.abspath(args.pptx)}")
    print(f"  출력: {os.path.abspath(args.output)}")
    print(f"  볼륨: {args.vol}")
    if args.merge_to:
        print(f"  머지: {os.path.abspath(args.merge_to)}")
    print(f"{'='*60}\n")

    run_pipeline(args.pptx, args.output, args.vol, args.merge_to)


if __name__ == '__main__':
    main()
