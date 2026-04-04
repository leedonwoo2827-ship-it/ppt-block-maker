# -*- coding: utf-8 -*-
"""
PPT Block Maker — 원본 PPTX → 프로젝트 ready 변환 도구

사용법:
    python run.py <원본.pptx> <출력_프로젝트_폴더> [--vol 2] [--template-map auto]

예시:
    python run.py input/기술부문.pptx C:/Users/.../260404-2 --vol 2
    python run.py input/사업관리부문.pptx C:/Users/.../260404-2 --vol 3

결과:
    출력폴더/
    ├── docs/          ← GUIDE.md + T0~T9.md (스니펫 포함)
    ├── templates/
    │   └── slides/    ← S2001.pptx, S2002.pptx, ... (1장짜리)
"""
import argparse
import json
import os
import shutil
import sys
import time
from pathlib import Path
from collections import defaultdict

# 플러그인 src 참조
PLUGIN_DIR = Path(r"D:\00work\260403-a3verticle-ppt")
sys.path.insert(0, str(PLUGIN_DIR / "src"))


# =========================================================================
# Step 1: 분석 (slide_index.json 생성)
# =========================================================================
def step_analyze(pptx_path, vol_num, output_dir):
    """PPTX를 분석하여 slide_index.json 생성"""
    from template_extractor import extract_slide_index, TEMPLATE_MAP, TEMPLATE_MAP_VOL3, TEMPLATE_NAMES
    from template_matcher import analyze_and_match

    pptx_path = os.path.abspath(pptx_path)
    slide_offset = vol_num * 1000

    # 기존 맵이 있으면 사용, 없으면 자동 분류
    if vol_num == 2:
        template_map = TEMPLATE_MAP
    elif vol_num == 3:
        template_map = TEMPLATE_MAP_VOL3
    else:
        # 자동 분류
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

    # slide_index 텍스트 블록처리
    from template_sanitizer import sanitize_slide_index
    sanitize_slide_index(idx_path)

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
# Step 4: MD 생성 (T0~T9.md + GUIDE.md)
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

    def build_snippet(s, tmpl):
        rm = s.get('role_map', {})
        sn = s['slide_number']
        lines = ['---slide', '# [SXXX] (제목)', f'template: {tmpl}', f'ref_slide: {sn}']
        if sn >= 3000:
            lines.append('reference_pptx: templates/placeholder_vol3.pptx')
        lines.append('---')
        if tmpl != 'T0':
            lines.extend(['@governing_message: (핵심 메시지, 200자)', '@breadcrumb: (섹션 경로)'])
        if tmpl == 'T0':
            lines.append('@content_1: (섹션 제목)')
        elif tmpl in ('T1', 'T2'):
            cards = len(rm.get('card_table', []))
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            if content > 0:
                lines.append('@content_1: (상단 요약)')
            for i in range(1, cards + 1):
                lines.extend([f'@카드{i}_제목: (카드{i} 제목, 15자)', f'@카드{i}_내용: (카드{i} 본문, 300자)'])
        elif tmpl == 'T3':
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            lines.append('@heading_1: (핵심 문구)')
            for i in range(1, min(max(content, 3) + 1, 7)):
                lines.append(f'@content_{i}: (영역{i} 설명)')
        elif tmpl in ('T4', 'T5', 'T6'):
            dtbl = len(rm.get('data_table', []))
            lines.append('')
            for i in range(1, max(dtbl, 1) + 1):
                lines.extend(['| 항목 | 내용 | 비고 |', '|---|---|---|', '| ... | ... | ... |'])
                if i < max(dtbl, 1):
                    lines.append('')
            if tmpl == 'T5':
                content = len(rm.get('content_box', []) + rm.get('content_shape', []))
                for i in range(1, min(content + 1, 4)):
                    lines.append(f'@content_{i}: (설명 텍스트)')
        elif tmpl == 'T7':
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            headings = len(rm.get('heading_box', []))
            for i in range(1, max(headings, 1) + 1):
                lines.append(f'@heading_{i}: (단계{i} 제목)')
            for i in range(1, min(max(content, 1) + 1, 4)):
                lines.append(f'@content_{i}: (단계{i} 설명)')
            lines.extend(['', '| 항목 | 내용 | 비고 |', '|---|---|---|', '| ... | ... | ... |'])
        elif tmpl == 'T8':
            lines.append('@content_1: (이미지 설명 - 수작업)')
        elif tmpl == 'T9':
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            labels = len(rm.get('label_box', []) + rm.get('label_shape', []))
            n = max(content, labels, 3)
            for i in range(1, min(n + 1, 7)):
                lines.append(f'@content_{i}: (핵심 문구{i}, 50~100자)')
        return '\n'.join(lines)

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
        L.extend(['', '---', '', '## 복사용 스니펫', ''])
        for s in slides:
            sn = s['slide_number']
            rm = s.get('role_map', {})
            cards = len(rm.get('card_table', []))
            dtbl = len(rm.get('data_table', []))
            content = len(rm.get('content_box', []) + rm.get('content_shape', []))
            L.extend([
                f'### ref_slide: {sn} — shapes:{s["shape_count"]}, 카드:{cards}, 테이블:{dtbl}, 콘텐츠:{content}',
                '', '```markdown', build_snippet(s, tmpl), '```', ''
            ])
            total_snippets += 1

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
2. T?.md에서 ref_slide 선택 (카드 수, 테이블 수 매칭)
3. 스니펫 복사 → proposal-body-extended.md 조립
4. @필드 채우기 + 출처 주석
5. create_pptx(md_file=..., project_dir=...) 호출

## 템플릿 MD 파일
| 파일 | 용도 |
|---|---|
{chr(10).join(f'| {t}.md | {NAMES.get(t, t)} — {WHEN.get(t, "")} |' for t in sorted(by_tmpl.keys()))}

## 글쓰기 규칙
- 거버닝 메시지: 200자, 카드 제목: 15자, 카드 본문: 300자, 핵심 문구: 50~100자

## 출처 주석
`<!-- [rawdata] 파일명, p.페이지 -->` / `<!-- [ref] 파일명, p.페이지 -->` / `<!-- [AI] 설명 -->`

## 참조 PPTX 번호 체계
{chr(10).join(ref_lines)}
"""
    with open(os.path.join(docs_dir, 'GUIDE.md'), 'w', encoding='utf-8') as f:
        f.write(guide)

    print(f"  MD 생성 완료: {total_snippets}개 스니펫, {len(by_tmpl)}개 템플릿 파일")


# =========================================================================
# 메인: 전체 파이프라인
# =========================================================================
def main():
    parser = argparse.ArgumentParser(description='PPT Block Maker — 원본 PPTX → 프로젝트 ready 변환')
    parser.add_argument('pptx', help='원본 PPTX 파일 경로')
    parser.add_argument('output', help='출력 프로젝트 폴더 (docs/, templates/slides/ 생성)')
    parser.add_argument('--vol', type=int, default=0, help='볼륨 번호 (2=II권, 3=III권, 0=자동분류)')
    args = parser.parse_args()

    pptx_path = os.path.abspath(args.pptx)
    output_dir = os.path.abspath(args.output)
    vol_num = args.vol

    if not os.path.exists(pptx_path):
        print(f"오류: 파일 없음: {pptx_path}")
        sys.exit(1)

    docs_dir = os.path.join(output_dir, "docs")
    slides_dir = os.path.join(output_dir, "templates", "slides")
    temp_dir = os.path.join(os.path.dirname(pptx_path), "_temp")
    os.makedirs(docs_dir, exist_ok=True)
    os.makedirs(slides_dir, exist_ok=True)
    os.makedirs(temp_dir, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"PPT Block Maker")
    print(f"  입력: {pptx_path}")
    print(f"  출력: {output_dir}")
    print(f"  볼륨: {vol_num}")
    print(f"{'='*60}\n")

    # Step 1: 분석
    print("[Step 1/4] 슬라이드 분석...")
    idx_path = os.path.join(temp_dir, "slide_index.json")
    step_analyze(pptx_path, vol_num, temp_dir)

    # Step 2: 블록처리
    print("\n[Step 2/4] 텍스트 블록처리...")
    sanitized_pptx = os.path.join(temp_dir, "sanitized.pptx")
    step_sanitize(pptx_path, sanitized_pptx)

    # Step 3: 분할
    print("\n[Step 3/4] 슬라이드 분할...")
    step_split(sanitized_pptx, slides_dir, vol_num)

    # Step 4: MD 생성
    print("\n[Step 4/4] 템플릿 MD 생성...")

    # slide_index를 플러그인 templates/에도 머지
    plugin_idx = str(PLUGIN_DIR / "templates" / "slide_index.json")
    if os.path.exists(plugin_idx):
        with open(plugin_idx, 'r', encoding='utf-8') as f:
            existing = json.load(f)
        with open(idx_path, 'r', encoding='utf-8') as f:
            new_data = json.load(f)
        # 같은 번호대 슬라이드 제거 후 추가
        vol_min = vol_num * 1000
        vol_max = (vol_num + 1) * 1000
        existing['slides'] = [s for s in existing['slides']
                              if not (vol_min <= s['slide_number'] < vol_max)]
        existing['slides'].extend(new_data['slides'])
        existing['slides'].sort(key=lambda s: s['slide_number'])
        with open(plugin_idx, 'w', encoding='utf-8') as f:
            json.dump(existing, f, ensure_ascii=False, indent=2)
        merged_idx = plugin_idx
        print(f"  플러그인 slide_index.json 머지 완료")
    else:
        shutil.copy2(idx_path, plugin_idx)
        merged_idx = plugin_idx

    step_generate_md(merged_idx, docs_dir)

    # 임시 파일 정리
    shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"\n{'='*60}")
    print(f"완료!")
    print(f"  docs/     : {len(os.listdir(docs_dir))}개 파일")
    print(f"  slides/   : {len([f for f in os.listdir(slides_dir) if f.endswith('.pptx')])}개 파일")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
