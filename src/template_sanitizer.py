"""
참조 PPTX의 모든 텍스트를 블록문자(████)로 치환하여
민감 정보 없는 깨끗한 템플릿 PPTX를 생성.

사용법:
    python src/template_sanitizer.py <원본.pptx> <출력.pptx>
    python src/template_sanitizer.py <원본.pptx> <slide_index.json> <출력.pptx>
"""
import json
import os
import shutil
import sys
import comtypes.client

# 블록처리하지 않을 텍스트 (구조 요소)
KEEP_TEXTS = {
    'Chapter', 'Governing\nMessage', 'Governing Message',
}

# 페이지 번호 패턴
PAGE_NUMBER_PREFIXES = ('Ⅱ -', 'Ⅲ -', 'Ⅳ -', '‹#›')


def make_block_text(length):
    """원본 길이에 비례하는 블록 문자(████) 생성"""
    if length <= 5:
        return '████'
    elif length <= 15:
        return '████ ████'
    elif length <= 30:
        return '████ ████████ ████'
    else:
        blocks_per_line = 5
        num_lines = min(length // 40 + 1, 5)
        line = '████ ' * blocks_per_line
        return '\n'.join([line.strip()] * num_lines)


def should_skip_text(text):
    """블록처리를 건너뛸 텍스트인지 판단"""
    t = text.strip()
    if not t or len(t) <= 1:
        return True
    if t in KEEP_TEXTS:
        return True
    for prefix in PAGE_NUMBER_PREFIXES:
        if t.startswith(prefix):
            return True
    # 이미 블록처리된 텍스트
    if all(c in '████ \n' for c in t):
        return True
    return False


def sanitize_shape_text(shape, stats):
    """shape의 텍스트를 블록처리"""
    try:
        if shape.HasTextFrame:
            text = shape.TextFrame.TextRange.Text
            if not should_skip_text(text):
                shape.TextFrame.TextRange.Text = make_block_text(len(text))
                stats['replaced'] += 1
            else:
                stats['skipped'] += 1
    except Exception:
        stats['error'] += 1

    try:
        if shape.HasTable:
            table = shape.Table
            for ri in range(1, table.Rows.Count + 1):
                for ci in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(ri, ci)
                        text = cell.Shape.TextFrame.TextRange.Text
                        if not should_skip_text(text):
                            cell.Shape.TextFrame.TextRange.Text = make_block_text(len(text))
                            stats['replaced'] += 1
                    except Exception:
                        pass
    except Exception:
        pass

    # 그룹 내부 shape도 재귀 처리
    try:
        if shape.Type == 6:  # msoGroup
            for i in range(1, shape.GroupItems.Count + 1):
                sanitize_shape_text(shape.GroupItems(i), stats)
    except Exception:
        pass


def sanitize_pptx_aggressive(ref_pptx_path, output_pptx_path):
    """모든 슬라이드의 모든 shape 텍스트를 블록처리"""

    ref_pptx_path = os.path.abspath(ref_pptx_path)
    output_pptx_path = os.path.abspath(output_pptx_path)

    shutil.copy2(ref_pptx_path, output_pptx_path)

    pp = comtypes.client.CreateObject('Powerpoint.Application')
    pp.Visible = 1

    stats = {'replaced': 0, 'skipped': 0, 'error': 0}

    try:
        prs = pp.Presentations.Open(output_pptx_path, WithWindow=False)
        total = prs.Slides.Count

        # 1. 모든 슬라이드의 모든 shape 처리
        for sn in range(1, total + 1):
            slide = prs.Slides(sn)
            for si in range(1, slide.Shapes.Count + 1):
                sanitize_shape_text(slide.Shapes(si), stats)
            if sn % 20 == 0:
                print(f'  Slides: {sn}/{total}...')

        print(f'  Slides: {total}/{total} done')

        # 2. 슬라이드 마스터/레이아웃의 텍스트도 처리
        master_count = 0
        try:
            # Designs > SlideMaster 경로로 접근
            for di in range(1, prs.Designs.Count + 1):
                master = prs.Designs(di).SlideMaster
                for si in range(1, master.Shapes.Count + 1):
                    sanitize_shape_text(master.Shapes(si), stats)
                    master_count += 1
                # 레이아웃
                for li in range(1, master.CustomLayouts.Count + 1):
                    layout = master.CustomLayouts(li)
                    for si in range(1, layout.Shapes.Count + 1):
                        sanitize_shape_text(layout.Shapes(si), stats)
                        master_count += 1
        except Exception as e:
            print(f'  Master processing error: {e}')

        print(f'  Master/Layout shapes: {master_count}')

        prs.Save()
        prs.Close()

        print(f'\nSanitized: {output_pptx_path}')
        print(f'  Replaced: {stats["replaced"]}')
        print(f'  Skipped: {stats["skipped"]}')
        print(f'  Errors: {stats["error"]}')

    finally:
        try:
            pp.Quit()
        except Exception:
            pass


def sanitize_slide_index(slide_index_path, output_path=None):
    """slide_index.json의 텍스트 필드도 블록처리"""
    with open(slide_index_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    for slide in data.get('slides', []):
        for shape in slide.get('shapes', []):
            text = shape.get('text', '')
            length = shape.get('text_length', len(text))

            if text and not should_skip_text(text):
                shape['text'] = make_block_text(min(length, 100))

            if 'table_preview' in shape:
                shape['table_preview'] = [
                    make_block_text(len(cell)) if cell and not should_skip_text(cell) else cell
                    for cell in shape['table_preview']
                ]

    out = output_path or slide_index_path
    with open(out, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f'Sanitized slide_index: {out}')


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage:')
        print('  python template_sanitizer.py <ref.pptx> <output.pptx>')
        print('  python template_sanitizer.py <ref.pptx> <slide_index.json> <output.pptx>')
        sys.exit(1)

    if sys.argv[2].endswith('.json'):
        # 구버전 호환: slide_index 기반
        ref_pptx = sys.argv[1]
        slide_idx = sys.argv[2]
        output = sys.argv[3] if len(sys.argv) > 3 else 'output_placeholder.pptx'
        sanitize_pptx_aggressive(ref_pptx, output)
        sanitize_slide_index(slide_idx)
    else:
        # 신버전: slide_index 없이 직접 처리
        ref_pptx = sys.argv[1]
        output = sys.argv[2]
        sanitize_pptx_aggressive(ref_pptx, output)
