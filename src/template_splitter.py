"""
블록처리된 placeholder PPTX를 1장짜리 개별 파일로 분할.

사용법:
    python src/template_splitter.py <placeholder.pptx> <output_dir> [slide_offset]

예시:
    python src/template_splitter.py templates/placeholder_vol2.pptx templates/slides/ 0
    python src/template_splitter.py templates/placeholder_vol3.pptx templates/slides/ 1000
"""
import os
import sys
import shutil
import time
import comtypes.client


def split_placeholder(pptx_path, output_dir, slide_offset=0):
    """placeholder PPTX를 1장짜리 개별 파일로 분할"""
    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)
    os.makedirs(output_dir, exist_ok=True)

    pp = comtypes.client.CreateObject('Powerpoint.Application')
    pp.Visible = 1

    try:
        prs = pp.Presentations.Open(pptx_path, WithWindow=False)
        total = prs.Slides.Count
        print(f'Source: {pptx_path} ({total} slides)')
        print(f'Output: {output_dir}')
        print(f'Offset: {slide_offset}')

        for si in range(1, total + 1):
            slide_num = si + slide_offset
            out_file = os.path.join(output_dir, f'S{slide_num:04d}.pptx')

            # 원본을 임시 파일로 복사
            tmp = out_file + '.tmp'
            shutil.copy2(pptx_path, tmp)

            # 임시 파일 열기
            tmp_prs = pp.Presentations.Open(os.path.abspath(tmp), WithWindow=False)

            # 해당 슬라이드만 남기고 나머지 삭제 (뒤에서부터)
            for di in range(tmp_prs.Slides.Count, 0, -1):
                if di != si:
                    tmp_prs.Slides(di).Delete()

            # SaveAs
            tmp_prs.SaveAs(os.path.abspath(out_file), 24)
            tmp_prs.Close()

            # 임시 파일 삭제
            try:
                os.remove(tmp)
            except Exception:
                pass

            if si % 10 == 0:
                print(f'  {si}/{total}...')

        prs.Close()
        print(f'\nDone: {total} files in {output_dir}')

    finally:
        try:
            pp.Quit()
        except Exception:
            pass


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('Usage: python template_splitter.py <placeholder.pptx> <output_dir> [slide_offset]')
        sys.exit(1)

    pptx = sys.argv[1]
    outdir = sys.argv[2]
    offset = int(sys.argv[3]) if len(sys.argv) > 3 else 0

    split_placeholder(pptx, outdir, offset)
