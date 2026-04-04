# -*- coding: utf-8 -*-
"""
PPT Block Maker MCP Server

원본 PPTX를 분석 → 블록처리 → 분할 → MD 생성하여
pptx-vertical-writer에서 바로 사용할 수 있는 프로젝트 폴더를 만듭니다.

실행: python server.py
"""
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
PLUGIN_DIR = Path(r"D:\00work\260403-a3verticle-ppt")
sys.path.insert(0, str(PLUGIN_DIR / "src"))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("pptx-blockmaker")


@mcp.tool()
def prepare_project(
    pptx_path: str,
    output_dir: str,
    vol: int = 0
) -> str:
    """원본 PPTX를 분석하여 프로젝트 폴더를 준비합니다.

    전체 파이프라인을 실행합니다:
    1. 슬라이드 분석 (shape 역할 분류)
    2. 텍스트 블록처리 (████)
    3. 1장짜리 PPTX로 분할
    4. 템플릿 MD 파일 생성 (T0~T9.md + GUIDE.md)

    Args:
        pptx_path: 원본 PPTX 파일 경로
        output_dir: 출력 프로젝트 폴더 경로 (docs/, templates/slides/ 생성)
        vol: 볼륨 번호 (2=II권, 3=III권, 0=자동분류)

    Returns:
        생성 결과 요약
    """
    # run.py의 함수들을 import
    sys.path.insert(0, str(BASE_DIR))
    from run import step_analyze, step_sanitize, step_split, step_generate_md

    import os
    import json
    import shutil

    pptx_path = os.path.abspath(pptx_path)
    output_dir = os.path.abspath(output_dir)

    if not os.path.exists(pptx_path):
        return f"오류: 파일 없음: {pptx_path}"

    docs_dir = os.path.join(output_dir, "docs")
    slides_dir = os.path.join(output_dir, "templates", "slides")
    temp_dir = os.path.join(str(BASE_DIR), "_temp")

    for d in [docs_dir, slides_dir, temp_dir]:
        os.makedirs(d, exist_ok=True)

    try:
        # Step 1: 분석
        idx_path = os.path.join(temp_dir, "slide_index.json")
        step_analyze(pptx_path, vol, temp_dir)

        # Step 2: 블록처리
        sanitized = os.path.join(temp_dir, "sanitized.pptx")
        step_sanitize(pptx_path, sanitized)

        # Step 3: 분할
        count = step_split(sanitized, slides_dir, vol)

        # Step 4: MD 생성 (플러그인 slide_index에 머지)
        plugin_idx = str(PLUGIN_DIR / "templates" / "slide_index.json")
        if os.path.exists(plugin_idx):
            with open(plugin_idx, 'r', encoding='utf-8') as f:
                existing = json.load(f)
            with open(idx_path, 'r', encoding='utf-8') as f:
                new_data = json.load(f)
            vol_min = vol * 1000
            vol_max = (vol + 1) * 1000
            existing['slides'] = [s for s in existing['slides']
                                  if not (vol_min <= s['slide_number'] < vol_max)]
            existing['slides'].extend(new_data['slides'])
            existing['slides'].sort(key=lambda s: s['slide_number'])
            with open(plugin_idx, 'w', encoding='utf-8') as f:
                json.dump(existing, f, ensure_ascii=False, indent=2)
            step_generate_md(plugin_idx, docs_dir)
        else:
            shutil.copy2(idx_path, plugin_idx)
            step_generate_md(idx_path, docs_dir)

        # 정리
        shutil.rmtree(temp_dir, ignore_errors=True)

        docs_count = len([f for f in os.listdir(docs_dir) if f.endswith('.md')])
        slides_count = len([f for f in os.listdir(slides_dir) if f.endswith('.pptx')])

        return (
            f"프로젝트 준비 완료!\n"
            f"  출력: {output_dir}\n"
            f"  docs/: {docs_count}개 MD 파일\n"
            f"  slides/: {slides_count}개 PPTX 파일\n"
            f"  볼륨: vol{vol} (S{vol}001~)"
        )

    except Exception as e:
        shutil.rmtree(temp_dir, ignore_errors=True)
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def add_volume(
    pptx_path: str,
    output_dir: str,
    vol: int
) -> str:
    """기존 프로젝트에 새 볼륨의 슬라이드를 추가합니다.

    이미 docs/와 templates/slides/가 있는 프로젝트에
    새 PPTX 볼륨을 추가 분석하여 병합합니다.

    Args:
        pptx_path: 추가할 원본 PPTX 파일 경로
        output_dir: 기존 프로젝트 폴더 경로
        vol: 볼륨 번호 (1~9)

    Returns:
        추가 결과 요약
    """
    return prepare_project(pptx_path, output_dir, vol)


if __name__ == "__main__":
    mcp.run()
