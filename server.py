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
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP

mcp = FastMCP("pptx-blockmaker")


@mcp.tool()
def prepare_project(
    pptx_path: str,
    output_dir: str,
    vol: int = 0,
    merge_to: str = ""
) -> str:
    """원본 PPTX를 분석하여 프로젝트 폴더를 준비합니다.

    전체 파이프라인을 실행합니다:
    1. 슬라이드 분석 (shape 역할 분류)
    2. 텍스트 블록처리 (████)
    3. 1장짜리 PPTX로 분할
    4. 개별 슬라이드 MD 생성 (S????.md — 원본 텍스트 포함)
    5. 그룹 MD 생성 (T0~T9.md + GUIDE.md)

    Args:
        pptx_path: 원본 PPTX 파일 경로
        output_dir: 출력 프로젝트 폴더 경로
        vol: 볼륨 번호 (2=II권, 3=III권, 0=자동분류)
        merge_to: 외부 slide_index.json 경로 (선택, 추가 머지)

    Returns:
        생성 결과 요약
    """
    sys.path.insert(0, str(BASE_DIR))
    from run import run_pipeline
    return run_pipeline(pptx_path, output_dir, vol, merge_to or None)


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
