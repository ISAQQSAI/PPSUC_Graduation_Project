#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
LATEX_DIR = ROOT / "latex-template"
MAIN_TEX = LATEX_DIR / "main.tex"
COMPILED_PDF = LATEX_DIR / "main.pdf"
DEFAULT_OUTPUT = ROOT / "main.pdf"


def ensure_compiled_pdf() -> str:
    latexmk = shutil.which("latexmk")
    if latexmk:
        subprocess.run(
            [latexmk, "-xelatex", "-interaction=nonstopmode", "-halt-on-error", MAIN_TEX.name],
            cwd=LATEX_DIR,
            check=True,
        )
        return "compiled"

    xelatex = shutil.which("xelatex")
    if xelatex is not None:
        for _ in range(2):
            subprocess.run(
                [xelatex, "-interaction=nonstopmode", "-halt-on-error", MAIN_TEX.name],
                cwd=LATEX_DIR,
                check=True,
            )
        return "compiled"

    if COMPILED_PDF.exists():
        return "existing"

    raise FileNotFoundError("未找到 latexmk 或 xelatex，且没有可复用的现有 PDF 产物。")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export PDF from the PPSUC LaTeX template."
    )
    parser.add_argument(
        "output",
        nargs="?",
        default=str(DEFAULT_OUTPUT),
        help="output .pdf path, default: ./main.pdf",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    output = Path(args.output).resolve()

    if not MAIN_TEX.exists():
        print(f"找不到主模板文件: {MAIN_TEX}", file=sys.stderr)
        return 1

    try:
        build_status = ensure_compiled_pdf()
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 1
    except subprocess.CalledProcessError as exc:
        tool = Path(exc.cmd[0]).name if exc.cmd else "latex"
        print(f"{tool} 编译失败。", file=sys.stderr)
        return exc.returncode or 1

    if not COMPILED_PDF.exists():
        print(f"未找到编译产物: {COMPILED_PDF}", file=sys.stderr)
        return 1

    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(COMPILED_PDF, output)
    output.touch()
    if build_status == "existing":
        print("未检测到 LaTeX 编译器，已复用现有 PDF 编译产物。")
    print(f"已导出 PDF 文件: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
