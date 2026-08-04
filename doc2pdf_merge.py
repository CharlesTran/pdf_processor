#!/usr/bin/env python3
"""批量把 doc/docx 转成 pdf，并按顺序清单文件拼接成一个 PDF。

用法:
    python3 doc2pdf_merge.py order.txt [输出.pdf]

order.txt 每行一个文件名，支持 # 注释和空行。
相对路径相对于 order.txt 所在目录解析。
清单里也可以直接写已有的 .pdf 文件（跳过转换，直接拼接）。
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

WORD_APPLESCRIPT = """
on run argv
	set inFile to POSIX file (item 1 of argv)
	set outFile to (item 2 of argv) as text
	tell application "Microsoft Word"
		open inFile
		set theDoc to active document
		save as theDoc file name outFile file format format PDF
		close theDoc saving no
	end tell
end run
"""


def convert_to_pdf(doc_path, pdf_path):
    cmd = ["osascript", "-e", WORD_APPLESCRIPT, doc_path, pdf_path]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not os.path.exists(pdf_path):
        raise RuntimeError(f"转换失败: {doc_path}\n{proc.stderr}")


def merge_pdfs(pdf_files, out_pdf):
    cmd = ["gs", "-q", "-dNOPAUSE", "-dBATCH", "-dSAFER",
           "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.5",
           "-sOutputFile=" + out_pdf] + pdf_files
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"合并失败: {proc.stderr}")


def doc2pdf_merge(list_file, output):
    """按清单把 doc/docx 转成 pdf 并按顺序拼接为一个 pdf。"""
    base_dir = os.path.dirname(os.path.abspath(list_file))
    items = []
    with open(list_file, encoding="utf-8") as f:
        for lineno, raw in enumerate(f, 1):
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            p = line if os.path.isabs(line) else os.path.join(base_dir, line)
            if not os.path.exists(p):
                raise RuntimeError(f"第 {lineno} 行文件不存在: {p}")
            items.append(p)

    if not items:
        raise RuntimeError("清单为空，无事可做。")

    tmpdir = tempfile.mkdtemp(prefix="doc2pdf_")
    try:
        pdfs = []
        total = len(items)
        for i, src in enumerate(items, 1):
            if src.lower().endswith(".pdf"):
                pdfs.append(src)
                print(f"[{i}/{total}] 直接使用 {os.path.basename(src)}")
                continue
            out_pdf = os.path.join(tmpdir, f"{i:03d}.pdf")
            print(f"[{i}/{total}] 转换 {os.path.basename(src)} ... ", end="", flush=True)
            convert_to_pdf(src, out_pdf)
            pdfs.append(out_pdf)
            print("完成")

        print(f"拼接 {len(pdfs)} 个 PDF ... ", end="", flush=True)
        merge_pdfs(pdfs, output)
        print(f"完成: {output}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(
        description="doc/docx 批量转 pdf 并按清单顺序拼接",
        epilog="示例: doc2pdf_merge.py order.txt 结果.pdf",
    )
    parser.add_argument("list_file", help="顺序清单文件，每行一个文件名")
    parser.add_argument("output", nargs="?", default="merged.pdf",
                        help="输出 pdf 路径（默认 merged.pdf）")
    args = parser.parse_args()
    try:
        doc2pdf_merge(args.list_file, args.output)
    except RuntimeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
