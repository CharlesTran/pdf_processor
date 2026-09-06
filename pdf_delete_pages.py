#!/usr/bin/env python3
"""删除PDF的指定页码，通过复用 pdf_split 拆分与 merge_pdfs 合并实现。

用法:
    python pdf_delete_pages.py <PDF文件> -d 3,5 [-o 输出.pdf]
    python pdf_delete_pages.py <PDF文件> -r 2-4 [-o 输出.pdf]
"""

import argparse
import glob
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from pdf_split import split_pdf
from merge_pdfs import merge_pdfs


def parse_pages(text):
    """把 '1,3,5' 或 '2-4' 解析成页码集合，支持两种混用。"""
    pages = set()
    for part in text.split(","):
        part = part.strip()
        if "-" in part:
            start, end = map(int, part.split("-"))
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    return pages


def delete_pages(pdf_path, pages_to_delete, output_path):
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    tmpdir = tempfile.mkdtemp(prefix="pdf_delete_")
    try:
        split_pdf(pdf_path, output_dir=tmpdir, mode="single")

        kept = []
        for f in sorted(glob.glob(os.path.join(tmpdir, "*.pdf")),
                        key=lambda p: int(os.path.basename(p).split("第")[1].split("页")[0])):
            name = os.path.basename(f)
            page_num = int(name.split("第")[1].split("页")[0])
            if page_num not in pages_to_delete:
                kept.append(f)
            else:
                print(f"已删除第 {page_num} 页")

        if not kept:
            print("删除后没有剩余页面，未生成输出。")
            return

        merge_pdfs(kept, output_path)
        print(f"已保存: {output_path} (剩余 {len(kept)} 页)")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def main():
    parser = argparse.ArgumentParser(description="删除PDF指定页码")
    parser.add_argument("pdf", help="PDF文件路径")
    parser.add_argument("-d", "--delete", help="要删除的页码，如 3,5 或 2-4")
    parser.add_argument("-r", "--range", help="要删除的页码范围，如 2-4")
    parser.add_argument("-o", "--output", help="输出路径（默认 <原名>_删除后.pdf）")
    args = parser.parse_args()

    if not os.path.exists(args.pdf):
        print(f"文件不存在: {args.pdf}")
        sys.exit(1)

    pages = set()
    if args.delete:
        pages |= parse_pages(args.delete)
    if args.range:
        pages |= parse_pages(args.range)
    if not pages:
        parser.error("请用 -d 或 -r 指定要删除的页码")

    if args.output is None:
        base, ext = os.path.splitext(args.pdf)
        args.output = f"{base}_删除后{ext}"

    delete_pages(args.pdf, pages, args.output)


if __name__ == "__main__":
    main()
