#!/usr/bin/env python3
"""PDF处理工具整合入口: 拆分 / 合并 / 删除页码 / doc转pdf合并"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pdf_split
import merge_pdfs
import pdf_delete_pages
import doc2pdf_merge


def build_parser():
    parser = argparse.ArgumentParser(description="PDF处理工具: 拆分/合并/删除页码/doc转pdf")
    sub = parser.add_subparsers(dest="command")

    p_split = sub.add_parser("split", help="拆分PDF")
    p_split.add_argument("pdf")
    p_split.add_argument("-r", "--range", help="页码范围，如 1-5")
    p_split.add_argument("-p", "--pages", help="指定页，如 1,3,5")
    p_split.add_argument("-o", "--output", help="输出目录")

    p_merge = sub.add_parser("merge", help="合并PDF(不跟文件时打开交互界面调整顺序)")
    p_merge.add_argument("files", nargs="*")
    p_merge.add_argument("-o", "--output", default="merged.pdf")

    p_del = sub.add_parser("delete", help="删除指定页码")
    p_del.add_argument("pdf")
    p_del.add_argument("-d", "--delete", help="要删除的页码，如 3,5")
    p_del.add_argument("-r", "--range", help="要删除的页码范围，如 2-4")
    p_del.add_argument("-o", "--output")

    p_doc = sub.add_parser("doc2pdf", help="doc/docx转pdf并按清单合并")
    p_doc.add_argument("list_file")
    p_doc.add_argument("output", nargs="?", default="merged.pdf")

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    if not args.command:
        parser.print_help()
        sys.exit(1)

    if args.command == "split":
        if args.pages:
            pages = [int(p) for p in args.pages.split(",")]
            pdf_split.split_pdf(args.pdf, args.output, pages, "custom")
        elif args.range:
            start, end = map(int, args.range.split("-"))
            pdf_split.split_pdf(args.pdf, args.output, (start, end), "range")
        else:
            pdf_split.split_pdf(args.pdf, args.output, mode="single")

    elif args.command == "merge":
        if not args.files:
            merge_pdfs.merge_gui()
        else:
            merge_pdfs.merge_pdfs(args.files, args.output)
            print(f"已保存: {args.output} (共 {len(args.files)} 个PDF)")

    elif args.command == "delete":
        pages = set()
        if args.delete:
            pages |= pdf_delete_pages.parse_pages(args.delete)
        if args.range:
            pages |= pdf_delete_pages.parse_pages(args.range)
        if not pages:
            parser.error("请用 -d 或 -r 指定要删除的页码")
        if args.output is None:
            base, ext = os.path.splitext(args.pdf)
            args.output = f"{base}_删除后{ext}"
        pdf_delete_pages.delete_pages(args.pdf, pages, args.output)

    elif args.command == "doc2pdf":
        try:
            doc2pdf_merge.doc2pdf_merge(args.list_file, args.output)
        except RuntimeError as e:
            print(f"[错误] {e}", file=sys.stderr)
            sys.exit(1)


if __name__ == "__main__":
    main()
