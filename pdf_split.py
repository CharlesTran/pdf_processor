#!/usr/bin/env python3
"""简易PDF拆分工具"""

import sys
import os
from pypdf import PdfReader, PdfWriter


def get_pdf_info(pdf_path):
    reader = PdfReader(pdf_path)
    return len(reader.pages)


def split_pdf(pdf_path, output_dir=None, pages=None, mode="single"):
    if not os.path.exists(pdf_path):
        print(f"文件不存在: {pdf_path}")
        return

    if output_dir is None:
        output_dir = os.path.dirname(pdf_path) or "."
    os.makedirs(output_dir, exist_ok=True)

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    print(f"PDF: {pdf_path}")
    print(f"总页数: {total_pages}")
    print("-" * 40)

    if mode == "single":
        for i in range(total_pages):
            writer = PdfWriter()
            writer.add_page(reader.pages[i])

            output_path = os.path.join(output_dir, f"{base_name}_第{i+1}页.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)

            print(f"已保存: {output_path}")

    elif mode == "range" and pages:
        start, end = pages
        start = max(1, min(start, total_pages))
        end = max(start, min(end, total_pages))

        writer = PdfWriter()
        for i in range(start - 1, end):
            writer.add_page(reader.pages[i])

        output_path = os.path.join(output_dir, f"{base_name}_第{start}-{end}页.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

        print(f"已保存: {output_path}")

    elif mode == "custom" and pages:
        writer = PdfWriter()
        for page_num in pages:
            if 1 <= page_num <= total_pages:
                writer.add_page(reader.pages[page_num - 1])

        output_path = os.path.join(output_dir, f"{base_name}_自定义.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)

        print(f"已保存: {output_path}")

    print("-" * 40)
    print("拆分完成!")


def parse_page_spec(text):
    """把 '2,4-6,8' 这类文本解析为有序页号列表。

    支持逗号分隔的单个页码与 a-b 范围混写；空文本返回 []（表示全部页）。
    格式非法时抛 ValueError。
    """
    text = (text or "").strip()
    if not text:
        return []
    pages = []
    for part in text.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            try:
                start, end = int(a), int(b)
            except ValueError:
                raise ValueError(f"页码范围格式不对：{part}（应为 a-b，如 1-5）")
            if start < 1 or end < start:
                raise ValueError(f"页码范围格式不对：{part}（应为正数且 a<=b）")
            pages.extend(range(start, end + 1))
        else:
            try:
                n = int(part)
            except ValueError:
                raise ValueError(f"页码格式不对：{part}（应为数字）")
            if n < 1:
                raise ValueError(f"页码格式不对：{part}（页码从 1 开始）")
            pages.append(n)
    return pages


def unique_pages(pages):
    """去重并保持原有顺序。"""
    seen = set()
    out = []
    for p in pages:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def extract_pages(pdf_path, page_numbers, output_path):
    """把指定页(1 起)写成一个 PDF 文件；page_numbers 为空列表表示整本。

    超出实际总页数的页码自动忽略。返回实际写入的页号列表。
    """
    reader = PdfReader(pdf_path)
    total = len(reader.pages)
    if page_numbers:
        nums = [n for n in unique_pages(page_numbers) if 1 <= n <= total]
    else:
        nums = list(range(1, total + 1))
    if not nums:
        raise ValueError("没有有效的页码可提取（超出总页数？）")

    writer = PdfWriter()
    for n in nums:
        writer.add_page(reader.pages[n - 1])
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with open(output_path, "wb") as f:
        writer.write(f)
    print(f"已保存: {output_path} ({len(nums)} 页)")
    return nums


def main():
    if len(sys.argv) > 1:
        if sys.argv[1] == "-h" or sys.argv[1] == "--help":
            print("""
用法:
    python pdf_split.py <PDF文件>                    # 每页拆分
    python pdf_split.py <PDF文件> -r 1-5             # 拆分1-5页
    python pdf_split.py <PDF文件> -p 1,3,5           # 拆分指定页
""")
        else:
            pdf_path = sys.argv[1]
            output_dir = None
            pages = None
            mode = "single"

            i = 2
            while i < len(sys.argv):
                if sys.argv[i] in ["-r", "--range"]:
                    start, end = map(int, sys.argv[i+1].split("-"))
                    pages = (start, end)
                    mode = "range"
                    i += 2
                elif sys.argv[i] in ["-p", "--pages"]:
                    pages = [int(p) for p in sys.argv[i+1].split(",")]
                    mode = "custom"
                    i += 2
                elif sys.argv[i] in ["-o", "--output"]:
                    output_dir = sys.argv[i+1]
                    i += 2
                else:
                    i += 1

            split_pdf(pdf_path, output_dir, pages, mode)
    else:
        print("用法: python pdf_split.py <PDF文件> [-r 1-5] [-p 1,3,5]")


if __name__ == "__main__":
    main()
