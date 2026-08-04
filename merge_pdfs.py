#!/usr/bin/env python3
"""macOS PDF合并工具 - 手动选择文件后合并"""

import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfWriter
import os


def merge_pdfs(files, output_path):
    """按顺序把多个PDF合并为一个。"""
    merger = PdfWriter()
    for f in files:
        merger.append(f)

    with open(output_path, "wb") as f:
        merger.write(f)

    merger.close()
    return output_path


def main():
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="选择要合并的PDF文件",
        filetypes=[("PDF files", "*.pdf")],
    )

    if not files:
        return

    output_path = filedialog.asksaveasfilename(
        title="保存合并后的PDF",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile="merged.pdf",
    )

    if not output_path:
        return

    merge_pdfs(files, output_path)

    messagebox.showinfo("完成", f"已合并 {len(files)} 个PDF文件到:\n{output_path}")

if __name__ == "__main__":
    main()
