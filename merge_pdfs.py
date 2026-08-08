#!/usr/bin/env python3
"""PDF合并工具 - 支持选择文件后用上下按钮自定义合并顺序"""

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


def merge_gui():
    """交互式合并: 选择文件 -> 上下按钮调整顺序 -> 合并保存"""
    root = tk.Tk()
    root.withdraw()

    files = list(filedialog.askopenfilenames(
        title="选择要合并的PDF文件",
        filetypes=[("PDF files", "*.pdf")],
    ))
    if not files:
        return

    root.deiconify()
    root.title("调整合并顺序")
    root.geometry("480x420")

    frame = tk.Frame(root, padx=10, pady=10)
    frame.pack(fill="both", expand=True)

    displayed = list(files)

    listbox = tk.Listbox(frame, font=("Helvetica", 13))
    listbox.pack(side="left", fill="both", expand=True)

    scrollbar = tk.Scrollbar(frame, orient="vertical", command=listbox.yview)
    scrollbar.pack(side="left", fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    btn_frame = tk.Frame(frame)
    btn_frame.pack(side="left", padx=10)

    def refresh():
        listbox.delete(0, "end")
        for f in displayed:
            listbox.insert("end", os.path.basename(f))
        if displayed:
            listbox.selection_set(0)

    def move(delta):
        sel = listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if not (0 <= j < len(displayed)):
            return
        displayed[i], displayed[j] = displayed[j], displayed[i]
        refresh()
        listbox.selection_set(j)

    tk.Button(btn_frame, text="上移 ▲", width=8,
              command=lambda: move(-1)).pack(pady=5)
    tk.Button(btn_frame, text="下移 ▼", width=8,
              command=lambda: move(1)).pack(pady=5)

    def do_merge():
        output_path = filedialog.asksaveasfilename(
            title="保存合并后的PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="merged.pdf",
        )
        if not output_path:
            return
        merge_pdfs(displayed, output_path)
        messagebox.showinfo("完成", f"已按自定义顺序合并 {len(displayed)} 个PDF文件到:\n{output_path}")

    tk.Button(root, text="开始合并", width=15,
              command=do_merge).pack(pady=10)

    refresh()
    root.mainloop()


def main():
    merge_gui()

if __name__ == "__main__":
    main()
