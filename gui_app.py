#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""PDF 处理工具 —— 图形界面版（tkinter，跨平台，主要为 Windows 打包）。

整合五个功能：
    1. 拆分 PDF（每页 / 页码范围 / 指定页）
    2. 合并 PDF（可调整顺序）
    3. 删除指定页
    4. doc/docx 按清单批量转 PDF 并拼接（Windows 用 Word COM）
    5. 批量把 PDF 页面等比缩放、居中到标准 A4

打包方式见 .github/workflows/build-windows.yml（PyInstaller -> Windows exe）。
"""

import io
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pdf_split
import merge_pdfs
import pdf_delete_pages
import doc2pdf_merge
import pdf_to_a4

APP_NAME = "PDF 处理工具"
APP_VERSION = "1.0.0"

PDF_TYPES = [("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
LIST_TYPES = [("文本文件", "*.txt"), ("所有文件", "*.*")]


class _QueueWriter(io.TextIOBase):
    """把 print 输出同时转发到日志队列（供界面显示）和真正的 stdout。"""

    def __init__(self, q: queue.Queue, real=None):
        self.q = q
        self.real = real

    def write(self, s: str) -> int:
        self.q.put(("text", s))
        if self.real is not None:
            try:
                self.real.write(s)
                self.real.flush()
            except Exception:
                pass
        return len(s)

    def flush(self) -> None:
        if self.real is not None:
            try:
                self.real.flush()
            except Exception:
                pass


class PdfToolApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("860x680")
        self.minsize(760, 580)

        self.log_queue: queue.Queue = queue.Queue()
        self.worker: threading.Thread | None = None
        self._busy = False
        self._run_buttons: list[tk.Widget] = []

        self._build_ui()
        self.after(120, self._poll_log)
        self.log_line(f"{APP_NAME} v{APP_VERSION} 已启动。\n"
                      "请选择功能页签，填好参数后点击“开始”。\n"
                      "“doc转PDF”在 Windows 上需要本机安装 Microsoft Word。\n")

    # ------------------------------------------------------------- UI 搭建
    def _build_ui(self):
        pad = dict(padx=10, pady=8)
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, **pad)

        self._tab_split = self._build_split_tab(nb)
        self._tab_merge = self._build_merge_tab(nb)
        self._tab_delete = self._build_delete_tab(nb)
        self._tab_doc = self._build_doc_tab(nb)
        self._tab_a4 = self._build_a4_tab(nb)

        nb.add(self._tab_split, text="① 拆分")
        nb.add(self._tab_merge, text="② 合并")
        nb.add(self._tab_delete, text="③ 删除页")
        nb.add(self._tab_doc, text="④ doc转PDF")
        nb.add(self._tab_a4, text="⑤ 转A4")

        log_frame = ttk.LabelFrame(self, text="运行日志")
        log_frame.pack(fill="both", expand=False, padx=10, pady=(0, 8))
        self.log_text = tk.Text(log_frame, height=12, wrap="word",
                                font=("Microsoft YaHei UI", 9))
        sb = ttk.Scrollbar(log_frame, orient="vertical",
                           command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log_text.pack(side="left", fill="both", expand=True)

    def _file_row(self, parent, label, textvar, types,
                  action_label, browse_cmd, width=64):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label, width=12).pack(side="left")
        ent = ttk.Entry(row, textvariable=textvar, width=width)
        ent.pack(side="left", fill="x", expand=True, padx=(0, 6))
        ttk.Button(row, text=action_label, width=8,
                   command=browse_cmd).pack(side="left")
        return row

    def _row(self, parent, widget, pady=3):
        w = widget(parent)
        w.pack(fill="x", pady=pady)
        return w

    # ---------- 拆分
    def _build_split_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)

        self.split_pdf = tk.StringVar()
        self._file_row(tab, "PDF文件", self.split_pdf, PDF_TYPES, "选择…",
                       lambda: self._pick_open(
                           self.split_pdf, "选择要拆分的 PDF", PDF_TYPES))

        mode = ttk.Frame(tab)
        mode.pack(fill="x", pady=6)
        self.split_mode = tk.StringVar(value="single")
        ttk.Radiobutton(mode, text="按每页拆分", value="single",
                        variable=self.split_mode).pack(side="left", padx=(0, 14))
        ttk.Radiobutton(mode, text="页码范围(如 1-5)", value="range",
                        variable=self.split_mode).pack(side="left", padx=(0, 14))
        ttk.Radiobutton(mode, text="指定页(如 1,3,5)", value="custom",
                        variable=self.split_mode).pack(side="left")

        self.split_range = tk.StringVar()
        self.split_custom = tk.StringVar()
        row = ttk.Frame(tab)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text="范围/页", width=12).pack(side="left")
        self.split_range_ent = ttk.Entry(row, textvariable=self.split_range, width=22)
        self.split_range_ent.pack(side="left", padx=(0, 14))
        self.split_custom_ent = ttk.Entry(row, textvariable=self.split_custom, width=22)
        self.split_custom_ent.pack(side="left")

        self.split_out = tk.StringVar()
        self._file_row(tab, "输出目录", self.split_out, None, "选择…",
                       lambda: self._pick_dir(self.split_out,
                                              "选择输出目录（默认与 PDF 同目录）"))
        self._make_run_button(tab, "开始拆分", self._run_split)
        return tab

    # ---------- 合并
    def _build_merge_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)

        ttk.Label(tab, text="合并顺序（从上到下）：").pack(anchor="w")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True, pady=4)
        self.merge_listbox = tk.Listbox(body, font=("Microsoft YaHei UI", 10))
        sb = ttk.Scrollbar(body, orient="vertical",
                           command=self.merge_listbox.yview)
        self.merge_listbox.configure(yscrollcommand=sb.set)
        self.merge_listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        btns = ttk.Frame(body)
        btns.pack(side="left", padx=8, fill="y")
        ttk.Button(btns, text="添加文件…", width=10,
                   command=self._merge_add).pack(pady=2, fill="x")
        ttk.Button(btns, text="移除选中", width=10,
                   command=self._merge_remove).pack(pady=2, fill="x")
        ttk.Button(btns, text="上移 ▲", width=10,
                   command=lambda: self._merge_move(-1)).pack(pady=2, fill="x")
        ttk.Button(btns, text="下移 ▼", width=10,
                   command=lambda: self._merge_move(1)).pack(pady=2, fill="x")

        self.merge_out = tk.StringVar()
        self._file_row(tab, "输出文件", self.merge_out, PDF_TYPES, "另存为…",
                       lambda: self._pick_save(
                           self.merge_out, "保存合并后的 PDF",
                           "merged.pdf", PDF_TYPES))
        self._make_run_button(tab, "开始合并", self._run_merge)
        return tab

    # ---------- 删除页
    def _build_delete_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)

        self.del_pdf = tk.StringVar()
        self._file_row(tab, "PDF文件", self.del_pdf, PDF_TYPES, "选择…",
                       lambda: self._pick_open(
                           self.del_pdf, "选择要删除页的 PDF", PDF_TYPES))

        self.del_pages = tk.StringVar()
        self._file_row(tab, "删除页", self.del_pages, None, "",
                       lambda: None)
        ttk.Label(tab, text="支持页码与范围混写，例如：2,6,20-22",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        self.del_out = tk.StringVar()
        self._file_row(tab, "输出文件", self.del_out, PDF_TYPES, "另存为…",
                       lambda: self._pick_save(
                           self.del_out, "保存删除后的 PDF", "", PDF_TYPES))
        ttk.Label(tab, text="不填输出文件时，自动生成在 PDF 同目录：<原名>_删除后.pdf",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        self._make_run_button(tab, "开始删除", self._run_delete)
        return tab

    # ---------- doc 转 PDF
    def _build_doc_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)

        self.doc_list = tk.StringVar()
        self._file_row(tab, "清单文件", self.doc_list, LIST_TYPES, "选择…",
                       lambda: self._pick_open(
                           self.doc_list, "选择清单文件(order.txt)", LIST_TYPES))
        ttk.Label(tab, text="清单每行一个文件名；# 开头为注释；.pdf 直接拼接；"
                            "其余按 doc/docx 转成 PDF 后按行序拼接。",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        self.doc_out = tk.StringVar()
        self._file_row(tab, "输出文件", self.doc_out, PDF_TYPES, "另存为…",
                       lambda: self._pick_save(
                           self.doc_out, "保存拼接后的 PDF", "merged.pdf",
                           PDF_TYPES))
        ttk.Label(tab, text="不填时自动保存为清单所在目录下的 merged.pdf。"
                            "Windows 转换依赖本机安装的 Microsoft Word。",
                  foreground="#666").pack(anchor="w", pady=(0, 6))

        self._make_run_button(tab, "开始转换并拼接", self._run_doc)
        return tab

    # ---------- 转 A4
    def _build_a4_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)

        self.a4_in = tk.StringVar()
        self._file_row(tab, "输入目录", self.a4_in, None, "选择…",
                       lambda: self._pick_dir(self.a4_in, "选择存放 PDF 的目录"))

        self.a4_out = tk.StringVar()
        self._file_row(tab, "输出目录", self.a4_out, None, "选择…",
                       lambda: self._pick_dir(
                           self.a4_out, "选择输出目录（不填则为 输入目录/A4标准版）"))

        self.a4_overwrite = tk.BooleanVar(value=False)
        ttk.Checkbutton(tab, text="覆盖已存在的同名输出文件",
                        variable=self.a4_overwrite).pack(anchor="w", pady=4)
        ttk.Label(tab, text="页面将保持原比例并居中到标准 A4 画布(595.28×841.89pt)。",
                  foreground="#666").pack(anchor="w")

        self._make_run_button(tab, "开始转换", self._run_a4)
        return tab

    def _make_run_button(self, parent, text, cmd):
        frame = ttk.Frame(parent)
        frame.pack(fill="x", pady=(10, 2))
        btn = ttk.Button(frame, text=text, command=cmd)
        btn.pack(anchor="e")
        self._run_buttons.append(btn)

    # ----------------------------------------------------------- 运行控制
    def _pick_open(self, var, title, types):
        f = filedialog.askopenfilename(title=title, filetypes=types)
        if f:
            var.set(f)

    def _pick_save(self, var, title, initialfile, types):
        f = filedialog.asksaveasfilename(
            title=title, defaultextension=".pdf",
            filetypes=types, initialfile=initialfile)
        if f:
            var.set(f)

    def _pick_dir(self, var, title):
        d = filedialog.askdirectory(title=title)
        if d:
            var.set(d)

    def _merge_add(self):
        files = filedialog.askopenfilenames(
            title="选择要合并的 PDF", filetypes=PDF_TYPES)
        for f in files:
            if f not in self.merge_listbox.get(0, "end"):
                self.merge_listbox.insert("end", f)

    def _merge_remove(self):
        for idx in reversed(self.merge_listbox.curselection()):
            self.merge_listbox.delete(idx)

    def _merge_move(self, delta):
        sel = self.merge_listbox.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if not (0 <= j < self.merge_listbox.size()):
            return
        a, b = self.merge_listbox.get(i), self.merge_listbox.get(j)
        self.merge_listbox.delete(i)
        self.merge_listbox.insert(i, b)
        self.merge_listbox.delete(j)
        self.merge_listbox.insert(j, a)
        self.merge_listbox.selection_set(j)

    def _start(self, job):
        if self._busy:
            messagebox.showwarning(APP_NAME, "已有任务在处理中，请稍候。")
            return
        self._busy = True
        for b in self._run_buttons:
            b.configure(state="disabled")
        self.log_line("\n" + "=" * 50 + "\n开始任务…\n")

        def target():
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = _QueueWriter(self.log_queue, old_out)
            sys.stderr = _QueueWriter(self.log_queue, old_err)
            try:
                job()
                self.log_queue.put(("text", "\n[完成] 任务执行完毕。\n"))
            except Exception as exc:
                tb = traceback.format_exc()
                self.log_queue.put(("text",
                                    f"\n[错误] {exc}\n{tb[-1200:]}\n"))
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                self.log_queue.put(("done", None))

        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()

    def _poll_log(self):
        try:
            while True:
                kind, payload = self.log_queue.get_nowait()
                if kind == "text":
                    self.log_text.insert("end", payload)
                    self.log_text.see("end")
                elif kind == "done":
                    self._busy = False
                    for b in self._run_buttons:
                        b.configure(state="normal")
                    self.log_text.see("end")
        except queue.Empty:
            pass
        self.after(120, self._poll_log)

    def log_line(self, msg):
        self.log_text.insert("end", msg)
        self.log_text.see("end")

    def _require_file(self, path, label):
        path = path.strip()
        if not path:
            messagebox.showwarning(APP_NAME, f"请先选择{label}。")
            return None
        if not os.path.isfile(path):
            messagebox.showwarning(APP_NAME, f"{label}不存在:\n{path}")
            return None
        return path

    # ----------------------------------------------------------- 各功能任务
    def _run_split(self):
        pdf = self._require_file(self.split_pdf.get(), "PDF文件")
        if pdf is None:
            return
        mode = self.split_mode.get()
        pages = None
        if mode == "range":
            text = self.split_range.get().strip()
            try:
                start, end = map(int, text.split("-", 1))
                pages = (start, end)
            except ValueError:
                messagebox.showwarning(APP_NAME, "页码范围格式应为 1-5")
                return
        elif mode == "custom":
            text = self.split_custom.get().strip()
            try:
                pages = [int(p) for p in text.split(",") if p.strip()]
            except ValueError:
                messagebox.showwarning(APP_NAME, "指定页格式应为 1,3,5")
                return
            if not pages:
                messagebox.showwarning(APP_NAME, "请填写要提取的页码。")
                return

        out_dir = self.split_out.get().strip() or os.path.dirname(pdf) or os.getcwd()

        def job():
            print(f"输出目录: {out_dir}\n")
            pdf_split.split_pdf(pdf, out_dir, pages, mode)
        self._start(job)

    def _run_merge(self):
        files = list(self.merge_listbox.get(0, "end"))
        if not files:
            messagebox.showwarning(APP_NAME, "请先添加要合并的 PDF 文件。")
            return
        out = self.merge_out.get().strip()
        if not out:
            out = os.path.join(os.path.dirname(files[0]), "merged.pdf")
        if os.path.isdir(out):
            out = os.path.join(out, "merged.pdf")

        def job():
            merge_pdfs.merge_pdfs(files, out)
            print(f"已保存: {out}（共 {len(files)} 个 PDF）")
        self._start(job)

    def _run_delete(self):
        pdf = self._require_file(self.del_pdf.get(), "PDF文件")
        if pdf is None:
            return
        text = self.del_pages.get().strip()
        if not text:
            messagebox.showwarning(APP_NAME, "请填写要删除的页码，如 2,6,20-22。")
            return
        try:
            pages = pdf_delete_pages.parse_pages(text)
        except ValueError:
            messagebox.showwarning(APP_NAME, "删除页格式不对，示例：2,6,20-22")
            return
        out = self.del_out.get().strip()
        if not out:
            base, ext = os.path.splitext(pdf)
            out = f"{base}_删除后{ext}"

        def job():
            pdf_delete_pages.delete_pages(pdf, pages, out)
            print(f"已保存: {out}")
        self._start(job)

    def _run_doc(self):
        lst = self._require_file(self.doc_list.get(), "清单文件")
        if lst is None:
            return
        out = self.doc_out.get().strip()
        if not out:
            out = os.path.join(os.path.dirname(os.path.abspath(lst)),
                               "merged.pdf")

        def job():
            doc2pdf_merge.doc2pdf_merge(lst, out)
            print(f"已保存: {out}")
        self._start(job)

    def _run_a4(self):
        in_dir = self.a4_in.get().strip()
        if not in_dir:
            messagebox.showwarning(APP_NAME, "请选择存放 PDF 的输入目录。")
            return
        if not os.path.isdir(in_dir):
            messagebox.showwarning(APP_NAME, f"输入目录不存在:\n{in_dir}")
            return
        out_dir = self.a4_out.get().strip() or None
        overwrite = self.a4_overwrite.get()

        def job():
            pdf_to_a4.batch_convert_to_a4(in_dir, out_dir, overwrite)
        self._start(job)


def main():
    app = PdfToolApp()
    app.mainloop()


if __name__ == "__main__":
    main()
