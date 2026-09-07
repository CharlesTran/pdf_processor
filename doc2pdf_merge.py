#!/usr/bin/env python3
"""批量把 doc/docx 转成 pdf，并按顺序清单文件拼接成一个 PDF。

用法:
    python3 doc2pdf_merge.py order.txt [输出.pdf]

order.txt 每行一个文件名，支持 # 注释和空行。
相对路径相对于 order.txt 所在目录解析。
清单里也可以直接写已有的 .pdf 文件（跳过转换，直接拼接）。

转换后端按平台选择：
- Windows: Microsoft Word COM（需安装 MS Word），通过 pywin32 调用
- macOS:   AppleScript 调用 Microsoft Word（保持原有实现）
最终拼接统一使用 pypdf（不再依赖 Ghostscript）。
"""

import argparse
import os
import shutil
import subprocess
import sys
import tempfile

from pypdf import PdfWriter

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


def _word_com_convert(doc_path, pdf_path):
    """Windows: 通过 Word COM 接口把 doc/docx 另存为 PDF。"""
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise RuntimeError(
            "缺少 pywin32（Word COM 需要）。请执行: pip install pywin32"
        ) from exc

    pythoncom.CoInitialize()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0  # wdAlertsNone
        doc = word.Documents.Open(str(doc_path), ReadOnly=True)
        try:
            # wdFormatPDF = 17
            doc.SaveAs(str(pdf_path), FileFormat=17)
        finally:
            doc.Close(False)
    except Exception as exc:
        raise RuntimeError(f"Word 转换失败: {doc_path}\n{exc}") from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not os.path.exists(pdf_path):
        raise RuntimeError(f"转换失败（未生成输出文件）: {doc_path}")


def _word_applescript_convert(doc_path, pdf_path):
    """macOS: 通过 AppleScript 调用 Word 另存为 PDF。"""
    cmd = ["osascript", "-e", WORD_APPLESCRIPT, doc_path, pdf_path]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not os.path.exists(pdf_path):
        raise RuntimeError(f"转换失败: {doc_path}\n{proc.stderr}")


def convert_to_pdf(doc_path, pdf_path):
    if sys.platform.startswith("win"):
        _word_com_convert(doc_path, pdf_path)
    elif sys.platform == "darwin":
        _word_applescript_convert(doc_path, pdf_path)
    else:
        raise RuntimeError(
            "当前系统不支持 doc 自动转 PDF，请手动转好后在清单里写 .pdf 文件。"
        )


def merge_pdfs(pdf_files, out_pdf):
    """用 pypdf 按顺序拼接 PDF（跨平台，替代原来的 Ghostscript）。"""
    writer = PdfWriter()
    for f in pdf_files:
        writer.append(f)
    try:
        with open(out_pdf, "wb") as fh:
            writer.write(fh)
    finally:
        writer.close()
    if not os.path.exists(out_pdf):
        raise RuntimeError("合并失败：未生成输出文件")


def _load_items(list_file):
    """读取清单文件，返回按顺序展开的文件路径列表（含错误检查）。"""
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
    return items


def _unique_path(path):
    """目标路径已存在时自动追加 _2/_3…，避免覆盖。"""
    if not os.path.exists(path):
        return path
    stem, ext = os.path.splitext(path)
    n = 2
    while True:
        cand = f"{stem}_{n}{ext}"
        if not os.path.exists(cand):
            return cand
        n += 1


def _process_items(items, output, merge):
    """核心处理：给定已解析好的文件列表，合并或逐份输出。"""
    total = len(items)
    if not merge:
        os.makedirs(output, exist_ok=True)
        results = []
        for i, src in enumerate(items, 1):
            target = os.path.join(
                output, os.path.splitext(os.path.basename(src))[0] + ".pdf")
            target = _unique_path(target)
            if src.lower().endswith(".pdf"):
                print(f"[{i}/{total}] 复制 {os.path.basename(src)} ... ",
                      end="", flush=True)
                shutil.copy2(src, target)
            else:
                print(f"[{i}/{total}] 转换 {os.path.basename(src)} ... ",
                      end="", flush=True)
                convert_to_pdf(src, target)
            print("完成")
            results.append(target)
        print(f"处理结束：共 {len(results)} 个 PDF，输出目录: {output}")
        return results

    tmpdir = tempfile.mkdtemp(prefix="doc2pdf_")
    try:
        pdfs = []
        for i, src in enumerate(items, 1):
            if src.lower().endswith(".pdf"):
                pdfs.append(src)
                print(f"[{i}/{total}] 直接使用 {os.path.basename(src)}")
                continue
            out_pdf = os.path.join(tmpdir, f"{i:03d}.pdf")
            print(f"[{i}/{total}] 转换 {os.path.basename(src)} ... ",
                  end="", flush=True)
            convert_to_pdf(src, out_pdf)
            pdfs.append(out_pdf)
            print("完成")

        print(f"拼接 {len(pdfs)} 个 PDF ... ", end="", flush=True)
        merge_pdfs(pdfs, output)
        print(f"完成: {output}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def doc2pdf_files(files, output, merge=True):
    """GUI：直接传入文件列表（绝对路径）转 PDF。

    merge=True : 全部拼接为一个 PDF（output 为文件路径）
    merge=False: 每份单独输出一个 PDF（output 为输出目录）
    """
    _process_items([os.path.abspath(f) for f in files], output, merge)


def doc2pdf_merge(list_file, output, merge=True):
    """按清单文件把 doc/docx 转成 pdf（命令行入口用）。

    merge=True : 全部拼接为一个 PDF（output 为文件路径）
    merge=False: 每份单独输出一个 PDF（output 为输出目录）
    """
    items = _load_items(list_file)
    _process_items(items, output, merge)


def main():
    parser = argparse.ArgumentParser(
        description="doc/docx 批量转 pdf：默认按清单顺序拼接为一个 PDF；"
                    "加 --no-merge 则每份单独输出",
        epilog="示例:\n"
               "  doc2pdf_merge.py order.txt 结果.pdf          # 合并为一个\n"
               "  doc2pdf_merge.py order.txt 输出目录/ --no-merge # 逐份输出到目录",
    )
    parser.add_argument("list_file", help="顺序清单文件，每行一个文件名")
    parser.add_argument("output", nargs="?", default="merged.pdf",
                        help="合并时: 输出 pdf 路径；--no-merge 时: 输出目录")
    parser.add_argument("--no-merge", action="store_true",
                        help="不合并：每份文档单独转成 PDF 放入输出目录")
    args = parser.parse_args()
    try:
        doc2pdf_merge(args.list_file, args.output, merge=not args.no_merge)
    except RuntimeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
