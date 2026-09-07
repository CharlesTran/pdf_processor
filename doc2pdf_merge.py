#!/usr/bin/env python3
"""批量把 doc/docx 转成 pdf，并按顺序清单文件拼接成一个 PDF。

用法:
    python3 doc2pdf_merge.py order.txt [输出.pdf]

order.txt 每行一个文件名，支持 # 注释和空行。
相对路径相对于 order.txt 所在目录解析。
清单里也可以直接写已有的 .pdf 文件（跳过转换，直接拼接）。

转换引擎（convert_to_pdf 的 engine 参数）：
- auto (默认): 按本机可用能力自动选择
  Windows: Word(2010+) 优先 → LibreOffice → 明确报错指引
  macOS  : LibreOffice 优先 → Word(AppleScript)
- word:       强制 Microsoft Word（Windows 走 COM；macOS 走 AppleScript）
- libreoffice:强制 LibreOffice（--headless，不依赖 Office/WPS）

LibreOffice 查找顺序：环境变量 PDFPROC_SOFFICE →
系统 PATH → 常见安装目录 → 软件旁的 libreoffice/ 便携目录。
最终拼接统一使用 pypdf。
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


# ------------------------------------------------------------- 引擎探测
def _module_dir():
    return os.path.dirname(os.path.abspath(__file__))


def _app_dir():
    """打包后为 exe 所在目录，开发时为本脚本所在目录。"""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        # onefile 临时解压目录；程序真正所在目录用 sys.executable 的目录
        return os.path.dirname(os.path.abspath(sys.executable))
    return _module_dir()


def _find_soffice():
    """按多种途径查找 LibreOffice 可执行文件，找不到返回 None。"""
    env = os.environ.get("PDFPROC_SOFFICE")
    if env and os.path.isfile(env):
        return env

    for name in ("soffice", "soffice.bin", "soffice.com"):
        p = shutil.which(name)
        if p:
            return p

    # 常见 Windows 安装位置与“软件旁的便携目录”
    bases = [_app_dir()]
    bases += [os.path.join(_app_dir(), "libreoffice"),
              os.path.join(_app_dir(), "LibreOffice")]
    bases += [r"C:\Program Files\LibreOffice\program",
              r"C:\Program Files (x86)\LibreOffice\program",
              "/Applications/LibreOffice.app/Contents/MacOS",
              "/usr/bin", "/usr/local/bin"]
    for base in bases:
        for name in ("soffice", "soffice.exe", "soffice.com", "soffice.bin"):
            cand = os.path.join(base, name)
            if os.path.isfile(cand):
                return cand
    return None


def _windows_word_major():
    """探测已安装 Microsoft Word 的主版本号（2010=14 …）；未安装返回 None。"""
    if not sys.platform.startswith("win"):
        return None
    try:
        import winreg
    except ImportError:
        return None
    majors = (16, 15, 14, 12)   # 2016+/2013/2010/2007
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for major in majors:
            key = rf"Software\Microsoft\Office\{major}.0\Word\InstallRoot"
            try:
                with winreg.OpenKey(hive, key):
                    return major
            except OSError:
                continue
    return None


# ------------------------------------------------------------- 各引擎转换
def _word_com_convert(doc_path, pdf_path, hint_2007=True):
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
        doc = word.Documents.Open(
            str(doc_path),
            ConfirmConversions=False,     # 不弹“格式转换”确认框
            ReadOnly=True,
            AddToRecentFiles=False,
            NoEncodingDialog=True,        # 不弹编码选择框
            Visible=False,
        )
        try:
            # wdFormatPDF = 17（Word 2007 需安装官方“另存为PDF/XPS”插件）
            doc.SaveAs(str(pdf_path), FileFormat=17)
        finally:
            try:
                doc.Close(False)
            except Exception:
                pass
    except Exception as exc:
        msg = f"Word 转换失败: {doc_path}\n{exc}"
        if hint_2007 and "PDF" in str(exc) or hint_2007 and "17" in str(exc):
            msg += ("\n[提示] 若为 Office 2007：其不自带“另存为PDF”，"
                    "请安装官方插件 “Microsoft Save as PDF or XPS Add-in”，"
                    "或在软件里把转换方式改为 LibreOffice。")
        raise RuntimeError(msg) from exc
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()

    if not _valid_pdf_output(pdf_path):
        raise RuntimeError(
            f"转换失败（未生成有效 PDF 文件）: {doc_path}\n"
            "[提示] Office 2007 需安装官方“另存为PDF/XPS”插件，"
            "或改用 LibreOffice 转换方式。")


def _word_applescript_convert(doc_path, pdf_path):
    """macOS: 通过 AppleScript 调用 Word 另存为 PDF。"""
    cmd = ["osascript", "-e", WORD_APPLESCRIPT, doc_path, pdf_path]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    if proc.returncode != 0 or not _valid_pdf_output(pdf_path):
        raise RuntimeError(f"Word 转换失败: {doc_path}\n{proc.stderr}")


def _soffice_convert(soffice, doc_path, pdf_path):
    """LibreOffice headless 转换（不依赖 Office/WPS，支持 .doc/.docx）。"""
    out_dir = os.path.dirname(os.path.abspath(pdf_path)) or "."
    os.makedirs(out_dir, exist_ok=True)
    # 独立 profile，避免与本机正开着的 LibreOffice 抢锁
    profile = tempfile.mkdtemp(prefix="lo_profile_")
    try:
        cmd = [soffice, "--headless", "--norestore", "--nolockcheck",
               "-env:UserInstallation=file:///" +
               profile.replace("\\", "/").lstrip("/"),
               "--convert-to", "pdf:writer_pdf_Export",
               "--outdir", out_dir,
               os.path.abspath(doc_path)]
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
        produced = os.path.join(
            out_dir, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
        if (os.path.exists(produced)
                and os.path.abspath(produced) != os.path.abspath(pdf_path)):
            os.replace(produced, pdf_path)
        if not _valid_pdf_output(pdf_path):
            detail = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(
                f"LibreOffice 转换失败: {os.path.basename(doc_path)}\n"
                f"{detail[-500:]}")
    finally:
        shutil.rmtree(profile, ignore_errors=True)


def _valid_pdf_output(path):
    return (os.path.exists(path) and os.path.getsize(path) > 0)


# ------------------------------------------------------------- 引擎选择
_LIBREOFFICE_MISSING = (
    "未找到 LibreOffice，无法转换。\n"
    "LibreOffice 免费开源、不依赖 Office/WPS，下载："
    "https://www.libreoffice.org/download/download-libreoffice/\n"
    "或下载便携版，把 program\\soffice.exe 放入本程序目录的 "
    "libreoffice 文件夹（或设置环境变量 PDFPROC_SOFFICE）。"
)

_NO_ENGINE_HELP = (
    "未找到可用的转换引擎（不支持 doc 自动转 PDF）。\n"
    "任选其一：\n"
    "1) 安装免费开源的 LibreOffice（https://www.libreoffice.org/ "
    "下载安装，不依赖 Office/WPS，支持 doc/docx）；\n"
    "2) 安装 Microsoft Word（2010 及以上版本，排版最接近原样）；\n"
    "3) 使用 LibreOffice 便携版：把其中的 program/soffice.exe "
    "放进本程序目录下的 libreoffice 文件夹（或设置环境变量 PDFPROC_SOFFICE）。"
)

_OFFICE2007_HELP = (
    "检测到本机为 Office 2007：其不自带“另存为PDF”功能，直接调用可能卡死。\n"
    "请安装官方插件 “Microsoft Save as PDF or XPS Add-in”，"
    "或在软件界面把转换方式改为 LibreOffice（推荐，无需 Office）。"
)


def convert_to_pdf(doc_path, pdf_path, engine="auto"):
    """把 doc/docx 转成 PDF。engine: auto / word / libreoffice。"""
    if engine not in ("auto", "word", "libreoffice"):
        raise RuntimeError(f"未知转换引擎: {engine}")
    soffice = _find_soffice()
    is_win = sys.platform.startswith("win")

    if engine == "libreoffice":
        if soffice:
            _soffice_convert(soffice, doc_path, pdf_path)
        else:
            raise RuntimeError(_LIBREOFFICE_MISSING)
        return

    if engine == "word":
        if is_win:
            major = _windows_word_major()
            if major == 12:
                print("[提示] 检测到 Office 2007；如转换卡住请先安装官方"
                      "“另存为PDF/XPS”插件，或改用 LibreOffice。",
                      flush=True)
            _word_com_convert(doc_path, pdf_path)
        elif sys.platform == "darwin":
            _word_applescript_convert(doc_path, pdf_path)
        else:
            raise RuntimeError(_NO_ENGINE_HELP)
        return

    # ---- auto ----
    if is_win:
        major = _windows_word_major()
        if major is not None and major >= 14:      # Word 2010+ 效果最好
            _word_com_convert(doc_path, pdf_path)
            return
        if soffice:                                 # LibreOffice 兜底
            _soffice_convert(soffice, doc_path, pdf_path)
            return
        if major == 12:                             # 2007 无插件易卡死，先给指引
            raise RuntimeError(_OFFICE2007_HELP)
        raise RuntimeError(_NO_ENGINE_HELP)
    if sys.platform == "darwin":
        if soffice:
            _soffice_convert(soffice, doc_path, pdf_path)
        else:
            _word_applescript_convert(doc_path, pdf_path)
        return
    if soffice:
        _soffice_convert(soffice, doc_path, pdf_path)
        return
    raise RuntimeError(_NO_ENGINE_HELP)


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


def _process_items(items, output, merge, engine="auto"):
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
                convert_to_pdf(src, target, engine=engine)
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
            convert_to_pdf(src, out_pdf, engine=engine)
            pdfs.append(out_pdf)
            print("完成")

        print(f"拼接 {len(pdfs)} 个 PDF ... ", end="", flush=True)
        merge_pdfs(pdfs, output)
        print(f"完成: {output}")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


def doc2pdf_files(files, output, merge=True, engine="auto"):
    """GUI：直接传入文件列表（绝对路径）转 PDF。

    merge=True : 全部拼接为一个 PDF（output 为文件路径）
    merge=False: 每份单独输出一个 PDF（output 为输出目录）
    engine: auto / word / libreoffice
    """
    _process_items([os.path.abspath(f) for f in files], output, merge, engine)


def doc2pdf_merge(list_file, output, merge=True, engine="auto"):
    """按清单文件把 doc/docx 转成 pdf（命令行入口用）。

    merge=True : 全部拼接为一个 PDF（output 为文件路径）
    merge=False: 每份单独输出一个 PDF（output 为输出目录）
    engine: auto / word / libreoffice
    """
    items = _load_items(list_file)
    _process_items(items, output, merge, engine)


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
    parser.add_argument("--engine", choices=["auto", "word", "libreoffice"],
                        default="auto",
                        help="转换引擎：auto 自动选择 / word / libreoffice")
    args = parser.parse_args()
    try:
        doc2pdf_merge(args.list_file, args.output, merge=not args.no_merge,
                      engine=args.engine)
    except RuntimeError as e:
        print(f"[错误] {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
