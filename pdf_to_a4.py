#!/usr/bin/env python3
"""批量将目录中的 PDF 页面等比缩放、居中到标准 A4 画布。"""

from __future__ import annotations

import argparse
import os
import tempfile
from pathlib import Path

from pypdf import PageObject, PdfReader, PdfWriter, Transformation


# PDF 使用 point 作为单位：1 inch = 72 points，A4 = 210 mm x 297 mm。
A4_WIDTH = 210 / 25.4 * 72
A4_HEIGHT = 297 / 25.4 * 72
# 命令行不传目录时默认处理当前工作目录（GUI 中始终由用户选择目录）。
DEFAULT_INPUT_DIR = None
DEFAULT_OUTPUT_DIR_NAME = "A4标准版"


def _copy_metadata(reader: PdfReader, writer: PdfWriter) -> None:
    """复制可安全写入的文档元数据。"""
    if not reader.metadata:
        return
    metadata = {
        str(key): str(value)
        for key, value in reader.metadata.items()
        if key and value is not None
    }
    if metadata:
        writer.add_metadata(metadata)


def convert_pdf_to_a4(input_path: Path, output_path: Path) -> int:
    """转换单个 PDF，返回处理的页数。"""
    reader = PdfReader(input_path)
    if reader.is_encrypted and not reader.decrypt(""):
        raise ValueError("PDF 已加密，无法用空密码打开")

    writer = PdfWriter()
    # pypdf 默认会使用 PDF 1.3 文件头；明确使用较新的版本以兼容上传系统。
    writer.pdf_header = b"%PDF-1.7"
    _copy_metadata(reader, writer)

    for source_page in reader.pages:
        # 将 /Rotate 合并进页面内容，后续可统一按实际可见方向计算。
        if source_page.rotation:
            source_page.transfer_rotation_to_content()

        source_box = source_page.cropbox
        source_width = float(source_box.width)
        source_height = float(source_box.height)
        if source_width <= 0 or source_height <= 0:
            raise ValueError("发现宽度或高度为 0 的页面")

        scale = min(A4_WIDTH / source_width, A4_HEIGHT / source_height)
        offset_x = (A4_WIDTH - source_width * scale) / 2
        offset_y = (A4_HEIGHT - source_height * scale) / 2

        # cropbox 左下角不一定是 (0, 0)，先抵消原点偏移，再缩放和居中。
        transform = (
            Transformation()
            .translate(-float(source_box.left), -float(source_box.bottom))
            .scale(scale)
            .translate(offset_x, offset_y)
        )
        a4_page = PageObject.create_blank_page(width=A4_WIDTH, height=A4_HEIGHT)
        a4_page.merge_transformed_page(source_page, transform, expand=False)
        writer.add_page(a4_page)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temp_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="wb",
            prefix=f".{output_path.stem}_",
            suffix=".tmp",
            dir=output_path.parent,
            delete=False,
        ) as temp_file:
            temp_path = Path(temp_file.name)
            writer.write(temp_file)
        os.replace(temp_path, output_path)
    finally:
        if temp_path is not None and temp_path.exists():
            temp_path.unlink()

    return len(reader.pages)


def batch_convert_to_a4(
    input_dir: Path | str | None = None,
    output_dir: Path | str | None = None,
    overwrite: bool = False,
) -> tuple[int, int, int]:
    """批量转换目录中的 PDF，返回 (成功数, 跳过数, 失败数)。"""
    if input_dir is None:
        input_dir = Path.cwd()
    input_dir = Path(input_dir).expanduser().resolve()
    if not input_dir.is_dir():
        raise NotADirectoryError(f"输入目录不存在: {input_dir}")

    if output_dir is None:
        output_dir = input_dir / DEFAULT_OUTPUT_DIR_NAME
    else:
        output_dir = Path(output_dir).expanduser().resolve()

    pdf_files = sorted(
        (path for path in input_dir.iterdir() if path.is_file() and path.suffix.lower() == ".pdf"),
        key=lambda path: path.name,
    )
    if not pdf_files:
        print(f"未找到 PDF: {input_dir}")
        return 0, 0, 0

    print(f"输入目录: {input_dir}")
    print(f"输出目录: {output_dir}")
    print(f"找到 {len(pdf_files)} 个 PDF")

    succeeded = skipped = failed = 0
    for input_path in pdf_files:
        output_path = output_dir / input_path.name
        if output_path.exists() and not overwrite:
            skipped += 1
            print(f"[跳过] {output_path.name}（已存在；可加 --overwrite 覆盖）")
            continue

        try:
            page_count = convert_pdf_to_a4(input_path, output_path)
            succeeded += 1
            print(f"[完成] {input_path.name} -> {output_path.name}（{page_count} 页）")
        except Exception as exc:
            failed += 1
            print(f"[失败] {input_path.name}: {exc}")

    print(f"处理结束：成功 {succeeded}，跳过 {skipped}，失败 {failed}")
    return succeeded, skipped, failed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="批量将 PDF 页面等比缩放、居中到标准 A4（595.28 x 841.89 pt）"
    )
    parser.add_argument(
        "input_dir",
        nargs="?",
        default=None,
        help="PDF 所在目录（默认: 当前目录）",
    )
    parser.add_argument(
        "-o",
        "--output",
        help=f"输出目录（默认: 输入目录/{DEFAULT_OUTPUT_DIR_NAME}）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="覆盖输出目录中已有的同名 PDF",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        input_dir = Path(args.input_dir).expanduser() if args.input_dir else Path.cwd()
        _, _, failed = batch_convert_to_a4(input_dir, args.output, args.overwrite)
    except (OSError, ValueError) as exc:
        print(f"[错误] {exc}")
        return 1
    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
