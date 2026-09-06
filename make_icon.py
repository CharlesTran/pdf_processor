#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用纯 Python 标准库生成应用图标 assets/app.ico（免 Pillow）。

Windows 图标文件（.ico）内嵌 PNG 图像（Vista+ 支持）。
图标主体：蓝色圆角方块 + 白色文档 + 深蓝块状 “PDF” 字样 + 红色底条。
用法: python make_icon.py [输出路径]
"""

import os
import struct
import sys
import zlib

SIZE = 256
R = 56             # 背景圆角半径
BG = (45, 125, 233, 255)        # 主蓝
PAGE = (255, 255, 255, 255)
INK = (38, 62, 96, 255)         # 深蓝字
BAR = (203, 211, 224, 255)      # 浅灰条
RED = (222, 78, 78, 255)        # 红条

# 5x7 点阵字（'#' 为像素）
FONT = {
    "P": ["11110", "10001", "10001", "11110", "10000", "10000", "10000"],
    "D": ["11110", "10001", "10001", "10001", "10001", "10001", "11110"],
    "F": ["11111", "10000", "10000", "11110", "10000", "10000", "10000"],
}


def rounded_rect_alpha(x, y, x0, y0, x1, y1, r):
    """点在圆角矩形内的覆盖率(0~1)，做简单 2x2 超采样抗锯齿。"""
    cover = 0.0
    for sx in (0.25, 0.75):
        for sy in (0.25, 0.75):
            px, py = x + sx, y + sy
            if not (x0 <= px <= x1 and y0 <= py <= y1):
                continue
            cx = min(max(px, x0 + r), x1 - r)
            cy = min(max(py, y0 + r), y1 - r)
            dx, dy = px - cx, py - cy
            if dx * dx + dy * dy <= r * r:
                cover += 1.0
    return cover / 4.0


def blend(dst, src, a):
    """把 (r,g,b) src 以透明度 a 叠加到 dst (r,g,b,a) 上。"""
    dr, dg, db, da = dst
    out_a = a + da * (1 - a)
    if out_a <= 0:
        return (0, 0, 0, 0)
    return (
        (src[0] * a + dr * da * (1 - a)) / out_a,
        (src[1] * a + dg * da * (1 - a)) / out_a,
        (src[2] * a + db * da * (1 - a)) / out_a,
        out_a,
    )


def draw_canvas():
    """返回 SIZE*SIZE 的像素列表（每项 RGBA 浮点/整数）。"""
    px = [[(0.0, 0.0, 0.0, 0.0)] * SIZE for _ in range(SIZE)]

    # 背景蓝圆角方块
    for yy in range(SIZE):
        for xx in range(SIZE):
            a = rounded_rect_alpha(xx, yy, 0, 0, SIZE - 1, SIZE - 1, R)
            if a > 0:
                px[yy][xx] = blend(px[yy][xx], BG, a)

    # 白色文档页（圆角矩形）
    page = (64.0, 48.0, 208.0, 208.0)
    for yy in range(int(page[1]) - 2, int(page[3]) + 2):
        for xx in range(int(page[0]) - 2, int(page[2]) + 2):
            a = rounded_rect_alpha(xx, yy, page[0], page[1], page[2], page[3], 14)
            if a > 0:
                px[yy][xx] = blend(px[yy][xx], PAGE, a)

    def fill_rect(x0, y0, x1, y1, color, only_on_page=True):
        for yy in range(y0, y1):
            for xx in range(x0, x1):
                if not (0 <= xx < SIZE and 0 <= yy < SIZE):
                    continue
                if only_on_page and px[yy][xx][3] < 0.5:
                    continue  # 只画在不透明(白色页面)区域内
                px[yy][xx] = blend(px[yy][xx], color, 1.0)

    # 顶部灰色标题条
    fill_rect(84, 86, 188, 92, BAR)
    # 底部红色条
    fill_rect(84, 192, 188, 198, RED)

    # “PDF” 点阵字
    letters = ["P", "D", "F"]
    cell, gap = 6, 2
    total_cols = len(letters) * 5 + (len(letters) - 1) * gap
    start_x = (SIZE - total_cols * cell) // 2
    top = 106
    for li, ch in enumerate(letters):
        base_x = start_x + li * (5 + gap) * cell
        for row_i, row in enumerate(FONT[ch]):
            for col_i, bit in enumerate(row):
                if bit == "1":
                    fill_rect(base_x + col_i * cell, top + row_i * cell,
                              base_x + (col_i + 1) * cell,
                              top + (row_i + 1) * cell, INK)

    out = bytearray(SIZE * SIZE * 4)
    for yy in range(SIZE):
        for xx in range(SIZE):
            r, g, b, a = px[yy][xx]
            i = (yy * SIZE + xx) * 4
            out[i] = int(r + 0.5)
            out[i + 1] = int(g + 0.5)
            out[i + 2] = int(b + 0.5)
            out[i + 3] = int(a * 255 + 0.5)
    return bytes(out)


def downscale(rgba, size):
    """从 256 画布降采样到 size（取最近点即可，图标允许）。"""
    scale = SIZE // size
    out = bytearray(size * size * 4)
    for yy in range(size):
        by = (yy * scale + scale // 2) * SIZE
        for xx in range(size):
            i = (yy * size + xx) * 4
            bi = (by + xx * scale + scale // 2) * 4
            out[i:i + 4] = rgba[bi:bi + 4]
    return bytes(out)


def png_chunk(tag: bytes, data: bytes) -> bytes:
    return (struct.pack(">I", len(data)) + tag + data +
            struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))


def encode_png(rgba: bytes, w: int, h: int) -> bytes:
    raw = bytearray()
    for y in range(h):
        raw.append(0)
        raw += rgba[y * w * 4:(y + 1) * w * 4]
    return (b"\x89PNG\r\n\x1a\n" +
            png_chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 6, 0, 0, 0)) +
            png_chunk(b"IDAT", zlib.compress(bytes(raw), 9)) +
            png_chunk(b"IEND", b""))


def make_ico(out_path: str) -> None:
    canvas = draw_canvas()
    pngs = []
    for size in (16, 32, 48, 64, 128, 256):
        data = canvas if size == 256 else downscale(canvas, size)
        pngs.append((size, encode_png(data, size, size)))

    header = struct.pack("<HHH", 0, 1, len(pngs))
    entries = b""
    offset = 6 + 16 * len(pngs)
    for size, png in pngs:
        b = 0 if size == 256 else size
        entries += struct.pack("<BBBBHHII", b, b, 0, 0, 1, 32, len(png), offset)
        offset += len(png)
    with open(out_path, "wb") as f:
        f.write(header + entries + b"".join(p for _, p in pngs))
    print(f"已生成: {out_path} ({os.path.getsize(out_path)} bytes)")


def main() -> None:
    out = sys.argv[1] if len(sys.argv) > 1 else "assets/app.ico"
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    make_ico(out)


if __name__ == "__main__":
    main()
