#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用纯 Python 标准库生成应用图标（免 Pillow）。

- assets/app.ico   Windows 图标（内嵌 PNG，Vista+ 支持）
- assets/app.icns  macOS 图标（通过 iconutil 打包，需 macOS 系统自带工具）

图标主体：蓝色圆角方块 + 白色文档 + 深蓝块状 “PDF” 字样 + 红色底条。
用法:
    python make_icon.py            # 重新生成 assets 下两个图标
"""

import os
import struct
import subprocess
import sys
import tempfile
import zlib

BASE = 256
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


def draw_canvas(size=BASE):
    """绘制 size×size 的图标位图，返回 RGBA 字节。"""
    f = size / BASE
    px = [[(0.0, 0.0, 0.0, 0.0)] * size for _ in range(size)]

    def F(*vals):
        return [round(v * f) for v in vals]

    # 背景蓝圆角方块
    br = 56 * f
    for yy in range(size):
        for xx in range(size):
            a = rounded_rect_alpha(xx, yy, 0, 0, size - 1, size - 1, br)
            if a > 0:
                px[yy][xx] = blend(px[yy][xx], BG, a)

    # 白色文档页（圆角矩形）
    px0, py0, px1, py1 = F(64, 48, 208, 208)
    pr = 14 * f
    for yy in range(max(0, py0 - 2), min(size, py1 + 2)):
        for xx in range(max(0, px0 - 2), min(size, px1 + 2)):
            a = rounded_rect_alpha(xx, yy, px0, py0, px1, py1, pr)
            if a > 0:
                px[yy][xx] = blend(px[yy][xx], PAGE, a)

    def fill_rect(x0, y0, x1, y1, color):
        x0, y0, x1, y1 = F(x0, y0, x1, y1)
        for yy in range(max(0, y0), min(size, y1)):
            for xx in range(max(0, x0), min(size, x1)):
                if px[yy][xx][3] < 0.5:
                    continue  # 只画在不透明(白色页面)区域内
                px[yy][xx] = blend(px[yy][xx], color, 1.0)

    fill_rect(84, 86, 188, 92, BAR)     # 顶部灰色标题条
    fill_rect(84, 192, 188, 198, RED)   # 底部红色条

    # “PDF” 点阵字：坐标按 256 基准计算，交给 fill_rect 统一缩放
    letters = ["P", "D", "F"]
    cell0, gap0 = 6, 2
    total_cols = len(letters) * 5 + (len(letters) - 1) * gap0
    start_x0 = (BASE - total_cols * cell0) // 2
    top0 = 106
    for li, ch in enumerate(letters):
        base_x0 = start_x0 + li * (5 + gap0) * cell0
        for row_i, row in enumerate(FONT[ch]):
            for col_i, bit in enumerate(row):
                if bit == "1":
                    fill_rect(base_x0 + col_i * cell0, top0 + row_i * cell0,
                              base_x0 + (col_i + 1) * cell0,
                              top0 + (row_i + 1) * cell0, INK)

    out = bytearray(size * size * 4)
    for yy in range(size):
        for xx in range(size):
            r, g, b, a = px[yy][xx]
            i = (yy * size + xx) * 4
            out[i] = int(r + 0.5)
            out[i + 1] = int(g + 0.5)
            out[i + 2] = int(b + 0.5)
            out[i + 3] = int(a * 255 + 0.5)
    return bytes(out)


def resize_nearest(rgba, src, dst):
    """最近邻缩放到 dst×dst。"""
    out = bytearray(dst * dst * 4)
    for yy in range(dst):
        sy = yy * src // dst
        for xx in range(dst):
            sx = xx * src // dst
            i = (yy * dst + xx) * 4
            j = (sy * src + sx) * 4
            out[i:i + 4] = rgba[j:j + 4]
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


def make_ico(out_path: str, canvas256: bytes) -> None:
    pngs = []
    for size in (16, 32, 48, 64, 128, 256):
        data = canvas256 if size == 256 else resize_nearest(canvas256, 256, size)
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


def make_icns(out_path: str, canvas1024: bytes) -> None:
    """先生成 .iconset（10 张标准尺寸 PNG），再用 iconutil 转 .icns。"""
    if sys.platform != "darwin":
        print("提示: .icns 需在 macOS 上生成（依赖 iconutil），跳过。")
        return
    sets = [
        ("icon_16x16.png", 16),
        ("icon_16x16@2x.png", 32),
        ("icon_32x32.png", 32),
        ("icon_32x32@2x.png", 64),
        ("icon_128x128.png", 128),
        ("icon_128x128@2x.png", 256),
        ("icon_256x256.png", 256),
        ("icon_256x256@2x.png", 512),
        ("icon_512x512.png", 512),
        ("icon_512x512@2x.png", 1024),
    ]
    with tempfile.TemporaryDirectory(prefix="pdfproc_iconset_") as td:
        iconset = os.path.join(td, "app.iconset")
        os.makedirs(iconset)
        for name, size in sets:
            data = canvas1024 if size == 1024 else resize_nearest(canvas1024, 1024, size)
            with open(os.path.join(iconset, name), "wb") as fh:
                fh.write(encode_png(data, size, size))
        proc = subprocess.run(
            ["iconutil", "-c", "icns", iconset, "-o", out_path],
            capture_output=True, text=True)
        if proc.returncode != 0:
            raise RuntimeError(f"iconutil 失败: {proc.stderr}")
    print(f"已生成: {out_path} ({os.path.getsize(out_path)} bytes)")


def main() -> None:
    assets = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets")
    os.makedirs(assets, exist_ok=True)
    canvas256 = draw_canvas(BASE)
    make_ico(os.path.join(assets, "app.ico"), canvas256)
    if sys.platform == "darwin":
        canvas1024 = draw_canvas(1024)
        make_icns(os.path.join(assets, "app.icns"), canvas1024)


if __name__ == "__main__":
    main()
