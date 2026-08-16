#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成程序图标 交易汇总工具.ico（32x32，蓝底白色对勾）。

纯标准库实现：手工编码 PNG，再封装为 ICO（Windows Vista+ 支持 PNG 压缩的 ICO）。
"""

import struct
import zlib
from pathlib import Path

SIZE = 32
BG = (0x2F, 0x6F, 0xDB)   # 主蓝
FG = (0xFF, 0xFF, 0xFF)   # 白色对勾


def rounded_rect_inside(x: int, y: int, r: int = 7) -> bool:
    cx = cy = SIZE / 2.0
    half = SIZE / 2.0
    qx = abs(x - cx) - (half - r)
    qy = abs(y - cy) - (half - r)
    dx = max(qx, 0.0)
    dy = max(qy, 0.0)
    dist = (dx * dx + dy * dy) ** 0.5 + min(max(qx, qy), 0.0)
    return dist <= 0.0


def seg_dist(px, py, ax, ay, bx, by):
    vx, vy = bx - ax, by - ay
    wx, wy = px - ax, py - ay
    c1 = vx * wx + vy * wy
    if c1 <= 0:
        return ((px - ax) ** 2 + (py - ay) ** 2) ** 0.5
    c2 = vx * vx + vy * vy
    if c2 <= c1:
        return ((px - bx) ** 2 + (py - by) ** 2) ** 0.5
    t = c1 / c2
    return ((px - (ax + t * vx)) ** 2 + (py - (ay + t * vy)) ** 2) ** 0.5


CHECK_POINTS = [(7, 16), (13, 22), (25, 8)]


def check_inside(x: int, y: int, thickness: float = 1.8) -> bool:
    return any(
        seg_dist(x, y,
                 CHECK_POINTS[i][0], CHECK_POINTS[i][1],
                 CHECK_POINTS[i + 1][0], CHECK_POINTS[i + 1][1]) < thickness
        for i in range(len(CHECK_POINTS) - 1))


def make_png() -> bytes:
    raw = bytearray()
    for y in range(SIZE):
        raw.append(0)  # filter: None
        for x in range(SIZE):
            if not rounded_rect_inside(x, y):
                color = (0, 0, 0, 0)  # 透明
            elif check_inside(x, y):
                color = FG
            else:
                color = BG
            raw.extend(color[:3])

    def chunk(typ: bytes, data: bytes) -> bytes:
        return (struct.pack(">I", len(data)) + typ + data
                + struct.pack(">I", zlib.crc32(typ + data) & 0xFFFFFFFF))

    ihdr = struct.pack(">IIBBBBB", SIZE, SIZE, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
            + chunk(b"IEND", b""))


def make_ico() -> bytes:
    png = make_png()
    header = struct.pack("<HHH", 0, 1, 1)
    entry = struct.pack("<BBBBHHII", SIZE, SIZE, 0, 0, 1, 32, len(png), 22)
    return header + entry + png


if __name__ == "__main__":
    out = Path(__file__).with_name("交易汇总工具.ico")
    out.write_bytes(make_ico())
    print("已生成图标：", out, out.stat().st_size, "bytes")
