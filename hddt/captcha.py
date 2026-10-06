"""Giải CAPTCHA SVG của trang hoadondientu.gdt.gov.vn.

Trang thuế trả CAPTCHA dưới dạng JSON ``{"key": ..., "content": "<svg ...>"}``.
Mỗi ký tự trong ảnh là một thẻ ``<path d="...">`` vẽ bằng font cố định, nên
chỉ cần rút gọn chuỗi lệnh vẽ về dãy ``M``/``Q``/``Z`` là nhận ra được ký tự
(cách làm quen thuộc của các tool tải hóa đơn: TaiHoaDonDienTu, hddt-downloader).
Các path còn lại (nét nhiễu) không khớp bảng nên bị bỏ qua. Thứ tự ký tự lấy
theo tọa độ x của lệnh ``M`` đầu tiên.
"""

from __future__ import annotations

import re
from dataclasses import dataclass

# Bảng keyword M/Q/Z của font CAPTCHA. Index 0-25 -> A-Z, 26-35 -> 0-9.
# Các ký tự dễ nhầm (I, L, O, U, 0, 1) không có trong font nên để rỗng.
_KEYWORDS: list[str] = [
    "MQQQQQZMQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZMQQZ",  # A
    "MQQQQQQQQQZMQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQZMQQQQQQQQZ",  # B
    "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ",  # C
    "MQQQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQZ",  # D
    "MQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # E
    "MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ",  # F
    "MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # G
    "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQZ",  # H
    "",  # I
    "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQZ",  # J
    "MQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ",  # K
    "",  # L
    "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # M
    "MQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQZ",  # N
    "",  # O
    "MQQQQQQZMQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ",  # P
    "MQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQZ",  # Q
    "MQQQQQQZMQQQQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQZ",  # R
    "MQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # S
    "MQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQZ",  # T
    "",  # U
    "MQQQQQQQQQQZMQQQQQQQQQQQQQQQQZ",  # V
    "MQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # W
    "MQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZ",  # X
    "MQQQQQQQQQZMQQQQQQQQQQQQQZ",  # Y
    "MQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZ",  # Z
    "",  # 0
    "",  # 1
    "MQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # 2
    "MQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # 3
    "MQQQQZMQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQZ",  # 4
    "MQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQQQQQQQQQZ",  # 5
    "MQQQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQZ",  # 6
    "MQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQQZ",  # 7
    "MQQQQQQQQZMQQQQQQQZMQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQZMQQQQQQQZ",  # 8
    "MQQQQQQQQZMQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQQQQQQQQQQZMQQQQQQQQQQQZ",  # 9
]

_ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
_KEYWORD_TO_CHAR: dict[str, str] = {
    kw: _ALPHABET[i] for i, kw in enumerate(_KEYWORDS) if kw
}

_PATH_D = re.compile(r'\bd="([^"]*)"')
_COMMAND = re.compile(r"([MQZ])([^MQZ]*)")
_LEADING_NUMBER = re.compile(r"^\s*(-?(?:\d+(?:\.\d+)?|\.\d+))")


@dataclass
class CaptchaGlyph:
    x: float
    char: str


def _leading_number(text: str) -> float:
    m = _LEADING_NUMBER.match(text or "")
    return float(m.group(1)) if m else 0.0


def solve_svg_captcha(svg: str) -> str:
    """Trả về mã CAPTCHA đọc được từ nội dung SVG, hoặc chuỗi rỗng nếu không nhận ra."""
    if not svg:
        return ""
    text = svg.replace('\\"', '"')
    glyphs: list[CaptchaGlyph] = []
    for d in _PATH_D.findall(text):
        matches = _COMMAND.findall(d)
        if not matches:
            continue
        keyword = "".join(cmd for cmd, _ in matches)
        char = _KEYWORD_TO_CHAR.get(keyword)
        if char is None:
            continue
        glyphs.append(CaptchaGlyph(x=_leading_number(matches[0][1]), char=char))
    glyphs.sort(key=lambda g: g.x)
    return "".join(g.char for g in glyphs)
