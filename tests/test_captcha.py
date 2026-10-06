from hddt.captcha import _KEYWORDS, solve_svg_captcha


def _path_from_keyword(keyword: str, x: float) -> str:
    """Dựng chuỗi d từ keyword M/Q/Z với tọa độ x cho lệnh M đầu tiên."""
    parts = []
    first = True
    for cmd in keyword:
        if cmd == "M":
            parts.append(f"M{x if first else 5} 20" if first else "M5 20")
            first = False
        elif cmd == "Q":
            parts.append("Q1 2 3 4")
        else:
            parts.append("Z")
    return " ".join(parts)


def _svg(chars: str, noise: bool = True) -> str:
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    paths = []
    # Đặt các ký tự theo thứ tự x ngược để kiểm tra sắp xếp.
    for i, ch in enumerate(reversed(chars)):
        kw = _KEYWORDS[alphabet.index(ch)]
        x = 10 + (len(chars) - 1 - i) * 25.5
        paths.append(f'<path d="{_path_from_keyword(kw, x)}" fill="#111"/>')
    if noise:
        paths.append('<path d="M1 2 L3 4 C5 6 7 8 9 10" stroke="#ccc"/>')
        paths.append('<path d="M9 9 Q1 1 2 2 Q3 3 4 4 Z" stroke="#ccc"/>')
    return '<svg xmlns="http://www.w3.org/2000/svg" width="150" height="50">' + "".join(paths) + "</svg>"


def test_solve_sorted_by_x_and_ignores_noise():
    assert solve_svg_captcha(_svg("A7KX2")) == "A7KX2"


def test_solve_handles_json_escaped_quotes():
    svg = _svg("B3").replace('"', '\\"')
    assert solve_svg_captcha(svg) == "B3"


def test_empty_or_unknown_returns_empty():
    assert solve_svg_captcha("") == ""
    assert solve_svg_captcha('<svg><path d="M1 1 L2 2"/></svg>') == ""


def test_all_known_chars_round_trip():
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    known = "".join(c for c, kw in zip(alphabet, _KEYWORDS) if kw)
    assert len(known) == 30
    assert solve_svg_captcha(_svg(known, noise=False)) == known
