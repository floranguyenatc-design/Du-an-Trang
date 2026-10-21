"""Tách tiền hàng chưa thuế và tiền thuế của một hóa đơn theo từng mức thuế suất.

Nguồn số liệu, theo thứ tự ưu tiên:
1. Bảng tổng hợp theo thuế suất ghi trên chính hóa đơn (XML: TToan/THTTLTSuat/LTSuat;
   chi tiết trên cổng thuế: thttltsuat).
2. Bảng tổng hợp theo thuế suất trong danh sách hóa đơn cổng thuế trả về (thttltsuat).
3. Cộng từ các dòng hàng (chiết khấu thương mại ghi giảm, dòng ghi chú bỏ qua).
Không có nguồn nào thì toàn bộ tiền hàng được xếp vào "Chưa rõ thuế suất".
"""

from __future__ import annotations

import re
from typing import Any

from .xmlparse import ParsedInvoice, parse_number

BUCKETS: tuple[str, ...] = ("0%", "5%", "8%", "10%", "KCT", "KKKNT", "Khác")
TAXED_BUCKETS: tuple[str, ...] = ("5%", "8%", "10%", "Khác")
UNKNOWN = "Chưa rõ"

SRC_INVOICE = "Bảng thuế suất trên hóa đơn"
SRC_LIST = "Danh sách cổng thuế"
SRC_LINES = "Cộng từ dòng hàng"


def misa_vat_rate(text: Any) -> tuple[str, float | None]:
    """Chuẩn hóa thuế suất: trả về (mã, tỷ lệ cho thuế suất khác). Mã: 0, 5, 8, 10, KCT, KKKNT, KHAC."""
    t = str(text if text is not None else "").strip().upper().replace(" ", "")
    if not t:
        return "", None
    if t in ("KCT", "KKKNT"):
        return t, None
    if t.startswith("KHAC"):
        m = re.search(r"(\d+(?:[.,]\d+)?)", t)
        return "KHAC", (float(m.group(1).replace(",", ".")) if m else None)
    m = re.fullmatch(r"(\d+(?:[.,]\d+)?)%?", t)
    if not m:
        return "", None
    v = float(m.group(1).replace(",", "."))
    if 0 < v < 1 and "%" not in t:
        v *= 100  # cổng thuế có khi ghi 0.1 thay cho 10%
    v = round(v, 4)
    if v in (0, 5, 8, 10):
        return str(int(v)), None
    return "KHAC", v


def bucket_of(text: Any) -> str | None:
    code, _ = misa_vat_rate(text)
    if not code:
        return None
    if code in ("KCT", "KKKNT"):
        return code
    if code == "KHAC":
        return "Khác"
    return f"{code}%"


def _num(v: Any) -> float:
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    return parse_number(str(v)) or 0.0


def _from_summary(items: list[dict[str, Any]], rate_keys: tuple[str, ...], amt_keys: tuple[str, ...], tax_keys: tuple[str, ...]) -> dict[str, list[float]]:
    sums: dict[str, list[float]] = {}
    for it in items:
        if not isinstance(it, dict):
            continue
        rate = next((it[k] for k in rate_keys if it.get(k) not in (None, "")), None)
        b = bucket_of(rate)
        if b is None:
            continue
        amt = next((it[k] for k in amt_keys if it.get(k) not in (None, "")), 0)
        tax = next((it[k] for k in tax_keys if it.get(k) not in (None, "")), 0)
        s = sums.setdefault(b, [0.0, 0.0])
        s[0] += _num(amt)
        s[1] += _num(tax)
    return sums


def breakdown(raw: dict[str, Any] | None, parsed: ParsedInvoice | None) -> tuple[dict[str, list[float]], str]:
    """Trả về ({mức thuế suất: [tiền chưa thuế, tiền thuế]}, nguồn số liệu)."""
    if parsed is not None and parsed.thue_theo_suat:
        sums = _from_summary(parsed.thue_theo_suat, ("thue_suat",), ("thanh_tien",), ("tien_thue",))
        if sums:
            return sums, SRC_INVOICE
    items = (raw or {}).get("thttltsuat")
    if isinstance(items, list) and items:
        sums = _from_summary(items, ("tsuat", "ltsuat"), ("thtien", "tthtien"), ("tthue",))
        if sums:
            return sums, SRC_LIST
    if parsed is not None and parsed.lines:
        sums: dict[str, list[float]] = {}
        for ln in parsed.lines:
            if ln.tchat == "4":  # dòng ghi chú, diễn giải
                continue
            if ln.thanh_tien is None and ln.tien_thue is None:
                continue
            sign = -1 if ln.tchat == "3" else 1  # chiết khấu thương mại ghi giảm
            s = sums.setdefault(bucket_of(ln.thue_suat) or UNKNOWN, [0.0, 0.0])
            s[0] += sign * (ln.thanh_tien or 0)
            s[1] += sign * (ln.tien_thue or 0)
        if sums:
            return sums, SRC_LINES
    return {}, ""
