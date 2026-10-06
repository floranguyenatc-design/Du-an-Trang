"""Dựng ParsedInvoice từ JSON chi tiết hóa đơn của GDT (endpoint invoices/detail).

Dùng cho hóa đơn không có XML gốc trên cổng thuế (hóa đơn không mã của viễn thông,
ngân hàng, điện nước... chỉ gửi dữ liệu tổng hợp): export-xml trả HTTP 500 nhưng
màn hình "Xem chi tiết" của GDT vẫn có dòng hàng hóa trong trường ``hdhhdvu``.
"""

from __future__ import annotations

from typing import Any

from .xmlparse import InvoiceLine, ParsedInvoice, parse_date, parse_number


def _s(d: dict[str, Any], *keys: str) -> str:
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return str(v)
    return ""


def _n(d: dict[str, Any], *keys: str) -> float | None:
    for k in keys:
        v = d.get(k)
        if v in (None, ""):
            continue
        if isinstance(v, (int, float)):
            return float(v)
        n = parse_number(str(v))
        if n is not None:
            return n
    return None


def _extra_list(value: Any) -> dict[str, str]:
    out: dict[str, str] = {}
    if isinstance(value, list):
        for item in value:
            if isinstance(item, dict):
                key = _s(item, "ttruong", "TTruong")
                if key:
                    out[key] = _s(item, "dlieu", "DLieu")
    return out


def _tax_rate_text(item: dict[str, Any]) -> str:
    """Thuế suất dạng chữ: ưu tiên ltsuat ("10%", "KCT"), rồi tsuat số."""
    lt = _s(item, "ltsuat")
    if lt:
        return lt
    ts = item.get("tsuat")
    if ts in (None, ""):
        return ""
    try:
        f = float(ts)
    except (TypeError, ValueError):
        return str(ts)
    # GDT ghi thuế suất dạng tỷ lệ (0.1 = 10%, 0.08 = 8%).
    if 0 < f < 1:
        f = round(f * 100, 4)
    return f"{int(f)}%" if float(f).is_integer() else f"{f:g}%"


def _tax_rate_value(text: str) -> float | None:
    t = text.strip().replace("%", "")
    try:
        return float(t)
    except ValueError:
        return None


def parsed_from_detail(detail: dict[str, Any]) -> ParsedInvoice:
    inv = ParsedInvoice(
        invoice_id=_s(detail, "id"),
        loai_hd=_s(detail, "tlhdon", "hdon"),
        khmshdon=_s(detail, "khmshdon"),
        khhdon=_s(detail, "khhdon"),
        shdon=_s(detail, "shdon"),
        ngay_lap=parse_date(_s(detail, "tdlap")),
        dvtte=_s(detail, "dvtte"),
        ty_gia=_n(detail, "tgia"),
        httt=_s(detail, "thtttoan", "htttoan"),
        mst_tcgp=_s(detail, "msttcgp"),
        nb_ten=_s(detail, "nbten"),
        nb_mst=_s(detail, "nbmst"),
        nb_dchi=_s(detail, "nbdchi"),
        nb_sdt=_s(detail, "nbsdthoai"),
        nb_stk=_s(detail, "nbstkhoan"),
        nm_ten=_s(detail, "nmten"),
        nm_mst=_s(detail, "nmmst"),
        nm_dchi=_s(detail, "nmdchi"),
        nm_email=_s(detail, "nmdctdtu"),
        mccqt=_s(detail, "mhdon"),
        tong_tien_chua_thue=_n(detail, "tgtcthue"),
        tong_tien_thue=_n(detail, "tgtthue"),
        tong_ck=_n(detail, "ttcktmai"),
        tong_tien_tt=_n(detail, "tgtttbso"),
        tong_tien_chu=_s(detail, "tgtttbchu"),
        ngay_ky_nb=parse_date(_s(detail, "nky")),
        ngay_ky_cqt=parse_date(_s(detail, "ncma")),
    )
    inv.extra.update(_extra_list(detail.get("ttkhac")))
    for lts in detail.get("thttltsuat") or []:
        if isinstance(lts, dict):
            inv.thue_theo_suat.append(
                {"thue_suat": _s(lts, "tsuat"), "thanh_tien": _n(lts, "thtien"), "tien_thue": _n(lts, "tthue")}
            )

    for item in detail.get("hdhhdvu") or []:
        if not isinstance(item, dict):
            continue
        extra = _extra_list(item.get("ttkhac"))
        rate_text = _tax_rate_text(item)
        thanh_tien = _n(item, "thtien")
        tien_thue = _n(item, "tthue")
        if tien_thue is None:
            for k in ("TThue", "TienThueGTGT", "ThueGTGT", "TThueVAT"):
                if k in extra:
                    tien_thue = parse_number(extra[k])
                    break
        if tien_thue is None and thanh_tien is not None:
            rate = _tax_rate_value(rate_text)
            if rate is not None:
                tien_thue = round(thanh_tien * rate / 100.0, 2)
        co_thue = None
        for k in ("THTienCoVAT", "ThanhTienCoVAT", "TongTienCoThue"):
            if k in extra:
                co_thue = parse_number(extra[k])
                break
        if co_thue is None and thanh_tien is not None and tien_thue is not None:
            co_thue = thanh_tien + tien_thue
        inv.lines.append(
            InvoiceLine(
                stt=_s(item, "stt"),
                tchat=_s(item, "tchat"),
                ma_hang=_s(item, "mhhdvu"),
                ten_hang=_s(item, "ten", "thhdvu"),
                dvt=_s(item, "dvtinh"),
                so_luong=_n(item, "sluong"),
                don_gia=_n(item, "dgia"),
                tl_ck=_n(item, "tlckhau"),
                st_ck=_n(item, "stckhau"),
                thanh_tien=thanh_tien,
                thue_suat=rate_text,
                tien_thue=tien_thue,
                thanh_tien_co_thue=co_thue,
                extra=extra,
            )
        )
    return inv
