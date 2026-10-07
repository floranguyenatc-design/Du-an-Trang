"""Đọc file XML hóa đơn điện tử theo chuẩn Thông tư 78 (định dạng của Tổng cục Thuế).

Cấu trúc chính::

    HDon
      DLHDon (Id=...)
        TTChung : KHMSHDon, KHHDon, SHDon, NLap, DVTTe, TGia, HTTToan, MSTTCGP, TTKhac...
        NDHDon
          NBan  : Ten, MST, DChi, SDThoai, ...
          NMua  : Ten, MST, DChi, ...
          DSHHDVu/HHDVu : STT, TChat, MHHDVu, THHDVu, DVTinh, SLuong, DGia, TLCKhau, STCKhau, ThTien, TSuat, TTKhac...
          TToan : THTTLTSuat/LTSuat (ThTien, TSuat, TThue), TgTCThue, TgTThue, TTCKTMai, TgTTTBSo, TgTTTBChu
      MCCQT : mã của cơ quan thuế
      DSCKS : chữ ký số (NBan, CQT -> SigningTime)

Namespace không cố định giữa các nhà cung cấp nên tra theo local-name.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any
from xml.etree import ElementTree as ET

_NUM_CLEAN = re.compile(r"[^0-9,.\-]")


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _child(node: ET.Element | None, name: str) -> ET.Element | None:
    if node is None:
        return None
    for ch in node:
        if _local(ch.tag) == name:
            return ch
    return None


def _children(node: ET.Element | None, name: str) -> list[ET.Element]:
    if node is None:
        return []
    return [ch for ch in node if _local(ch.tag) == name]


def _find(node: ET.Element | None, name: str) -> ET.Element | None:
    """Tìm phần tử con (mọi độ sâu) theo local-name."""
    if node is None:
        return None
    for el in node.iter():
        if el is not node and _local(el.tag) == name:
            return el
    return None


def _text(node: ET.Element | None, name: str = "", default: str = "") -> str:
    if node is None:
        return default
    target = node if not name else _child(node, name)
    if target is None or target.text is None:
        return default
    return target.text.strip()


def parse_number(value: str | None) -> float | None:
    """Đổi chuỗi số của hóa đơn (có thể dùng dấu , hoặc .) thành float."""
    if value is None:
        return None
    s = _NUM_CLEAN.sub("", str(value).strip())
    if not s or s in ("-", ".", ","):
        return None
    if "," in s and "." in s:
        # Dấu nào đứng sau cùng là dấu thập phân.
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        # 1,5 -> thập phân ; 1,000,000 -> nghìn
        parts = s.split(",")
        s = s.replace(",", ".") if len(parts) == 2 and len(parts[1]) != 3 else s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def parse_date(value: str | None) -> datetime | None:
    if not value:
        return None
    v = value.strip()
    for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y"):
        try:
            dt = datetime.strptime(v.replace("Z", "+0000") if fmt.endswith("%z") else v, fmt)
            return dt.replace(tzinfo=None) if dt.tzinfo is None else dt.astimezone().replace(tzinfo=None)
        except ValueError:
            continue
    return None


def _extra_fields(node: ET.Element | None) -> dict[str, str]:
    """TTKhac: danh sách TTin{TTruong, KDLieu, DLieu}."""
    out: dict[str, str] = {}
    ttkhac = _child(node, "TTKhac")
    if ttkhac is None:
        return out
    for info in _children(ttkhac, "TTin"):
        key = _text(info, "TTruong")
        if key:
            out[key] = _text(info, "DLieu")
    return out


@dataclass
class InvoiceLine:
    stt: str = ""
    tchat: str = ""
    ma_hang: str = ""
    ten_hang: str = ""
    dvt: str = ""
    so_luong: float | None = None
    don_gia: float | None = None
    tl_ck: float | None = None
    st_ck: float | None = None
    thanh_tien: float | None = None
    thue_suat: str = ""
    tien_thue: float | None = None
    thanh_tien_co_thue: float | None = None
    extra: dict[str, str] = field(default_factory=dict)


@dataclass
class ParsedInvoice:
    invoice_id: str = ""
    loai_hd: str = ""
    ten_hd: str = ""
    khmshdon: str = ""
    khhdon: str = ""
    shdon: str = ""
    ngay_lap: datetime | None = None
    dvtte: str = ""
    ty_gia: float | None = None
    httt: str = ""
    mst_tcgp: str = ""
    nb_ten: str = ""
    nb_mst: str = ""
    nb_dchi: str = ""
    nb_sdt: str = ""
    nb_stk: str = ""
    nm_ten: str = ""
    nm_hvtn: str = ""  # Họ và tên người mua hàng (hóa đơn cho cá nhân thường chỉ có trường này)
    nm_mst: str = ""
    nm_dchi: str = ""
    nm_email: str = ""
    mccqt: str = ""
    tong_tien_chua_thue: float | None = None
    tong_tien_thue: float | None = None
    tong_ck: float | None = None
    tong_tien_tt: float | None = None
    tong_tien_chu: str = ""
    thue_theo_suat: list[dict[str, Any]] = field(default_factory=list)
    ngay_ky_nb: datetime | None = None
    ngay_ky_cqt: datetime | None = None
    extra: dict[str, str] = field(default_factory=dict)
    lines: list[InvoiceLine] = field(default_factory=list)


def parse_invoice_xml(data: bytes | str) -> ParsedInvoice:
    if isinstance(data, str):
        data = data.encode("utf-8")
    # Chặn DTD/thực thể ngoài: ElementTree mặc định không tải thực thể ngoài.
    root = ET.fromstring(data)
    hdon = root if _local(root.tag) == "HDon" else _find(root, "HDon")
    if hdon is None:
        raise ValueError("File XML không có phần tử HDon.")
    dl = _child(hdon, "DLHDon")
    ttchung = _child(dl, "TTChung")
    nd = _child(dl, "NDHDon")
    nban = _child(nd, "NBan")
    nmua = _child(nd, "NMua")
    ttoan = _child(nd, "TToan")

    inv = ParsedInvoice(
        invoice_id=(dl.get("Id") if dl is not None else "") or "",
        loai_hd=hdon.get("LoaiHD", "") or "",
        ten_hd=_text(ttchung, "THDon"),
        khmshdon=_text(ttchung, "KHMSHDon"),
        khhdon=_text(ttchung, "KHHDon"),
        shdon=_text(ttchung, "SHDon"),
        ngay_lap=parse_date(_text(ttchung, "NLap")),
        dvtte=_text(ttchung, "DVTTe"),
        ty_gia=parse_number(_text(ttchung, "TGia")),
        httt=_text(ttchung, "HTTToan"),
        mst_tcgp=_text(ttchung, "MSTTCGP"),
        nb_ten=_text(nban, "Ten"),
        nb_mst=_text(nban, "MST"),
        nb_dchi=_text(nban, "DChi"),
        nb_sdt=_text(nban, "SDThoai"),
        nb_stk=_text(nban, "STKNHang"),
        nm_ten=_text(nmua, "Ten"),
        nm_hvtn=_text(nmua, "HVTNMHang"),
        nm_mst=_text(nmua, "MST"),
        nm_dchi=_text(nmua, "DChi"),
        nm_email=_text(nmua, "DCTDTu"),
        mccqt=_text(_find(hdon, "MCCQT")),
        tong_tien_chua_thue=parse_number(_text(ttoan, "TgTCThue")),
        tong_tien_thue=parse_number(_text(ttoan, "TgTThue")),
        tong_ck=parse_number(_text(ttoan, "TTCKTMai")),
        tong_tien_tt=parse_number(_text(ttoan, "TgTTTBSo")),
        tong_tien_chu=_text(ttoan, "TgTTTBChu"),
    )
    inv.extra.update(_extra_fields(ttchung))
    inv.extra.update(_extra_fields(dl))

    for lts in _children(_child(ttoan, "THTTLTSuat"), "LTSuat"):
        inv.thue_theo_suat.append(
            {
                "thue_suat": _text(lts, "TSuat"),
                "thanh_tien": parse_number(_text(lts, "ThTien")),
                "tien_thue": parse_number(_text(lts, "TThue")),
            }
        )

    dscks = _child(hdon, "DSCKS")
    inv.ngay_ky_nb = parse_date(_text(_find(_child(dscks, "NBan"), "SigningTime")))
    inv.ngay_ky_cqt = parse_date(_text(_find(_child(dscks, "CQT"), "SigningTime")))

    for item in _children(_child(nd, "DSHHDVu"), "HHDVu"):
        extra = _extra_fields(item)
        tien_thue = parse_number(_text(item, "TThue"))
        if tien_thue is None:
            for k in ("TThue", "TienThueGTGT", "ThueGTGT", "TThueVAT", "Tiền thuế"):
                if k in extra:
                    tien_thue = parse_number(extra[k])
                    break
        co_thue = None
        for k in ("THTienCoVAT", "ThanhTienCoVAT", "TongTienCoThue", "Thành tiền có thuế"):
            if k in extra:
                co_thue = parse_number(extra[k])
                break
        thanh_tien = parse_number(_text(item, "ThTien"))
        if co_thue is None and thanh_tien is not None and tien_thue is not None:
            co_thue = thanh_tien + tien_thue
        inv.lines.append(
            InvoiceLine(
                stt=_text(item, "STT"),
                tchat=_text(item, "TChat"),
                ma_hang=_text(item, "MHHDVu"),
                ten_hang=_text(item, "THHDVu") or _text(item, "Ten"),
                dvt=_text(item, "DVTinh"),
                so_luong=parse_number(_text(item, "SLuong")),
                don_gia=parse_number(_text(item, "DGia")),
                tl_ck=parse_number(_text(item, "TLCKhau")),
                st_ck=parse_number(_text(item, "STCKhau")),
                thanh_tien=thanh_tien,
                thue_suat=_text(item, "TSuat") or _text(item, "LTSuat"),
                tien_thue=tien_thue,
                thanh_tien_co_thue=co_thue,
                extra=extra,
            )
        )
    return inv


def parse_invoice_file(path: str) -> ParsedInvoice:
    with open(path, "rb") as f:
        return parse_invoice_xml(f.read())
