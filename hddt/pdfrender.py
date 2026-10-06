"""Dựng file PDF (bản thể hiện) cho từng hóa đơn từ dữ liệu XML / chi tiết GDT.

Cổng hoadondientu.gdt.gov.vn không cung cấp file PDF; PDF ở đây do tool dựng lại
từ chính dữ liệu hóa đơn (XML gốc đã ký, hoặc màn hình xem chi tiết của GDT),
có đủ thông tin người bán, người mua, từng dòng hàng, tổng tiền, mã CQT, ngày ký.

Font: cần font có dấu tiếng Việt. Tự tìm Arial/Times (Windows), DejaVu (Linux),
Arial (macOS). Có thể chỉ định bằng biến môi trường HDDT_PDF_FONT / HDDT_PDF_FONT_BOLD.
"""

from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
from typing import Iterable
from xml.sax.saxutils import escape

from .xmlparse import ParsedInvoice

_FONT_LOCK = threading.Lock()
_FONTS: tuple[str, str] | None = None

_FONT_CANDIDATES: list[tuple[str, str]] = [
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf"),
    ("C:/Windows/Fonts/times.ttf", "C:/Windows/Fonts/timesbd.ttf"),
    ("C:/Windows/Fonts/tahoma.ttf", "C:/Windows/Fonts/tahomabd.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
    ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf"),
    ("/System/Library/Fonts/Supplemental/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial Bold.ttf"),
]

INVOICE_TITLES = {
    "1": "HÓA ĐƠN GIÁ TRỊ GIA TĂNG",
    "2": "HÓA ĐƠN BÁN HÀNG",
    "3": "HÓA ĐƠN BÁN TÀI SẢN CÔNG",
    "4": "HÓA ĐƠN BÁN HÀNG DỰ TRỮ QUỐC GIA",
    "5": "HÓA ĐƠN ĐIỆN TỬ",
    "6": "PHIẾU XUẤT KHO KIÊM VẬN CHUYỂN NỘI BỘ",
}


class PdfFontError(RuntimeError):
    pass


def _fonts() -> tuple[str, str]:
    """Đăng ký font tiếng Việt một lần; trả về (tên font thường, tên font đậm)."""
    global _FONTS
    with _FONT_LOCK:
        if _FONTS:
            return _FONTS
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        candidates = list(_FONT_CANDIDATES)
        env_regular = os.environ.get("HDDT_PDF_FONT")
        if env_regular:
            candidates.insert(0, (env_regular, os.environ.get("HDDT_PDF_FONT_BOLD") or env_regular))
        for regular, bold in candidates:
            if not Path(regular).is_file():
                continue
            pdfmetrics.registerFont(TTFont("HDDT", regular))
            pdfmetrics.registerFont(TTFont("HDDT-Bold", bold if Path(bold).is_file() else regular))
            # Để thẻ <b> trong Paragraph dùng đúng font đậm.
            pdfmetrics.registerFontFamily("HDDT", normal="HDDT", bold="HDDT-Bold", italic="HDDT", boldItalic="HDDT-Bold")
            _FONTS = ("HDDT", "HDDT-Bold")
            return _FONTS
        raise PdfFontError(
            "Không tìm thấy font tiếng Việt (Arial/Times/DejaVu) để tạo PDF. "
            "Đặt biến HDDT_PDF_FONT trỏ tới file .ttf có dấu tiếng Việt."
        )


def _fmt_num(v: float | None, decimals: int = 0) -> str:
    if v is None:
        return ""
    if decimals == 0 and float(v).is_integer():
        s = f"{int(v):,}"
    else:
        s = f"{v:,.{max(decimals, 2)}f}".rstrip("0").rstrip(".")
    # Kiểu Việt Nam: 1.234.567,5
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def _fmt_qty(v: float | None) -> str:
    if v is None:
        return ""
    return _fmt_num(v, 0 if float(v).is_integer() else 4)


def _fmt_date(d: datetime | None, with_time: bool = False) -> str:
    if not d:
        return ""
    return d.strftime("%d/%m/%Y %H:%M:%S" if with_time else "%d/%m/%Y")


def _p(text: str) -> str:
    return escape(text or "").replace("\n", "<br/>")


def render_invoice_pdf(
    inv: ParsedInvoice,
    path: str | os.PathLike[str],
    *,
    source_note: str = "",
    status_lines: Iterable[str] = (),
) -> str:
    """Ghi PDF cho một hóa đơn. Trả về đường dẫn file."""
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    regular, bold = _fonts()
    base = ParagraphStyle("base", fontName=regular, fontSize=9, leading=11.5)
    small = ParagraphStyle("small", parent=base, fontSize=7.5, leading=9.5, textColor=colors.HexColor("#444444"))
    cell = ParagraphStyle("cell", parent=base, fontSize=8.5, leading=10.5)
    cell_r = ParagraphStyle("cell_r", parent=cell, alignment=TA_RIGHT)
    cell_c = ParagraphStyle("cell_c", parent=cell, alignment=TA_CENTER)
    head = ParagraphStyle("head", parent=cell, fontName=bold, alignment=TA_CENTER)
    title = ParagraphStyle("title", parent=base, fontName=bold, fontSize=15, leading=19, alignment=TA_CENTER, textColor=colors.HexColor("#B71C1C"))
    center = ParagraphStyle("center", parent=base, alignment=TA_CENTER)
    label_b = ParagraphStyle("label_b", parent=base, fontName=bold)

    out = Path(path)
    out.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(out), pagesize=A4, leftMargin=12 * mm, rightMargin=12 * mm, topMargin=10 * mm, bottomMargin=12 * mm,
        title=f"Hóa đơn {inv.khmshdon}{inv.khhdon} số {inv.shdon}", author=inv.nb_ten or "",
    )
    width = A4[0] - 24 * mm
    story: list = []

    # --- tiêu đề + ký hiệu
    heading = (inv.ten_hd or INVOICE_TITLES.get(inv.loai_hd or inv.khmshdon, "HÓA ĐƠN ĐIỆN TỬ")).upper()
    left = [
        Paragraph(f"<b>{_p(heading)}</b>", title),
        Spacer(1, 2),
        Paragraph(f"Ngày lập: <b>{_fmt_date(inv.ngay_lap)}</b>", center),
    ]
    if inv.mccqt:
        left.append(Paragraph(f"Mã của cơ quan thuế: <b>{_p(inv.mccqt)}</b>", center))
    right = [
        Paragraph(f"Ký hiệu mẫu số: <b>{_p(inv.khmshdon)}</b>", base),
        Paragraph(f"Ký hiệu: <b>{_p(inv.khhdon)}</b>", base),
        Paragraph(f"Số: <b><font color='#B71C1C' size='12'>{_p(inv.shdon)}</font></b>", base),
    ]
    t = Table([[left, right]], colWidths=[width * 0.72, width * 0.28])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    story += [t, Spacer(1, 6)]

    # --- người bán / người mua
    def party(label: str, ten: str, mst: str, dchi: str, extra: list[tuple[str, str]]) -> Table:
        rows = [
            [Paragraph(label, label_b), Paragraph(f"<b>{_p(ten)}</b>", base)],
            [Paragraph("Mã số thuế:", base), Paragraph(f"<b>{_p(mst)}</b>", base)],
            [Paragraph("Địa chỉ:", base), Paragraph(_p(dchi), base)],
        ]
        rows += [[Paragraph(k, base), Paragraph(_p(v), base)] for k, v in extra if v]
        tb = Table(rows, colWidths=[30 * mm, width - 30 * mm])
        tb.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ]))
        return tb

    story.append(party("Đơn vị bán:", inv.nb_ten, inv.nb_mst, inv.nb_dchi, [("Điện thoại:", inv.nb_sdt), ("Số tài khoản:", inv.nb_stk)]))
    story.append(Spacer(1, 4))
    story.append(party("Đơn vị mua:", inv.nm_ten, inv.nm_mst, inv.nm_dchi, [("Email:", inv.nm_email), ("Hình thức TT:", inv.httt)]))
    story.append(Spacer(1, 6))

    # --- bảng hàng hóa
    cols = ["STT", "Tên hàng hóa, dịch vụ", "ĐVT", "Số lượng", "Đơn giá", "Thành tiền", "Thuế suất", "Tiền thuế"]
    cw = [9 * mm, None, 14 * mm, 17 * mm, 22 * mm, 26 * mm, 15 * mm, 22 * mm]
    fixed = sum(w for w in cw if w)
    cw[1] = width - fixed
    data = [[Paragraph(c, head) for c in cols]]
    for ln in inv.lines:
        name = ln.ten_hang
        if ln.tchat == "2":
            name += " (Khuyến mại)"
        elif ln.tchat == "3":
            name += " (Chiết khấu thương mại)"
        data.append([
            Paragraph(_p(ln.stt), cell_c),
            Paragraph(_p(name), cell),
            Paragraph(_p(ln.dvt), cell_c),
            Paragraph(_fmt_qty(ln.so_luong), cell_r),
            Paragraph(_fmt_num(ln.don_gia, 2), cell_r),
            Paragraph(_fmt_num(ln.thanh_tien), cell_r),
            Paragraph(_p(ln.thue_suat), cell_c),
            Paragraph(_fmt_num(ln.tien_thue), cell_r),
        ])
    if len(data) == 1:
        data.append([Paragraph("", cell), Paragraph("<i>(Không có dữ liệu dòng hàng)</i>", cell)] + [Paragraph("", cell)] * 6)
    items = Table(data, colWidths=cw, repeatRows=1)
    items.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#888888")),
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E8EEF6")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
    ]))
    story += [items, Spacer(1, 6)]

    # --- tổng hợp theo thuế suất + tổng tiền
    sum_rows = []
    for lts in inv.thue_theo_suat:
        sum_rows.append([
            Paragraph(f"Thuế suất {_p(str(lts.get('thue_suat') or ''))}:", base),
            Paragraph(f"Tiền hàng {_fmt_num(lts.get('thanh_tien'))}", cell_r),
            Paragraph(f"Tiền thuế {_fmt_num(lts.get('tien_thue'))}", cell_r),
        ])
    totals = [
        ("Cộng tiền hàng (chưa thuế):", _fmt_num(inv.tong_tien_chua_thue)),
        ("Chiết khấu thương mại:", _fmt_num(inv.tong_ck) if inv.tong_ck else ""),
        ("Tiền thuế GTGT:", _fmt_num(inv.tong_tien_thue)),
        ("Tổng tiền thanh toán:", _fmt_num(inv.tong_tien_tt)),
    ]
    tot_rows = [[Paragraph(k, label_b if "Tổng" in k else base), Paragraph(f"<b>{v}</b>" if "Tổng" in k else v, cell_r)] for k, v in totals if v]
    if inv.dvtte and inv.dvtte.upper() != "VND":
        tot_rows.append([Paragraph("Đơn vị tiền tệ / Tỷ giá:", base), Paragraph(f"{_p(inv.dvtte)} / {_fmt_num(inv.ty_gia, 2)}", cell_r)])
    block = []
    if sum_rows:
        st = Table(sum_rows, colWidths=[width * 0.3, width * 0.35, width * 0.35])
        st.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
        block += [st, Spacer(1, 3)]
    if tot_rows:
        tt = Table(tot_rows, colWidths=[width * 0.6, width * 0.4])
        tt.setStyle(TableStyle([
            ("LINEABOVE", (0, -1), (-1, -1), 0.6, colors.black),
            ("LEFTPADDING", (0, 0), (-1, -1), 2), ("TOPPADDING", (0, 0), (-1, -1), 1.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ]))
        block.append(tt)
    if inv.tong_tien_chu:
        block += [Spacer(1, 3), Paragraph(f"Số tiền viết bằng chữ: <i>{_p(inv.tong_tien_chu)}</i>", base)]
    if block:
        story.append(KeepTogether(block))
    story.append(Spacer(1, 10))

    # --- chữ ký
    sign_left = [Paragraph("<b>NGƯỜI MUA HÀNG</b>", center), Paragraph("(Chữ ký số, nếu có)", small)]
    sign_right = [Paragraph("<b>NGƯỜI BÁN HÀNG</b>", center)]
    if inv.ngay_ky_nb or inv.nb_ten:
        sign_right.append(Paragraph(
            f"<font color='#1B5E20'>Ký bởi: {_p(inv.nb_ten)}<br/>Ký ngày: {_fmt_date(inv.ngay_ky_nb, True) or _fmt_date(inv.ngay_lap)}</font>",
            ParagraphStyle("sig", parent=small, alignment=TA_CENTER),
        ))
    if inv.ngay_ky_cqt:
        sign_right.append(Paragraph(f"CQT cấp mã ngày: {_fmt_date(inv.ngay_ky_cqt, True)}", ParagraphStyle("sig2", parent=small, alignment=TA_CENTER)))
    sg = Table([[sign_left, sign_right]], colWidths=[width / 2, width / 2])
    sg.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story += [KeepTogether([sg]), Spacer(1, 10)]

    # --- ghi chú nguồn
    notes = [s for s in status_lines if s]
    notes.append(source_note or "Bản thể hiện do công cụ tạo từ dữ liệu hóa đơn điện tử tải về từ hoadondientu.gdt.gov.vn.")
    for n in notes:
        story.append(Paragraph(_p(n), small))

    doc.build(story)
    return str(out)
