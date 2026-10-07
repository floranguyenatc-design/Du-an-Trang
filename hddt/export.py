"""Xuất kết quả ra Excel (openpyxl): sheet HoaDon, ChiTiet, Loi."""

from __future__ import annotations

from datetime import datetime
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .client import TTHAI_LABELS, TTXLY_LABELS, InvoiceRef
from .xmlparse import ParsedInvoice, parse_date

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")

INVOICE_COLUMNS: list[tuple[str, str]] = [
    ("Chiều", "chieu"),
    ("Nguồn", "nguon"),
    ("Ký hiệu mẫu số", "khmshdon"),
    ("Ký hiệu hóa đơn", "khhdon"),
    ("Số hóa đơn", "shdon"),
    ("Ngày lập", "ngay_lap"),
    ("MST người bán", "nbmst"),
    ("Tên người bán", "nbten"),
    ("Địa chỉ người bán", "nbdchi"),
    ("MST người mua", "nmmst"),
    ("Tên người mua", "nmten"),
    ("Địa chỉ người mua", "nmdchi"),
    ("Tổng tiền chưa thuế", "tgtcthue"),
    ("Tổng tiền thuế", "tgtthue"),
    ("Tổng tiền chiết khấu", "ttcktmai"),
    ("Tổng tiền thanh toán", "tgtttbso"),
    ("Đơn vị tiền tệ", "dvtte"),
    ("Tỷ giá", "tgia"),
    ("Trạng thái hóa đơn", "tthai_label"),
    ("Kết quả kiểm tra", "ttxly_label"),
    ("Mã CQT", "mhdon"),
    ("Ngày ký", "nky"),
    ("Ngày cấp mã", "ncma"),
    ("Hóa đơn gốc (mẫu số/ký hiệu/số)", "hd_goc"),
    ("Ngày HĐ gốc", "tdlhdgoc"),
    ("Ghi chú HĐ gốc", "gchdgoc"),
    ("Hình thức thanh toán", "htttoan"),
    ("MST TCGP", "msttcgp"),
    ("File PDF", "pdf_file"),
    ("File XML/JSON", "xml_file"),
    ("Ghi chú tải XML", "xml_error"),
]

LINE_COLUMNS: list[tuple[str, str]] = [
    ("Chiều", "chieu"),
    ("Ký hiệu mẫu số", "khmshdon"),
    ("Ký hiệu hóa đơn", "khhdon"),
    ("Số hóa đơn", "shdon"),
    ("Ngày lập", "ngay_lap"),
    ("MST người bán", "nbmst"),
    ("Tên người bán", "nbten"),
    ("MST người mua", "nmmst"),
    ("Tên người mua", "nmten"),
    ("STT", "stt"),
    ("Tính chất", "tchat"),
    ("Mã hàng", "ma_hang"),
    ("Tên hàng hóa, dịch vụ", "ten_hang"),
    ("ĐVT", "dvt"),
    ("Số lượng", "so_luong"),
    ("Đơn giá", "don_gia"),
    ("Tỷ lệ CK", "tl_ck"),
    ("Số tiền CK", "st_ck"),
    ("Thành tiền", "thanh_tien"),
    ("Thuế suất", "thue_suat"),
    ("Tiền thuế", "tien_thue"),
    ("Thành tiền có thuế", "thanh_tien_co_thue"),
]

ERROR_COLUMNS = ["Chiều", "Nguồn", "MST người bán", "Mẫu số", "Ký hiệu", "Số HĐ", "Bước", "Lỗi"]

TCHAT_LABELS = {"1": "Hàng hóa, dịch vụ", "2": "Khuyến mại", "3": "Chiết khấu thương mại", "4": "Ghi chú/diễn giải"}


def _label(mapping: dict[int, str], value: Any) -> str:
    try:
        n = int(str(value))
    except (TypeError, ValueError):
        return str(value or "")
    return mapping.get(n, str(value))


def _num(value: Any) -> Any:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return value
    try:
        return float(str(value).replace(",", ""))
    except ValueError:
        return value


def _date(value: Any) -> Any:
    if isinstance(value, datetime):
        return value
    dt = parse_date(str(value)) if value else None
    return dt or (value or None)


def invoice_row(ref: InvoiceRef, parsed: ParsedInvoice | None = None, xml_file: str = "", xml_error: str = "") -> dict[str, Any]:
    r = ref.raw
    hd_goc = ""
    if r.get("khmshdgoc") or r.get("shdgoc"):
        hd_goc = f"{r.get('khmshdgoc') or ''}{r.get('khhdgoc') or ''} / {r.get('shdgoc') or ''}"
    row: dict[str, Any] = {
        "chieu": "Mua vào" if ref.direction == "purchase" else "Bán ra",
        "nguon": "Máy tính tiền" if ref.family == "sco-query" else "Hóa đơn điện tử",
        "khmshdon": ref.khmshdon,
        "khhdon": ref.khhdon,
        "shdon": _num(ref.shdon),
        "ngay_lap": _date(r.get("tdlap")),
        "nbmst": ref.nbmst,
        "nbten": r.get("nbten", ""),
        "nbdchi": r.get("nbdchi", ""),
        "nmmst": r.get("nmmst", ""),
        "nmten": r.get("nmten") or r.get("nmtnmua") or "",
        "nmdchi": r.get("nmdchi", ""),
        "tgtcthue": _num(r.get("tgtcthue")),
        "tgtthue": _num(r.get("tgtthue")),
        "ttcktmai": _num(r.get("ttcktmai")),
        "tgtttbso": _num(r.get("tgtttbso")),
        "dvtte": r.get("dvtte", ""),
        "tgia": _num(r.get("tgia")),
        "tthai_label": _label(TTHAI_LABELS, r.get("tthai")),
        "ttxly_label": _label(TTXLY_LABELS, r.get("ttxly")),
        "mhdon": r.get("mhdon", ""),
        "nky": _date(r.get("nky")),
        "ncma": _date(r.get("ncma")),
        "hd_goc": hd_goc,
        "tdlhdgoc": _date(r.get("tdlhdgoc")),
        "gchdgoc": r.get("gchdgoc", ""),
        "htttoan": r.get("thtttoan") or r.get("htttoan", ""),
        "msttcgp": r.get("msttcgp", ""),
        "xml_file": xml_file,
        "xml_error": xml_error,
    }
    if parsed is not None:
        # Ưu tiên số liệu trong XML gốc nếu danh sách thiếu.
        fill = {
            "nbten": parsed.nb_ten, "nbdchi": parsed.nb_dchi, "nmmst": parsed.nm_mst, "nmten": parsed.nm_ten or parsed.nm_hvtn,
            "nmdchi": parsed.nm_dchi, "tgtcthue": parsed.tong_tien_chua_thue, "tgtthue": parsed.tong_tien_thue,
            "ttcktmai": parsed.tong_ck, "tgtttbso": parsed.tong_tien_tt, "dvtte": parsed.dvtte, "tgia": parsed.ty_gia,
            "mhdon": parsed.mccqt, "nky": parsed.ngay_ky_nb, "ncma": parsed.ngay_ky_cqt, "htttoan": parsed.httt,
            "msttcgp": parsed.mst_tcgp, "ngay_lap": parsed.ngay_lap,
        }
        for k, v in fill.items():
            if (row.get(k) in (None, "")) and v not in (None, ""):
                row[k] = v
    return row


def line_rows(ref: InvoiceRef, parsed: ParsedInvoice) -> list[dict[str, Any]]:
    base = {
        "chieu": "Mua vào" if ref.direction == "purchase" else "Bán ra",
        "khmshdon": ref.khmshdon or parsed.khmshdon,
        "khhdon": ref.khhdon or parsed.khhdon,
        "shdon": _num(ref.shdon or parsed.shdon),
        "ngay_lap": parsed.ngay_lap or _date(ref.raw.get("tdlap")),
        "nbmst": ref.nbmst or parsed.nb_mst,
        "nbten": parsed.nb_ten or ref.raw.get("nbten", ""),
        "nmmst": parsed.nm_mst or ref.raw.get("nmmst", ""),
        "nmten": parsed.nm_ten or parsed.nm_hvtn or ref.raw.get("nmten") or ref.raw.get("nmtnmua") or "",
    }
    rows = []
    for ln in parsed.lines:
        rows.append(
            {
                **base,
                "stt": _num(ln.stt),
                "tchat": TCHAT_LABELS.get(ln.tchat, ln.tchat),
                "ma_hang": ln.ma_hang,
                "ten_hang": ln.ten_hang,
                "dvt": ln.dvt,
                "so_luong": ln.so_luong,
                "don_gia": ln.don_gia,
                "tl_ck": ln.tl_ck,
                "st_ck": ln.st_ck,
                "thanh_tien": ln.thanh_tien,
                "thue_suat": ln.thue_suat,
                "tien_thue": ln.tien_thue,
                "thanh_tien_co_thue": ln.thanh_tien_co_thue,
            }
        )
    return rows


def _write_sheet(ws, columns: list[tuple[str, str]], rows: Iterable[dict[str, Any]]) -> None:
    ws.append([h for h, _ in columns])
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    count = 0
    for row in rows:
        ws.append([row.get(key) for _, key in columns])
        count += 1
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(count + 1, 1)}"
    for idx, (header, key) in enumerate(columns, start=1):
        width = max(10, min(45, len(header) + 2))
        if key in ("nbten", "nmten", "ten_hang", "nbdchi", "nmdchi"):
            width = 40
        ws.column_dimensions[get_column_letter(idx)].width = width
        if key in ("ngay_lap", "nky", "ncma", "tdlhdgoc"):
            for cell in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                cell[0].number_format = "dd/mm/yyyy"
        if key in ("tgtcthue", "tgtthue", "ttcktmai", "tgtttbso", "thanh_tien", "tien_thue", "thanh_tien_co_thue", "don_gia", "st_ck"):
            for cell in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                cell[0].number_format = "#,##0.##"


def write_workbook(
    path: str,
    invoices: list[dict[str, Any]],
    lines: list[dict[str, Any]],
    errors: list[dict[str, Any]],
) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "HoaDon"
    _write_sheet(ws, INVOICE_COLUMNS, invoices)
    _write_sheet(wb.create_sheet("ChiTiet"), LINE_COLUMNS, lines)
    err_cols = [(h, h) for h in ERROR_COLUMNS]
    _write_sheet(wb.create_sheet("Loi"), err_cols, errors)
    wb.save(path)
