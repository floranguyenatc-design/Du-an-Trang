"""Chuẩn bị dữ liệu hóa đơn đã kéo về để đưa vào phần mềm kế toán MISA SME.

MISA SME có hai cách nhận hóa đơn mua vào, tool chuẩn bị sẵn cho cả hai:

1. Đọc trực tiếp file XML hóa đơn điện tử (Mua hàng > Lập chứng từ từ hóa đơn điện tử).
   Tool gom các XML mua vào hợp lệ vào một thư mục phẳng ``misa/xml_mua_vao`` để chọn
   nhiều file một lần.
2. Nhập từ Excel (Mua hàng > Thêm chứng từ mua hàng > Nhập từ excel). Bước "Ghép dữ liệu"
   của MISA cho chọn cột tương ứng; tên cột ở đây đặt theo tên trường trên màn hình
   chứng từ mua hàng của MISA để ghép tự động được nhiều nhất. Cách này dùng được cả
   cho hóa đơn không có XML gốc (dữ liệu chi tiết lấy từ cổng thuế).

Kèm theo là danh mục Nhà cung cấp và Vật tư hàng hóa để nhập trước, vì chứng từ cần
mã NCC và mã hàng đã có trong danh mục.

Hóa đơn bị hủy (tthai 6) và bị thay thế (tthai 4) không đưa vào; liệt kê ở sheet BoQua.
"""

from __future__ import annotations

import hashlib
import json
import logging
import re
import shutil
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from .client import TTHAI_LABELS
from .detail import parsed_from_detail
from .xmlparse import ParsedInvoice, parse_invoice_file

log = logging.getLogger("hddt")

HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
REQUIRED_FILL = PatternFill("solid", fgColor="B71C1C")
HEADER_FONT = Font(bold=True, color="FFFFFF")

# Trạng thái không hạch toán: 4 = bị thay thế, 6 = bị hủy.
SKIP_TTHAI = {4: "Hóa đơn đã bị thay thế (hạch toán theo hóa đơn thay thế)", 6: "Hóa đơn đã bị hủy"}

VOUCHER_COLUMNS: list[tuple[str, str, bool]] = [
    # (tên cột, khóa, bắt buộc)
    ("Ngày hạch toán", "ngay_ht", True),
    ("Ngày chứng từ", "ngay_ct", True),
    ("Số chứng từ", "so_ct", True),
    ("Mẫu số HĐ", "mau_so", False),
    ("Ký hiệu HĐ", "ky_hieu", False),
    ("Số hóa đơn", "so_hd", False),
    ("Ngày hóa đơn", "ngay_hd", False),
    ("Mã nhà cung cấp", "ma_ncc", True),
    ("Tên nhà cung cấp", "ten_ncc", False),
    ("Địa chỉ", "dia_chi", False),
    ("Mã số thuế", "mst", False),
    ("Diễn giải", "dien_giai", False),
    ("Mã hàng", "ma_hang", True),
    ("Tên hàng", "ten_hang", False),
    ("TK kho/TK chi phí", "tk_no", True),
    ("TK công nợ", "tk_co", True),
    ("ĐVT", "dvt", False),
    ("Số lượng", "so_luong", False),
    ("Đơn giá", "don_gia", False),
    ("Thành tiền", "thanh_tien", True),
    ("Tỷ lệ CK (%)", "tl_ck", False),
    ("Tiền chiết khấu", "tien_ck", False),
    ("% thuế GTGT", "thue_suat", False),
    ("Tiền thuế GTGT", "tien_thue", False),
    ("TK thuế GTGT", "tk_thue", False),
    ("Nhóm HHDV mua vào", "nhom_hhdv", False),
    ("Ghi chú (tool)", "ghi_chu", False),
]

SUPPLIER_COLUMNS: list[tuple[str, str, bool]] = [
    ("Mã nhà cung cấp", "ma", True),
    ("Tên nhà cung cấp", "ten", True),
    ("Địa chỉ", "dia_chi", False),
    ("Mã số thuế", "mst", False),
    ("Điện thoại", "dien_thoai", False),
    ("Tài khoản ngân hàng", "stk", False),
    ("Tổ chức/Cá nhân", "loai", False),
]

ITEM_COLUMNS: list[tuple[str, str, bool]] = [
    ("Mã", "ma", True),
    ("Tên", "ten", True),
    ("Tính chất", "tinh_chat", True),
    ("Đơn vị tính chính", "dvt", False),
    ("Thuế suất GTGT (%)", "thue_suat", False),
    ("TK kho", "tk_kho", False),
    ("TK chi phí", "tk_chi_phi", False),
    ("Mã hàng trên hóa đơn", "ma_goc", False),
]


@dataclass
class MisaSettings:
    tk_chi_phi: str = "642"
    tk_cong_no: str = "331"
    tk_thue: str = "1331"
    nhom_hhdv: str = "1"  # 1 = HHDV dùng riêng cho SXKD chịu thuế GTGT
    so_ct_prefix: str = "MH"


@dataclass
class MisaResult:
    folder: str = ""
    excel_path: str = ""
    xml_dir: str = ""
    vouchers: int = 0
    lines: int = 0
    xml_files: int = 0
    no_xml: int = 0
    skipped: int = 0
    suppliers: int = 0
    items: int = 0
    warnings: list[str] = field(default_factory=list)


def _ascii_slug(text: str) -> str:
    t = unicodedata.normalize("NFD", text or "")
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    t = t.replace("đ", "d").replace("Đ", "D")
    t = re.sub(r"[^A-Za-z0-9]+", "_", t).strip("_").upper()
    return t


def _norm_name(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").strip()).lower()


class ItemCatalog:
    """Gán mã hàng ổn định: dùng mã trên hóa đơn nếu có, không thì tạo từ tên hàng."""

    def __init__(self) -> None:
        self.by_key: dict[tuple[str, str, str], dict[str, Any]] = {}
        self.used: dict[str, tuple[str, str, str]] = {}

    def code_for(self, ma_goc: str, ten: str, dvt: str, thue_suat: str, s: MisaSettings) -> str:
        # Khóa theo tên + ĐVT: có người bán dùng một mã (vd SP.0000) cho nhiều mặt hàng.
        key = (_norm_name(ten), _norm_name(dvt))
        if key in self.by_key:
            return self.by_key[key]["ma"]
        goc = (ma_goc or "").strip()
        base = _ascii_slug(goc)[:20] if goc else (_ascii_slug(ten)[:18] or "HH")
        code = base
        if code in self.used and self.used[code] != key:
            h = hashlib.md5("|".join(key).encode("utf-8")).hexdigest()[:4].upper()
            code = f"{base[:20]}_{h}"
        self.used[code] = key
        name = _norm_name(ten)
        dich_vu = any(w in name for w in ("cước", "dịch vụ", "phí ", "phí", "thuê", "vận chuyển", "sửa chữa", "bảo trì"))
        self.by_key[key] = {
            "ma": code,
            "ten": ten,
            "tinh_chat": "Dịch vụ" if dich_vu else "Vật tư hàng hóa",
            "dvt": dvt,
            "thue_suat": thue_suat,
            "tk_kho": "" if dich_vu else "156",
            "tk_chi_phi": s.tk_chi_phi,
            "ma_goc": goc,
        }
        return code


def _load_raw_index(out_dir: Path) -> dict[str, dict[str, Any]]:
    """Ghép các file danh_sach_*.json: khóa = file_stem của hóa đơn."""
    from .client import InvoiceRef

    index: dict[str, dict[str, Any]] = {}
    for f in sorted(out_dir.glob("danh_sach_*.json")):
        try:
            items = json.loads(f.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        for item in items if isinstance(items, list) else []:
            if not isinstance(item, dict):
                continue
            for fam in ("query", "sco-query"):
                ref = InvoiceRef.from_json("purchase", fam, item)
                index[ref.file_stem] = item
    return index


def _collect(out_dir: Path, raw_index: dict[str, dict[str, Any]]) -> list[tuple[str, Path, ParsedInvoice, bool]]:
    """Trả về [(stem, file nguồn, ParsedInvoice, có XML gốc)] cho hóa đơn mua vào."""
    src = out_dir / "xml" / "purchase"
    found: dict[str, tuple[str, Path, ParsedInvoice, bool]] = {}
    if not src.is_dir():
        return []
    for xml in sorted(src.glob("*.xml")):
        stem = xml.stem
        # Gói có nhiều XML được lưu là <stem>_1.xml, <stem>_2.xml: lấy file đầu tiên.
        base = re.sub(r"_\d+$", "", stem)
        if stem not in raw_index and base in raw_index:
            stem = base
        if stem in found:
            continue
        try:
            found[stem] = (stem, xml, parse_invoice_file(str(xml)), True)
        except Exception as exc:  # noqa: BLE001
            log.warning("[MISA] Bỏ qua %s: đọc XML lỗi: %s", xml.name, exc)
    for js in sorted(src.glob("*.json")):
        if js.stem in found:
            continue
        try:
            found[js.stem] = (js.stem, js, parsed_from_detail(json.loads(js.read_text(encoding="utf-8"))), False)
        except Exception as exc:  # noqa: BLE001
            log.warning("[MISA] Bỏ qua %s: đọc chi tiết lỗi: %s", js.name, exc)
    return list(found.values())


def _write(ws, columns: list[tuple[str, str, bool]], rows: list[dict[str, Any]]) -> None:
    ws.append([c[0] for c in columns])
    for i, (_, _, req) in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=i)
        cell.fill = REQUIRED_FILL if req else HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append([r.get(c[1]) for c in columns])
    ws.freeze_panes = "A2"
    for i, (name, key, _) in enumerate(columns, start=1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = 40 if key in ("ten_ncc", "ten_hang", "dia_chi", "dien_giai", "ten", "ghi_chu") else max(11, len(name) + 3)
        if key.startswith("ngay"):
            for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                c.number_format = "dd/mm/yyyy"
        if key in ("don_gia", "thanh_tien", "tien_thue", "tien_ck"):
            for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                c.number_format = "#,##0.##"
        if key in ("ma_ncc", "mst", "ma", "so_ct", "so_hd", "mau_so", "tk_no", "tk_co", "tk_thue", "ma_hang"):
            for (c,) in ws.iter_rows(min_row=2, min_col=i, max_col=i):
                c.number_format = "@"


def export_misa(output_dir: str | Path, settings: MisaSettings | None = None) -> MisaResult:
    s = settings or MisaSettings()
    out_dir = Path(output_dir)
    misa_dir = out_dir / "misa"
    xml_out = misa_dir / "xml_mua_vao"
    if xml_out.exists():
        shutil.rmtree(xml_out)
    xml_out.mkdir(parents=True, exist_ok=True)
    res = MisaResult(folder=str(misa_dir), xml_dir=str(xml_out))

    raw_index = _load_raw_index(out_dir)
    invoices = _collect(out_dir, raw_index)
    if not invoices:
        raise FileNotFoundError(
            f"Không thấy hóa đơn mua vào trong {out_dir / 'xml' / 'purchase'}. Hãy kéo hóa đơn (có tải XML) trước."
        )
    invoices.sort(key=lambda t: (t[2].ngay_lap or datetime.min, t[2].nb_mst, t[2].khhdon, t[2].shdon))

    catalog = ItemCatalog()
    suppliers: dict[str, dict[str, Any]] = {}
    vouchers: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    used_so_ct: set[str] = set()

    for stem, path, inv, has_xml in invoices:
        raw = raw_index.get(stem, {})
        try:
            tthai = int(raw.get("tthai")) if raw.get("tthai") not in (None, "") else 0
        except (TypeError, ValueError):
            tthai = 0
        label = f"{inv.khmshdon}{inv.khhdon} số {inv.shdon} - {inv.nb_ten}"
        if tthai in SKIP_TTHAI:
            skipped.append({"hd": label, "mst": inv.nb_mst, "ly_do": SKIP_TTHAI[tthai], "file": path.name})
            continue

        ma_ncc = inv.nb_mst or _ascii_slug(inv.nb_ten)[:20] or "NCC"
        suppliers.setdefault(ma_ncc, {
            "ma": ma_ncc, "ten": inv.nb_ten, "dia_chi": inv.nb_dchi, "mst": inv.nb_mst,
            "dien_thoai": inv.nb_sdt, "stk": inv.nb_stk, "loai": "Tổ chức",
        })

        so_ct = f"{s.so_ct_prefix}{_ascii_slug(inv.khhdon)}{inv.shdon}"[:20]
        n = 1
        while so_ct in used_so_ct:
            n += 1
            so_ct = f"{s.so_ct_prefix}{_ascii_slug(inv.khhdon)}{inv.shdon}"[:18] + f"_{n}"
        used_so_ct.add(so_ct)

        note_parts = []
        if not has_xml:
            note_parts.append("Không có XML gốc; số liệu lấy từ chi tiết trên cổng thuế")
            res.no_xml += 1
        if tthai in (2, 3, 5):
            note_parts.append(TTHAI_LABELS.get(tthai, "") + " - kiểm tra lại trước khi hạch toán")
        note = "; ".join(p for p in note_parts if p)

        header = {
            "ngay_ht": inv.ngay_lap, "ngay_ct": inv.ngay_lap, "so_ct": so_ct,
            "mau_so": inv.khmshdon, "ky_hieu": inv.khhdon, "so_hd": inv.shdon, "ngay_hd": inv.ngay_lap,
            "ma_ncc": ma_ncc, "ten_ncc": inv.nb_ten, "dia_chi": inv.nb_dchi, "mst": inv.nb_mst,
            "dien_giai": f"Mua hàng của {inv.nb_ten} theo HĐ {inv.khhdon} số {inv.shdon}",
            "tk_no": s.tk_chi_phi, "tk_co": s.tk_cong_no, "tk_thue": s.tk_thue, "nhom_hhdv": s.nhom_hhdv,
            "ghi_chu": note,
        }

        lines = [ln for ln in inv.lines if ln.tchat != "4" and (ln.ten_hang or ln.thanh_tien)]
        if not lines:
            ma = catalog.code_for("", "Hàng hóa, dịch vụ theo hóa đơn", "", "", s)
            vouchers.append({
                **header, "ma_hang": ma, "ten_hang": "Hàng hóa, dịch vụ theo hóa đơn",
                "thanh_tien": inv.tong_tien_chua_thue, "tien_thue": inv.tong_tien_thue,
                "thue_suat": (inv.thue_theo_suat[0]["thue_suat"] if len(inv.thue_theo_suat) == 1 else ""),
                "ghi_chu": "; ".join(p for p in (note, "Hóa đơn không có dòng hàng chi tiết; ghi theo tổng tiền") if p),
            })
            res.lines += 1
        for ln in lines:
            ma = catalog.code_for(ln.ma_hang, ln.ten_hang, ln.dvt, ln.thue_suat, s)
            sign = -1 if ln.tchat == "3" else 1  # chiết khấu thương mại ghi giảm
            thanh_tien = None if ln.thanh_tien is None else sign * ln.thanh_tien
            tien_thue = None if ln.tien_thue is None else sign * ln.tien_thue
            line_note = note
            if ln.tchat == "2":
                line_note = "; ".join(p for p in (note, "Hàng khuyến mại") if p)
            elif ln.tchat == "3":
                line_note = "; ".join(p for p in (note, "Chiết khấu thương mại (ghi âm)") if p)
            vouchers.append({
                **header, "ma_hang": ma, "ten_hang": ln.ten_hang, "dvt": ln.dvt,
                "so_luong": ln.so_luong, "don_gia": ln.don_gia, "thanh_tien": thanh_tien,
                "tl_ck": ln.tl_ck, "tien_ck": ln.st_ck, "thue_suat": ln.thue_suat, "tien_thue": tien_thue,
                "ghi_chu": line_note,
            })
            res.lines += 1

        if has_xml:
            # Tên file: <MST người bán>_<ký hiệu>_<số>, lấy từ tên file gốc (purchase_<nguồn>_<MST>_<mẫu>_<ký hiệu>_<số>).
            parts = stem.split("_")
            nice = "_".join([parts[2], parts[4], "_".join(parts[5:])]) if len(parts) >= 6 else stem
            dest = xml_out / f"{nice}.xml"
            shutil.copyfile(path, dest)
            res.xml_files += 1
        res.vouchers += 1

    wb = Workbook()
    ws = wb.active
    ws.title = "ChungTuMuaHang"
    _write(ws, VOUCHER_COLUMNS, vouchers)
    _write(wb.create_sheet("DanhMuc_NhaCungCap"), SUPPLIER_COLUMNS, list(suppliers.values()))
    _write(wb.create_sheet("DanhMuc_VatTuHangHoa"), ITEM_COLUMNS, [v for v in catalog.by_key.values()])
    skip_ws = wb.create_sheet("BoQua")
    _write(skip_ws, [("Hóa đơn", "hd", False), ("MST người bán", "mst", False), ("Lý do không đưa vào", "ly_do", False), ("File", "file", False)], skipped)
    guide = wb.create_sheet("HuongDan")
    for line in HUONG_DAN.strip().splitlines():
        guide.append([line])
    guide.column_dimensions["A"].width = 120

    res.excel_path = str(misa_dir / "MISA_NhapKhau_MuaHang.xlsx")
    wb.save(res.excel_path)
    res.skipped = len(skipped)
    res.suppliers = len(suppliers)
    res.items = len(catalog.by_key)
    (misa_dir / "HUONG_DAN_NHAP_MISA.txt").write_text(HUONG_DAN.strip() + "\n", encoding="utf-8")
    log.info(
        "[MISA] %d chứng từ (%d dòng), %d XML gom cho MISA, %d hóa đơn không có XML, %d bỏ qua; %d NCC, %d mã hàng -> %s",
        res.vouchers, res.lines, res.xml_files, res.no_xml, res.skipped, res.suppliers, res.items, res.excel_path,
    )
    return res


HUONG_DAN = """
HƯỚNG DẪN ĐƯA HÓA ĐƠN MUA VÀO VÀO MISA SME
==========================================

Thư mục misa gồm:
- xml_mua_vao\\              : file XML gốc của các hóa đơn mua vào (đã bỏ hóa đơn bị hủy / bị thay thế).
- MISA_NhapKhau_MuaHang.xlsx : sheet ChungTuMuaHang (mỗi dòng hàng một dòng, các dòng cùng "Số chứng từ" là một chứng từ),
                               DanhMuc_NhaCungCap, DanhMuc_VatTuHangHoa, BoQua (hóa đơn không đưa vào và lý do).

CÁCH 1 - MISA ĐỌC TRỰC TIẾP FILE XML (khuyên dùng cho hóa đơn có XML)
1. Mở MISA SME > phân hệ Mua hàng > chức năng lập chứng từ từ hóa đơn điện tử (nhập khẩu hóa đơn điện tử).
2. Chọn nhập từ tệp hóa đơn điện tử, trỏ tới thư mục misa\\xml_mua_vao, chọn tất cả file .xml.
3. MISA tự đọc thông tin người bán, dòng hàng, thuế; chị chọn loại chứng từ (mua hàng hóa / mua dịch vụ),
   tài khoản hạch toán rồi lưu. Nhà cung cấp, mã hàng chưa có MISA sẽ gợi ý thêm mới.

CÁCH 2 - NHẬP TỪ EXCEL (dùng cho mọi hóa đơn, kể cả hóa đơn không có XML như Viettel, VNPT, ngân hàng)
1. Mở file MISA_NhapKhau_MuaHang.xlsx, kiểm tra cột "TK kho/TK chi phí" (mặc định 642), "TK công nợ" (331),
   "TK thuế GTGT" (1331). Sửa theo chế độ kế toán của đơn vị (vd 6422 với TT200, 156 với hàng hóa nhập kho).
   Cột tiêu đề màu đỏ là cột MISA bắt buộc.
2. Nhập danh mục trước: trong MISA vào Tệp > Nhập khẩu từ Excel, chọn Danh mục Nhà cung cấp, chọn file này,
   sheet DanhMuc_NhaCungCap. Làm tương tự với Vật tư hàng hóa / Dịch vụ, sheet DanhMuc_VatTuHangHoa.
3. Nhập chứng từ: phân hệ Mua hàng > tab Mua hàng hóa, dịch vụ > bấm mũi tên cạnh "Thêm chứng từ mua hàng"
   > Nhập từ excel. Chọn file này, sheet ChungTuMuaHang.
4. Ở bước "Ghép dữ liệu", kiểm tra mỗi thông tin của MISA đã ghép đúng cột trong file (tên cột trong file đặt
   giống tên trên màn hình chứng từ MISA nên phần lớn tự ghép). Cột nào chưa ghép thì chọn tay.
5. Bấm Thực hiện, xem kết quả kiểm tra, sửa các dòng MISA báo lỗi rồi nhập lại.

LƯU Ý
- Không dùng cả hai cách cho cùng một hóa đơn (sẽ bị nhập trùng). Gợi ý: Cách 1 cho hóa đơn có XML,
  Cách 2 lọc cột "Ghi chú (tool)" chứa "Không có XML gốc" cho phần còn lại.
- Số chứng từ được tạo dạng MH + ký hiệu + số hóa đơn để không trùng; có thể sửa theo quy tắc của đơn vị.
- Hóa đơn thay thế / điều chỉnh / đã bị điều chỉnh có ghi chú "kiểm tra lại trước khi hạch toán".
- Chiết khấu thương mại ghi âm; hàng khuyến mại giữ nguyên số liệu trên hóa đơn.
"""
