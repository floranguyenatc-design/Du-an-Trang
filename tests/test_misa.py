import json
import shutil
from pathlib import Path

from openpyxl import load_workbook

from hddt.misa import MisaSettings, export_misa

FIX = Path(__file__).parent / "fixtures" / "hoadon_mau.xml"


def _setup(tmp_path: Path) -> Path:
    out = tmp_path / "out"
    pur = out / "xml" / "purchase"
    pur.mkdir(parents=True)
    items = [
        # có XML, hợp lệ
        {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "125", "tthai": 1, "ttxly": 5},
        # không có XML, chi tiết từ GDT
        {"nbmst": "0100109106", "khmshdon": "1", "khhdon": "K24DAA", "shdon": "6928336", "tthai": 1, "ttxly": 6},
        # bị hủy -> bỏ qua
        {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "130", "tthai": 6, "ttxly": 5},
        # gói có 2 XML -> lưu _1, _2
        {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "140", "tthai": 3, "ttxly": 5},
    ]
    (out / "danh_sach_20251201_20251231.json").write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
    shutil.copy(FIX, pur / "purchase_query_0312345678_1_C25TAA_125.xml")
    shutil.copy(FIX, pur / "purchase_query_0312345678_1_C25TAA_130.xml")
    xml140 = FIX.read_text(encoding="utf-8").replace("<SHDon>125</SHDon>", "<SHDon>140</SHDon>").replace(
        "<TChat>1</TChat><STT>2</STT>", "<TChat>3</TChat><STT>2</STT>"
    )
    (pur / "purchase_query_0312345678_1_C25TAA_140_1.xml").write_text(xml140, encoding="utf-8")
    (pur / "purchase_query_0312345678_1_C25TAA_140_2.xml").write_text(xml140, encoding="utf-8")
    detail = {
        "nbmst": "0100109106", "nbten": "Tập đoàn Công nghiệp - Viễn thông Quân đội", "khmshdon": 1, "khhdon": "K24DAA",
        "shdon": 6928336, "tdlap": "2024-04-02T00:00:00", "tgtcthue": 435729, "tgtthue": 43573, "tgtttbso": 479302,
        "hdhhdvu": [{"stt": 1, "ten": "Cước dịch vụ viễn thông", "dvtinh": "Tháng", "sluong": 1, "dgia": 435729,
                     "thtien": 435729, "tsuat": 0.1, "mhhdvu": "SP.0000"}],
    }
    (pur / "purchase_query_0100109106_1_K24DAA_6928336.json").write_text(json.dumps(detail, ensure_ascii=False), encoding="utf-8")
    return out


def test_export_misa(tmp_path):
    from datetime import datetime

    from hddt.misa import MISA_COLUMNS

    out = _setup(tmp_path)
    res = export_misa(out, MisaSettings(tk_chi_phi="6422"))
    assert res.vouchers == 3 and res.skipped == 1 and res.no_xml == 1 and res.xml_files == 2
    assert Path(res.excel_path).name == "Mua_hang_khong_qua_kho_VND.xlsx"

    xmls = sorted(p.name for p in Path(res.xml_dir).glob("*.xml"))
    assert xmls == ["0312345678_C25TAA_125.xml", "0312345678_C25TAA_140.xml"]

    # File chứng từ: đúng mẫu MISA (1 sheet, tiêu đề 39 cột, ghi chú gợi ý của MISA vẫn còn).
    wb = load_workbook(res.excel_path)
    assert wb.sheetnames == ["Chứng từ mua hàng không qua kho"]
    ws = wb.active
    assert [ws[f"{c}1"].value for c, _ in MISA_COLUMNS] == [h for _, h in MISA_COLUMNS]
    prompts = {str(dv.sqref).split(":")[0]: dv.prompt for dv in ws.data_validations.dataValidation}
    assert "Nhập 10: 10%" in prompts["AA1"]

    header = [ws[f"{c}1"].value for c, _ in MISA_COLUMNS]
    rows = [dict(zip(header, r[: len(header)])) for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
    # 125: 2 dòng hàng (bỏ dòng ghi chú TChat=4); 140: 2 dòng; Viettel: 1 dòng
    assert len(rows) == 5
    so_ct = {r["Số hóa đơn"]: r["Số chứng từ (*)"] for r in rows}
    assert len(set(so_ct.values())) == 3 and all(len(v) <= 20 for v in so_ct.values())
    for r in rows:
        assert r["Hình thức mua hàng"] == "0" and r["Phương thức thanh toán"] == "0" and r["Nhận kèm hóa đơn"] == "1"
        assert r["TK chi phí (*)"] == "6422" and r["TK công nợ/TK tiền (*)"] == "331"
        assert isinstance(r["Ngày hạch toán (*)"], datetime) and r["Ngày hạch toán (*)"] == r["Ngày hóa đơn"]
        assert r["Mã hàng (*)"] and r["Mã nhà cung cấp"]

    viettel = [r for r in rows if r["Mã nhà cung cấp"] == "0100109106"][0]
    assert viettel["Tên hàng"] == "Cước dịch vụ viễn thông" and viettel["Đơn giá"] == 435729 and viettel["Số lượng"] == 1
    assert viettel["% thuế GTGT"] == "10" and viettel["Tiền thuế GTGT"] == 43573
    assert viettel["TK thuế GTGT"] == "1331" and viettel["Nhóm HHDV mua vào"] == "1"

    hd125 = [r for r in rows if r["Số hóa đơn"] == "125"]
    assert [(r["Tên hàng"], r["ĐVT"], r["Số lượng"], r["Đơn giá"], r["Thành tiền"], r["% thuế GTGT"], r["Tiền thuế GTGT"]) for r in hd125] == [
        ("Dịch vụ tư vấn", "Gói", 2, 1000000, 2000000, "10", 200000),
        ("Văn phòng phẩm", "Hộp", 10, 50000, 500000, "8", 40000),
    ]

    ck = [r for r in rows if r["Số hóa đơn"] == "140" and r["Tên hàng"] == "Văn phòng phẩm"][0]
    assert ck["Thành tiền"] == -500000 and ck["Tiền thuế GTGT"] == -40000

    # File danh mục + kiểm tra
    wb2 = load_workbook(res.catalog_path)
    assert wb2.sheetnames == ["KiemTra", "DanhMuc_NhaCungCap", "DanhMuc_VatTuHangHoa", "BoQua", "HuongDan"]
    checks = {r[0]: r for r in wb2["KiemTra"].iter_rows(min_row=2, values_only=True)}
    chk140 = [r for r in checks.values() if "số 140" in r[1]][0]
    assert "chiết khấu" in chk140[12] and "kiểm tra lại" in chk140[12] and "lệch" in chk140[12]
    chk125 = [r for r in checks.values() if "số 125" in r[1]][0]
    assert chk125[6] == 0 and chk125[9] == 0 and chk125[11] == "Có"
    assert res.warnings and "140" in res.warnings[0]

    ncc = list(wb2["DanhMuc_NhaCungCap"].iter_rows(min_row=2, values_only=True))
    assert sorted(r[0] for r in ncc) == ["0100109106", "0312345678"]
    items = {r[1]: r for r in wb2["DanhMuc_VatTuHangHoa"].iter_rows(min_row=2, values_only=True)}
    assert items["Cước dịch vụ viễn thông"][2] == "Dịch vụ" and items["Cước dịch vụ viễn thông"][0] == "SP_0000"
    assert items["Văn phòng phẩm"][2] == "Vật tư hàng hóa"
    codes = [r[0] for r in items.values()]
    assert len(codes) == len(set(codes))
    # mã hàng trên chứng từ đều có trong danh mục
    assert {r["Mã hàng (*)"] for r in rows} <= set(codes)

    skipped = list(wb2["BoQua"].iter_rows(min_row=2, values_only=True))
    assert len(skipped) == 1 and "hủy" in skipped[0][2]
    assert (Path(res.folder) / "HUONG_DAN_NHAP_MISA.txt").is_file()


def test_payment_method_changes_counter_account(tmp_path):
    out = _setup(tmp_path)
    res = export_misa(out, MisaSettings(phuong_thuc_tt="1"))
    ws = load_workbook(res.excel_path).active
    assert ws["C2"].value == "1" and ws["S2"].value == "1111"


def test_misa_vat_rate():
    from hddt.misa import misa_vat_rate

    assert misa_vat_rate("10%") == ("10", None)
    assert misa_vat_rate("8%") == ("8", None)
    assert misa_vat_rate("0%") == ("0", None)
    assert misa_vat_rate("5") == ("5", None)
    assert misa_vat_rate("0.1") == ("10", None)
    assert misa_vat_rate("KCT") == ("KCT", None)
    assert misa_vat_rate("kkknt") == ("KKKNT", None)
    assert misa_vat_rate("KHAC:3.5%") == ("KHAC", 3.5)
    assert misa_vat_rate("3,5%") == ("KHAC", 3.5)
    assert misa_vat_rate("") == ("", None)


def test_same_seller_code_for_different_items_gets_unique_codes():
    from hddt.misa import ItemCatalog

    cat = ItemCatalog()
    s = MisaSettings()
    a = cat.code_for("SP.0000", "Xăng RON 95-III", "Lít", "10%", s)
    b = cat.code_for("SP.0000", "Dầu DO 0,05S", "Lít", "10%", s)
    c = cat.code_for("SP.0000", "Xăng RON 95-III", "lít", "10%", s)
    assert a == "SP_0000" and b != a and c == a


def test_export_misa_cli(tmp_path, capsys):
    from hddt import cli

    out = _setup(tmp_path)
    assert cli.main(["--env", "x", "misa", "--thu-muc", str(out)]) == 0
    assert "3 chứng từ" in capsys.readouterr().out


def test_export_misa_empty_dir(tmp_path):
    from hddt import cli

    assert cli.main(["--env", "x", "misa", "--thu-muc", str(tmp_path)]) == 1
