import json
import shutil
from pathlib import Path

import pytest
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
    (out / "danh_sach_muavao_20251201_20251231.json").write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")
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


def res_template():
    from hddt.misa import resolve_template

    return resolve_template("2020")


def test_export_misa(tmp_path):
    from datetime import datetime

    from hddt.misa import template_info

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
    header = [c.value for c in ws[1] if c.value]
    assert len(header) == 38 and "Tỷ lệ tính thuế (Thuế suất KHAC)" not in header  # mẫu MISA SME 2020 R3
    assert header == list(template_info(res_template())[0])
    prompts = {str(dv.sqref).split(":")[0]: dv.prompt for dv in ws.data_validations.dataValidation}
    assert "Nhập 10: 10%" in prompts["AA1"]
    rows = [dict(zip(header, r[: len(header)])) for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
    # 125: 2 dòng hàng (bỏ dòng ghi chú TChat=4); 140: 2 dòng; Viettel: 1 dòng
    assert len(rows) == 5
    so_ct = {r["Số hóa đơn"]: r["Số chứng từ (*)"] for r in rows}
    assert len(set(so_ct.values())) == 3 and all(len(v) <= 20 for v in so_ct.values())
    for r in rows:
        assert r["Hình thức mua hàng"] == 0 and r["Phương thức thanh toán"] == 1 and r["Nhận kèm hóa đơn"] == 1
        assert r["TK chi phí (*)"] == "6422" and r["TK công nợ/TK tiền (*)"] == "1111"
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
    assert "chiết khấu" in chk140[13] and "kiểm tra lại" in chk140[13] and "lệch" in chk140[13]
    chk125 = [r for r in checks.values() if "số 125" in r[1]][0]
    assert chk125[7] == 0 and chk125[10] == 0 and chk125[12] == "Có"
    assert any("số 140" in w and "lệch" in w for w in res.warnings)

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


def test_payment_method_fixed_modes(tmp_path):
    out = _setup(tmp_path)
    ws = load_workbook(export_misa(out, MisaSettings(phuong_thuc_tt="0")).excel_path).active
    assert ws["C2"].value == 0 and ws["S2"].value == "331"
    ws = load_workbook(export_misa(out, MisaSettings(phuong_thuc_tt="2")).excel_path).active
    assert ws["C2"].value == 2 and ws["S2"].value == "1121"
    ws = load_workbook(export_misa(out, MisaSettings(phuong_thuc_tt="4")).excel_path).active
    assert ws["C2"].value == 4 and ws["S2"].value == "1111"  # séc tiền mặt -> TK tiền mặt


def test_bank_invoices_use_1121_and_others_1111(tmp_path):
    from hddt.misa import is_bank_seller

    out = _setup(tmp_path)
    pur = out / "xml" / "purchase"
    banks = [
        ("NGAN HANG TMCP CONG THUONG VIET NAM - CN SAM SON", "0100111948-125", "K24TBB", 11474),  # không dấu
        ("Ngân hàng TMCP Sài Gòn Thương Tín - Hội sở", "0301103908", "K24TAB", 12686),
        ("Ngân hàng Vietinbank chi nhánh Thanh Hóa", "0100111948-130", "K24TBC", 5),
        ("Sacombank - Chi nhánh Sầm Sơn", "0300000001", "K24TAD", 6),
    ]
    items = json.loads(next(out.glob("danh_sach_*.json")).read_text(encoding="utf-8"))
    for ten, mst, kh, so in banks:
        detail = {"nbmst": mst, "nbten": ten, "khmshdon": 1, "khhdon": kh, "shdon": so, "tdlap": "2024-05-29T00:00:00",
                  "tgtcthue": 35000, "tgtthue": 0,
                  "hdhhdvu": [{"stt": 1, "ten": "Phí dịch vụ ngân hàng", "thtien": 35000, "ltsuat": "KCT"}]}
        (pur / f"purchase_query_{mst}_1_{kh}_{so}.json").write_text(json.dumps(detail, ensure_ascii=False), encoding="utf-8")
        items.append({"nbmst": mst, "khmshdon": "1", "khhdon": kh, "shdon": str(so), "tthai": 1, "ttxly": 5})
    next(out.glob("danh_sach_*.json")).write_text(json.dumps(items, ensure_ascii=False), encoding="utf-8")

    res = export_misa(out)
    ws = load_workbook(res.excel_path).active
    header = [c.value for c in ws[1]]
    rows = [dict(zip(header, r)) for r in ws.iter_rows(min_row=2, values_only=True) if any(r)]
    by_seller = {r["Tên nhà cung cấp"]: r for r in rows}
    for ten, *_ in banks:
        r = by_seller[ten]
        assert r["TK công nợ/TK tiền (*)"] == "1121" and r["Phương thức thanh toán"] == 2, ten
    for r in rows:
        if r["Tên nhà cung cấp"] not in {b[0] for b in banks}:
            assert r["TK công nợ/TK tiền (*)"] == "1111" and r["Phương thức thanh toán"] == 1
    # sheet KiemTra nêu rõ từng hóa đơn đã xếp vào loại nào
    wb2 = load_workbook(res.catalog_path)
    labels = [r[4] for r in wb2["KiemTra"].iter_rows(min_row=2, values_only=True)]
    assert sum("TK 1121" in l for l in labels) == 4 and sum("TK 1111" in l for l in labels) == 3

    assert is_bank_seller("NGÂN HÀNG TMCP ĐẦU TƯ VÀ PHÁT TRIỂN") and is_bank_seller("MB Bank") and is_bank_seller("Vietcombank")
    assert not is_bank_seller("CÔNG TY TNHH ABC") and not is_bank_seller("Công ty Điện lực Thanh Hóa")


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


@pytest.mark.parametrize("mau", ["2020", "moi"])
def test_output_matches_template_layout(tmp_path, mau):
    """Tiêu đề, độ rộng cột, font của file xuất khớp từng mẫu MISA đi kèm."""
    from hddt.misa import resolve_template, template_info

    res = export_misa(_setup(tmp_path), MisaSettings(mau=mau))
    ws = load_workbook(res.excel_path).active
    tpl = load_workbook(resolve_template(mau)).active
    heads_out = [(c.column_letter, c.value) for c in ws[1] if c.value]
    heads_tpl = [(c.column_letter, c.value) for c in tpl[1] if c.value]
    assert heads_out == heads_tpl
    assert len(heads_out) == (38 if mau == "2020" else 39)
    for letter, _ in heads_out:
        assert abs((ws.column_dimensions[letter].width or 0) - (tpl.column_dimensions[letter].width or 0)) < 0.01, letter
        assert ws[f"{letter}1"].font.name == "Times New Roman" and ws[f"{letter}1"].font.b
    assert len(ws.data_validations.dataValidation) == len(tpl.data_validations.dataValidation)
    assert template_info(resolve_template(mau))[1] == ({"0", "5", "10"} if mau == "2020" else {"0", "5", "8", "10", "KCT", "KKKNT", "KHAC"})


def test_columns_follow_header_names_in_2020_template(tmp_path):
    """Mẫu 2020 R3 không có cột Tỷ lệ tính thuế nên Tiền thuế GTGT ở AB, TK thuế AD, Nhóm HHDV AE (mẫu 39 cột: AC, AE, AF)."""
    out = _setup(tmp_path)
    ws = load_workbook(export_misa(out).excel_path).active
    assert [ws[f"{c}1"].value for c in ("AA", "AB", "AC", "AD", "AE")] == [
        "% thuế GTGT", "Tiền thuế GTGT", "TKĐƯ thuế GTGT", "TK thuế GTGT", "Nhóm HHDV mua vào"]
    row = {ws[f"{c}1"].value: ws[f"{c}{r}"].value for r in range(2, 7) for c in ("AA", "AB", "AC", "AD", "AE", "J") if ws[f"J{r}"].value == "6928336"}
    assert row["% thuế GTGT"] == "10" and row["Tiền thuế GTGT"] == 43573 and row["TKĐƯ thuế GTGT"] is None
    assert row["TK thuế GTGT"] == "1331" and row["Nhóm HHDV mua vào"] == "1"

    ws39 = load_workbook(export_misa(out, MisaSettings(mau="moi")).excel_path).active
    assert [ws39[f"{c}1"].value for c in ("AA", "AB", "AC", "AD", "AE", "AF")] == [
        "% thuế GTGT", "Tỷ lệ tính thuế (Thuế suất KHAC)", "Tiền thuế GTGT", "TKĐƯ thuế GTGT", "TK thuế GTGT", "Nhóm HHDV mua vào"]
    vt = [r for r in range(2, 7) if ws39[f"J{r}"].value == "6928336"][0]
    assert ws39[f"AC{vt}"].value == 43573 and ws39[f"AE{vt}"].value == "1331" and ws39[f"AF{vt}"].value == "1"


def test_vat_rate_not_in_template_hint_is_flagged(tmp_path):
    """Mẫu 2020 R3 chỉ liệt kê thuế suất 0/5/10: hóa đơn có 8% vẫn ghi đúng nhưng bị đánh dấu ở KiemTra."""
    out = _setup(tmp_path)  # hóa đơn 125 có dòng thuế suất 8%
    res = export_misa(out)
    wb2 = load_workbook(res.catalog_path)
    notes = {r[1]: r[13] for r in wb2["KiemTra"].iter_rows(min_row=2, values_only=True)}
    n125 = [v for k, v in notes.items() if "số 125" in k][0]
    assert "Thuế suất 8 không có trong ghi chú của mẫu MISA" in n125 and "0, 5, 10" in n125
    assert any("mẫu MISA" in w for w in res.warnings)
    ws = load_workbook(res.excel_path).active
    assert "8" in {ws[f"AA{r}"].value for r in range(2, 8)}
    # mẫu 39 cột có liệt kê thuế suất 8 nên không đánh dấu
    res39 = export_misa(out, MisaSettings(mau="moi"))
    n125b = [r[13] for r in load_workbook(res39.catalog_path)["KiemTra"].iter_rows(min_row=2, values_only=True) if "số 125" in r[1]][0]
    assert "mẫu MISA" not in (n125b or "")


def test_template_choice_and_custom_path(tmp_path):
    import pytest as _pt

    from hddt.misa import BUNDLED_TEMPLATES, resolve_template

    assert resolve_template("2020") == BUNDLED_TEMPLATES["2020"] and resolve_template("MOI") == BUNDLED_TEMPLATES["moi"]
    custom = tmp_path / "mau_rieng.xlsx"
    custom.write_bytes(BUNDLED_TEMPLATES["2020"].read_bytes())
    assert resolve_template(str(custom)) == custom
    with _pt.raises(ValueError):
        resolve_template("khong-co")
    res = export_misa(_setup(tmp_path / "x"), MisaSettings(mau=str(custom)))
    assert Path(res.excel_path).is_file()
