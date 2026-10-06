from datetime import datetime
from pathlib import Path

from hddt.xmlparse import parse_date, parse_invoice_file, parse_number

FIX = Path(__file__).parent / "fixtures" / "hoadon_mau.xml"


def test_parse_sample_invoice():
    inv = parse_invoice_file(str(FIX))
    assert inv.khmshdon == "1"
    assert inv.khhdon == "C25TAA"
    assert inv.shdon == "125"
    assert inv.ngay_lap == datetime(2025, 12, 15)
    assert inv.nb_mst == "0312345678"
    assert inv.nm_mst == "0109876543"
    assert inv.nm_email == "kt@xyz.vn"
    assert inv.tong_tien_chua_thue == 2500000
    assert inv.tong_tien_thue == 240000
    assert inv.tong_tien_tt == 2740000
    assert inv.mccqt == "M1-25-ABCDE-00000000001"
    assert inv.ngay_ky_nb == datetime(2025, 12, 15, 10, 20, 30)
    assert inv.ngay_ky_cqt == datetime(2025, 12, 15, 10, 25, 0)
    assert inv.extra["MaTraCuu"] == "ABC123"
    assert [t["thue_suat"] for t in inv.thue_theo_suat] == ["10%", "8%"]


def test_parse_lines_including_ttkhac_tax():
    inv = parse_invoice_file(str(FIX))
    assert len(inv.lines) == 3
    l1, l2, l3 = inv.lines
    assert l1.ten_hang == "Dịch vụ tư vấn"
    assert l1.tien_thue == 200000  # lấy từ TTKhac/TThue
    assert l1.thanh_tien_co_thue == 2200000
    assert l2.tien_thue == 40000
    assert l2.so_luong == 10 and l2.don_gia == 50000
    assert l3.tchat == "4" and l3.thanh_tien is None


def test_parse_number_formats():
    assert parse_number("1,000,000") == 1000000
    assert parse_number("1.000.000,50") == 1000000.5
    assert parse_number("1,5") == 1.5
    assert parse_number("2500000.75") == 2500000.75
    assert parse_number("") is None
    assert parse_number("abc") is None


def test_parse_date_formats():
    assert parse_date("2025-12-31T00:00:00") == datetime(2025, 12, 31)
    assert parse_date("31/12/2025") == datetime(2025, 12, 31)
    assert parse_date("") is None


def test_parsed_from_detail_json():
    from hddt.detail import parsed_from_detail

    inv = parsed_from_detail({
        "nbmst": "0100109106", "nbten": "Viettel", "khmshdon": 1, "khhdon": "K24DAA", "shdon": 6928336,
        "tdlap": "2024-04-02T00:00:00", "tgtcthue": 435729, "tgtthue": 43573, "tgtttbso": 479302,
        "hdhhdvu": [
            {"stt": 1, "ten": "Cước di động", "dvtinh": "Tháng", "sluong": 1, "dgia": 435729, "thtien": 435729, "tsuat": 0.1},
            {"stt": 2, "thhdvu": "Khuyến mại", "tchat": 2, "thtien": 0, "ltsuat": "KCT",
             "ttkhac": [{"ttruong": "TThue", "dlieu": "0"}]},
        ],
    })
    assert inv.shdon == "6928336" and inv.nb_ten == "Viettel" and inv.tong_tien_tt == 479302
    l1, l2 = inv.lines
    assert l1.ten_hang == "Cước di động" and l1.dvt == "Tháng" and l1.so_luong == 1 and l1.don_gia == 435729
    assert l1.thue_suat == "10%" and l1.tien_thue == 43572.9
    assert l2.ten_hang == "Khuyến mại" and l2.thue_suat == "KCT" and l2.tien_thue == 0


def test_detail_tax_rate_formats():
    from hddt.detail import _tax_rate_text

    assert _tax_rate_text({"tsuat": 0.08}) == "8%"
    assert _tax_rate_text({"tsuat": 0.1}) == "10%"
    assert _tax_rate_text({"tsuat": 5}) == "5%"
    assert _tax_rate_text({"tsuat": 0.035}) == "3.5%"
    assert _tax_rate_text({"ltsuat": "KCT", "tsuat": 0}) == "KCT"
    assert _tax_rate_text({}) == ""


def test_render_pdf_from_sample_xml(tmp_path):
    from pypdf import PdfReader

    from hddt.pdfrender import render_invoice_pdf

    inv = parse_invoice_file(str(FIX))
    assert inv.ten_hd == "HÓA ĐƠN GIÁ TRỊ GIA TĂNG"
    out = render_invoice_pdf(inv, tmp_path / "a" / "hd.pdf", status_lines=["Trạng thái: Hóa đơn mới"])
    text = "".join(pg.extract_text() for pg in PdfReader(out).pages)
    for expected in ("HÓA ĐƠN GIÁ TRỊ GIA TĂNG", "C25TAA", "125", "CÔNG TY TNHH ABC", "0109876543",
                     "Dịch vụ tư vấn", "Văn phòng phẩm", "2.000.000", "2.740.000", "M1-25-ABCDE-00000000001",
                     "Hai triệu bảy trăm bốn mươi nghìn đồng", "Trạng thái: Hóa đơn mới"):
        assert expected in text, expected


def test_render_pdf_many_lines_spans_pages(tmp_path):
    from pypdf import PdfReader

    from hddt.pdfrender import render_invoice_pdf
    from hddt.xmlparse import InvoiceLine, ParsedInvoice

    inv = ParsedInvoice(khmshdon="1", khhdon="K24TAB", shdon="1", nb_ten="A", nm_ten="B")
    inv.lines = [InvoiceLine(stt=str(i), ten_hang=f"Mặt hàng số {i} " * 3, dvt="Cái", so_luong=i, don_gia=1000.5, thanh_tien=1000.5 * i) for i in range(1, 121)]
    out = render_invoice_pdf(inv, tmp_path / "long.pdf")
    reader = PdfReader(out)
    assert len(reader.pages) >= 3
    assert "Mặt hàng số 120" in reader.pages[-1].extract_text() or "Mặt hàng số 120" in "".join(p.extract_text() for p in reader.pages)
