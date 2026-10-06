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
