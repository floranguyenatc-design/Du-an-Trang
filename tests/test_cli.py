"""Chạy trọn luồng `pull` và `parse-xml` với server giả lập."""

from pathlib import Path
from unittest import mock

import requests
from openpyxl import load_workbook

from hddt import cli
from tests.test_client import FakeGdt

FIX = Path(__file__).parent / "fixtures" / "hoadon_mau.xml"
RealSession = requests.Session


def _run(argv, fake, monkeypatch):
    session = mock.Mock(spec=RealSession)
    session.request.side_effect = fake.request
    session.proxies = {}
    monkeypatch.setattr(requests, "Session", lambda: session)
    monkeypatch.setenv("GDT_USERNAME", "0109876543")
    monkeypatch.setenv("GDT_PASSWORD", "secret")
    monkeypatch.setenv("REQUEST_DELAY_MS", "0")
    monkeypatch.setenv("PAGE_SIZE", "2")
    return cli.main(argv)


def test_pull_end_to_end(tmp_path, monkeypatch):
    fake = FakeGdt()
    out = tmp_path / "out"
    code = _run(["--env", str(tmp_path / "no.env"), "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--luong", "2"], fake, monkeypatch)
    assert code == 0  # hóa đơn 126 không có XML nhưng đã lấy chi tiết từ GDT -> không tính là lỗi
    xlsx = out / "HoaDon_20251201_20251231.xlsx"
    assert xlsx.is_file()
    wb = load_workbook(xlsx)
    ws = wb["HoaDon"]
    rows = list(ws.iter_rows(values_only=True))
    assert rows[0][:5] == ("Chiều", "Nguồn", "Ký hiệu mẫu số", "Ký hiệu hóa đơn", "Số hóa đơn")
    assert len(rows) == 5  # 1 header + 4 hóa đơn
    by_no = {int(r[4]): r for r in rows[1:]}
    assert by_no[125][18] == "Hóa đơn mới" and by_no[125][19] == "Đã cấp mã hóa đơn"
    assert by_no[9][1] == "Máy tính tiền"
    assert "chi tiết lấy từ màn hình xem hóa đơn" in by_no[126][-1]
    detail = list(wb["ChiTiet"].iter_rows(min_row=2, values_only=True))
    hd126 = [r for r in detail if r[3] == 126]
    assert [r[12] for r in hd126] == ["Cước dịch vụ viễn thông", "Phí SIM"]
    assert hd126[0][13] == "Tháng" and hd126[0][14] == 1 and hd126[0][15] == 350000 and hd126[0][18] == 350000
    assert hd126[0][19] == "10%" and hd126[0][20] == 35000
    assert (out / "xml" / "purchase" / "purchase_query_0312345678_1_C25TAA_125.xml").is_file()
    assert (out / "xml" / "purchase" / "purchase_query_0312345678_1_C25TAA_125.html").is_file()
    err_rows = list(wb["Loi"].iter_rows(values_only=True))
    assert len(err_rows) == 1  # chỉ dòng tiêu đề
    assert (out / "xml" / "purchase" / "purchase_query_0312345678_1_C25TAA_126.json").is_file()
    assert (out / "danh_sach_20251201_20251231.json").is_file()


def test_pull_reuses_existing_xml(tmp_path, monkeypatch):
    fake = FakeGdt()
    out = tmp_path / "out"
    _run(["--env", "x", "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-mtt"], fake, monkeypatch)
    first = sum(1 for c in fake.calls if "export-xml" in c[1])
    _run(["--env", "x", "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-mtt"], fake, monkeypatch)
    second = sum(1 for c in fake.calls if "export-xml" in c[1]) - first
    assert first == 3 and second == 0  # lần 2 dùng lại cả XML lẫn JSON chi tiết đã lưu


def test_parse_xml_offline(tmp_path, monkeypatch):
    folder = tmp_path / "xml" / "purchase"
    folder.mkdir(parents=True)
    (folder / "a.xml").write_bytes(FIX.read_bytes())
    out = tmp_path / "local.xlsx"
    code = cli.main(["--env", "x", "parse-xml", str(tmp_path / "xml"), "--excel", str(out)])
    assert code == 0
    wb = load_workbook(out)
    assert wb["HoaDon"].max_row == 2
    assert wb["ChiTiet"].max_row == 4
    detail = list(wb["ChiTiet"].iter_rows(min_row=2, values_only=True))
    assert detail[0][12] == "Dịch vụ tư vấn" and detail[0][20] == 200000


def test_global_options_accepted_after_subcommand(tmp_path, monkeypatch):
    fake = FakeGdt()
    out = tmp_path / "out"
    log = tmp_path / "log.txt"
    code = _run(["pull", "--env", "x", "--log-file", str(log), "--thang", "01/12/2025-31/12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-xml"], fake, monkeypatch)
    assert code == 0
    assert (out / "HoaDon_20251201_20251231.xlsx").is_file()
    assert log.is_file()


def test_gui_helpers_without_tk(tmp_path):
    import keo_hoa_don as gui

    env = tmp_path / ".env"
    (tmp_path / ".env.example").write_text("# mau\nGDT_USERNAME=\nGDT_PASSWORD=\nOUTPUT_DIR=output\n", encoding="utf-8")
    gui.save_env({"GDT_USERNAME": "0109876543", "GDT_PASSWORD": "a=b", "XML_WORKERS": "2"}, env)
    vals = gui.read_env(env)
    assert vals["GDT_USERNAME"] == "0109876543" and vals["GDT_PASSWORD"] == "a=b" and vals["XML_WORKERS"] == "2"
    assert env.read_text(encoding="utf-8").startswith("# mau\n")

    opts = gui.build_options({"username": "0109876543", "password": "x", "start": "01/01/2024", "end": "31/12/2024", "direction": "purchase", "include_sco": False})
    assert opts.families == ["query"] and opts.directions == ["purchase"]
    assert str(opts.start) == "2024-01-01" and str(opts.end) == "2024-12-31"
    import pytest

    with pytest.raises(ValueError):
        gui.build_options({"username": "", "password": "x", "start": "01/01/2024", "end": "31/12/2024"})
    with pytest.raises(ValueError):
        gui.build_options({"username": "a", "password": "x", "start": "31/12/2024", "end": "01/01/2024"})


def test_no_xml_and_no_detail_goes_to_error_sheet(tmp_path, monkeypatch):
    fake = FakeGdt()
    fake.invoices[("query", "purchase")].append(
        {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "999", "tdlap": "2025-12-28T00:00:00", "tthai": 1, "ttxly": 6}
    )
    original = fake.request

    def no_xml_999(method, url, **kw):
        if "export-xml" in url and "shdon=999" in url:
            return __import__("tests.test_client", fromlist=["FakeResponse"]).FakeResponse(500, {"message": "x"})
        return original(method, url, **kw)

    fake.request = no_xml_999
    out = tmp_path / "out"
    code = _run(["--env", "x", "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-mtt"], fake, monkeypatch)
    assert code == 2
    wb = load_workbook(out / "HoaDon_20251201_20251231.xlsx")
    err_rows = list(wb["Loi"].iter_rows(min_row=2, values_only=True))
    assert [r[5] for r in err_rows] == ["999"]


def test_pull_creates_pdf_per_invoice(tmp_path, monkeypatch):
    from pypdf import PdfReader

    fake = FakeGdt()
    out = tmp_path / "out"
    code = _run(["--env", "x", "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-mtt"], fake, monkeypatch)
    assert code == 0
    pdf_dir = out / "pdf" / "purchase"
    pdfs = sorted(p.name for p in pdf_dir.glob("*.pdf"))
    assert pdfs == [
        "purchase_query_0312345678_1_C25TAA_125.pdf",
        "purchase_query_0312345678_1_C25TAA_126.pdf",
        "purchase_query_0399999999_2_C25TBB_7.pdf",
    ]
    # Hóa đơn 126 không có XML: PDF dựng từ chi tiết GDT, có đủ dòng hàng.
    text = "".join(pg.extract_text() for pg in PdfReader(str(pdf_dir / pdfs[1])).pages)
    assert "Cước dịch vụ viễn thông" in text and "350.000" in text and "Tập đoàn Công nghiệp" in text
    assert "không có file XML gốc" in text
    wb = load_workbook(out / "HoaDon_20251201_20251231.xlsx")
    header = [c.value for c in wb["HoaDon"][1]]
    col = header.index("File PDF")
    values = [r[col] for r in wb["HoaDon"].iter_rows(min_row=2, values_only=True)]
    assert all(v and v.endswith(".pdf") for v in values)


def test_pull_without_pdf(tmp_path, monkeypatch):
    fake = FakeGdt()
    out = tmp_path / "out"
    _run(["--env", "x", "pull", "--thang", "12/2025", "--chieu", "mua", "--thu-muc", str(out), "--khong-mtt", "--khong-pdf"], fake, monkeypatch)
    assert not (out / "pdf").exists()
