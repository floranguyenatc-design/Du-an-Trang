"""Giao diện dòng lệnh: ``python -m hddt <lệnh>``.

Lệnh chính:
    login       Kiểm tra đăng nhập (tự giải CAPTCHA), in thông tin đơn vị.
    pull        Kéo danh sách hóa đơn (mua vào / bán ra), tải XML và xuất Excel.
    excel-gdt   Tải file Excel do GDT xuất sẵn cho khoảng ngày.
    parse-xml   Đọc thư mục XML đã tải (offline) và xuất Excel.
    captcha     Lấy một CAPTCHA và in mã giải được (để kiểm tra bộ giải).
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from . import __version__
from .client import DEFAULT_USER_AGENT, FAMILIES, GdtClient, HddtError, InvoiceRef, LoginError
from .config import Settings, parse_vn_date
from .export import invoice_row, line_rows, write_workbook
from .pull import PullOptions, run_pull
from .xmlparse import parse_invoice_file

log = logging.getLogger("hddt")

DIRECTION_CHOICES = {
    "mua": ["purchase"],
    "ban": ["sold"],
    "ca-hai": ["purchase", "sold"],
    "purchase": ["purchase"],
    "sold": ["sold"],
    "both": ["purchase", "sold"],
}


def _setup_logging(verbose: bool, log_file: str | None = None) -> None:
    handlers: list[logging.Handler] = [logging.StreamHandler(sys.stdout)]
    if log_file:
        handlers.append(logging.FileHandler(log_file, encoding="utf-8"))
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.INFO,
        format="%(asctime)s %(levelname)-5s %(message)s",
        datefmt="%H:%M:%S",
        handlers=handlers,
    )
    logging.getLogger("urllib3").setLevel(logging.WARNING)


def _client(settings: Settings, args: argparse.Namespace) -> GdtClient:
    username = getattr(args, "user", None) or settings.username
    password = getattr(args, "password", None) or settings.password
    token = getattr(args, "token", None) or settings.token
    return GdtClient(
        username=username,
        password=password,
        token=token,
        user_agent=settings.user_agent or DEFAULT_USER_AGENT,
        proxies=settings.proxies(),
        timeout=settings.timeout,
        request_interval=settings.request_interval,
        verify_ssl=settings.verify_ssl,
    )


def _families(args: argparse.Namespace) -> list[str]:
    fams = list(FAMILIES)
    if getattr(args, "khong_mtt", False):
        fams.remove("sco-query")
    if getattr(args, "chi_mtt", False):
        fams = ["sco-query"]
    return fams


def _default_dates(args: argparse.Namespace) -> tuple[date, date]:
    if args.thang and "-" in args.thang.strip() and args.thang.count("/") >= 4:
        a, _, b = args.thang.partition("-")
        args.tu_ngay, args.den_ngay, args.thang = a.strip(), b.strip(), None
    if args.thang:
        y, m = _parse_month(args.thang)
        start = date(y, m, 1)
        end = (start + timedelta(days=32)).replace(day=1) - timedelta(days=1)
        return start, end
    if not args.tu_ngay or not args.den_ngay:
        raise SystemExit("Cần --tu-ngay và --den-ngay (dd/mm/yyyy) hoặc --thang mm/yyyy.")
    start, end = parse_vn_date(args.tu_ngay), parse_vn_date(args.den_ngay)
    if start > end:
        raise SystemExit("--tu-ngay phải nhỏ hơn hoặc bằng --den-ngay.")
    return start, end


def _parse_month(text: str) -> tuple[int, int]:
    for fmt in ("%m/%Y", "%Y-%m", "%m-%Y"):
        try:
            d = datetime.strptime(text.strip(), fmt)
            return d.year, d.month
        except ValueError:
            continue
    raise SystemExit(f"--thang không hợp lệ: {text!r} (dùng mm/yyyy)")


# ----------------------------------------------------------------- commands
def cmd_login(args: argparse.Namespace, settings: Settings) -> int:
    client = _client(settings, args)
    client.login()
    try:
        profile = client.get_profile()
        print("Đăng nhập thành công.")
        for key in ("name", "tin", "username", "email", "phone"):
            if profile.get(key):
                print(f"  {key}: {profile[key]}")
    except HddtError as exc:
        print(f"Đăng nhập thành công nhưng không đọc được profile: {exc}")
    if args.in_token:
        print("TOKEN:", client.token)
    return 0


def cmd_captcha(args: argparse.Namespace, settings: Settings) -> int:
    client = _client(settings, args)
    for _ in range(args.so_lan):
        key, code, svg = client.get_captcha()
        print(f"key={key}  giải được: {code or '(không nhận ra)'}")
        if args.luu:
            Path(args.luu).mkdir(parents=True, exist_ok=True)
            (Path(args.luu) / f"captcha_{key[:12] or 'x'}_{code or 'unknown'}.svg").write_text(svg, encoding="utf-8")
    return 0


def cmd_excel_gdt(args: argparse.Namespace, settings: Settings) -> int:
    client = _client(settings, args)
    start, end = _default_dates(args)
    out_dir = Path(args.thu_muc or settings.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    for direction in DIRECTION_CHOICES[args.chieu]:
        for family in _families(args):
            ttxly = args.ttxly if args.ttxly is not None else (8 if family == "sco-query" else None)
            data = client.download_excel(direction, start, end, family=family, ttxly=ttxly)
            name = f"GDT_{direction}_{family}_{start:%Y%m%d}_{end:%Y%m%d}.xlsx"
            (out_dir / name).write_bytes(data)
            print(f"Đã lưu {out_dir / name} ({len(data):,} bytes)")
    return 0


def cmd_pull(args: argparse.Namespace, settings: Settings) -> int:
    client = _client(settings, args)
    start, end = _default_dates(args)
    opts = PullOptions(
        start=start,
        end=end,
        directions=DIRECTION_CHOICES[args.chieu],
        families=_families(args),
        ttxly=args.ttxly,
        output_dir=args.thu_muc or settings.output_dir,
        excel_path=args.excel or "",
        download_xml=not args.khong_xml,
        redownload=args.tai_lai,
        make_pdf=not args.khong_pdf,
        workers=args.luong or settings.xml_workers,
        page_size=settings.page_size,
    )
    res = run_pull(client, opts)
    print(f"\nXong. Hóa đơn: {res.invoices} | Dòng hàng: {res.lines} | Lỗi: {res.errors}")
    print(f"Excel : {res.excel_path}")
    if opts.download_xml:
        print(f"XML   : {res.xml_dir}")
    if res.pdfs:
        print(f"PDF   : {res.pdf_dir} ({res.pdfs} file)")
    return 0 if not res.errors else 2


def cmd_misa(args: argparse.Namespace, settings: Settings) -> int:
    from .misa import export_misa

    res = export_misa(args.thu_muc or settings.output_dir, settings.misa_settings())
    print(f"Đã chuẩn bị {res.vouchers} chứng từ ({res.lines} dòng hàng) cho MISA SME.")
    print(f"  File nhập khẩu (mẫu MISA) : {res.excel_path}")
    print(f"  Danh mục + kiểm tra       : {res.catalog_path}")
    print(f"  XML gốc                   : {res.xml_dir} ({res.xml_files} file)")
    print(f"  Không có XML gốc          : {res.no_xml} hóa đơn")
    print(f"  Bỏ qua (hủy/thay thế)     : {res.skipped}")
    for w in res.warnings:
        print(f"  ! {w}")
    print(f"  Hướng dẫn                 : {Path(res.folder) / 'HUONG_DAN_NHAP_MISA.txt'}")
    return 0


def cmd_parse_xml(args: argparse.Namespace, settings: Settings) -> int:
    folder = Path(args.thu_muc_xml)
    if not folder.is_dir():
        raise SystemExit(f"Không thấy thư mục {folder}")
    rows: list[dict[str, Any]] = []
    lines: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    files = sorted(folder.rglob("*.xml"))
    for path in files:
        try:
            parsed = parse_invoice_file(str(path))
        except Exception as exc:  # noqa: BLE001
            errors.append({"Chiều": "", "Nguồn": "", "MST người bán": "", "Mẫu số": "", "Ký hiệu": "", "Số HĐ": str(path.name), "Bước": "Đọc XML", "Lỗi": str(exc)})
            continue
        direction = "purchase"
        lowered = str(path).lower()
        if "sold" in lowered or "ban" in lowered.split(os.sep):
            direction = "sold"
        if args.chieu in ("mua", "purchase"):
            direction = "purchase"
        elif args.chieu in ("ban", "sold"):
            direction = "sold"
        ref = InvoiceRef(
            direction=direction, family="query", nbmst=parsed.nb_mst, khmshdon=parsed.khmshdon,
            khhdon=parsed.khhdon, shdon=parsed.shdon, raw={},
        )
        rows.append(invoice_row(ref, parsed, str(path), ""))
        lines.extend(line_rows(ref, parsed))
    out = Path(args.excel or (folder / "HoaDon_Local.xlsx"))
    write_workbook(str(out), rows, lines, errors)
    print(f"Đã đọc {len(files)} file XML -> {out} ({len(rows)} hóa đơn, {len(lines)} dòng hàng, {len(errors)} lỗi)")
    return 0


# ------------------------------------------------------------------- parser
def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(prog="hddt", description="Kéo hóa đơn điện tử từ hệ thống của Cơ quan Thuế (hoadondientu.gdt.gov.vn).")
    p.add_argument("--version", action="version", version=f"hddt {__version__}")
    p.add_argument("--env", default=".env", help="File cấu hình .env (mặc định: .env)")
    p.add_argument("--user", help="MST / tài khoản đăng nhập GDT (ghi đè GDT_USERNAME)")
    p.add_argument("--password", help="Mật khẩu GDT (ghi đè GDT_PASSWORD)")
    p.add_argument("--token", help="Dùng sẵn token JWT lấy từ trình duyệt (bỏ qua bước đăng nhập)")
    p.add_argument("-v", "--verbose", action="store_true", help="In log chi tiết")
    p.add_argument("--log-file", help="Ghi log ra file")
    sub = p.add_subparsers(dest="cmd", required=True)
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("-v", "--verbose", action="store_true", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    common.add_argument("--log-file", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    common.add_argument("--env", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    common.add_argument("--user", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    common.add_argument("--password", default=argparse.SUPPRESS, help=argparse.SUPPRESS)
    common.add_argument("--token", default=argparse.SUPPRESS, help=argparse.SUPPRESS)

    s = sub.add_parser("login", parents=[common], help="Kiểm tra đăng nhập")
    s.add_argument("--in-token", action="store_true", help="In token ra màn hình")
    s.set_defaults(func=cmd_login)

    s = sub.add_parser("captcha", parents=[common], help="Lấy CAPTCHA và in mã giải được")
    s.add_argument("--so-lan", type=int, default=1)
    s.add_argument("--luu", help="Thư mục lưu SVG để kiểm tra")
    s.set_defaults(func=cmd_captcha)

    def add_range(sp: argparse.ArgumentParser) -> None:
        sp.add_argument("--tu-ngay", help="Từ ngày lập, dd/mm/yyyy")
        sp.add_argument("--den-ngay", help="Đến ngày lập, dd/mm/yyyy")
        sp.add_argument("--thang", help="Lấy trọn tháng, mm/yyyy (thay cho --tu-ngay/--den-ngay)")
        sp.add_argument("--chieu", choices=sorted(DIRECTION_CHOICES), default="ca-hai", help="mua | ban | ca-hai (mặc định ca-hai)")
        sp.add_argument("--ttxly", type=int, help="Lọc kết quả kiểm tra (5=đã cấp mã, 6=không mã, 8=máy tính tiền). Mặc định lấy tất cả.")
        sp.add_argument("--khong-mtt", action="store_true", help="Bỏ qua hóa đơn máy tính tiền (sco-query)")
        sp.add_argument("--chi-mtt", action="store_true", help="Chỉ lấy hóa đơn máy tính tiền")
        sp.add_argument("--thu-muc", help="Thư mục kết quả (mặc định OUTPUT_DIR hoặc ./output)")

    s = sub.add_parser("pull", parents=[common], help="Kéo danh sách + XML + xuất Excel")
    add_range(s)
    s.add_argument("--excel", help="Đường dẫn file Excel đầu ra")
    s.add_argument("--khong-xml", action="store_true", help="Chỉ lấy danh sách, không tải XML")
    s.add_argument("--tai-lai", action="store_true", help="Tải lại XML (và tạo lại PDF) dù đã có file")
    s.add_argument("--khong-pdf", action="store_true", help="Không tạo file PDF cho từng hóa đơn")
    s.add_argument("--luong", type=int, help="Số luồng tải XML song song (1-10, mặc định XML_WORKERS=3)")
    s.set_defaults(func=cmd_pull)

    s = sub.add_parser("excel-gdt", parents=[common], help="Tải file Excel do GDT xuất sẵn")
    add_range(s)
    s.set_defaults(func=cmd_excel_gdt)

    s = sub.add_parser("misa", parents=[common], help="Chuẩn bị file nhập khẩu MISA SME từ hóa đơn mua vào đã kéo")
    s.add_argument("--thu-muc", help="Thư mục kết quả đã kéo (mặc định OUTPUT_DIR hoặc ./output)")
    s.set_defaults(func=cmd_misa)

    s = sub.add_parser("parse-xml", parents=[common], help="Đọc thư mục XML có sẵn và xuất Excel (offline)")
    s.add_argument("thu_muc_xml", help="Thư mục chứa file .xml")
    s.add_argument("--chieu", choices=["auto", "mua", "ban", "purchase", "sold"], default="auto")
    s.add_argument("--excel", help="File Excel đầu ra")
    s.set_defaults(func=cmd_parse_xml)
    return p


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    _setup_logging(args.verbose, args.log_file)
    settings = Settings.from_env(args.env)
    try:
        return args.func(args, settings)
    except LoginError as exc:
        log.error("%s", exc)
        return 3
    except (HddtError, FileNotFoundError) as exc:
        log.error("%s", exc)
        return 1
    except KeyboardInterrupt:
        log.warning("Đã dừng theo yêu cầu.")
        return 130
