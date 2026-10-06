"""Luồng kéo hóa đơn dùng chung cho dòng lệnh (cli.py) và giao diện (keo_hoa_don.py)."""

from __future__ import annotations

import json
import logging
import os
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Callable

from .client import ApiError, FAMILIES, GdtClient, HddtError, InvoiceRef, StopRequested
from .detail import parsed_from_detail
from .pdfrender import render_invoice_pdf
from .export import invoice_row, line_rows, write_workbook
from .xmlparse import ParsedInvoice, parse_invoice_file, parse_invoice_xml

log = logging.getLogger("hddt")


@dataclass
class PullOptions:
    start: date
    end: date
    directions: list[str] = field(default_factory=lambda: ["purchase", "sold"])
    families: list[str] = field(default_factory=lambda: list(FAMILIES))
    ttxly: int | None = None
    output_dir: str = "output"
    excel_path: str = ""
    download_xml: bool = True
    redownload: bool = False
    # Khi GDT không có XML (hóa đơn không mã của viễn thông/ngân hàng...), lấy chi
    # tiết dòng hàng từ màn hình "Xem chi tiết" (invoices/detail) thay thế.
    detail_fallback: bool = True
    # Tạo file PDF (bản thể hiện) cho từng hóa đơn trong thư mục pdf/.
    make_pdf: bool = True
    workers: int = 3
    page_size: int = 50


@dataclass
class PullResult:
    excel_path: str = ""
    xml_dir: str = ""
    pdf_dir: str = ""
    pdfs: int = 0
    invoices: int = 0
    lines: int = 0
    errors: int = 0
    stopped: bool = False


def _err(direction: str, family: str, ref: InvoiceRef | None, step: str, message: str) -> dict[str, Any]:
    return {
        "Chiều": "Mua vào" if direction == "purchase" else ("Bán ra" if direction == "sold" else direction),
        "Nguồn": family,
        "MST người bán": ref.nbmst if ref else "",
        "Mẫu số": ref.khmshdon if ref else "",
        "Ký hiệu": ref.khhdon if ref else "",
        "Số HĐ": ref.shdon if ref else "",
        "Bước": step,
        "Lỗi": message,
    }


PDF_NOTE_XML = "Bản thể hiện do công cụ tạo từ file XML gốc (đã ký số) tải về từ hoadondientu.gdt.gov.vn."
PDF_NOTE_DETAIL = (
    "Bản thể hiện do công cụ tạo từ dữ liệu chi tiết hóa đơn trên hoadondientu.gdt.gov.vn "
    "(cổng thuế không có file XML gốc cho hóa đơn này)."
)


def _status_line(ref: InvoiceRef) -> str:
    from .client import TTHAI_LABELS, TTXLY_LABELS

    parts = []
    for key, labels, name in (("tthai", TTHAI_LABELS, "Trạng thái"), ("ttxly", TTXLY_LABELS, "Kết quả kiểm tra")):
        v = ref.raw.get(key)
        try:
            label = labels.get(int(v), str(v)) if v not in (None, "") else ""
        except (TypeError, ValueError):
            label = str(v)
        if label:
            parts.append(f"{name}: {label}")
    return " | ".join(parts)


NO_XML_NOTE = "Không có XML gốc trên cổng thuế; chi tiết lấy từ màn hình xem hóa đơn của GDT"


def run_pull(
    client: GdtClient,
    opts: PullOptions,
    stop_event: threading.Event | None = None,
    progress: Callable[[int, int], None] | None = None,
) -> PullResult:
    """Kéo danh sách -> tải XML -> đọc XML -> ghi Excel. Trả về PullResult."""
    stop_event = stop_event or threading.Event()
    client.stop_event = stop_event
    out_dir = Path(opts.output_dir)
    xml_dir = out_dir / "xml"
    pdf_dir = out_dir / "pdf"
    out_dir.mkdir(parents=True, exist_ok=True)
    result = PullResult(xml_dir=str(xml_dir), pdf_dir=str(pdf_dir))

    client.login()
    log.info(
        "Kéo hóa đơn %s từ %s đến %s (nguồn: %s)",
        "/".join("mua vào" if d == "purchase" else "bán ra" for d in opts.directions),
        f"{opts.start:%d/%m/%Y}", f"{opts.end:%d/%m/%Y}", ", ".join(opts.families),
    )

    refs: list[InvoiceRef] = []
    errors: list[dict[str, Any]] = []
    seen: set[tuple[str, ...]] = set()
    def on_error(d: str, f: str, s: date, e: date, exc: Exception) -> None:
        errors.append(_err(d, f, None, "Danh sách", f"Kỳ {s:%d/%m/%Y}-{e:%d/%m/%Y}: {exc}"))

    for direction in opts.directions:
        if stop_event.is_set():
            break
        try:
            for ref in client.iter_invoices(
                direction, opts.start, opts.end, families=opts.families, ttxly=opts.ttxly, page_size=opts.page_size,
                on_error=on_error,
                on_page=lambda d, f, s, e, p, n: log.info(
                    "[DANH SÁCH] %s/%s %s-%s trang %d: +%d hóa đơn",
                    "mua vào" if d == "purchase" else "bán ra", "MTT" if f == "sco-query" else "HĐĐT",
                    f"{s:%d/%m}", f"{e:%d/%m/%Y}", p, n,
                ),
            ):
                key = (ref.direction, ref.family, ref.nbmst, ref.khmshdon, ref.khhdon, ref.shdon)
                if key in seen:
                    continue
                seen.add(key)
                refs.append(ref)
        except StopRequested:
            break
        except HddtError as exc:
            log.error("Lỗi lấy danh sách %s: %s", direction, exc)
            errors.append(_err(direction, "", None, "Danh sách", str(exc)))
    log.info("Tổng cộng %d hóa đơn.", len(refs))

    raw_path = out_dir / f"danh_sach_{opts.start:%Y%m%d}_{opts.end:%Y%m%d}.json"
    raw_path.write_text(json.dumps([r.raw for r in refs], ensure_ascii=False, indent=1), encoding="utf-8")

    parsed_map: dict[int, ParsedInvoice] = {}
    xml_files: dict[int, str] = {}
    xml_errors: dict[int, str] = {}
    pdf_files: dict[int, str] = {}

    if opts.download_xml and refs and not stop_event.is_set():
        xml_dir.mkdir(parents=True, exist_ok=True)

        def work(idx: int, ref: InvoiceRef) -> tuple[int, str | None, ParsedInvoice | None, str, str, str]:
            idx, path, parsed, err = fetch(idx, ref)
            pdf_path, pdf_err = "", ""
            if parsed is not None and opts.make_pdf and not stop_event.is_set():
                target = pdf_dir / ref.direction / f"{ref.file_stem}.pdf"
                try:
                    if target.is_file() and not opts.redownload:
                        pdf_path = str(target)
                    else:
                        pdf_path = render_invoice_pdf(
                            parsed, target,
                            status_lines=[_status_line(ref)],
                            source_note=PDF_NOTE_XML if path and path.endswith(".xml") else PDF_NOTE_DETAIL,
                        )
                except Exception as exc:  # noqa: BLE001
                    pdf_err = f"Tạo PDF lỗi: {exc}"
            return idx, path, parsed, err, pdf_path, pdf_err

        def fetch(idx: int, ref: InvoiceRef) -> tuple[int, str | None, ParsedInvoice | None, str]:
            if stop_event.is_set():
                return idx, None, None, "Đã dừng"
            target_dir = xml_dir / ref.direction
            target_dir.mkdir(parents=True, exist_ok=True)
            existing = target_dir / f"{ref.file_stem}.xml"
            existing_json = target_dir / f"{ref.file_stem}.json"
            try:
                if existing.is_file() and not opts.redownload:
                    return idx, str(existing), parse_invoice_file(str(existing)), ""
                if existing_json.is_file() and not opts.redownload and opts.detail_fallback:
                    data = json.loads(existing_json.read_text(encoding="utf-8"))
                    return idx, str(existing_json), parsed_from_detail(data), NO_XML_NOTE
                try:
                    files = client.download_xml(ref)
                except ApiError as exc:
                    if not opts.detail_fallback or exc.status not in (400, 404, 500):
                        raise
                    # Không có XML gốc: lấy chi tiết từ GDT.
                    detail = client.get_detail(ref)
                    if not isinstance(detail, dict) or not detail:
                        raise ApiError("GDT không có hồ sơ XML và không trả chi tiết hóa đơn.", exc.status) from exc
                    existing_json.write_text(json.dumps(detail, ensure_ascii=False, indent=1), encoding="utf-8")
                    return idx, str(existing_json), parsed_from_detail(detail), NO_XML_NOTE
                xml_path = None
                parsed = None
                for name, data in files:
                    p = target_dir / name
                    p.write_bytes(data)
                    if name.endswith(".xml") and xml_path is None:
                        xml_path = str(p)
                        try:
                            parsed = parse_invoice_xml(data)
                        except Exception as exc:  # noqa: BLE001
                            return idx, xml_path, None, f"Đọc XML lỗi: {exc}"
                return idx, xml_path, parsed, ""
            except StopRequested:
                return idx, None, None, "Đã dừng"
            except HddtError as exc:
                msg = str(exc)
                if getattr(exc, "status", None) == 500:
                    msg = "GDT không có hồ sơ XML cho hóa đơn này (HTTP 500)"
                return idx, None, None, msg
            except Exception as exc:  # noqa: BLE001
                return idx, None, None, f"Lỗi không xác định: {exc}"

        workers = max(1, min(opts.workers, 10))
        done = 0
        with ThreadPoolExecutor(max_workers=workers) as pool:
            futures = [pool.submit(work, i, r) for i, r in enumerate(refs)]
            for fut in as_completed(futures):
                idx, path, parsed, err, pdf_path, pdf_err = fut.result()
                done += 1
                if pdf_path:
                    pdf_files[idx] = pdf_path
                if pdf_err:
                    errors.append(_err(refs[idx].direction, refs[idx].family, refs[idx], "Tạo PDF", pdf_err))
                    log.warning("[PDF] %s: %s", refs[idx].label, pdf_err)
                if path:
                    xml_files[idx] = path
                if parsed:
                    parsed_map[idx] = parsed
                if err:
                    xml_errors[idx] = err
                    if err == NO_XML_NOTE:
                        log.info("[XML] %d/%d %s: không có XML gốc, đã lấy chi tiết từ GDT.", done, len(refs), refs[idx].label)
                    elif err != "Đã dừng":
                        errors.append(_err(refs[idx].direction, refs[idx].family, refs[idx], "Tải XML", err))
                        log.warning("[XML] %d/%d %s: %s", done, len(refs), refs[idx].label, err)
                elif done % 10 == 0 or done == len(refs):
                    log.info("[XML] %d/%d đã tải.", done, len(refs))
                if progress:
                    progress(done, len(refs))

    rows = []
    for i, r in enumerate(refs):
        row = invoice_row(r, parsed_map.get(i), os.path.relpath(xml_files[i], out_dir) if i in xml_files else "", xml_errors.get(i, ""))
        if i in pdf_files:
            row["pdf_file"] = os.path.relpath(pdf_files[i], out_dir)
        rows.append(row)
    lines: list[dict[str, Any]] = []
    for i, r in enumerate(refs):
        if i in parsed_map:
            lines.extend(line_rows(r, parsed_map[i]))

    excel_path = Path(opts.excel_path) if opts.excel_path else out_dir / f"HoaDon_{opts.start:%Y%m%d}_{opts.end:%Y%m%d}.xlsx"
    write_workbook(str(excel_path), rows, lines, errors)
    log.info("Đã ghi Excel: %s (%d hóa đơn, %d dòng hàng, %d lỗi)", excel_path, len(rows), len(lines), len(errors))

    result.excel_path = str(excel_path)
    result.invoices = len(rows)
    result.lines = len(lines)
    result.pdfs = len(pdf_files)
    result.errors = len(errors)
    result.stopped = stop_event.is_set()
    return result
