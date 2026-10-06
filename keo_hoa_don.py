# -*- coding: utf-8 -*-
"""Giao diện cửa sổ: Tải hóa đơn điện tử từ Cơ quan Thuế (hoadondientu.gdt.gov.vn).

Chạy bằng Tai-hoa-don.bat (hoặc: python keo_hoa_don.py). Không cần gõ lệnh.
"""

from __future__ import annotations

import logging
import os
import queue
import subprocess
import sys
import threading
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path

HERE = Path(__file__).resolve().parent
os.chdir(HERE)
sys.path.insert(0, str(HERE))

ENV_FILE = HERE / ".env"
ERROR_FILE = HERE / "loi_giao_dien.txt"


# --------------------------------------------------------------------------- tiện ích không phụ thuộc Tk
def ensure_dependencies() -> str:
    """Cài requests/openpyxl nếu máy chưa có. Trả về thông báo (rỗng nếu không làm gì)."""
    missing = []
    for mod in ("requests", "openpyxl", "reportlab"):
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    if not missing:
        return ""
    cmd = [sys.executable, "-m", "pip", "install", "-q", *missing]
    subprocess.run(cmd, check=False, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return "Đã cài thư viện: " + ", ".join(missing)


def read_env(path: Path = ENV_FILE) -> dict[str, str]:
    values: dict[str, str] = {}
    if not path.is_file():
        return values
    for raw in path.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, _, v = line.partition("=")
        values[k.strip()] = v.strip()
    return values


def save_env(updates: dict[str, str], path: Path = ENV_FILE) -> None:
    """Cập nhật các khóa trong .env, giữ nguyên phần còn lại và chú thích."""
    lines: list[str] = []
    if path.is_file():
        lines = path.read_text(encoding="utf-8-sig").splitlines()
    elif (path.parent / ".env.example").is_file():
        lines = (path.parent / ".env.example").read_text(encoding="utf-8-sig").splitlines()
    done: set[str] = set()
    out: list[str] = []
    for line in lines:
        s = line.strip()
        if s and not s.startswith("#") and "=" in s:
            k = s.partition("=")[0].strip()
            if k in updates:
                out.append(f"{k}={updates[k]}")
                done.add(k)
                continue
        out.append(line)
    for k, v in updates.items():
        if k not in done:
            out.append(f"{k}={v}")
    path.write_text("\n".join(out) + "\n", encoding="utf-8")


def parse_vn_date(text: str) -> date:
    t = text.strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Ngày không hợp lệ: '{text}' (nhập dạng dd/mm/yyyy)")


def month_bounds(d: date) -> tuple[date, date]:
    first = d.replace(day=1)
    last = (first + timedelta(days=32)).replace(day=1) - timedelta(days=1)
    return first, last


def build_options(form: dict):
    """Chuyển giá trị trên form thành PullOptions; ném ValueError nếu nhập sai."""
    from hddt.pull import PullOptions

    username = (form.get("username") or "").strip()
    password = form.get("password") or ""
    if not username or not password:
        raise ValueError("Chưa nhập MST hoặc mật khẩu đăng nhập trang thuế.")
    start = parse_vn_date(form.get("start") or "")
    end = parse_vn_date(form.get("end") or "")
    if start > end:
        raise ValueError("'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
    chieu = form.get("direction") or "both"
    directions = {"purchase": ["purchase"], "sold": ["sold"], "both": ["purchase", "sold"]}[chieu]
    families = ["query", "sco-query"] if form.get("include_sco", True) else ["query"]
    out_dir = (form.get("output_dir") or "output").strip() or "output"
    return PullOptions(
        start=start,
        end=end,
        directions=directions,
        families=families,
        output_dir=out_dir,
        download_xml=bool(form.get("download_xml", True)),
        redownload=bool(form.get("redownload", False)),
        make_pdf=bool(form.get("make_pdf", True)),
        export_misa=bool(form.get("export_misa", False)),
        workers=int(form.get("workers") or 3),
    )


# --------------------------------------------------------------------------- giao diện
class QueueLogHandler(logging.Handler):
    def __init__(self, q: "queue.Queue[str]"):
        super().__init__()
        self.q = q

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self.q.put(self.format(record))
        except Exception:  # noqa: BLE001
            pass


def run_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    from hddt.client import GdtClient, HddtError, LoginError
    from hddt.pull import run_pull

    env = read_env()
    today = date.today()
    m_first, m_last = month_bounds(today)

    root = tk.Tk()
    from hddt import __version__

    root.title(f"Tải hóa đơn điện tử từ Cơ quan Thuế - phiên bản {__version__}")
    root.minsize(760, 620)
    try:
        root.iconbitmap(default="")
    except Exception:  # noqa: BLE001
        pass

    pad = {"padx": 6, "pady": 4}
    frm = ttk.Frame(root, padding=10)
    frm.pack(fill="both", expand=True)
    frm.columnconfigure(1, weight=1)
    frm.columnconfigure(3, weight=1)

    # --- tài khoản
    ttk.Label(frm, text="MST / tài khoản:").grid(row=0, column=0, sticky="e", **pad)
    v_user = tk.StringVar(value=env.get("GDT_USERNAME", ""))
    ttk.Entry(frm, textvariable=v_user, width=28).grid(row=0, column=1, sticky="we", **pad)

    ttk.Label(frm, text="Mật khẩu:").grid(row=0, column=2, sticky="e", **pad)
    v_pass = tk.StringVar(value=env.get("GDT_PASSWORD", ""))
    e_pass = ttk.Entry(frm, textvariable=v_pass, width=22, show="*")
    e_pass.grid(row=0, column=3, sticky="we", **pad)
    v_show = tk.BooleanVar(value=False)
    ttk.Checkbutton(frm, text="Hiện", variable=v_show, command=lambda: e_pass.config(show="" if v_show.get() else "*")).grid(row=0, column=4, sticky="w")

    # --- khoảng ngày
    ttk.Label(frm, text="Từ ngày:").grid(row=1, column=0, sticky="e", **pad)
    v_start = tk.StringVar(value=f"{m_first:%d/%m/%Y}")
    ttk.Entry(frm, textvariable=v_start, width=14).grid(row=1, column=1, sticky="w", **pad)
    ttk.Label(frm, text="Đến ngày:").grid(row=1, column=2, sticky="e", **pad)
    v_end = tk.StringVar(value=f"{m_last:%d/%m/%Y}")
    ttk.Entry(frm, textvariable=v_end, width=14).grid(row=1, column=3, sticky="w", **pad)

    quick = ttk.Frame(frm)
    quick.grid(row=2, column=1, columnspan=4, sticky="w", **pad)

    def set_range(a: date, b: date) -> None:
        v_start.set(f"{a:%d/%m/%Y}")
        v_end.set(f"{b:%d/%m/%Y}")

    prev_first, prev_last = month_bounds(m_first - timedelta(days=1))
    q = (today.month - 1) // 3
    q_first = date(today.year, q * 3 + 1, 1)
    pq_last = q_first - timedelta(days=1)
    pq_first = date(pq_last.year, ((pq_last.month - 1) // 3) * 3 + 1, 1)
    for text, a, b in (
        ("Tháng này", m_first, m_last),
        ("Tháng trước", prev_first, prev_last),
        ("Quý trước", pq_first, pq_last),
        ("Năm nay", date(today.year, 1, 1), date(today.year, 12, 31)),
        ("Năm trước", date(today.year - 1, 1, 1), date(today.year - 1, 12, 31)),
    ):
        ttk.Button(quick, text=text, command=lambda a=a, b=b: set_range(a, b)).pack(side="left", padx=2)

    # --- loại hóa đơn
    ttk.Label(frm, text="Loại hóa đơn:").grid(row=3, column=0, sticky="e", **pad)
    v_dir = tk.StringVar(value="both")
    dir_frame = ttk.Frame(frm)
    dir_frame.grid(row=3, column=1, columnspan=4, sticky="w", **pad)
    for text, val in (("Mua vào", "purchase"), ("Bán ra", "sold"), ("Cả hai", "both")):
        ttk.Radiobutton(dir_frame, text=text, value=val, variable=v_dir).pack(side="left", padx=4)

    v_sco = tk.BooleanVar(value=True)
    v_xml = tk.BooleanVar(value=True)
    v_redo = tk.BooleanVar(value=False)
    v_pdf = tk.BooleanVar(value=True)
    v_misa = tk.BooleanVar(value=True)
    v_remember = tk.BooleanVar(value=True)
    opt_frame = ttk.Frame(frm)
    opt_frame.grid(row=4, column=1, columnspan=4, sticky="w", **pad)
    line1 = ttk.Frame(opt_frame)
    line1.pack(anchor="w")
    line2 = ttk.Frame(opt_frame)
    line2.pack(anchor="w", pady=(2, 0))
    ttk.Checkbutton(line1, text="Gồm hóa đơn máy tính tiền", variable=v_sco).pack(side="left", padx=4)
    ttk.Checkbutton(line1, text="Tải XML gốc (lấy chi tiết hàng hóa)", variable=v_xml).pack(side="left", padx=4)
    ttk.Checkbutton(line1, text="Tạo file PDF", variable=v_pdf).pack(side="left", padx=4)
    ttk.Checkbutton(line2, text="Tạo file nhập MISA", variable=v_misa).pack(side="left", padx=4)
    ttk.Checkbutton(line2, text="Tải lại XML đã có", variable=v_redo).pack(side="left", padx=4)
    ttk.Checkbutton(line2, text="Ghi nhớ mật khẩu", variable=v_remember).pack(side="left", padx=4)

    # --- thư mục kết quả
    ttk.Label(frm, text="Thư mục kết quả:").grid(row=5, column=0, sticky="e", **pad)
    v_out = tk.StringVar(value=env.get("OUTPUT_DIR", "output") or "output")
    ttk.Entry(frm, textvariable=v_out).grid(row=5, column=1, columnspan=3, sticky="we", **pad)

    def choose_dir() -> None:
        d = filedialog.askdirectory(initialdir=str(HERE), title="Chọn thư mục lưu kết quả")
        if d:
            v_out.set(d)

    ttk.Button(frm, text="Chọn...", command=choose_dir).grid(row=5, column=4, sticky="w", **pad)

    # --- nút
    btns = ttk.Frame(frm)
    btns.grid(row=6, column=0, columnspan=5, sticky="we", pady=(8, 4))
    b_login = ttk.Button(btns, text="Kiểm tra đăng nhập")
    b_start = ttk.Button(btns, text="▶  Bắt đầu tải hóa đơn")
    b_stop = ttk.Button(btns, text="■  Dừng", state="disabled")
    b_open = ttk.Button(btns, text="Mở thư mục kết quả")
    b_excel = ttk.Button(btns, text="Mở file Excel", state="disabled")
    b_misa = ttk.Button(btns, text="Xuất sang MISA")
    for b in (b_login, b_start, b_stop, b_open, b_excel, b_misa):
        b.pack(side="left", padx=4)

    prog = ttk.Progressbar(frm, mode="determinate")
    prog.grid(row=7, column=0, columnspan=5, sticky="we", padx=6)
    v_status = tk.StringVar(value="Sẵn sàng. Điền thông tin rồi bấm 'Bắt đầu tải hóa đơn'.")
    ttk.Label(frm, textvariable=v_status).grid(row=8, column=0, columnspan=5, sticky="w", padx=6)

    txt = tk.Text(frm, height=18, wrap="word", state="disabled", font=("Consolas", 9))
    txt.grid(row=9, column=0, columnspan=5, sticky="nsew", padx=6, pady=6)
    frm.rowconfigure(9, weight=1)
    sb = ttk.Scrollbar(frm, command=txt.yview)
    sb.grid(row=9, column=5, sticky="ns")
    txt.config(yscrollcommand=sb.set)

    # --- log
    log_q: "queue.Queue[str]" = queue.Queue()
    ui_q: "queue.Queue" = queue.Queue()

    def ui(fn) -> None:
        """Lên lịch chạy fn trên luồng giao diện (an toàn khi gọi từ luồng nền)."""
        ui_q.put(fn)

    handler = QueueLogHandler(log_q)
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)-5s %(message)s", "%H:%M:%S"))
    logger = logging.getLogger("hddt")
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)

    def append(line: str) -> None:
        txt.config(state="normal")
        txt.insert("end", line + "\n")
        txt.see("end")
        txt.config(state="disabled")

    def poll() -> None:
        try:
            while True:
                append(log_q.get_nowait())
        except queue.Empty:
            pass
        try:
            while True:
                fn = ui_q.get_nowait()
                try:
                    fn()
                except Exception as exc:  # noqa: BLE001
                    append(f"Lỗi giao diện: {exc}")
        except queue.Empty:
            pass
        root.after(150, poll)

    poll()

    state = {"thread": None, "stop": threading.Event(), "excel": "", "file_handler": None}

    def form() -> dict:
        return {
            "username": v_user.get(),
            "password": v_pass.get(),
            "start": v_start.get(),
            "end": v_end.get(),
            "direction": v_dir.get(),
            "include_sco": v_sco.get(),
            "download_xml": v_xml.get(),
            "redownload": v_redo.get(),
            "make_pdf": v_pdf.get() and v_xml.get(),
            "export_misa": v_misa.get() and v_xml.get(),
            "output_dir": v_out.get(),
            "workers": 3,
        }

    def persist() -> None:
        updates = {"GDT_USERNAME": v_user.get().strip(), "OUTPUT_DIR": v_out.get().strip() or "output"}
        if v_remember.get():
            updates["GDT_PASSWORD"] = v_pass.get()
        try:
            save_env(updates)
        except OSError as exc:
            append(f"Không lưu được .env: {exc}")

    def set_busy(busy: bool, stoppable: bool = True) -> None:
        for b in (b_login, b_start, b_misa):
            b.config(state="disabled" if busy else "normal")
        b_stop.config(state="normal" if busy and stoppable else "disabled")

    def make_client() -> GdtClient:
        return GdtClient(
            v_user.get().strip(), v_pass.get(),
            request_interval=float(env.get("REQUEST_DELAY_MS", "600") or 600) / 1000.0,
            proxies={"http": env["PROXY_URL"], "https": env["PROXY_URL"]} if env.get("PROXY_URL") else None,
        )

    def attach_file_log(out_dir: str) -> None:
        if state["file_handler"]:
            logger.removeHandler(state["file_handler"])
        Path(out_dir).mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(Path(out_dir) / f"log_{datetime.now():%Y%m%d_%H%M%S}.txt", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)-5s %(message)s"))
        logger.addHandler(fh)
        state["file_handler"] = fh

    # --- hành động
    def do_login() -> None:
        if not v_user.get().strip() or not v_pass.get():
            messagebox.showwarning("Thiếu thông tin", "Nhập MST và mật khẩu trước.")
            return
        persist()
        set_busy(True)
        v_status.set("Đang đăng nhập (tự giải CAPTCHA)...")

        def worker() -> None:
            try:
                c = make_client()
                c.login()
                info = ""
                try:
                    p = c.get_profile()
                    info = " - " + " / ".join(str(p[k]) for k in ("name", "tin") if p.get(k))
                except HddtError:
                    pass
                ui(lambda: (v_status.set("Đăng nhập thành công" + info), messagebox.showinfo("OK", "Đăng nhập thành công" + info)))
            except LoginError as exc:
                msg = str(exc)
                ui(lambda m=msg: (v_status.set(f"Đăng nhập thất bại: {m}"), messagebox.showerror("Đăng nhập thất bại", m)))
            except Exception as exc:  # noqa: BLE001
                msg = str(exc) or type(exc).__name__
                logger.error("Lỗi: %s", msg)
                logger.debug("Chi tiết lỗi:\n%s", traceback.format_exc())
                ui(lambda m=msg: (v_status.set(f"Lỗi: {m}"), messagebox.showerror("Lỗi", m)))
            finally:
                ui(lambda: set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def do_start() -> None:
        try:
            opts = build_options(form())
        except ValueError as exc:
            messagebox.showwarning("Kiểm tra lại", str(exc))
            return
        persist()
        if opts.export_misa:
            from hddt.config import Settings

            opts.misa_settings = Settings.from_env(str(ENV_FILE)).misa_settings()
        attach_file_log(opts.output_dir)
        state["stop"] = threading.Event()
        state["excel"] = ""
        b_excel.config(state="disabled")
        prog.config(value=0, maximum=1)
        set_busy(True)
        v_status.set("Đang lấy danh sách hóa đơn...")
        append("=" * 70)

        def progress(done: int, total: int) -> None:
            ui(lambda: (prog.config(maximum=max(total, 1), value=done), v_status.set(f"Đang tải XML {done}/{total}...")))

        def worker() -> None:
            try:
                res = run_pull(make_client(), opts, stop_event=state["stop"], progress=progress)
                state["excel"] = res.excel_path
                msg = (
                    f"{'ĐÃ DỪNG. ' if res.stopped else 'XONG. '}"
                    f"Hóa đơn: {res.invoices} | Dòng hàng hóa: {res.lines} | PDF: {res.pdfs} | Lỗi: {res.errors}\n\nExcel: {res.excel_path}"
                    + (f"\nPDF: {res.pdf_dir}" if res.pdfs else "")
                    + (
                        f"\nMISA: {res.misa_folder} ({res.misa_vouchers} chứng từ, file Mua_hang_khong_qua_kho_VND.xlsx)"
                        if res.misa_folder and not res.misa_error else ""
                    )
                    + (f"\n\nKHÔNG tạo được file MISA: {res.misa_error}\nXem file LOI_XUAT_MISA.txt trong thư mục misa." if res.misa_error else "")
                )
                def finish() -> None:
                    v_status.set(msg.splitlines()[0])
                    b_excel.config(state="normal")
                    messagebox.showinfo("Kết quả", msg + ("\n\nXem sheet 'Loi' trong Excel để biết hóa đơn nào lỗi." if res.errors else ""))
                ui(finish)
            except LoginError as exc:
                msg = str(exc)
                ui(lambda m=msg: (v_status.set(f"Đăng nhập thất bại: {m}"), messagebox.showerror("Đăng nhập thất bại", m)))
            except Exception as exc:  # noqa: BLE001
                msg = str(exc) or type(exc).__name__
                logger.error("Lỗi: %s", msg)
                logger.debug("Chi tiết lỗi:\n%s", traceback.format_exc())
                ui(lambda m=msg: (v_status.set(f"Lỗi: {m}"), messagebox.showerror("Lỗi", m)))
            finally:
                ui(lambda: set_busy(False))

        state["thread"] = threading.Thread(target=worker, daemon=True)
        state["thread"].start()

    def do_misa() -> None:
        from hddt.config import Settings
        from hddt.misa import export_misa

        out_dir = v_out.get().strip() or "output"
        set_busy(True, stoppable=False)
        v_status.set("Đang chuẩn bị dữ liệu cho MISA SME...")
        append("=" * 70)

        def worker() -> None:
            try:
                res = export_misa(out_dir, Settings.from_env(str(ENV_FILE)).misa_settings())
                msg = (
                    f"Đã chuẩn bị {res.vouchers} hóa đơn mua vào ({res.lines} dòng hàng) cho MISA SME.\n\n"
                    "• Mua_hang_khong_qua_kho_VND.xlsx: đúng mẫu MISA, dùng để nhập chứng từ.\n"
                    f"• MISA_DanhMuc_va_KiemTra.xlsx: {res.suppliers} nhà cung cấp, {res.items} mã hàng, bảng kiểm tra.\n"
                    f"• {res.no_xml} hóa đơn không có XML gốc (số liệu lấy từ cổng thuế).\n"
                    f"• {res.skipped} hóa đơn bị hủy/bị thay thế: không đưa vào (sheet BoQua).\n"
                    + (f"• {len(res.warnings)} hóa đơn có tổng dòng hàng lệch tổng hóa đơn: xem sheet KiemTra.\n" if res.warnings else "")
                    + "\nCác bước nhập vào MISA: xem file HUONG_DAN_NHAP_MISA.txt trong thư mục misa (sẽ mở ngay)."
                )

                def finish() -> None:
                    v_status.set(f"Đã chuẩn bị {res.vouchers} hóa đơn cho MISA: {res.folder}")
                    messagebox.showinfo("Xuất sang MISA", msg)
                    open_path(str(Path(res.folder).resolve()))
                ui(finish)
            except Exception as exc:  # noqa: BLE001
                m = str(exc) or type(exc).__name__
                logger.error("Lỗi xuất MISA: %s", m)
                ui(lambda m=m: (v_status.set(f"Lỗi: {m}"), messagebox.showerror("Lỗi xuất MISA", m)))
            finally:
                ui(lambda: set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def do_stop() -> None:
        state["stop"].set()
        v_status.set("Đang dừng... (chờ request hiện tại kết thúc, kết quả đã tải vẫn được ghi ra Excel)")

    def open_path(p: str) -> None:
        if not p:
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(p)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", p])
            else:
                subprocess.Popen(["xdg-open", p])
        except OSError as exc:
            messagebox.showerror("Không mở được", str(exc))

    def do_open_dir() -> None:
        d = Path(v_out.get().strip() or "output")
        d.mkdir(parents=True, exist_ok=True)
        open_path(str(d.resolve()))

    b_login.config(command=do_login)
    b_start.config(command=do_start)
    b_stop.config(command=do_stop)
    b_open.config(command=do_open_dir)
    b_excel.config(command=lambda: open_path(state["excel"]))
    b_misa.config(command=do_misa)

    def on_close() -> None:
        if state["thread"] and state["thread"].is_alive():
            if not messagebox.askyesno("Đang tải", "Đang tải hóa đơn. Dừng và thoát?"):
                return
            state["stop"].set()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    append("Điền MST, mật khẩu, chọn khoảng ngày rồi bấm 'Bắt đầu tải hóa đơn'.")
    append("Kết quả: file Excel (sheet HoaDon = bảng kê, ChiTiet = từng dòng hàng hóa/ĐVT/SL/đơn giá/thành tiền, Loi), thư mục pdf (mỗi hóa đơn một file) và thư mục xml.")
    root.mainloop()


def main() -> None:
    try:
        note = ensure_dependencies()
        if note:
            print(note)
        run_gui()
    except Exception:  # noqa: BLE001
        err = traceback.format_exc()
        try:
            ERROR_FILE.write_text(err, encoding="utf-8")
        except OSError:
            pass
        try:
            import tkinter as tk
            from tkinter import messagebox

            r = tk.Tk()
            r.withdraw()
            messagebox.showerror("Lỗi khởi động", f"Không mở được chương trình.\n\n{err[-1500:]}\n\nChi tiết đã ghi vào {ERROR_FILE.name}")
        except Exception:  # noqa: BLE001
            print(err)
        raise SystemExit(1)


if __name__ == "__main__":
    main()
