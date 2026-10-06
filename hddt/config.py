"""Đọc cấu hình từ file .env (không cần thư viện ngoài) và biến môi trường."""

from __future__ import annotations

import os
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path


def load_env_file(path: str | os.PathLike[str] = ".env") -> dict[str, str]:
    """Đọc file KEY=VALUE; không ghi đè biến môi trường đã có."""
    p = Path(path)
    values: dict[str, str] = {}
    if not p.is_file():
        return values
    for raw in p.read_text(encoding="utf-8-sig").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key = key.strip()
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in "\"'":
            value = value[1:-1]
        values[key] = value
        os.environ.setdefault(key, value)
    return values


def parse_vn_date(text: str) -> date:
    """Nhận dd/mm/yyyy hoặc yyyy-mm-dd."""
    t = text.strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y"):
        try:
            return datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Ngày không hợp lệ: {text!r} (dùng dạng dd/mm/yyyy)")


def _bool(value: str | None, default: bool) -> bool:
    if value is None or value == "":
        return default
    return value.strip().lower() in ("1", "true", "yes", "y", "co", "có")


@dataclass
class Settings:
    username: str = ""
    password: str = ""
    token: str = ""
    proxy_url: str = ""
    output_dir: str = "output"
    page_size: int = 50
    request_interval: float = 0.6
    xml_workers: int = 3
    timeout: float = 90.0
    verify_ssl: bool = True
    user_agent: str = ""
    misa_tk_chi_phi: str = "642"
    misa_tk_cong_no: str = "331"
    misa_tk_thue: str = "1331"
    misa_nhom_hhdv: str = "1"
    misa_phuong_thuc_tt: str = "0"

    @classmethod
    def from_env(cls, env_file: str = ".env") -> "Settings":
        load_env_file(env_file)
        g = os.environ.get
        return cls(
            username=g("GDT_USERNAME", ""),
            password=g("GDT_PASSWORD", ""),
            token=g("GDT_TOKEN", ""),
            proxy_url=g("PROXY_URL", ""),
            output_dir=g("OUTPUT_DIR", "output"),
            page_size=int(g("PAGE_SIZE", "50") or 50),
            request_interval=float(g("REQUEST_DELAY_MS", "600") or 600) / 1000.0,
            xml_workers=int(g("XML_WORKERS", "3") or 3),
            timeout=float(g("HTTP_TIMEOUT_SECONDS", "90") or 90),
            verify_ssl=_bool(g("VERIFY_SSL"), True),
            user_agent=g("BROWSER_USER_AGENT", ""),
            misa_tk_chi_phi=g("MISA_TK_CHI_PHI", "642") or "642",
            misa_tk_cong_no=g("MISA_TK_CONG_NO", "331") or "331",
            misa_tk_thue=g("MISA_TK_THUE", "1331") or "1331",
            misa_nhom_hhdv=g("MISA_NHOM_HHDV", "1") or "1",
            misa_phuong_thuc_tt=g("MISA_PHUONG_THUC_TT", "0") or "0",
        )

    def misa_settings(self):
        from .misa import MisaSettings

        return MisaSettings(
            tk_chi_phi=self.misa_tk_chi_phi, tk_cong_no=self.misa_tk_cong_no, tk_thue=self.misa_tk_thue,
            nhom_hhdv=self.misa_nhom_hhdv, phuong_thuc_tt=self.misa_phuong_thuc_tt,
        )

    def proxies(self) -> dict[str, str] | None:
        if not self.proxy_url:
            return None
        return {"http": self.proxy_url, "https": self.proxy_url}
