"""Client gọi API của hệ thống hóa đơn điện tử Tổng cục Thuế.

Luồng hoạt động (giống cách các tool tải hóa đơn phổ biến đang làm):

1. ``GET  /api/captcha``                        -> {"key", "content": "<svg>"}; tự giải SVG.
2. ``POST /api/security-taxpayer/authenticate`` -> {"token": JWT}.
3. ``GET  /api/{query|sco-query}/invoices/{purchase|sold}?sort=tdlap:desc&size=50&search=...&state=...``
   - ``query``     : hóa đơn điện tử thông thường.
   - ``sco-query`` : hóa đơn khởi tạo từ máy tính tiền.
   - ``search``    : ``tdlap=ge=dd/MM/yyyyT00:00:00;tdlap=le=dd/MM/yyyyT23:59:59[;ttxly==5]``
   - phân trang bằng ``state`` do server trả về (không có số trang).
4. ``GET  /api/{family}/invoices/export-xml?nbmst=&khhdon=&shdon=&khmshdon=`` -> ZIP chứa XML (+HTML).
5. ``GET  /api/{family}/invoices/{detail|related|relative}?...`` -> chi tiết / thông tin liên quan.

Token chỉ giữ trong bộ nhớ. Gặp 401/403 giữa chừng thì tự đăng nhập lại một lần.
Gặp 429 thì nghỉ theo ``Retry-After`` (mặc định 15s, tăng dần) rồi thử lại.
"""

from __future__ import annotations

import io
import logging
import threading
import time
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import date, timedelta
from typing import Any, Callable, Iterable, Iterator
from urllib.parse import quote

import requests

from .captcha import solve_svg_captcha

log = logging.getLogger("hddt")

BASE_URL = "https://hoadondientu.gdt.gov.vn"
API_URL = BASE_URL + "/api"
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36 Edg/124.0.0.0"
)

FAMILIES = ("query", "sco-query")
DIRECTIONS = ("purchase", "sold")

# Nhãn tiếng Việt cho các mã trạng thái GDT trả về trong danh sách.
TTHAI_LABELS = {
    1: "Hóa đơn mới",
    2: "Hóa đơn thay thế",
    3: "Hóa đơn điều chỉnh",
    4: "Hóa đơn bị thay thế",
    5: "Hóa đơn đã bị điều chỉnh",
    6: "Hóa đơn bị hủy",
}
TTXLY_LABELS = {
    0: "Tổng cục thuế đã nhận",
    1: "Đang tiến hành kiểm tra điều kiện cấp mã",
    2: "CQT từ chối theo từng lần phát sinh",
    3: "Hóa đơn đủ điều kiện cấp mã",
    4: "Hóa đơn không đủ điều kiện cấp mã",
    5: "Đã cấp mã hóa đơn",
    6: "Tổng cục thuế đã nhận không mã",
    7: "Đã kiểm tra HĐĐT định kỳ không có mã",
    8: "Tổng cục thuế đã nhận hóa đơn có mã khởi tạo từ máy tính tiền",
}


class HddtError(Exception):
    """Lỗi chung của tool."""


class LoginError(HddtError):
    """Đăng nhập thất bại (sai tài khoản/mật khẩu, CAPTCHA, ...)."""


class StopRequested(HddtError):
    """Người dùng yêu cầu dừng."""


class ApiError(HddtError):
    def __init__(self, message: str, status: int | None = None, body: str = ""):
        super().__init__(message)
        self.status = status
        self.body = body


@dataclass
class InvoiceRef:
    """Khóa định danh một hóa đơn trên GDT."""

    direction: str  # purchase | sold
    family: str  # query | sco-query
    nbmst: str  # MST người bán
    khmshdon: str  # ký hiệu mẫu số
    khhdon: str  # ký hiệu hóa đơn
    shdon: str  # số hóa đơn
    raw: dict[str, Any] = field(default_factory=dict)

    @property
    def label(self) -> str:
        return f"{self.nbmst} {self.khmshdon}{self.khhdon} {self.shdon}"

    @property
    def file_stem(self) -> str:
        return "_".join(
            _safe(part) for part in (self.direction, self.family, self.nbmst, self.khmshdon, self.khhdon, self.shdon)
        )

    @classmethod
    def from_json(cls, direction: str, family: str, data: dict[str, Any]) -> "InvoiceRef":
        return cls(
            direction=direction,
            family=family,
            nbmst=str(data.get("nbmst") or ""),
            khmshdon=str(data.get("khmshdon") or ""),
            khhdon=str(data.get("khhdon") or ""),
            shdon=str(data.get("shdon") or ""),
            raw=data,
        )

    def query_string(self) -> str:
        return (
            f"nbmst={quote(self.nbmst)}&khhdon={quote(self.khhdon)}"
            f"&shdon={quote(self.shdon)}&khmshdon={quote(self.khmshdon)}"
        )


def _safe(text: str) -> str:
    return "".join(c if c.isalnum() or c in "-_." else "_" for c in str(text))


def month_ranges(start: date, end: date) -> Iterator[tuple[date, date]]:
    """Chia khoảng ngày thành từng tháng (GDT giới hạn số bản ghi mỗi lần truy vấn)."""
    cur = start
    while cur <= end:
        nxt = (cur.replace(day=1) + timedelta(days=32)).replace(day=1)
        last = min(nxt - timedelta(days=1), end)
        yield cur, last
        cur = last + timedelta(days=1)


def build_search(start: date, end: date, ttxly: int | None = None, extra: str = "") -> str:
    s = f"tdlap=ge={start:%d/%m/%Y}T00:00:00;tdlap=le={end:%d/%m/%Y}T23:59:59"
    if ttxly is not None:
        s += f";ttxly=={ttxly}"
    if extra:
        s += ";" + extra.strip(";")
    return s


class RateLimiter:
    """Giãn cách tối thiểu giữa hai request, dùng chung cho nhiều luồng."""

    def __init__(self, interval: float):
        self.interval = max(0.0, interval)
        self._lock = threading.Lock()
        self._next = 0.0
        self._pause_until = 0.0

    def wait(self) -> None:
        while True:
            with self._lock:
                now = time.monotonic()
                target = max(self._next, self._pause_until)
                if now >= target:
                    self._next = now + self.interval
                    return
                sleep_for = target - now
            time.sleep(min(sleep_for, 0.25))

    def pause(self, seconds: float) -> None:
        with self._lock:
            self._pause_until = max(self._pause_until, time.monotonic() + seconds)
            # Sau mỗi lần bị giới hạn tốc độ thì nới giãn cách thêm 50%.
            self.interval = min(self.interval * 1.5 if self.interval else 0.5, 10.0)


class GdtClient:
    def __init__(
        self,
        username: str = "",
        password: str = "",
        token: str = "",
        *,
        user_agent: str = DEFAULT_USER_AGENT,
        proxies: dict[str, str] | None = None,
        timeout: float = 90.0,
        request_interval: float = 0.6,
        max_retries: int = 4,
        verify_ssl: bool = True,
        session: requests.Session | None = None,
        sleep: Callable[[float], None] = time.sleep,
    ):
        self.username = username.strip()
        self.password = password
        self.token = token.strip()
        self.user_agent = user_agent
        self.timeout = timeout
        self.max_retries = max_retries
        self.limiter = RateLimiter(request_interval)
        self._sleep = sleep
        self._login_lock = threading.Lock()
        self.session = session or requests.Session()
        if proxies:
            self.session.proxies.update(proxies)
        self.session.verify = verify_ssl
        self.profile: dict[str, Any] | None = None
        self.stop_event = threading.Event()

    def _check_stop(self) -> None:
        if self.stop_event.is_set():
            raise StopRequested("Đã dừng theo yêu cầu.")

    # ------------------------------------------------------------------ HTTP
    def _headers(self, *, auth: bool, referer: str, accept: str, action: str = "", origin: bool = False) -> dict[str, str]:
        h = {
            "User-Agent": self.user_agent,
            "Accept": accept,
            "Accept-Language": "vi",
            "Referer": referer,
            "Request-Id": str(uuid.uuid4()),
            "Sec-Fetch-Site": "same-origin",
            "Sec-Fetch-Mode": "cors",
            "Sec-Fetch-Dest": "empty",
        }
        if origin:
            h["Origin"] = BASE_URL
        if action:
            h["Action"] = quote(action)
            h["End-Point"] = "/tra-cuu/tra-cuu-hoa-don"
        if auth:
            if not self.token:
                self.login()
            h["Authorization"] = "Bearer " + self.token
        return h

    def _request(
        self,
        method: str,
        url: str,
        *,
        auth: bool = True,
        referer: str = BASE_URL + "/tra-cuu/tra-cuu-hoa-don",
        accept: str = "application/json, text/plain, */*",
        action: str = "",
        origin: bool = False,
        json_body: Any = None,
        allow_relogin: bool = True,
        retry_5xx: bool = True,
    ) -> requests.Response:
        attempt = 0
        rate_limit_hits = 0
        relogged = False
        while True:
            self._check_stop()
            attempt += 1
            self.limiter.wait()
            self._check_stop()
            headers = self._headers(auth=auth, referer=referer, accept=accept, action=action, origin=origin)
            try:
                resp = self.session.request(method, url, headers=headers, json=json_body, timeout=self.timeout)
            except requests.RequestException as exc:
                if attempt > self.max_retries:
                    raise ApiError(f"Lỗi mạng khi gọi {url}: {exc}") from exc
                wait = min(2.0 * attempt, 10.0)
                log.warning("Lỗi mạng (%s), thử lại sau %.0fs (%d/%d)", exc, wait, attempt, self.max_retries)
                self._sleep(wait)
                continue

            if resp.status_code == 429:
                rate_limit_hits += 1
                if rate_limit_hits > 12:
                    raise ApiError("GDT giới hạn tốc độ (HTTP 429) quá nhiều lần.", 429, resp.text[:300])
                wait = _retry_after(resp) or min(15.0 * (2 ** (rate_limit_hits - 1)), 600.0)
                log.warning("GDT trả 429 (giới hạn tốc độ); nghỉ %.0fs rồi thử lại.", wait)
                self.limiter.pause(wait)
                continue

            if resp.status_code in (401, 403) and auth and allow_relogin and not relogged and self.username and self.password:
                log.warning("Token hết hạn/bị từ chối (HTTP %s); đăng nhập lại.", resp.status_code)
                relogged = True
                self.token = ""
                self.login()
                continue

            if resp.status_code >= 500 and retry_5xx and attempt <= self.max_retries:
                wait = min(3.0 * attempt, 15.0)
                log.warning("GDT trả HTTP %s; thử lại sau %.0fs (%d/%d)", resp.status_code, wait, attempt, self.max_retries)
                self._sleep(wait)
                continue

            if resp.status_code >= 400:
                raise ApiError(f"HTTP {resp.status_code} khi gọi {url}", resp.status_code, resp.text[:500])
            return resp

    # ----------------------------------------------------------------- Login
    def get_captcha(self) -> tuple[str, str, str]:
        """Trả về (key, mã đã giải, svg gốc)."""
        resp = self._request("GET", API_URL + "/captcha", auth=False, referer=BASE_URL + "/")
        try:
            payload = resp.json()
        except ValueError as exc:
            raise ApiError("Phản hồi CAPTCHA không phải JSON.", resp.status_code, resp.text[:300]) from exc
        key = str(payload.get("key") or "")
        content = str(payload.get("content") or "")
        return key, solve_svg_captcha(content or resp.text), content

    def login(self, max_attempts: int = 5) -> str:
        """Đăng nhập bằng MST/mật khẩu; tự giải CAPTCHA. Trả về JWT token."""
        with self._login_lock:
            if self.token:
                return self.token
            if not self.username or not self.password:
                raise LoginError("Chưa có tài khoản GDT (GDT_USERNAME / GDT_PASSWORD) và cũng không có token.")
            last = ""
            for i in range(1, max_attempts + 1):
                key, code, _ = self.get_captcha()
                if not key:
                    last = "CAPTCHA không có key"
                    continue
                if len(code) < 4:
                    last = f"không nhận diện được CAPTCHA (đọc được '{code}')"
                    log.info("Đăng nhập %d/%d: %s; lấy CAPTCHA khác.", i, max_attempts, last)
                    continue
                log.info("Đăng nhập %d/%d: tài khoản %s, CAPTCHA tự đọc: %s", i, max_attempts, self.username, code)
                body = {"username": self.username, "password": self.password, "cvalue": code, "ckey": key}
                try:
                    resp = self._request(
                        "POST",
                        API_URL + "/security-taxpayer/authenticate",
                        auth=False,
                        referer=BASE_URL + "/",
                        origin=True,
                        json_body=body,
                    )
                except ApiError as exc:
                    if exc.status in (400, 401, 403):
                        msg = _extract_message(exc.body)
                        if "captcha" in msg.lower():
                            last = msg
                            log.info("CAPTCHA sai, thử lại: %s", msg)
                            continue
                        raise LoginError(f"Đăng nhập thất bại: {msg or exc}") from exc
                    raise
                payload = resp.json()
                message = str(payload.get("message") or "")
                if message:
                    if "captcha" in message.lower():
                        last = message
                        continue
                    raise LoginError(f"Đăng nhập thất bại: {message}")
                token = str(payload.get("token") or "")
                if not token:
                    raise LoginError("Phản hồi đăng nhập không có token.")
                self.token = token
                log.info("Đăng nhập thành công.")
                return token
            raise LoginError(f"Đăng nhập thất bại sau {max_attempts} lần: {last}")

    def get_profile(self) -> dict[str, Any]:
        url = API_URL + "/security-taxpayer/profile"
        if self.username:
            url += "?smiUsername=" + quote(self.username)
        self.profile = self._request("GET", url, referer=BASE_URL + "/").json()
        return self.profile

    # ------------------------------------------------------------ Danh sách
    def iter_invoices(
        self,
        direction: str,
        start: date,
        end: date,
        *,
        families: Iterable[str] = FAMILIES,
        ttxly: int | None = None,
        extra_search: str = "",
        page_size: int = 50,
        on_page: Callable[[str, str, date, date, int, int], None] | None = None,
    ) -> Iterator[InvoiceRef]:
        """Duyệt toàn bộ hóa đơn mua vào (purchase) hoặc bán ra (sold) trong khoảng ngày."""
        if direction not in DIRECTIONS:
            raise ValueError(f"direction phải là purchase|sold, nhận {direction!r}")
        action = "Tìm kiếm"
        for family in families:
            for m_start, m_end in month_ranges(start, end):
                search = build_search(m_start, m_end, ttxly, extra_search)
                state = ""
                seen_states: set[str] = set()
                page = 0
                total_in_period = 0
                while True:
                    page += 1
                    url = (
                        f"{API_URL}/{family}/invoices/{direction}"
                        f"?sort=tdlap%3Adesc%2Ckhmshdon%3Aasc%2Cshdon%3Adesc&size={page_size}&search={quote(search)}"
                    )
                    if state:
                        url += "&state=" + quote(state)
                    payload = self._request("GET", url, action=action).json()
                    datas = payload.get("datas")
                    if datas is None:
                        raise ApiError("Phản hồi danh sách không có trường 'datas'.", body=str(payload)[:300])
                    for item in datas:
                        if isinstance(item, dict):
                            total_in_period += 1
                            yield InvoiceRef.from_json(direction, family, item)
                    if on_page:
                        on_page(direction, family, m_start, m_end, page, len(datas))
                    state = str(payload.get("state") or "").strip()
                    if not state or state in seen_states:
                        break
                    seen_states.add(state)
                log.info(
                    "[DANH SÁCH] %s/%s kỳ %s-%s: %d hóa đơn (%d trang).",
                    direction, family, f"{m_start:%d/%m/%Y}", f"{m_end:%d/%m/%Y}", total_in_period, page,
                )

    def list_invoices(self, direction: str, start: date, end: date, **kw: Any) -> list[InvoiceRef]:
        return list(self.iter_invoices(direction, start, end, **kw))

    # -------------------------------------------------------------- Chi tiết
    def _direction_text(self, ref: InvoiceRef) -> str:
        src = "hóa đơn máy tính tiền" if ref.family == "sco-query" else "hóa đơn"
        d = "mua vào" if ref.direction == "purchase" else "bán ra"
        return f"({src} {d})"

    def get_detail(self, ref: InvoiceRef) -> dict[str, Any]:
        url = f"{API_URL}/{ref.family}/invoices/detail?{ref.query_string()}"
        return self._request("GET", url, action=f"Xem hóa đơn {self._direction_text(ref)}").json()

    def get_related(self, ref: InvoiceRef) -> Any:
        """Thông tin liên quan (thông báo sai sót, ...)."""
        url = f"{API_URL}/{ref.family}/invoices/related?{ref.query_string()}"
        return self._request("GET", url, action=f"Xem thông tin liên quan {self._direction_text(ref)}").json()

    def get_relative(self, ref: InvoiceRef) -> Any:
        """Chuỗi hóa đơn thay thế / điều chỉnh (tthai 2-5)."""
        url = f"{API_URL}/{ref.family}/invoices/relative?{ref.query_string()}"
        return self._request("GET", url, action=f"Xem hóa đơn liên quan {self._direction_text(ref)}").json()

    # ------------------------------------------------------------------- XML
    def download_xml(self, ref: InvoiceRef) -> list[tuple[str, bytes]]:
        """Tải gói XML của hóa đơn. Trả về danh sách (tên file, nội dung) các file trong gói.

        GDT trả về ZIP (thường gồm invoice.xml và bản thể hiện .html). Một số
        trường hợp trả thẳng XML. Hóa đơn chưa có hồ sơ gốc sẽ bị HTTP 500.
        """
        url = f"{API_URL}/{ref.family}/invoices/export-xml?{ref.query_string()}"
        resp = self._request(
            "GET",
            url,
            accept="application/zip, application/xml, application/octet-stream, */*",
            action=f"Xuất xml {self._direction_text(ref)}",
            # export-xml trả HTTP 500 khi hóa đơn không có hồ sơ gốc: thử lại vô ích.
            retry_5xx=False,
        )
        return unpack_xml_payload(resp.content, ref.file_stem)

    def download_excel(self, direction: str, start: date, end: date, family: str = "query", ttxly: int | None = None) -> bytes:
        """Tải file Excel do chính GDT xuất (nút 'Xuất Excel' trên web)."""
        search = build_search(start, end, ttxly)
        endpoint = "export-excel" if direction == "purchase" else "export-excel-sold"
        url = (
            f"{API_URL}/{family}/invoices/{endpoint}"
            f"?sort=tdlap%3Adesc%2Ckhmshdon%3Aasc%2Cshdon%3Adesc&search={quote(search)}"
        )
        if direction == "sold":
            url += "&type=purchase"
        d = "mua vào" if direction == "purchase" else "bán ra"
        resp = self._request("GET", url, accept="*/*", action=f"Xuất hóa đơn (hóa đơn {d})")
        return resp.content


def unpack_xml_payload(content: bytes, stem: str) -> list[tuple[str, bytes]]:
    """Giải nén phản hồi export-xml thành danh sách (tên file, bytes)."""
    if not content:
        raise ApiError("Phản hồi export-xml rỗng.")
    if content[:2] == b"PK":
        out: list[tuple[str, bytes]] = []
        with zipfile.ZipFile(io.BytesIO(content)) as zf:
            for info in zf.infolist():
                if info.is_dir():
                    continue
                name = info.filename.rsplit("/", 1)[-1]
                ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
                if ext not in ("xml", "html", "htm", "pdf"):
                    continue
                data = zf.read(info)
                if ext == "xml" and b"<" not in data[:200]:
                    continue
                out.append((f"{stem}.{ext}" if ext != "xml" else f"{stem}.xml", data))
        # Nếu có nhiều XML thì đánh số để không ghi đè.
        xml_count = sum(1 for n, _ in out if n.endswith(".xml"))
        if xml_count > 1:
            idx = 0
            renamed = []
            for name, data in out:
                if name.endswith(".xml"):
                    idx += 1
                    name = f"{stem}_{idx}.xml"
                renamed.append((name, data))
            out = renamed
        if not any(n.endswith(".xml") for n, _ in out):
            raise ApiError("Gói ZIP không chứa file XML.")
        return out
    head = content.lstrip()[:200]
    if head.startswith(b"<?xml") or head.startswith(b"<"):
        return [(f"{stem}.xml", content)]
    text = content[:300].decode("utf-8", "replace")
    raise ApiError(f"Phản hồi export-xml không phải ZIP/XML: {text}")


def _retry_after(resp: requests.Response) -> float | None:
    value = resp.headers.get("Retry-After")
    if not value:
        return None
    try:
        return max(1.0, float(value))
    except ValueError:
        return None


def _extract_message(body: str) -> str:
    if not body:
        return ""
    try:
        import json

        data = json.loads(body)
        if isinstance(data, dict):
            return str(data.get("message") or data.get("error") or body)
    except ValueError:
        pass
    return body
