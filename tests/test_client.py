"""Kiểm tra client với server GDT giả lập (không gọi mạng thật)."""

import io
import json
import zipfile
from datetime import date
from unittest import mock
from urllib.parse import parse_qs, unquote, urlparse

import pytest
import requests

from hddt.captcha import _KEYWORDS
from hddt.client import GdtClient, InvoiceRef, LoginError, build_search, month_ranges, unpack_xml_payload
from tests.test_captcha import _svg


class FakeResponse:
    def __init__(self, status=200, body=b"", headers=None):
        self.status_code = status
        self.content = body if isinstance(body, bytes) else json.dumps(body).encode()
        self.headers = headers or {}

    @property
    def text(self):
        return self.content.decode("utf-8", "replace")

    def json(self):
        return json.loads(self.content)


class FakeGdt:
    """Mô phỏng các endpoint cần thiết."""

    def __init__(self):
        self.calls = []
        self.captcha_code = "A7KX2"
        self.token = "jwt-token-1"
        self.fail_captcha_once = False
        self.expire_token_once = False
        self.rate_limit_once = False
        self.invoices = {
            ("query", "purchase"): [
                {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "125", "tdlap": "2025-12-15T00:00:00", "tthai": 1, "ttxly": 5, "tgtttbso": 2740000},
                {"nbmst": "0312345678", "khmshdon": "1", "khhdon": "C25TAA", "shdon": "126", "tdlap": "2025-12-16T00:00:00", "tthai": 1, "ttxly": 5, "tgtttbso": 100},
                {"nbmst": "0399999999", "khmshdon": "2", "khhdon": "C25TBB", "shdon": "7", "tdlap": "2025-12-20T00:00:00", "tthai": 1, "ttxly": 6, "tgtttbso": 55},
            ],
            ("sco-query", "purchase"): [
                {"nbmst": "0355555555", "khmshdon": "1", "khhdon": "C25MAA", "shdon": "9", "tdlap": "2025-12-02T00:00:00", "tthai": 1, "ttxly": 8, "tgtttbso": 9},
            ],
            ("query", "sold"): [],
            ("sco-query", "sold"): [],
        }

    def request(self, method, url, headers=None, json=None, timeout=None):
        self.calls.append((method, url, headers, json))
        u = urlparse(url)
        q = parse_qs(u.query)
        path = u.path
        if path == "/api/captcha":
            return FakeResponse(200, {"key": "ckey-1", "content": _svg(self.captcha_code)})
        if path == "/api/security-taxpayer/authenticate":
            assert headers["Origin"] == "https://hoadondientu.gdt.gov.vn"
            if self.fail_captcha_once:
                self.fail_captcha_once = False
                return FakeResponse(200, {"message": "Captcha không hợp lệ"})
            if json["cvalue"] != self.captcha_code or json["ckey"] != "ckey-1":
                return FakeResponse(200, {"message": "Captcha không hợp lệ"})
            if json["password"] != "secret":
                return FakeResponse(401, {"message": "Sai thông tin đăng nhập"})
            return FakeResponse(200, {"token": self.token})
        # Các endpoint cần token
        auth = (headers or {}).get("Authorization", "")
        if self.expire_token_once:
            self.expire_token_once = False
            return FakeResponse(401, {"message": "expired"})
        if auth != "Bearer " + self.token:
            return FakeResponse(401, {"message": "unauthorized"})
        if self.rate_limit_once:
            self.rate_limit_once = False
            return FakeResponse(429, b"slow down", {"Retry-After": "1"})
        if path == "/api/security-taxpayer/profile":
            return FakeResponse(200, {"name": "CÔNG TY CP XYZ", "tin": "0109876543"})
        parts = path.split("/")  # ['', 'api', family, 'invoices', endpoint]
        family, endpoint = parts[2], parts[4]
        if endpoint in ("purchase", "sold"):
            assert q["size"] == ["2"]
            search = q["search"][0]
            assert search.startswith("tdlap=ge=01/12/2025T00:00:00;tdlap=le=31/12/2025T23:59:59"), search
            items = self.invoices[(family, endpoint)]
            if "ttxly==" in search:
                want = int(search.split("ttxly==")[1].split(";")[0])
                items = [i for i in items if i["ttxly"] == want]
            start = int(q.get("state", ["0"])[0])
            page = items[start : start + 2]
            nxt = start + 2
            state = str(nxt) if nxt < len(items) else ""
            return FakeResponse(200, {"datas": page, "state": state, "total": len(items)})
        if endpoint == "export-xml":
            if q["shdon"] == ["126"]:
                return FakeResponse(500, {"message": "no record"})
            buf = io.BytesIO()
            with zipfile.ZipFile(buf, "w") as zf:
                zf.writestr("invoice.xml", f'<?xml version="1.0"?><HDon><DLHDon><TTChung><SHDon>{q["shdon"][0]}</SHDon></TTChung></DLHDon></HDon>')
                zf.writestr("invoice.html", "<html></html>")
            return FakeResponse(200, buf.getvalue())
        if endpoint == "detail":
            return FakeResponse(200, {"shdon": q["shdon"][0], "hdhhdvu": []})
        return FakeResponse(404, {"message": "not found"})


@pytest.fixture
def fake():
    return FakeGdt()


@pytest.fixture
def client(fake):
    session = mock.Mock(spec=requests.Session)
    session.request.side_effect = fake.request
    session.proxies = {}
    return GdtClient("0109876543", "secret", session=session, request_interval=0, sleep=lambda s: None)


def test_login_solves_captcha_and_gets_token(client, fake):
    assert client.login() == "jwt-token-1"
    auth_call = [c for c in fake.calls if c[1].endswith("/authenticate")][0]
    assert auth_call[3] == {"username": "0109876543", "password": "secret", "cvalue": "A7KX2", "ckey": "ckey-1"}
    assert client.get_profile()["tin"] == "0109876543"


def test_login_retries_on_wrong_captcha(client, fake):
    fake.fail_captcha_once = True
    assert client.login() == "jwt-token-1"
    assert sum(1 for c in fake.calls if c[1].endswith("/authenticate")) == 2


def test_login_wrong_password_raises(fake):
    session = mock.Mock(spec=requests.Session)
    session.request.side_effect = fake.request
    session.proxies = {}
    c = GdtClient("0109876543", "wrong", session=session, request_interval=0, sleep=lambda s: None)
    with pytest.raises(LoginError):
        c.login()


def test_list_invoices_paginates_by_state_over_both_families(client, fake):
    refs = client.list_invoices("purchase", date(2025, 12, 1), date(2025, 12, 31), page_size=2)
    assert [(r.family, r.shdon) for r in refs] == [("query", "125"), ("query", "126"), ("query", "7"), ("sco-query", "9")]
    list_calls = [c for c in fake.calls if "/invoices/purchase" in c[1]]
    assert len(list_calls) == 3  # query: 2 trang, sco-query: 1 trang
    assert all(c[2]["Authorization"] == "Bearer jwt-token-1" for c in list_calls)
    assert "state=2" in list_calls[1][1]


def test_list_invoices_ttxly_filter(client):
    refs = client.list_invoices("purchase", date(2025, 12, 1), date(2025, 12, 31), page_size=2, ttxly=6, families=["query"])
    assert [r.shdon for r in refs] == ["7"]


def test_relogin_on_401_and_retry_on_429(client, fake):
    client.login()
    fake.expire_token_once = True
    fake.rate_limit_once = True
    refs = client.list_invoices("purchase", date(2025, 12, 1), date(2025, 12, 31), page_size=2, families=["query"])
    assert len(refs) == 3
    assert sum(1 for c in fake.calls if c[1].endswith("/authenticate")) == 2


def test_download_xml_unpacks_zip(client):
    ref = InvoiceRef("purchase", "query", "0312345678", "1", "C25TAA", "125")
    files = client.download_xml(ref)
    names = [n for n, _ in files]
    assert names == ["purchase_query_0312345678_1_C25TAA_125.xml", "purchase_query_0312345678_1_C25TAA_125.html"]
    assert b"<SHDon>125</SHDon>" in files[0][1]


def test_download_xml_missing_record_raises_500(client):
    from hddt.client import ApiError

    ref = InvoiceRef("purchase", "query", "0312345678", "1", "C25TAA", "126")
    with pytest.raises(ApiError) as exc:
        client.download_xml(ref)
    assert exc.value.status == 500


def test_unpack_raw_xml_payload():
    out = unpack_xml_payload(b'<?xml version="1.0"?><HDon/>', "x")
    assert out == [("x.xml", b'<?xml version="1.0"?><HDon/>')]


def test_month_ranges_and_search():
    rs = list(month_ranges(date(2025, 11, 15), date(2026, 1, 10)))
    assert rs == [(date(2025, 11, 15), date(2025, 11, 30)), (date(2025, 12, 1), date(2025, 12, 31)), (date(2026, 1, 1), date(2026, 1, 10))]
    assert build_search(date(2025, 12, 1), date(2025, 12, 31), 8) == "tdlap=ge=01/12/2025T00:00:00;tdlap=le=31/12/2025T23:59:59;ttxly==8"
