# hddt – Tool kéo hóa đơn điện tử từ Cơ quan Thuế

Tool dòng lệnh (Python) tự động đăng nhập **hoadondientu.gdt.gov.vn**, kéo danh sách
hóa đơn **mua vào / bán ra** (cả hóa đơn thường và hóa đơn **máy tính tiền**), tải
**XML gốc** từng hóa đơn và xuất ra **Excel** (bảng kê + chi tiết hàng hóa + bảng lỗi).

Cách làm giống các tool tải hóa đơn đang lưu hành (gọi thẳng API mà web của Tổng cục
Thuế dùng, tự đọc CAPTCHA SVG, phân trang theo `state`, tải XML qua `export-xml`).

## 1. Cài đặt

Cần Python 3.9 trở lên (Windows: tải tại python.org, nhớ tick *Add Python to PATH*).

```bash
pip install -r requirements.txt
```

Sao chép `.env.example` thành `.env` và điền tài khoản tra cứu hóa đơn do CQT cấp:

```
GDT_USERNAME=0101234567        # MST / tài khoản đăng nhập trang hoadondientu
GDT_PASSWORD=matkhau
```

> Không có mật khẩu? Có thể đăng nhập trang web bằng trình duyệt, bấm F12 → Network,
> copy giá trị header `Authorization: Bearer ...` dán vào `GDT_TOKEN=` (token sống
> khoảng vài giờ).

## 2. Sử dụng

```bash
# Kiểm tra đăng nhập (tool tự giải CAPTCHA)
python -m hddt login

# Kéo toàn bộ hóa đơn mua vào + bán ra tháng 12/2025: danh sách + XML + Excel
python -m hddt pull --thang 12/2025

# Chỉ hóa đơn mua vào, theo khoảng ngày, 5 luồng tải XML song song
python -m hddt pull --chieu mua --tu-ngay 01/10/2025 --den-ngay 31/12/2025 --luong 5

# Chỉ lấy danh sách (nhanh), không tải XML
python -m hddt pull --thang 12/2025 --khong-xml

# Bỏ qua hóa đơn máy tính tiền / chỉ lấy hóa đơn máy tính tiền
python -m hddt pull --thang 12/2025 --khong-mtt
python -m hddt pull --thang 12/2025 --chi-mtt

# Tải file Excel do chính GDT xuất (giống nút "Xuất Excel" trên web)
python -m hddt excel-gdt --thang 12/2025 --chieu mua

# Đọc lại thư mục XML đã tải (offline) ra Excel
python -m hddt parse-xml output/xml --excel output/HoaDon_Local.xlsx

# Kiểm tra bộ giải CAPTCHA (lấy 5 CAPTCHA, in mã đọc được, lưu SVG để đối chiếu)
python -m hddt captcha --so-lan 5 --luu output/captcha
```

Kết quả nằm trong thư mục `output/` (đổi bằng `--thu-muc` hoặc `OUTPUT_DIR`):

```
output/
  HoaDon_20251201_20251231.xlsx   # Sheet HoaDon (bảng kê), ChiTiet (dòng hàng), Loi
  danh_sach_20251201_20251231.json# JSON thô GDT trả về (để đối chiếu)
  xml/purchase/<chieu>_<nguon>_<MST bán>_<mẫu số>_<ký hiệu>_<số>.xml (+ .html bản thể hiện)
  xml/sold/...
```

Chạy lại cùng khoảng ngày sẽ **dùng lại XML đã có**, chỉ tải phần còn thiếu
(thêm `--tai-lai` để tải lại toàn bộ). Mã thoát `2` nghĩa là có hóa đơn lỗi, xem sheet `Loi`.

## 3. Tham số hay dùng

| Tham số | Ý nghĩa |
|---|---|
| `--chieu mua\|ban\|ca-hai` | Hóa đơn mua vào, bán ra hoặc cả hai (mặc định cả hai) |
| `--thang mm/yyyy` | Lấy trọn tháng |
| `--tu-ngay / --den-ngay dd/mm/yyyy` | Khoảng ngày lập hóa đơn (tool tự chia theo tháng khi gọi GDT) |
| `--ttxly 5\|6\|8` | Lọc kết quả kiểm tra: 5 đã cấp mã, 6 không mã, 8 máy tính tiền. Mặc định lấy tất cả |
| `--luong N` | Số luồng tải XML song song (1–10) |
| `--token ...` | Dùng token sẵn, bỏ qua đăng nhập |
| `-v` / `--log-file` | Log chi tiết / ghi log ra file |

Các biến trong `.env`: `PAGE_SIZE` (50), `REQUEST_DELAY_MS` (600), `XML_WORKERS` (3),
`HTTP_TIMEOUT_SECONDS` (90), `PROXY_URL`, `BROWSER_USER_AGENT`, `VERIFY_SSL`.

## 4. Cơ chế hoạt động (tóm tắt kỹ thuật)

| Bước | Endpoint GDT |
|---|---|
| Lấy CAPTCHA | `GET /api/captcha` → `{key, content: <svg>}` – mỗi ký tự là một `<path>`; rút gọn lệnh vẽ về chuỗi `M/Q/Z`, tra bảng font để ra ký tự, sắp theo tọa độ x |
| Đăng nhập | `POST /api/security-taxpayer/authenticate` `{username, password, ckey, cvalue}` → `{token}` (JWT, giữ trong RAM) |
| Danh sách | `GET /api/{query\|sco-query}/invoices/{purchase\|sold}?sort=tdlap:desc&size=50&search=tdlap=ge=dd/MM/yyyyT00:00:00;tdlap=le=...&state=...` – lặp tới khi `state` rỗng |
| XML | `GET /api/{query\|sco-query}/invoices/export-xml?nbmst=&khhdon=&shdon=&khmshdon=` → ZIP chứa `.xml` + `.html` |
| Chi tiết / liên quan | `.../invoices/detail`, `.../invoices/related`, `.../invoices/relative` (có sẵn trong `GdtClient`) |

Xử lý lỗi: HTTP 429 nghỉ theo `Retry-After` (mặc định 15 s, tăng dần) rồi tự giảm tốc;
401/403 giữa chừng → tự đăng nhập lại; 5xx thử lại tối đa 4 lần (riêng `export-xml`
trả 500 nghĩa là GDT không có hồ sơ XML, ghi vào sheet `Loi`).

## 5. Dùng như thư viện

```python
from datetime import date
from hddt.client import GdtClient
from hddt.xmlparse import parse_invoice_xml

c = GdtClient("0101234567", "matkhau")
for ref in c.iter_invoices("purchase", date(2025, 12, 1), date(2025, 12, 31)):
    files = c.download_xml(ref)              # [(tên file, bytes), ...]
    inv = parse_invoice_xml(files[0][1])     # ParsedInvoice: tổng tiền, người bán, dòng hàng...
    print(ref.label, inv.tong_tien_tt)
```

## 6. Kiểm thử

```bash
pip install pytest
python -m pytest -q
```

Bộ test chạy hoàn toàn offline với server GDT giả lập (đăng nhập + CAPTCHA, phân trang,
tải XML, hết hạn token, 429, xuất Excel).

## Lưu ý

- Tool chỉ dùng tài khoản của chính đơn vị, dữ liệu tải về là dữ liệu đơn vị được
  phép tra cứu trên cổng của Tổng cục Thuế. Không chia sẻ file `.env`, token hay log.
- GDT có thể đổi API/CAPTCHA bất kỳ lúc nào; khi đó chạy `python -m hddt captcha --luu ...`
  để lấy mẫu kiểm tra và cập nhật bảng trong `hddt/captcha.py`.
