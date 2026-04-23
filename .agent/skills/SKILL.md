# Skill: Tự động hóa Zalo Web - Tìm kiếm Theo Tên (Relative Search)

Bộ skill này là hướng dẫn chuẩn để vận hành công cụ tự động hóa nhắn tin Zalo Web theo tên. Script được thiết kế để xử lý danh sách "Cá nhân" trong kết quả tìm kiếm, tự động loại bỏ "Nhóm", và hỗ trợ gửi cả văn bản + hình ảnh.

---

## 1. Yêu Cầu Cài Đặt Bắt Buộc (Setup)

Trước khi chạy, AI và Người dùng PHẢI đảm bảo các bước sau đã hoàn tất:

### Bước 1: Môi trường Python & Thư viện
Mở Terminal tại thư mục gốc `f:\SELENIUM_ZALO_LIST` và chạy các lệnh sau:
```powershell
# Tạo môi trường ảo
python -m venv venv

# Kích hoạt môi trường ảo
.\venv\Scripts\Activate.ps1

# Cài đặt thư viện cần thiết
pip install selenium gspread oauth2client
```

### Bước 2: Cấu hình File
- **Credentials**: File `gen-lang-client-*.json` phải nằm ở thư mục gốc.
- **Ảnh**: Các file ảnh phải nằm trong thư mục `images/` (ví dụ: `hoc-python.png`).
- **Profile**: Thư mục `zalo-chrome-profile/` sẽ tự động lưu phiên đăng nhập Zalo sau lần quét QR đầu tiên.

---

## 2. Quy Tắc Vận Hành Của AI (Execution Rule)

Khi người dùng yêu cầu "chạy file", "gửi tin nhắn", hoặc bất kỳ yêu cầu nào tương đương, AI **BẮT BUỘC** phải tự động thực thi lệnh sau mà **KHÔNG CẦN HỎI LẠI**:

```powershell
.\venv\Scripts\python.exe OpenZaloSendListRelative.py
```

AI cần lưu ý:
1. Script được thiết kế để chạy **1 dòng dữ liệu mỗi lần**.
2. Sau khi chạy xong 1 dòng, script sẽ tự đóng trình duyệt. AI có thể yêu cầu người dùng chạy tiếp nếu muốn xử lý dòng tiếp theo.
3. Nếu Zalo yêu cầu quét mã QR, hãy báo người dùng thực hiện trên màn hình trình duyệt hiện ra.

---

## 3. Cấu Trúc Google Sheet (Worksheet: "Danh Sách Theo Tên")

Script đọc dữ liệu từ sheet theo cấu trúc sau:

| Cột | Mô tả |
|-----|-------|
| `Name` | Từ khóa tìm kiếm (Ví dụ: "Hải", "Nam"). |
| `Tiêu đề` | Tiêu đề tin nhắn. |
| `Nội Dung` | Nội dung chính. |
| `Hình ảnh` | Tên file ảnh (không cần đuôi .png/.jpg). |
| `Status` | `UNAPPROVED` (Mới), `ĐÃ GỞI LẦN 1` (Đã gửi 4 người), `APPROVED` (Xong). |

---

## 4. Logic Xử Lý Kỹ Thuật (Dành cho AI)

- **Root Script**: File thực thi chính nằm tại `f:\SELENIUM_ZALO_LIST\OpenZaloSendListRelative.py`.
- **Editor ID**: Luôn sử dụng `richInput` để tương tác với khung soạn thảo Zalo.
- **Search Overlay**: Sau khi click vào một người trong danh sách tìm kiếm, script sẽ nhấn `Escape` để đóng overlay search, giúp Editor lộ ra.
- **Image Sending**: Sử dụng PowerShell để bơm ảnh vào Clipboard và `Ctrl+V` để dán vào Zalo.
- **Per-Person Cycle**: Với mỗi người trong danh sách 4 người, script sẽ thực hiện một vòng lặp Search mới để đảm bảo độ ổn định của DOM.
- **Exclusion**: Chỉ tương tác với các ID bắt đầu bằng `friend-item-`. Tuyệt đối không click vào `group-item-`.

---

> [!IMPORTANT]
> **Lệnh chạy nhanh**: `.\venv\Scripts\python.exe OpenZaloSendListRelative.py`
