# Skill: Tự động hóa Zalo Web - Tìm kiếm Theo Tên (Relative Search)

Bộ skill này hướng dẫn cách cài đặt và vận hành công cụ tự động hóa nhắn tin Zalo Web bằng cách tìm kiếm từ khóa (tên). Script được tối ưu hóa để gửi tin nhắn hàng loạt cho nhiều người có cùng tên, xử lý danh sách ảo (Virtual List).

---

## 1. Các Tính Năng Nổi Bật

- **Tìm kiếm tương đối**: Gõ một tên (VD: "Nam") và gửi cho nhiều kết quả hiện ra.
- **Quy trình thông minh**:
    - Tự động gộp **Tiêu đề** và **Nội dung** tin nhắn.
    - **Loại bỏ Nhóm**: Chỉ gửi cho cá nhân, tự động dừng khi chạm đến vùng "Nhóm".
    - **Chạy từng dòng một**: Mỗi lần chạy chỉ xử lý đúng 1 dòng dữ liệu rồi dừng (để đảm bảo an toàn).
- **Hack Vượt Rào**: Bơm ảnh vào Clipboard hệ điều hành qua PowerShell rồi nhấn Ctrl+V.

---

## 2. Thiết Lập Môi Trường (Setup)

### Bước 1: Chuẩn bị Credentials & File chính
- Đảm bảo có file `gen-lang-client-*.json` (Google Service Account) trong thư mục gốc.
- File code chính là: `OpenZaloSendListRelative.py`.

### Bước 2: Cài đặt thư viện
```powershell
py -m venv venv
.\venv\Scripts\Activate.ps1
pip install selenium gspread oauth2client
```

### Bước 3: Thư mục Images & Profile
- Tạo thư mục `images/` và để ảnh vào (VD: `hoc-python.png`).
- Thư mục `zalo-chrome-profile/` sẽ tự sinh ra để lưu phiên đăng nhập.

---

## 3. Cách Vận Hành

### Chạy chương trình
Mỗi lần chạy sẽ xử lý **1 dòng** chưa hoàn thành trong Google Sheet:
```powershell
.\venv\Scripts\python.exe OpenZaloSendListRelative.py
```
> **Lưu ý**: Lần đầu chạy cần quét mã QR Zalo. Các lần sau script sẽ tự động vào thẳng.

---

## 4. Cấu Hình Google Sheet

Script tìm worksheet có tên **"Danh Sách Theo Tên"** với các cột sau:

| Cột | Mô tả |
|-----|-------|
| `Name` | Từ khóa tìm kiếm trên Zalo (VD: "Hải", "Nam") |
| `Tiêu đề` | Phần mở đầu tin nhắn (sẽ tự động xuống dòng sau tiêu đề) |
| `Nội Dung` | Nội dung chính của tin nhắn |
| `Hình ảnh` | Tên file ảnh trong thư mục `images/` (VD: `hoc-python`) |
| `Status` | `UNAPPROVED` (chưa gửi), `ĐÃ GỞI LẦN 1` (đã gửi xong 4 người đầu) hoặc `APPROVED` (hoàn thành) |

> **Lưu ý**: Script sẽ tự động ghi nhớ các thành viên đã gửi vào các cột ẩn phía sau để tránh gửi trùng khi chạy lại.

---

## 5. Cơ Chế Xử Lý Thông Minh (Logic)

1. **Search-Send-Research**: Mỗi khi gửi xong cho 1 người, script thực hiện tìm kiếm lại từ đầu để đảm bảo ổn định UI.
2. **Exclusion (Loại trừ)**: Chỉ chọn `friend-item-` (Cá nhân), bỏ qua `group-item-` (Nhóm).
3. **Clipboard Injection**: Sử dụng lệnh PowerShell để nạp ảnh trực tiếp vào Clipboard, giúp gửi ảnh mà không cần mở hộp thoại Windows.
