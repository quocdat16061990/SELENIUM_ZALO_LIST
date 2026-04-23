"""
OpenZaloSendListRelative.py
Tự động gửi tin nhắn Zalo Web theo danh sách tên từ Google Sheet.
Logic: Search -> Click cá nhân -> Gửi text+ảnh -> Search lại -> Click người tiếp -> ...
Tối đa 4 người/dòng. Chỉ gửi Cá nhân, KHÔNG gửi Nhóm.
"""
import os, sys, time, subprocess
import gspread
from google.oauth2.service_account import Credentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import (
    TimeoutException, StaleElementReferenceException,
    ElementClickInterceptedException
)

# === Unicode Fix Windows ===
if sys.platform == "win32":
    try: sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except: pass

# ============================================================
# CONFIGURATION
# ============================================================
SHEET_ID = "1SFAr1CFMzMPQXFToZEAwA2U1FaHpeCQqv7CyMa-f-0w"
WORKSHEET_NAME = "Danh Sách Theo Tên"
CREDENTIALS_FILE = "gen-lang-client-0450618162-54ea7d476a02.json"
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CHROME_PROFILE = os.path.join(BASE_DIR, "zalo-chrome-profile")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
MAX_PER_ROW = 4  # Tối đa 4 người mỗi dòng

# ============================================================
# 1. ENVIRONMENT
# ============================================================
def validate_environment():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(f"Thiếu credentials: {CREDENTIALS_FILE}")
    os.makedirs(CHROME_PROFILE, exist_ok=True)
    os.makedirs(IMAGES_DIR, exist_ok=True)
    print("[ENV] OK - credentials, profile, images dir sẵn sàng.")

# ============================================================
# 2. DRIVER
# ============================================================
def build_driver():
    opts = Options()
    opts.add_argument(f"--user-data-dir={CHROME_PROFILE}")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"])
    opts.add_experimental_option("useAutomationExtension", False)
    d = webdriver.Chrome(options=opts)
    d.execute_script(
        "Object.defineProperty(navigator,'webdriver',{get:()=>undefined})"
    )
    d.get("https://chat.zalo.me/")
    return d

# ============================================================
# 3. GOOGLE SHEET
# ============================================================
def load_sheet_data():
    """Đọc Google Sheet, trả về (worksheet, records, headers, status_col)."""
    print("[SHEET] Đang kết nối Google Sheets...")
    creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    gc = gspread.authorize(creds)
    ws = gc.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)
    records = ws.get_all_records()
    headers = ws.row_values(1)
    # Đảm bảo cột Status
    if 'Status' not in headers:
        ws.update_cell(1, len(headers) + 1, 'Status')
        headers = ws.row_values(1)
    status_col = headers.index('Status') + 1
    # Đảm bảo cột Thành viên 1-8
    for i in range(8):
        h = f"Thành viên {i+1}"
        if h not in headers:
            ws.update_cell(1, status_col + 1 + i, h)
    headers = ws.row_values(1)
    print(f"[SHEET] Đọc được {len(records)} dòng dữ liệu.")
    return ws, records, headers, status_col

def find_matching_image(raw_name):
    """Tìm file ảnh trong thư mục images/ khớp với tên từ Sheet."""
    if not raw_name or not os.path.exists(IMAGES_DIR):
        return ""
    for fn in os.listdir(IMAGES_DIR):
        base = os.path.splitext(fn)[0]
        if raw_name.lower() == base.lower() or raw_name.lower() in fn.lower():
            path = os.path.join(IMAGES_DIR, fn)
            print(f"    [IMG] Tìm thấy: {path}")
            return path
    print(f"    [IMG] KHÔNG tìm thấy ảnh cho '{raw_name}' trong {IMAGES_DIR}")
    return ""

# ============================================================
# 4. SEARCH ZALO - Lấy danh sách Cá nhân
# ============================================================
def search_keyword(driver, keyword):
    """
    Gõ keyword vào ô search Zalo.
    Trả về list dict: [{"id": "friend-item-XXX", "name": "Tên"}]
    CHỈ lấy friend-item (Cá nhân), KHÔNG lấy group-item (Nhóm).
    """
    print(f"    [SEARCH] Đang tìm: '{keyword}'...")

    # Focus ô search bằng JS (bypass overlay)
    driver.execute_script("""
        var si = document.getElementById('contact-search-input');
        if (si) { si.focus(); si.click(); }
    """)
    time.sleep(0.5)
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('a') \
        .key_up(Keys.CONTROL).perform()
    time.sleep(0.2)
    ActionChains(driver).send_keys(Keys.BACKSPACE).perform()
    time.sleep(0.3)
    # Gõ keyword qua ActionChains (không cần element reference)
    ActionChains(driver).send_keys(keyword).perform()
    time.sleep(3)  # Chờ kết quả load

    # Thử nhấn "Xem tất cả cá nhân" nếu có
    try:
        btns = driver.find_elements(
            By.XPATH,
            "//*[contains(text(),'Xem tất cả cá nhân') or "
            "contains(text(),'Xem tất cả') or "
            "contains(text(),'Xem thêm')]"
        )
        for b in btns:
            if b.is_displayed():
                driver.execute_script("arguments[0].click();", b)
                time.sleep(2)
                break
    except:
        pass

    # Thu thập KẾT QUẢ CÁ NHÂN
    # Selector: chỉ lấy [id^='friend-item-'] trong #global_search_list
    # Đây là cách an toàn nhất để tránh click nhầm vào Nhóm
    items = driver.find_elements(
        By.CSS_SELECTOR,
        "#global_search_list [id^='friend-item-']"
    )
    friends = []
    for item in items:
        try:
            iid = item.get_attribute("id") or ""
            if not iid.startswith("friend-item-"):
                continue
            # Lấy tên
            try:
                name_el = item.find_element(
                    By.CSS_SELECTOR, ".conv-item-title__name span.truncate"
                )
                name = name_el.text.strip()
            except:
                name = item.text.split("\n")[0].strip()
            if name:
                friends.append({"id": iid, "name": name})
        except StaleElementReferenceException:
            continue
        except:
            continue

    print(f"    [SEARCH] Tìm thấy {len(friends)} cá nhân.")
    return friends

# ============================================================
# 5. CLICK VÀO CÁ NHÂN
# ============================================================
def click_friend(driver, friend_id):
    """
    Click vào friend-item bằng nhiều cách fallback.
    Trả về True nếu editor sẵn sàng sau khi click.
    """
    print(f"    [CLICK] Đang click {friend_id}...")

    # Thử 3 cách click
    clicked = False
    for method in ["actionchains", "selenium", "javascript"]:
        try:
            el = driver.find_element(By.ID, friend_id)
            if method == "actionchains":
                ActionChains(driver).move_to_element(el).click().perform()
            elif method == "selenium":
                el.click()
            else:
                driver.execute_script("arguments[0].click();", el)
            clicked = True
            print(f"    [CLICK] Thành công bằng {method}")
            break
        except Exception as e:
            print(f"    [CLICK] {method} thất bại: {type(e).__name__}")
            continue

    if not clicked:
        print(f"    [CLICK] KHÔNG THỂ click {friend_id}")
        return False

    time.sleep(2)

    # Nhấn Escape để đóng search overlay (cho editor lộ ra)
    print(f"    [CLICK] Đóng search overlay (Escape)...")
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(2)

    # Kiểm tra editor có xuất hiện không (polling JS)
    for i in range(20):
        try:
            ok = driver.execute_script("""
                var e = document.getElementById('richInput');
                if (!e) return false;
                var rect = e.getBoundingClientRect();
                return rect.width > 0 && rect.height > 0;
            """)
            if ok:
                print(f"    [CLICK] Editor sẵn sàng! (sau {i+1}s)")
                return True
        except:
            pass
        time.sleep(1)

    print(f"    [CLICK] Editor KHÔNG xuất hiện sau 20 giây.")
    return False

# ============================================================
# 6. GỬI TIN NHẮN + ẢNH
# ============================================================
def send_message_with_image(driver, text, img_path=""):
    """
    Gửi text + ảnh vào khung chat đang mở.
    Ảnh: bơm vào Clipboard Windows bằng PowerShell → Ctrl+V.
    Text: JS insertText.
    Enter: lấy fresh editor rồi Enter.
    """
    wait = WebDriverWait(driver, 10)

    # --- BƯỚC 1: GỬI ẢNH ---
    if img_path and os.path.exists(img_path):
        print(f"      [MSG] Bước 1: Chuẩn bị ảnh {os.path.basename(img_path)}")
        abs_path = os.path.abspath(img_path).replace("'", "''")

        # 1a. Bơm vào clipboard
        print(f"      [MSG] 1a: Bơm clipboard...")
        cmd = (
            f'powershell -command "'
            f'Add-Type -AssemblyName System.Windows.Forms; '
            f'Add-Type -AssemblyName System.Drawing; '
            f"[System.Windows.Forms.Clipboard]::SetImage("
            f"[System.Drawing.Image]::FromFile('{abs_path}'))"
            f'"'
        )
        r = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if r.returncode != 0:
            print(f"      [MSG] PowerShell lỗi: {r.stderr[:200]}")
        else:
            print(f"      [MSG] 1a: Clipboard OK")
        time.sleep(1)

        # 1b. Focus editor
        print(f"      [MSG] 1b: Focus editor...")
        driver.execute_script("""
            var e = document.getElementById('richInput');
            if(e) { e.focus(); e.click(); }
        """)
        time.sleep(0.5)

        # 1c. Paste (Ctrl+V)
        print(f"      [MSG] 1c: Paste ảnh (Ctrl+V)...")
        ActionChains(driver).key_down(Keys.CONTROL).send_keys('v') \
            .key_up(Keys.CONTROL).perform()
        print(f"      [MSG] 1c: Đã paste")

        # 1d. Chờ thumbnail render
        print(f"      [MSG] 1d: Chờ thumbnail render (3s)...")
        time.sleep(3)
    elif img_path:
        print(f"      [MSG] [!] File ảnh không tồn tại: {img_path}")

    # --- BƯỚC 2: CHÈN TEXT ---
    if text:
        print(f"      [MSG] Bước 2: Chèn text...")
        driver.execute_script("""
            var e = document.getElementById('richInput');
            if(e) { e.focus(); }
        """)
        time.sleep(0.3)
        driver.execute_script(
            "document.execCommand('insertText', false, arguments[0]);", text
        )
        print(f"      [MSG] Bước 2: Text OK")
        time.sleep(0.5)

    # --- BƯỚC 3: LẤY FRESH EDITOR VÀ ENTER ---
    print(f"      [MSG] Bước 3: Lấy fresh editor + Enter...")
    try:
        fresh = wait.until(
            EC.presence_of_element_located((By.ID, "richInput"))
        )
        fresh.send_keys(Keys.ENTER)
        print(f"      [MSG] Bước 3: ĐÃ GỬI THÀNH CÔNG!")
        time.sleep(1.5)
        return True
    except Exception as e:
        print(f"      [MSG] Bước 3: LỖI Enter - {e}")
        # Fallback: ActionChains Enter
        try:
            driver.execute_script("""
                var e = document.getElementById('richInput');
                if(e) { e.focus(); }
            """)
            time.sleep(0.3)
            ActionChains(driver).send_keys(Keys.ENTER).perform()
            print(f"      [MSG] Bước 3: Gửi bằng ActionChains OK")
            time.sleep(1.5)
            return True
        except Exception as e2:
            print(f"      [MSG] Bước 3: Fallback cũng lỗi - {e2}")
            return False

# ============================================================
# 7. XỬ LÝ 1 DÒNG SHEET
# ============================================================
def process_row(driver, ws, idx, row, status_col, headers):
    """
    Xử lý 1 dòng: search → click từng người → gửi → cập nhật sheet.
    Mỗi người = 1 vòng search riêng (KHÔNG giữ DOM cũ).
    """
    status = str(row.get('Status', '')).strip()
    status_up = status.upper()

    # === Kiểm tra trạng thái ===
    is_new = "UNAPPROVED" in status_up
    is_phase1 = "LẦN 1" in status_up
    if not (is_new or is_phase1):
        return False  # Không cần xử lý

    # === Lấy danh sách người đã gửi (để skip) ===
    skip_names = []
    member_offset = 0
    next_status = "ĐÃ GỞI LẦN 1"

    if is_phase1:
        member_offset = 4
        next_status = "APPROVED"
        for i in range(4):
            n = str(row.get(f"Thành viên {i+1}", "")).strip()
            if n:
                skip_names.append(n)
        print(f"\n{'='*60}")
        print(f"[DÒNG {idx}] LẦN 2 - Keyword: '{row.get('Name')}'")
        print(f"  Bỏ qua: {skip_names}")
    else:
        print(f"\n{'='*60}")
        print(f"[DÒNG {idx}] LẦN 1 - Keyword: '{row.get('Name')}'")

    # === Đọc dữ liệu từ cột (linh hoạt tên tiếng Việt) ===
    search_term = title = message = raw_img = ""
    for k, v in row.items():
        kl = k.lower()
        if "name" in kl:
            search_term = str(v).strip()
        elif "tiêu đề" in kl or "title" in kl:
            title = str(v).strip()
        elif "nội dung" in kl or "message" in kl:
            message = str(v).strip()
        elif "hình" in kl or "image" in kl:
            raw_img = str(v).strip()

    print(f"  Keyword: '{search_term}'")
    print(f"  Tiêu đề: '{title}'")
    print(f"  Nội dung: '{message[:50]}...' " if len(message) > 50 else f"  Nội dung: '{message}'")
    print(f"  Ảnh (raw): '{raw_img}'")

    # Gộp tiêu đề + nội dung
    final_text = ""
    if title:
        final_text += title + "\n"
    if message:
        final_text += message

    # Tìm file ảnh
    img_path = find_matching_image(raw_img)

    # === VÒNG LẶP: Search → Click → Send (mỗi người 1 vòng) ===
    sent_names = []
    sent_ids = set()

    for round_num in range(MAX_PER_ROW):
        print(f"\n  --- Vòng {round_num + 1}/{MAX_PER_ROW} ---")

        # SEARCH LẠI TỪ ĐẦU (DOM luôn fresh)
        friends = search_keyword(driver, search_term)

        if not friends:
            print(f"  [!] Không có kết quả cá nhân.")
            break

        # Tìm người chưa gửi
        target = None
        for f in friends:
            if f["name"] in skip_names or f["name"] in sent_names:
                continue
            if f["id"] in sent_ids:
                continue
            target = f
            break

        if not target:
            print(f"  [!] Đã hết người cá nhân mới.")
            break

        sent_ids.add(target["id"])
        print(f"  >> Chọn: {target['name']} ({target['id']})")

        # CLICK VÀO CONTACT
        if not click_friend(driver, target["id"]):
            print(f"  [!] Click thất bại, thử người tiếp theo.")
            continue

        # GỬI TIN NHẮN
        ok = send_message_with_image(driver, final_text, img_path)
        if ok:
            sent_names.append(target["name"])
            print(f"  ✓ Đã gửi cho {target['name']} ({len(sent_names)}/{MAX_PER_ROW})")
        else:
            print(f"  ✗ Gửi thất bại cho {target['name']}")

    # === CẬP NHẬT SHEET ===
    print(f"\n  [SHEET] Cập nhật kết quả...")
    exhausted = len(sent_names) < MAX_PER_ROW

    if sent_names:
        # Lưu tên thành viên
        for i, name in enumerate(sent_names[:4]):
            col = status_col + 1 + member_offset + i
            ws.update_cell(idx, col, name)
            print(f"    Thành viên {member_offset + i + 1}: {name}")

        # Cập nhật status
        final_status = "APPROVED" if exhausted else next_status
        ws.update_cell(idx, status_col, final_status)
        print(f"    Status: {final_status}")
    else:
        ws.update_cell(idx, status_col, "APPROVED")
        print(f"    Không gửi được ai -> APPROVED")

    print(f"  [SHEET] Cập nhật xong!")
    return True  # Đã xử lý 1 dòng

# ============================================================
# 8. MAIN
# ============================================================
def main():
    try:
        validate_environment()
    except Exception as e:
        print(f"[!] {e}")
        return

    ws, records, headers, status_col = load_sheet_data()

    driver = None
    try:
        print("[DRIVER] Đang mở Zalo Web...")
        driver = build_driver()

        print("[LOGIN] Chờ đăng nhập (quét QR nếu cần, tối đa 5 phút)...")
        try:
            WebDriverWait(driver, 300).until(
                EC.presence_of_element_located((By.ID, "contact-search-input"))
            )
            print("[LOGIN] Đã đăng nhập thành công!")
        except TimeoutException:
            print("[LOGIN] Hết thời gian chờ.")
            return

        # Xử lý từng dòng (CHỈ 1 DÒNG RỒI DỪNG)
        processed = False
        for idx, row in enumerate(records, start=2):
            if process_row(driver, ws, idx, row, status_col, headers):
                processed = True
                print(f"\n=> Đã xử lý xong dòng {idx}. Dừng chương trình.")
                break

        if not processed:
            print("\n=> Không có dòng nào cần xử lý (tất cả đã APPROVED).")

    except Exception as e:
        print(f"\n[LỖI HỆ THỐNG] {e}")
    finally:
        if driver:
            print("\nĐóng trình duyệt sau 5 giây...")
            time.sleep(5)
            driver.quit()

if __name__ == "__main__":
    main()
