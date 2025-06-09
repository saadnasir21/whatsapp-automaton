import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pyperclip
from collections import defaultdict

# ========== CONFIGURATION ==========
SHEET_NAME           = 'AUTOMATIC ORDER CONFIRMATION'  
SERVICE_ACCOUNT_FILE = r'C:\Users\hp\Downloads\divine-clone-458223-j3-f0ac4852a414.json'

CHROMEDRIVER_PATH    = r'C:\Users\hp\Downloads\chromedriver.exe'
CHROME_BINARY        = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
USER_DATA_DIR        = r'C:\Users\hp\AppData\Local\Google\Chrome\User Data'
PROFILE_DIR          = 'Profile 2'
# ====================================

# --- Google Sheets setup ---
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file",
    "https://www.googleapis.com/auth/drive"
]
creds  = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
gc     = gspread.authorize(creds)
sheet  = gc.open(SHEET_NAME).sheet1

# Ensure "Msg Status" is in column G (7)
headers = sheet.row_values(1)
if len(headers) < 7 or headers[6] != 'Msg Status':
    sheet.update_cell(1, 7, 'Msg Status')

# # --- Selenium setup ---
options = webdriver.ChromeOptions()

# 1) Use the *cloned* profile directory (not the original)
options.add_argument(r"--user-data-dir=C:/selenium/chrome_profile2")

# 2) You no longer need --profile-directory
#   (the cloned folder is itself the profile)
# options.add_argument(r"--profile-directory=Profile 2")

# 3) Remote debugging port (required by new Chrome)
options.add_argument("--remote-debugging-port=9222")

# 4) Suppress os_crypt / sync errors
options.add_argument("--password-store=basic")
options.add_argument("--disable-sync")

# 5) Optional: start maximized
options.add_argument("--start-maximized")

service = Service(CHROMEDRIVER_PATH)
driver  = webdriver.Chrome(service=service, options=options)
wait    = WebDriverWait(driver, 30)

# ========== CORE FUNCTIONS ==========

def get_pending_orders():
    """Return list of (row_num, record) where Msg Status is blank."""
    records = sheet.get_all_records()
    pending = []
    for idx, rec in enumerate(records, start=2):
        if not str(rec.get('Msg Status', '')).strip():
            pending.append((idx, rec))
    return pending

def group_by_customer(pending):
    """Group pending rows by Billing Name."""
    groups = {}
    for row, rec in pending:
        name = rec['Billing Name']
        groups.setdefault(name, []).append((row, rec))
    return groups

def aggregate_items(entries):
    """Combine rows with the same item name and sum their quantities."""
    item_totals = defaultdict(int)
    for _, rec in entries:
        try:
            qty = int(rec['Lineitem Quantity'])
        except Exception:
            try:
                qty = int(float(rec['Lineitem Quantity']))
            except Exception:
                qty = 0
        item_totals[rec['Lineitem Name']] += qty
    return [(name, qty) for name, qty in item_totals.items()]

def open_whatsapp_chat(phone):
    driver.get(f"https://web.whatsapp.com/send?phone={phone}")
    print(f"🔄 Opening chat for {phone}")
    wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']")
    ))
    time.sleep(1)

def send_whatsapp_message(name, items, total, phone, rows):
    """Send a grouped multi-item message and mark each row’s Msg Status."""
    item_lines = "\n".join(f"{it} x {qty}" for it, qty in items)
    msg = f"""Hi {name},

Thank you for your order of the following item(s):
{item_lines}

Your total is Rs {total:.2f}. Could you kindly confirm your order so we can proceed with dispatching it?

Best regards,
ASH Homes Customer Support"""

    try:
        open_whatsapp_chat(phone)
        box = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']")
        ))
        box.click()
        time.sleep(0.5)

        # Paste the message in one go to avoid mis-typed newlines
        pyperclip.copy(msg)
        box.send_keys(Keys.CONTROL, 'v')
        box.send_keys(Keys.ENTER)
        send_confirmation_poll()

        print(f"✅ Sent to {phone}")
        for r in rows:
            sheet.update_cell(r, 7, "Message Sent ✅")
            time.sleep(0.2)

    except Exception as e:
        print(f"❌ Failed to send to {phone}: {e}")
        for r in rows:
            sheet.update_cell(r, 7, "Failed ❌")
            time.sleep(0.2)

def send_confirmation_poll():
    """Create a simple Yes/No poll to confirm the order."""
    try:
        attach_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='clip']"))
        )
        attach_btn.click()

        poll_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[title='Poll']"))
        )
        poll_btn.click()

        question = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div[contenteditable='true'][data-testid='poll-question']"))
        )
        question.send_keys("Have you confirmed your order?")

        options = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-testid='poll-create-option']"))
        )
        if len(options) >= 2:
            options[0].send_keys("Yes")
            options[1].send_keys("No")

        send_btn = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']"))
        )
        send_btn.click()
    except Exception as e:
        print(f"⚠️ Could not send poll: {e}")

def process_orders():
    while True:
        pending = get_pending_orders()
        if not pending:
            print("⏸️ No pending orders.")
        else:
            groups = group_by_customer(pending)
            for name, entries in groups.items():
                rows  = [r for r, _ in entries]
                items = aggregate_items(entries)
                total = sum(float(rec['Total']) for _, rec in entries)
                phone = "+92" + str(entries[0][1]['Shipping Phone'])[-10:]
                send_whatsapp_message(name, items, total, phone, rows)

        print("🕒 Sleeping for 5 minutes…")
        time.sleep(300)

if __name__ == "__main__":
    process_orders()
