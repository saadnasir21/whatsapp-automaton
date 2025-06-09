"""Simple marketing automation script for WhatsApp Web.

This script reads contacts from a CSV file and sends each of them a
customisable message. An optional image can be attached to every
message.

Usage:
    python whatsapp_marketing.py --message "Hi {name}!" --image path/to/img.jpg

The CSV file should have columns named "Name" and "Phone".
"""

import argparse
import csv
import time

import pyperclip
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- Selenium configuration ---
CHROMEDRIVER_PATH = r'C:\\Users\\hp\\Downloads\\chromedriver.exe'

options = webdriver.ChromeOptions()
options.add_argument(r"--user-data-dir=C:/selenium/chrome_profile2")
options.add_argument("--remote-debugging-port=9222")
options.add_argument("--password-store=basic")
options.add_argument("--disable-sync")
options.add_argument("--start-maximized")

service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)
wait = WebDriverWait(driver, 30)


def load_contacts(path: str):
    """Return list of (name, phone) tuples from a CSV file."""
    contacts = []
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = row.get('Name')
            phone = row.get('Phone')
            if name and phone:
                contacts.append((name.strip(), str(phone).strip()))
    return contacts


def open_whatsapp_chat(phone: str):
    """Navigate to the chat for the given phone number."""
    driver.get(f"https://web.whatsapp.com/send?phone={phone}")
    wait.until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']"))
    )
    time.sleep(1)


def send_whatsapp_message(name: str, phone: str, message: str, image_path: str | None = None):
    """Send a customised message (and optional image) to the given phone."""
    formatted = message.format(name=name)
    open_whatsapp_chat(phone)

    box = wait.until(
        EC.element_to_be_clickable((By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']"))
    )
    box.click()
    pyperclip.copy(formatted)
    box.send_keys(Keys.CONTROL, 'v')
    box.send_keys(Keys.ENTER)

    if image_path:
        try:
            attach = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='clip']")))
            attach.click()
            file_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='file']")))
            file_input.send_keys(image_path)
            time.sleep(1)
            send_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "span[data-icon='send']")))
            send_btn.click()
        except Exception as e:
            print(f"Failed to attach image for {phone}: {e}")

    print(f"Sent message to {phone}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Send marketing messages via WhatsApp Web")
    parser.add_argument("--contacts", default="contacts.csv", help="CSV file with Name and Phone columns")
    parser.add_argument("--message", required=True, help="Message template. Use {name} for the recipient name")
    parser.add_argument("--image", help="Optional path to an image to attach")
    args = parser.parse_args()

    recipients = load_contacts(args.contacts)
    for name, phone in recipients:
        send_whatsapp_message(name, phone, args.message, args.image)
        time.sleep(1)

    print("Done.")
