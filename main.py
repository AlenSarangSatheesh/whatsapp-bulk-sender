import time
import random
import pandas as pd
import logging
import sys
import os

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# ================= SETTINGS =================
# Uses file in the same folder as the script
EXCEL_FILENAME = "num.xlsx" 
MESSAGE = "Asian Paint and PPL pharma good for put option buy. Try"

# delay AFTER message is sent (per message, random)
POST_SEND_MIN = 1
POST_SEND_MAX = 3

# delay BETWEEN numbers (random)
BETWEEN_MIN = 10
BETWEEN_MAX = 25

# human typing speed (seconds per character)
TYPE_MIN = 0.04
TYPE_MAX = 0.12

# TIMEOUT SETTING
WAIT_TIMEOUT = 10 

MAX_RETRIES = 2
LOG_FILE = "whatsapp_log.txt"
FAILED_FILE = "failed_numbers.txt"
SUCCESS_FILE = "success_numbers.txt"

# Creates profile folder inside the project directory
PROFILE_DIR = os.path.join(os.getcwd(), "whatsapp_selenium_profile")
# ============================================

# Windows UTF-8 fix
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

os.makedirs(PROFILE_DIR, exist_ok=True)

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

def clean_phone_number(raw):
    """
    Cleans phone number and ensures it has a country code.
    Defaults to 91 (India) if only 10 digits are found.
    """
    num = str(raw).strip()
    if num.endswith(".0"):
        num = num[:-2]
    
    num = num.replace(" ", "").replace("-", "").replace("(", "").replace(")", "")
    
    # === MOBILE FILTER LOGIC ===
    if len(num) == 10 and num.isdigit():
        if num[0] in ['6', '7', '8', '9']:
            return "+91" + num
        else:
            return num 
            
    if len(num) == 11 and num.startswith("0") and num[1:].isdigit():
        if num[1] in ['6', '7', '8', '9']:
            return "+91" + num[1:]

    if not num.startswith("+"):
        num = "+" + num
    return num

def validate_phone_number(num):
    if not num.startswith("+91"): 
        if num.startswith("+"):
            return 10 <= len(num[1:]) <= 15
        return False

    digits = num[3:]
    
    if len(digits) == 10 and digits.isdigit():
        if digits[0] in ['6', '7', '8', '9']:
            return True
        else:
            logging.warning(f"Skipped {num}: Starts with '{digits[0]}' (Not a mobile number)")
            return False
            
    return False

def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument(f"--user-data-dir={PROFILE_DIR}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_argument("--disable-blink-features=AutomationControlled")

    return webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )

def wait_for_whatsapp(driver):
    driver.get("https://web.whatsapp.com/")
    print("Waiting for QR scan (60s timeout)...")
    WebDriverWait(driver, 60).until(
        EC.presence_of_element_located((By.ID, "pane-side"))
    )
    logging.info("✓ WhatsApp ready")

def open_chat_same_tab(driver, phone):
    driver.get(f"https://web.whatsapp.com/send?phone={phone[1:]}")
    
    try:
        WebDriverWait(driver, WAIT_TIMEOUT).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//footer//div[@contenteditable='true']")),
                EC.presence_of_element_located((By.XPATH, "//div[@data-testid='popup-controls-ok']")),
                EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'url is invalid')]"))
            )
        )
    except Exception:
        raise Exception(f"Chat load timed out ({WAIT_TIMEOUT}s)")

    popups = driver.find_elements(By.XPATH, "//div[@data-testid='popup-controls-ok']")
    if len(popups) > 0:
        try:
            popups[0].click()
        except:
            pass
        raise Exception("Invalid Number (Popup Detected)")
        
    if len(driver.find_elements(By.XPATH, "//footer//div[@contenteditable='true']")) == 0:
         raise Exception("Chat box did not load")

def human_typing(element, text):
    for char in text:
        element.send_keys(char)
        time.sleep(random.uniform(TYPE_MIN, TYPE_MAX))

def send_message(driver, message):
    box = WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located(
            (By.XPATH, "//footer//div[@contenteditable='true']")
        )
    )

    time.sleep(1)
    box.click()
    human_typing(box, message)
    time.sleep(0.5)
    box.send_keys(Keys.ENTER)

    delay = random.uniform(POST_SEND_MIN, POST_SEND_MAX)
    logging.info(f"Post-send delay: {delay:.1f}s")
    time.sleep(delay)

def save_failed_number(phone):
    try:
        with open(FAILED_FILE, "a", encoding="utf-8") as f:
            f.write(f"{phone}\n")
    except Exception as e:
        logging.error(f"Could not save failed number: {e}")

def save_success_number(phone):
    try:
        with open(SUCCESS_FILE, "a", encoding="utf-8") as f:
            f.write(f"{phone}\n")
    except Exception as e:
        logging.error(f"Could not save success number to TXT: {e}")

def main():
    logging.info("=" * 60)
    logging.info("WhatsApp Bulk Sender")
    
    # Path construction
    excel_path = os.path.join(os.getcwd(), EXCEL_FILENAME)
    
    logging.info(f"Looking for Excel file at: {excel_path}")
    logging.info("=" * 60)

    try:
        df = pd.read_excel(excel_path, header=None)
        numbers = df.iloc[:, 0].dropna()
        total = len(numbers)
    except Exception as e:
        logging.error(f"Error reading Excel file: {e}")
        logging.error(f"Make sure '{EXCEL_FILENAME}' is in the same folder as this script.")
        return

    print(f"\n⚠️ About to send messages to {total} numbers")
    print(f"Message: {MESSAGE}")
    if input("Continue? (y/n): ").strip().lower() != "y":
        return

    driver = init_driver()
    try:
        wait_for_whatsapp(driver)
    except Exception:
        logging.error("Timeout waiting for WhatsApp to load. Please restart.")
        driver.quit()
        return

    success = failed = skipped = 0

    for i, raw in enumerate(numbers, 1):
        phone = clean_phone_number(raw)
        print(f"\n[{i}/{total}] Processing: {phone}")

        if not validate_phone_number(phone):
            skipped += 1
            logging.warning(f"Skipped invalid/landline: {phone}")
            continue

        sent = False
        current_retries = MAX_RETRIES 
        
        for attempt in range(current_retries):
            try:
                open_chat_same_tab(driver, phone)
                send_message(driver, MESSAGE)
                logging.info(f"✓ Message sent to {phone}")
                save_success_number(phone)
                
                success += 1
                sent = True
                break
            except Exception as e:
                err_msg = str(e).split('\n')[0]
                logging.error(f"Attempt {attempt+1} failed for {phone}: {err_msg}")
                
                if "Invalid Number" in err_msg or "Popup Detected" in err_msg:
                    break
                    
                time.sleep(2)

        if not sent:
            failed += 1
            logging.error(f"❌ Failed to send to {phone}.")
            save_failed_number(phone)

        if i < total:
            wait_time = random.uniform(BETWEEN_MIN, BETWEEN_MAX)
            logging.info(f"Waiting {wait_time:.1f}s before next number")
            time.sleep(wait_time)

    logging.info("Final wait before closing browser...")
    time.sleep(5)

    logging.info("=" * 60)
    logging.info(f"✓ Sent: {success}")
    logging.info(f"✗ Failed: {failed}")
    logging.info(f"⊘ Skipped: {skipped}")
    logging.info("=" * 60)

    driver.quit()

if __name__ == "__main__":
    main()