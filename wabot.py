#!/usr/bin/env python3
import os
import re
import json
import time
import random
import urllib.parse
from datetime import datetime

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

CONFIG_FILE = "config.json"
CONTACTS_FILE = "contacts.xlsx"   # You can rename your Excel to this or pass a different path via CLI
SENT_LOG = "sent_log.csv"

def load_config(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Missing {path}.")
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    # Defaults
    cfg.setdefault("message_text", "")
    cfg.setdefault("media_path", "")
    cfg.setdefault("default_country_code", "91")
    cfg.setdefault("batch_limit", 500)
    cfg.setdefault("sleep_min_seconds", 8)
    cfg.setdefault("sleep_max_seconds", 16)
    cfg.setdefault("long_pause_every", 30)
    cfg.setdefault("long_pause_range_seconds", [30, 60])
    cfg.setdefault("dry_run", False)
    return cfg

def normalize_number(raw: str, default_cc: str = "91") -> str:
    if raw is None:
        return ""
    s = re.sub(r"\D+", "", str(raw))
    if s.startswith("0"):
        s = s.lstrip("0")
    # If it's a 10-digit local number, prefix default country code
    if len(s) == 10 and not s.startswith(default_cc):
        s = default_cc + s
    return s

def load_contacts(xlsx_path: str, default_cc: str = "91"):
    if not os.path.exists(xlsx_path):
        raise FileNotFoundError(f"Missing {xlsx_path}.")
    df = pd.read_excel(xlsx_path)
    if "Phone" not in df.columns:
        raise ValueError("Excel must contain a 'Phone' column.")
    if "Name" not in df.columns:
        df["Name"] = ""
    df["PhoneNorm"] = df["Phone"].apply(lambda x: normalize_number(x, default_cc))
    df = df.dropna(subset=["PhoneNorm"])
    df = df[df["PhoneNorm"].str.len() >= 10]
    df = df.drop_duplicates(subset=["PhoneNorm"])
    return df[["Name", "PhoneNorm"]].rename(columns={"PhoneNorm": "Phone"})

def load_sent_log(path: str):
    if not os.path.exists(path):
        return set()
    try:
        log_df = pd.read_csv(path)
        sent_set = set(log_df[log_df["status"] == "sent"]["phone"].astype(str).tolist())
        return sent_set
    except Exception:
        return set()

def append_log(path: str, phone: str, status: str, note: str = ""):
    exists = os.path.exists(path)
    with open(path, "a", encoding="utf-8") as f:
        if not exists:
            f.write("timestamp,phone,status,note\n")
        f.write(f"{datetime.now().isoformat(timespec='seconds')},{phone},{status},{note.replace(',', ';')}\n")

def build_driver():
    options = webdriver.ChromeOptions()
    # Keep a persistent profile so you don't need to scan QR every time
    options.add_argument("--user-data-dir=" + os.path.abspath("chrome-data"))
    options.add_argument("--profile-directory=Default")
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    # Headful mode is recommended
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def wait_for_composer(driver, timeout=30):
    # Try several possible selectors for the message composer
    candidates = [
        (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='10']"),
        (By.CSS_SELECTOR, "div[contenteditable='true'][data-tab='6']"),
        (By.CSS_SELECTOR, "div[contenteditable='true'][role='textbox']"),
        (By.CSS_SELECTOR, "div[contenteditable='true']")
    ]
    end = time.time() + timeout
    while time.time() < end:
        for by, sel in candidates:
            try:
                el = driver.find_element(by, sel)
                if el.is_displayed():
                    return el
            except Exception:
                pass
        time.sleep(0.5)
    raise TimeoutException("Composer not found. Are you logged in?")

def send_text_only(driver, phone: str, message: str, timeout=30):
    url = f"https://web.whatsapp.com/send?phone={phone}&text={urllib.parse.quote(message)}"
    driver.get(url)
    try:
        box = wait_for_composer(driver, timeout=timeout)
        time.sleep(1.2)
        
        # First try ENTER
        box.send_keys(Keys.ENTER)
        time.sleep(1.5)
        
        # Check if message really got sent, otherwise click send button
        try:
            send_btn = driver.find_element(By.CSS_SELECTOR, "span[data-icon='send'], div[aria-label='Send']")
            if send_btn.is_displayed():
                send_btn.click()
        except Exception:
            pass

        return True, ""
    except Exception as e:
        return False, f"text_send_error: {e}"


def attach_and_send_media(driver, file_path: str, timeout=60):
    # Try to open the attach menu
    attach_candidates = [
        (By.CSS_SELECTOR, "div[title='Attach']"),
        (By.CSS_SELECTOR, "span[data-icon='attach-menu-plus']"),
        (By.CSS_SELECTOR, "div[aria-label='Attach']"),
        (By.CSS_SELECTOR, "div[data-testid='clip']"),
    ]
    clicked = False
    for by, sel in attach_candidates:
        try:
            el = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((by, sel)))
            el.click()
            clicked = True
            break
        except Exception:
            continue
    # Whether clicked or not, try finding file inputs
    file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    if not file_inputs:
        # Sometimes inputs are not immediately in DOM; wait a bit
        end = time.time() + 10
        while time.time() < end and not file_inputs:
            time.sleep(0.5)
            file_inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='file']")
    if not file_inputs:
        raise RuntimeError("Could not find file input to upload media.")
    abs_path = os.path.abspath(file_path)
    file_inputs[-1].send_keys(abs_path)
    # Wait for the send button and click
    send_candidates = [
        (By.CSS_SELECTOR, "span[data-icon='send']"),
        (By.CSS_SELECTOR, "div[aria-label='Send']"),
        (By.CSS_SELECTOR, "button[aria-label*='Send']"),
        (By.CSS_SELECTOR, "button[data-testid='compose-btn-send']")
    ]
    for by, sel in send_candidates:
        try:
            btn = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, sel)))
            btn.click()
            return True
        except Exception:
            continue
    # As a fallback, try ENTER in the composer
    try:
        box = wait_for_composer(driver, timeout=10)
        box.send_keys(Keys.ENTER)
        return True
    except Exception:
        pass
    return False

def send_text_and_media(driver, phone: str, message: str, media_path: str, timeout=60):
    url = f"https://web.whatsapp.com/send?phone={phone}&text={urllib.parse.quote(message)}"
    driver.get(url)
    try:
        _ = wait_for_composer(driver, timeout=timeout)
        time.sleep(1.0)
        ok = attach_and_send_media(driver, media_path, timeout=timeout)
        if ok:
            # Ensure text is sent too (if not already sent with media caption).
            # For safety, send the prefilled text as a message after media.
            try:
                box = wait_for_composer(driver, timeout=10)
                time.sleep(0.5)
                box.send_keys(Keys.ENTER)
            except Exception:
                pass
            return True, ""
        return False, "media_send_failed"
    except Exception as e:
        return False, f"media_send_error: {e}"

def main():
    import argparse
    parser = argparse.ArgumentParser(description="Bulk sender via WhatsApp Web (Selenium).")
    parser.add_argument("--contacts", default=CONTACTS_FILE, help="Path to Excel with columns: Phone[, Name]. Default: contacts.xlsx")
    parser.add_argument("--config", default=CONFIG_FILE, help="Path to config.json")
    args = parser.parse_args()

    cfg = load_config(args.config)
    msg = cfg["message_text"]
    media_path = cfg.get("media_path", "").strip()
    default_cc = cfg.get("default_country_code", "91")
    batch_limit = int(cfg.get("batch_limit", 500))
    sleep_min = float(cfg.get("sleep_min_seconds", 8))
    sleep_max = float(cfg.get("sleep_max_seconds", 16))
    long_every = int(cfg.get("long_pause_every", 30))
    long_range = cfg.get("long_pause_range_seconds", [30, 60])
    dry_run = bool(cfg.get("dry_run", False))

    contacts = load_contacts(args.contacts, default_cc=default_cc)
    sent_already = load_sent_log(SENT_LOG)

    if len(contacts) == 0:
        print("No contacts found. Ensure your Excel has a 'Phone' column.")
        return

    if dry_run:
        print(f"[DRY RUN] Would send to up to {batch_limit} contacts.")
        print(contacts.head(10))
        return

    driver = build_driver()
    driver.get("https://web.whatsapp.com")
    input("Scan the QR code (first run) and wait until your chats are visible. Then press ENTER here to continue...")

    sent_count = 0
    failures_in_a_row = 0

    for idx, row in contacts.iterrows():
        name = str(row.get("Name") or "").strip()
        phone = str(row["Phone"]).strip()

        if phone in sent_already:
            continue

        personalized = msg.replace("{name}", name if name else "there")
        ok = False
        note = ""

        try:
            if media_path:
                ok, note = send_text_and_media(driver, phone, personalized, media_path)
            else:
                ok, note = send_text_only(driver, phone, personalized)
        except Exception as e:
            ok = False
            note = f"exception: {e}"

        if ok:
            append_log(SENT_LOG, phone, "sent", note)
            sent_count += 1
            failures_in_a_row = 0
            print(f"[{sent_count}] Sent to {phone}")
        else:
            append_log(SENT_LOG, phone, "failed", note)
            failures_in_a_row += 1
            print(f"[FAIL] {phone} -> {note}")

        # pacing
        if sent_count >= batch_limit:
            print(f"Reached batch_limit={batch_limit}. Stopping.")
            break

        # If too many consecutive failures, pause/stop
        if failures_in_a_row >= 5:
            print("Too many consecutive failures. Stopping to avoid account risk.")
            break

        # Sleep with jitter
        pause = random.uniform(sleep_min, sleep_max)
        time.sleep(pause)
        if sent_count > 0 and sent_count % max(1, long_every) == 0:
            extra = random.uniform(float(long_range[0]), float(long_range[1]))
            print(f"Taking a longer break for {extra:.1f}s...")
            time.sleep(extra)

    driver.quit()
    print("Done. See sent_log.csv for results.")

if __name__ == "__main__":
    main()
