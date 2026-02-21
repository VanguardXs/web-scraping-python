import os
import time
import random
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.core.os_manager import ChromeType

# ====== CONFIG ======
CATEGORY_URL = "https://www.mechta.kz/section/smartfony/apple-iphone/"
LIMIT = 50
OUT_FILE = "mechta_iphone_catalog_50.xlsx"
BASE = "https://www.mechta.kz"
PROFILE_DIR = os.path.abspath("./selenium_profile")


def get_driver():
    """
    Initialize Chrome driver with a persistent profile.
    Persistent profile allows manual Cloudflare / location verification once.
    """
    options = Options()
    options.binary_location = "/usr/sbin/chromium-browser"

    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)

 
    service = Service(ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install())

    driver = webdriver.Chrome(service=service, options=options)

    return driver


def scroll_page(driver, steps=8):
    """
    Scroll page gradually to trigger lazy loading of products.
    """
    for _ in range(steps):
        driver.execute_script("window.scrollBy(0, 1400);")
        time.sleep(random.uniform(0.5, 0.9))


def clean_name(name: str) -> str:
    """
    Remove unwanted store suffixes from product titles.
    """
    for tail in ["| Интернет-магазин Mechta.kz", "| Mechta.kz", "Интернет-магазин Mechta.kz"]:
        name = name.replace(tail, "").strip()
    return " ".join(name.split())


def parse_listing(html: str):
    """
    Parse product cards directly from the category page.

    Extracts:
    - product URL
    - product name
    - price from nearby span containing ₸ symbol
    """
    soup = BeautifulSoup(html, "html.parser")
    items = []
    seen = set()

    for a in soup.select('a[href^="/product/"]'):
        href = a.get("href")
        if not href:
            continue

        url = urljoin(BASE, href).split("?")[0]
        if url in seen:
            continue

        name = clean_name(a.get_text(" ", strip=True))
        if not name or len(name) < 5:
            continue

        # Locate parent card container
        card = a.find_parent("div")
        if not card:
            continue

        # Find price element nearby
        price_el = card.find_next("span", string=lambda x: x and "₸" in x)
        price = price_el.get_text(strip=True) if price_el else "N/A"
        price = " ".join(price.split())

        items.append({
            "url": url,
            "name": name,
            "price": price,
            "currency": "KZT" if "₸" in price else "N/A"
        })
        seen.add(url)

    return items


def main():
    driver = get_driver()
    rows = []
    seen_urls = set()

    try:
        print("[start] Opening category page...")
        driver.get(CATEGORY_URL)
        time.sleep(2)

        input("Complete Cloudflare/location verification if required, wait for products to load, then press Enter...")

        page_num = 1
        while len(rows) < LIMIT:
            page_url = f"{CATEGORY_URL}?page={page_num}"
            print(f"[page] {page_num} -> {page_url}")

            driver.get(page_url)
            time.sleep(2)
            scroll_page(driver, steps=8)

            items = parse_listing(driver.page_source)
            if not items:
                print("[stop] No products detected. Possibly verification or location selection required.")
                print("Open the page in the browser, ensure products are visible, then press Enter.")
                input("Enter...")
                driver.refresh()
                time.sleep(2)
                continue

            added = 0
            for it in items:
                if len(rows) >= LIMIT:
                    break
                if it["url"] in seen_urls:
                    continue
                rows.append(it)
                seen_urls.add(it["url"])
                added += 1

                # Autosave progress every 10 items
                if len(rows) % 10 == 0:
                    pd.DataFrame(rows).to_excel(OUT_FILE, index=False)
                    print(f"[autosave] {len(rows)} items saved -> {OUT_FILE}")

            print(f"[ok] Added from this page: {added}. Total: {len(rows)}/{LIMIT}")

            if added == 0:
                print("[warn] No new products on this page, moving to next page...")

            page_num += 1
            time.sleep(1)

        df = pd.DataFrame(rows[:LIMIT])
        df.to_excel(OUT_FILE, index=False)
        print(f"\n[done] Saved {len(df)} products -> {OUT_FILE}")

    finally:
        input("Press Enter to close the browser...")
        driver.quit()


if __name__ == "__main__":
    main()
