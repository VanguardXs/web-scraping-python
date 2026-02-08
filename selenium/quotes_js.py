from pathlib import Path

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.firefox.service import Service


URL = "https://quotes.toscrape.com/js/"


def create_driver() -> webdriver.Firefox:
    options = Options()
    # options.add_argument("-headless") 

    service = Service(GeckoDriverManager().install())
    return webdriver.Firefox(service=service, options=options)


def parse_page(driver) -> list[dict]:
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, "quote"))
    )

    rows = []
    for q in driver.find_elements(By.CLASS_NAME, "quote"):
        text = q.find_element(By.CLASS_NAME, "text").text
        author = q.find_element(By.CLASS_NAME, "author").text
        tags = [t.text for t in q.find_elements(By.CLASS_NAME, "tag")]
        rows.append({"text": text, "author": author, "tags": ", ".join(tags)})
    return rows


def save_csv(rows: list[dict], path: str) -> None:
    import csv

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)


def save_xlsx(rows: list[dict], path: str) -> None:
    import openpyxl

    Path(path).parent.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "quotes"
    ws.append(list(rows[0].keys()))
    for r in rows:
        ws.append(list(r.values()))
    wb.save(path)


def main():
    driver = create_driver()
    driver.get(URL)

    all_rows = []
    for _ in range(3):
        all_rows.extend(parse_page(driver))
        try:
            driver.find_element(By.CSS_SELECTOR, "li.next a").click()
        except Exception:
            break

    driver.quit()

    save_csv(all_rows, "data/quotes_firefox.csv")
    save_xlsx(all_rows, "data/quotes_firefox.xlsx")
    print(f"Saved {len(all_rows)} quotes")


if __name__ == "__main__":
    main()
