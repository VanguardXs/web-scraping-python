import re
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup


BASE_URL = "https://books.toscrape.com/"
PAGE_URL = urljoin(BASE_URL, "catalogue/page-1.html")


def parse_price(text: str):
    m = re.search(r"(\d+(\.\d+)?)", text.replace(",", "."))
    return float(m.group(1)) if m else None


def get_rating(article):
    p = article.select_one("p.star-rating")
    if not p:
        return None
    for c in p.get("class", []):
        if c != "star-rating":
            return c
    return None


def fetch_html(url: str) -> str:
    r = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=15,
    )
    r.raise_for_status()
    return r.text


def parse_books(html: str):
    soup = BeautifulSoup(html, "lxml")
    books = []

    for a in soup.select("article.product_pod"):
        title_el = a.select_one("h3 a")
        price_el = a.select_one(".price_color")

        books.append(
            {
                "title": title_el.get("title") if title_el else None,
                "price": parse_price(price_el.text) if price_el else None,
                "rating": get_rating(a),
                "link": urljoin(PAGE_URL, title_el.get("href")) if title_el else None,
            }
        )

    return books


def save_csv(rows, path):
    import csv
    from pathlib import Path

    Path(path).parent.mkdir(parents=True, exist_ok=True)

    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=rows[0].keys())
        w.writeheader()
        w.writerows(rows)


def save_xlsx(rows, path):
    from pathlib import Path

    import openpyxl

    Path(path).parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "books"

    ws.append(list(rows[0].keys()))
    for r in rows:
        ws.append(list(r.values()))

    wb.save(path)


def main():
    html = fetch_html(PAGE_URL)
    books = parse_books(html)

    save_csv(books, "data/books_page1.csv")
    save_xlsx(books, "data/books_page1.xlsx")

    print(f"Saved {len(books)} items")


if __name__ == "__main__":
    main()
