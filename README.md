# 🕷️ Web Scraper — Python Data Extraction Tool

A flexible Python web scraper that extracts structured data from websites and exports to CSV, Excel, or JSON.

## Features
- Extract data from any website (product listings, directories, job boards, etc.)
- Automatic pagination support
- Export to CSV, Excel, or JSON
- Polite scraping with configurable delays
- Error handling and retry logic
- Command-line interface

## Quick Start

```bash
pip install requests beautifulsoup4 openpyxl

# Scrape a single page
python scraper.py --url "https://books.toscrape.com" --output books.csv

# Scrape multiple pages
python scraper.py --url "https://books.toscrape.com/catalogue/page-{}.html" --pages 5 --output books.xlsx
```

## Example Output

See `sample_output/` for example results:
- `books_sample.csv` — 20 books scraped from books.toscrape.com
- `books_sample.json` — Same data in JSON format

## Project Structure
```
├── scraper.py          # Main scraper with CLI
├── extractors.py       # Site-specific extraction logic
├── exporters.py        # CSV/Excel/JSON export functions
├── sample_output/      # Example results
│   ├── books_sample.csv
│   └── books_sample.json
├── requirements.txt
└── README.md
```

## How It Works

1. Define extraction rules in `extractors.py` (CSS selectors for the data you need)
2. Run `scraper.py` with the target URL
3. Data is automatically cleaned and exported in your chosen format

## Customization

Each client project gets a custom extractor. Example:

```python
def extract_products(soup):
    results = []
    for item in soup.select(".product_pod"):
        results.append({
            "title": item.select_one("h3 a")["title"],
            "price": item.select_one(".price_color").text,
            "rating": item.select_one("p")["class"][1],
            "available": "In stock" in item.select_one(".availability").text,
        })
    return results
```

## Technologies
Python 3.10+ • requests • BeautifulSoup4 • openpyxl • argparse
