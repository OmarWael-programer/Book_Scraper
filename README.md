# Book Scraper

A Python web scraping project that extracts book information from Books to Scrape.

## Features

- Scrapes all book pages
- Extracts:
  - Title
  - Price
  - Availability
  - Rating
  - UPC
  - Product Type
  - Tax
  - Number of Reviews
  - Description
- Saves data to:
  - CSV
  - Excel (.xlsx)

## Technologies

- Python
- Requests
- BeautifulSoup
- openpyxl
- csv

## How to Run

1. Install the requirements:

```bash
pip install -r requirements.txt
```

2. Run:

```bash
python scraper.py
```

## Output

- books.csv
- books.xlsx
