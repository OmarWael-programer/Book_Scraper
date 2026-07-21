# 📚 Book Scraper

A Python web scraper that collects book information from Books to Scrape and exports the data to multiple file formats.

## Features

- Scrapes all books from every page
- Collects:
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
- BeautifulSoup4
- openpyxl

## Installation

```bash
pip install -r requirements.txt
```

## Run

```bash
python scraper.py
```

## Output

- books.csv
- books.xlsx