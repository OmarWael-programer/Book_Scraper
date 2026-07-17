# Book Scraper

## Description
A Python web scraper that collects book information from Books to Scrape and exports the data to CSV and Excel files.

## Features
- Scrapes all catalogue pages
- Collects detailed information for every book
- Handles failed requests with retries
- Exports data to CSV
- Exports data to Excel
- Formats the Excel header
- Displays scraping statistics

## Technologies
- Python
- Requests
- BeautifulSoup4
- OpenPyXL

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