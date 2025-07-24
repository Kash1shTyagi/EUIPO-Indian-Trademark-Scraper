# EUIPO Indian Trademark Scraper

A headless Selenium-based Python scraper for extracting trademark data (Nice Classes 1–10) from the EUIPO search portal, filtered for the India (CGPDTM) office.

## Features

* Runs in **headless** mode (no browser UI).
* Automatically closes EUIPO disclaimer modal.
* Iterates through Nice Classes 1–10 (or configurable range).
* Extracts and saves up to 10 pages per class to individual Excel files.
* Auto-adjusts column widths for readability.

## Prerequisites

* Python 3.8 or newer
* Google Chrome browser installed

## Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/Kash1shTyagi/EUIPO-Indian-Trademark-Scraper.git
   cd euipo-trademark-scraper
   ```

2. Create and activate a virtual environment:

   ```bash
   python -m venv venv
   source venv/bin/activate   # Linux/macOS
   venv\Scripts\activate    # Windows CMD
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

## Configuration

By default, the script scrapes Nice Classes 1–10 and extracts up to 10 pages per class. You can adjust these limits in the main script:

```python
# Range of Nice Classes to iterate over:
for nice_class in range(1, 11):  # 1 to 10
    # ...

# Maximum pages per class:
max_pages = 10
```

## Usage

Simply run the main Python file:

```bash
python scrape_trademarks.py
```

After completion, you will find Excel files named:

```
Indian_Trademark_Class1.xlsx
Indian_Trademark_Class2.xlsx
…
Indian_Trademark_Class10.xlsx
```

Each workbook has its columns auto-adjusted for easy reading.

## Headless Mode

The scraper uses Chrome in headless mode:

```python
from selenium.webdriver.chrome.options import Options

chrome_opts = Options()
chrome_opts.add_argument("--headless=new")

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_opts)
```

No browser window will appear; all actions occur in the background.

## Troubleshooting

* **Disallowed by EUIPO**: If EUIPO blocks headless traffic, try removing `--headless` or adding a custom user agent.
* **Timeouts**: Increase explicit wait times in the script (via `WebDriverWait`).
* **Stale Elements**: Retry logic is built in; adjust retry counts or delays as needed.

## Dependencies

See `requirements.txt`:

```
selenium
webdriver-manager
pandas
openpyxl
```

## License

This project is licensed under the **MIT License**.
See the [LICENSE](LICENSE) file for full details.