# 📈 Financial News Scraper (Investing.com & AAStocks)

A robust, concurrent, and crash-proof scraper suite designed to extract historical financial news for multiple stocks from Investing.com and AAStocks. It supports multi-threading, smart resuming, and real-time status reporting.

## ✨ Key Features

- ⚡ **Concurrent Scraping**: Scrape multiple stocks simultaneously (configurable limit).
- 🌍 **Multi-Source**: Dedicated scrapers for both Investing.com (Global/US focus) and AAStocks (Hong Kong focus).
- 🛡️ **Crash-Proof**: Saves progress after every page. If the script stops, it resumes exactly where it left off next time.
- 🧠 **Smart Resume**: Detects existing data. If a stock was partially scraped, it continues from the last saved link. If complete, it skips it.
- 🕵️ **Cloudflare Bypass**: Uses a stealth Firefox browser instance to minimize detection.
- 📊 **Real-Time Reporting**: Updates a summary CSV instantly as stocks progress.

### Two-Phase Collection (AAStocks):
Safely collects all links first, then scrapes content, preventing data loss during long jobs.

## 🛠️ Installation

### 1. Prerequisites
- Python 3.8 or higher.
- Google Chrome or Firefox installed on your machine.

### 2. Install Dependencies
Create a `requirements.txt` file with these contents:

```
playwright
pandas
pyyaml
openpyxl
```

Then run:

```bash
pip install -r requirements.txt
```

### 3. Install Playwright Browsers
This is critical for the scraper to work:

```bash
playwright install firefox
```

## ⚙️ Configuration

There are two separate configuration files depending on which site you are scraping.

### 1. For Investing.com (`config.yaml`)
Use this for Global & US stocks. Investing.com requires "URL Slugs" (e.g., tesla-motors).

```yaml
start_date: "2025-01-01"
end_date: "2026-02-05"
limit_items: 1000
headless: false
max_concurrent: 2

categories:
  Auto:
    - "TSLA"          # Ticker (Script tries to resolve this)
    - "tesla-motors"  # Slug (Recommended - Faster/Safer)
  AI_Chips:
    - "nvidia-corp"   # Correct Slug for NVDA
```

### 2. For AAStocks (`config_aastocks.yaml`)
Use this for Hong Kong stocks and simple US Tickers. AAStocks requires numeric codes for HK stocks.

```yaml
start_date: "2025-01-01"
end_date: "2026-02-05"
limit_items: 1000
headless: false
max_concurrent: 3

categories:
  Auto:
    - "1810"  # Xiaomi (HK: Use Numbers)
    - "TSLA"  # Tesla (US: Use Ticker)
  Bank:
    - "5"     # HSBC (HK: Use "5" or "00005")
```

## 🚀 How to Run

### Option A: Scrape Investing.com

#### Validate Slugs (Recommended):
Check if your stock codes resolve to valid Investing.com pages.

```bash
python check_slugs.py
```

**Tip**: If you see "Invalid", replace the code in `config.yaml` with the correct URL slug (e.g., change NVDA to nvidia-corp).

#### Run Scraper:

```bash
python scraper_investing.py
```

### Option B: Scrape AAStocks

#### Prepare Config:
Ensure `config_aastocks.yaml` is set up with HK numeric codes or US tickers.

#### Run Scraper:

```bash
python scraper_aastocks.py
```

This runs in two phases: Phase 1 collects all news links, Phase 2 scrapes the content.

## 🛑 How to Stop & Resume

### Stopping

- **Press Ctrl + C ONCE**:
  - The script will finish the current active pages/articles, save the data, and exit gracefully.
  - **Safe Stop**: No data corruption.
- **Press Ctrl + C TWICE**:
  - **Force Kill**: The script stops immediately. Some data from the currently active page might be lost, but previous pages are safe.

### Resuming

Just run the python script again.

- **Completed Stocks**: Skipped automatically (⏭️ Skipping...).
- **Partial Stocks**:
  - **Investing.com**: Resumes adding new articles until the limit is reached.
  - **AAStocks**: Checks existing links and only scrapes the missing content.

## 📂 Output Files

For every stock, the script generates:

- `temp_... .jsonl`
  - **What is it?** The raw backup data.
  - **Why?** Updated instantly. If the script crashes, this file keeps your data safe.
- `Category_Code_Name.xlsx`
  - **What is it?** The final, clean Excel file.
  - **When does it appear?** Only when a stock is successfully finished.
- `..._summary_report.csv`
  - **What is it?** A master log of all jobs. Open this in Excel/VS Code to see which stocks passed, failed, or are in progress.

## ❓ Troubleshooting

| Issue                  | Solution                                                                 |
|------------------------|-------------------------------------------------------------------------|
| "404 Page Not Found"  | Investing.com: The slug is wrong. Use `check_slugs.py` to find the correct one. |
|                        | AAStocks: Ensure HK stocks are numeric (e.g., "700") and US stocks are tickers (e.g., "TSLA"). |
| "Timeout 30000ms"     | Internet lag or Cloudflare. The script usually retries. If persistent, increase `max_concurrent` to 1 to reduce load. |
| Script Freezes on Stop | You pressed Ctrl+C once. It is waiting for the current page (which can take 45s) to finish loading. Press Ctrl+C again to force quit. |
| Excel Cell Warning     | Some news articles are huge. The script automatically truncates text to 32,000 chars to prevent Excel corruption. |