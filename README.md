# 📈 Investing.com Financial News Scraper

A robust, concurrent, and crash-proof scraper designed to extract historical financial news for multiple stocks from Investing.com. It supports multi-threading, smart resuming, and real-time status reporting.

## ✨ Key Features

* **⚡ Concurrent Scraping:** Scrape multiple stocks simultaneously (configurable limit).
* **🛡️ Crash-Proof:** Saves progress after every page. If the script stops, it resumes exactly where it left off next time.
* **🧠 Smart Resume:** Detects existing data. If a stock was partially scraped, it continues from the last saved link. If complete, it skips it.
* **🕵️ Cloudflare Bypass:** Uses a stealth Firefox browser instance to minimize detection.
* **📊 Real-Time Reporting:** Updates a `scraper_summary_report.csv` instantly as stocks progress.
* **🧹 Auto-Sanitization:** Handles messy filenames and Excel cell limits automatically.

---

## 🛠️ Installation

### 1. Prerequisites
* Python 3.8 or higher.
* Google Chrome or Firefox installed on your machine.

### 2. Install Dependencies

Create a `requirements.txt` file (if you haven't already) with these contents:

```text
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

## ⚙️ Configuration (`config.yaml`)

Control everything from the `config.yaml` file without touching the code.

```yaml
# --- DATE RANGE ---
start_date: "2025-01-01"
end_date: "2026-02-05"

# --- SCRAPING LIMITS ---
limit_items: 20       # How many articles to fetch per stock
max_concurrent: 2     # How many browsers to open at once (Keep low: 2-3 is safe)
headless: false       # true = Invisible (Faster), false = Visible (Easier to debug)

# --- STOCK LIST ---
categories:
  Auto:
    - "TSLA"          # Ticker (Will search automatically)
    - "tesla-motors"  # Direct Slug (Faster/Safer)
  
  AI_Chips:
    - "981"           # Stock Code
    - "nvidia-corp"   # Correct Slug for NVDA
```

💡 **Tip: Slug vs. Ticker**

- **Slug (Recommended):** The part of the URL after `investing.com/equities/`. E.g., `tesla-motors` or `tencent-holdings-ltd`. This is 100% accurate.
- **Ticker/Code:** E.g., `TSLA` or `700`. The script will try to search for it, but sometimes search results are ambiguous (e.g., finding an ETF instead of the stock).

## 🚀 How to Run

### 1. Validate Your Slugs (Recommended First Step)

Before running a long job, check if your stock codes are valid links.

```bash
python check_slugs.py
```

Check the output: It will tell you which stocks are `[Valid]` and which are `[Invalid]`.

Fix Invalid ones: Update `config.yaml` with the correct URL slug for any failing stocks.

### 2. Start the Scraper

```bash
python scraper_investing.py
```

The script will open browser windows (if `headless: false`).

Check the terminal for progress updates (e.g., `+ TSLA: 5/20`).

### 🛑 How to Stop & Resume

#### Stopping

- **Press `Ctrl + C` ONCE.**
  - The script will finish the current page it is scraping, save the data, and exit gracefully.
  - Status: "Stopped" in the report.
- **Press `Ctrl + C` TWICE.**
  - Force Kill. The script stops immediately. Some data from the current active page might be lost, but previous pages are safe.

#### Resuming

Just run `python scraper_investing.py` again.

- **Completed Stocks:** Skipped automatically (`⏭️ Skipping...`).
- **Partial Stocks:** Resumes adding new articles until the limit is reached.
- **Failed Stocks:** Retried automatically.

## 📂 Output Files

For every stock, the script generates:

1. **`temp_Category_Code.jsonl`**
   - **What is it?** The raw backup data.
   - **Why?** It is updated instantly. If the script crashes, this file keeps your data safe.

2. **`Category_Code_Name.xlsx`**
   - **What is it?** The final, clean Excel file.
   - **When does it appear?** Only when a stock is successfully finished.

3. **`scraper_summary_report.csv`**
   - **What is it?** A master log of all jobs. Open this in Excel/VS Code to see which stocks passed or failed.

## ❓ Troubleshooting

| Issue                  | Solution                                                                                     |
|------------------------|---------------------------------------------------------------------------------------------|
| `404 Page Not Found`   | The stock code in `config.yaml` is wrong. Use the `check_slugs.py` script to find the correct URL slug (e.g., change `NVDA` to `nvidia-corp`). |
| `Timeout 30000ms`      | Internet lag or Cloudflare. The script usually retries. If persistent, increase `max_concurrent` to `1` to reduce load. |
| Script Freezes on Stop | You pressed `Ctrl+C` once. It is waiting for the current page (which can take 45s) to finish loading. Press `Ctrl+C` again to force quit. |
| Excel Cell Warning     | Some news articles are huge. The script automatically truncates text to 32,000 chars to prevent corruption. |