import yaml
import pandas as pd
from datetime import datetime, timedelta
import re
import random
import time
import json
import os
import signal
import sys
import concurrent.futures
import threading
from playwright.sync_api import sync_playwright

# --- GLOBAL SETTINGS ---
TMP_DIR = "aa_tmp"          # <--- Folder for temp JSONL files
OUTPUT_DIR = "aa_output"    # <--- Folder for Excel and Summary files

LOG_FILE = "aastocks_progress_log.txt"
# Summary file now lives in the output directory
SUMMARY_FILE = os.path.join(OUTPUT_DIR, "aastocks_summary_report.csv")
CONFIG_FILE = "config_aastocks.yaml"
stop_requested = False

log_lock = threading.Lock()
csv_lock = threading.Lock()

# --- SIGNAL HANDLER ---
def signal_handler(sig, frame):
    global stop_requested
    if not stop_requested:
        print("\n🛑 STOP REQUESTED! Finishing active tasks... (Press Ctrl+C again to FORCE KILL)")
        stop_requested = True
    else:
        print("\n💀 FORCE QUITTING...")
        os._exit(1)

signal.signal(signal.SIGINT, signal_handler)

def load_config():
    with open(CONFIG_FILE, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_lock:
        print(f"[{timestamp}] {message}")
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(f"[{timestamp}] {message}\n")
        except: pass

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', str(name)).strip()

# --- REPORTING ---
def update_summary(category, stock_code, status, details=""):
    record = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Category": category,
        "Stock Code": stock_code,
        "Status": status,
        "Details": details
    }
    df = pd.DataFrame([record])
    with csv_lock:
        header = not os.path.exists(SUMMARY_FILE)
        try: df.to_csv(SUMMARY_FILE, mode='a', header=header, index=False)
        except: pass

# --- DATA IO ---
def get_filenames(category, stock_code):
    s_cat = sanitize_filename(category)
    s_code = sanitize_filename(stock_code)
    
    # Updated paths to use directories
    file_links = os.path.join(TMP_DIR, f"temp_links_aa_{s_cat}_{s_code}.jsonl")
    file_data = os.path.join(TMP_DIR, f"temp_data_aa_{s_cat}_{s_code}.jsonl")
    file_excel = os.path.join(OUTPUT_DIR, f"AA_{s_cat}_{s_code}.xlsx")
    
    return file_links, file_data, file_excel

def append_jsonl(filename, data_list):
    if not data_list: return
    try:
        with open(filename, "a", encoding="utf-8") as f:
            for item in data_list:
                f.write(json.dumps(item, default=str) + "\n")
    except Exception as e:
        log_message(f"❌ JSON Write Error: {e}")

def read_jsonl(filename):
    data = []
    if os.path.exists(filename):
        try:
            with open(filename, 'r', encoding="utf-8") as f:
                for line in f:
                    if line.strip(): data.append(json.loads(line))
        except: pass
    return data

def save_excel(data, filename):
    if not data: return
    df = pd.DataFrame(data)
    df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
    if 'Content' in df.columns:
        df['Content'] = df['Content'].astype(str).str.slice(0, 32000)
    
    cols = ['Category', 'Stock Code', 'Date', 'Title', 'Link', 'Recommend', 'Positive', 'Negative', 'Content']
    for c in cols: 
        if c not in df.columns: df[c] = ""
    
    try: df[cols].to_excel(filename, index=False)
    except: pass

# --- UTILS ---
def parse_aa_date(date_text):
    if not date_text: return None
    try:
        clean = re.search(r'(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2})', date_text)
        if clean: return datetime.strptime(clean.group(1), '%Y/%m/%d %H:%M')
    except: pass
    return None

# --- PHASE 1: COLLECT LINKS ---
def collect_links_phase(page, category, stock_code, is_hk, config, links_file):
    start_dt = datetime.strptime(config['start_date'], '%Y-%m-%d')
    end_dt = datetime.strptime(config['end_date'], '%Y-%m-%d')
    limit = config['limit_items']
    
    existing_links = read_jsonl(links_file)
    collected_urls = set(item['Link'] for item in existing_links)
    items_found = len(existing_links)
    
    if items_found >= limit:
        log_message(f"✅ {stock_code}: Phase 1 already done ({items_found} links).")
        return existing_links

    log_message(f"🔍 {stock_code}: Phase 1 - Collecting Links (Target: {limit})")
    
    page.route("**/*.{png,jpg,jpeg,svg,woff,woff2,gif,ico,css}", lambda route: route.abort())

    base_url = ""
    if is_hk:
        base_url = f"http://www.aastocks.com/en/stocks/analysis/stock-aafn/{stock_code}/0/hk-stock-news/"
    else:
        base_url = f"http://www.aastocks.com/en/usq/quote/stock-news.aspx?symbol={stock_code}"

    page_num = 1
    retry_scroll_count = 0 
    MAX_RETRIES = 10

    while items_found < limit and not stop_requested:
        current_url = f"{base_url}{page_num}" if is_hk else base_url
        
        try:
            if page.url != current_url:
                page.goto(current_url, timeout=60000, wait_until="domcontentloaded")
            
            # --- SCROLL LOGIC FOR ALL STOCKS ---
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)
            
            try: 
                loading = page.locator('#divLoading')
                if loading.is_visible():
                    loading.wait_for(state='hidden', timeout=10000)
            except: pass
            
            if retry_scroll_count > 2:
                page.evaluate("window.scrollBy(0, -500)")
                time.sleep(0.5)
                page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
                time.sleep(2)

            # --- SELECTOR ---
            news_items = page.locator('div.newshead4, div.newshead5').all()
            if not news_items:
                news_items = page.locator('a[href*="stock-news-content"]').all()

            if not news_items:
                if retry_scroll_count > MAX_RETRIES:
                    if is_hk and page_num > 1: break 
                    if not is_hk: break 
                
                retry_scroll_count += 1
                time.sleep(2)
                continue

            batch_links = []
            new_items_in_this_pass = 0
            
            for item in news_items:
                if items_found >= limit: break
                
                try:
                    tag_name = item.evaluate("el => el.tagName")
                    if tag_name == "DIV":
                        link_el = item.locator('a').first
                    else:
                        link_el = item
                    
                    title = link_el.inner_text().strip()
                    href = link_el.get_attribute('href')
                    
                    if not href or len(title) < 2: continue
                    
                    if not href.startswith('http'):
                        full_link = "http://www.aastocks.com" + href
                    else:
                        full_link = href
                    
                    if full_link in collected_urls: continue

                    date_text = ""
                    try:
                        parent_text = item.locator('..').inner_text() 
                        match = re.search(r'(\d{4}/\d{2}/\d{2}\s+\d{2}:\d{2})', parent_text)
                        if match: date_text = match.group(1)
                    except: pass

                    parsed_date = parse_aa_date(date_text)
                    final_date_str = parsed_date.strftime('%Y-%m-%d %H:%M:%S') if parsed_date else ""

                    if parsed_date:
                        if parsed_date < start_dt:
                            items_found = limit + 999 
                            break
                        if parsed_date > end_dt: continue

                    link_obj = {
                        'Category': category,
                        'Stock Code': stock_code,
                        'Date': final_date_str,
                        'Title': title,
                        'Link': full_link
                    }
                    batch_links.append(link_obj)
                    collected_urls.add(full_link)
                    items_found += 1
                    new_items_in_this_pass += 1
                    
                except: pass
            
            if new_items_in_this_pass > 0:
                append_jsonl(links_file, batch_links)
                log_message(f"   + {stock_code}: Found {new_items_in_this_pass} new links. Total: {len(collected_urls)}")
                retry_scroll_count = 0
            else:
                retry_scroll_count += 1
                if retry_scroll_count > MAX_RETRIES:
                    if not is_hk: break
                    # HK might need next page

            if is_hk: page_num += 1

        except Exception as e:
            log_message(f"⚠️ Link Collection Warning {stock_code}: {e}")
            break
            
    update_summary(category, stock_code, "Links Collected", f"Found {len(collected_urls)}")
    return read_jsonl(links_file)

# --- PHASE 2: SCRAPE CONTENT ---
def scrape_content_phase(context, category, stock_code, link_list, data_file, excel_file):
    existing_data = read_jsonl(data_file)
    completed_urls = set(item['Link'] for item in existing_data)
    
    links_to_scrape = [x for x in link_list if x['Link'] not in completed_urls]
    
    if not links_to_scrape:
        log_message(f"✅ {stock_code}: Content complete. Generating Excel.")
        save_excel(existing_data, excel_file)
        update_summary(category, stock_code, "Success", f"Total {len(existing_data)}")
        return

    log_message(f"📖 {stock_code}: Scraping content for {len(links_to_scrape)} links...")
    
    page = context.new_page()
    count = 0
    
    for item in links_to_scrape:
        if stop_requested: break
        
        try:
            page.goto(item['Link'], timeout=30000, wait_until="domcontentloaded")
            
            content = ""
            try:
                content_div = page.locator('#spanContent, .newscontent4, .newscontent5').first
                if content_div.count() > 0:
                    paras = content_div.locator('p').all_inner_texts()
                    content = "\n\n".join([p.strip() for p in paras if p.strip()])
            except: pass

            rec = "0"; pos = "0"; neg = "0"
            try:
                if page.locator('.divRecommend .value').count() > 0:
                    rec = page.locator('.divRecommend .value').first.inner_text()
                if page.locator('.divBullish .value').count() > 0:
                    pos = page.locator('.divBullish .value').first.inner_text()
                if page.locator('.divBearish .value').count() > 0:
                    neg = page.locator('.divBearish .value').first.inner_text()
            except: pass

            item['Content'] = content
            item['Recommend'] = rec
            item['Positive'] = pos
            item['Negative'] = neg
            
            append_jsonl(data_file, [item])
            existing_data.append(item)
            count += 1
            print(f"   -> {stock_code}: Saved article {count}/{len(links_to_scrape)}")
            
            if count % 5 == 0:
                update_summary(category, stock_code, "Scraping", f"{count}/{len(links_to_scrape)}")

            time.sleep(random.uniform(0.5, 1.2))

        except: pass
    
    page.close()
    
    if count > 0 or existing_data:
        save_excel(existing_data, excel_file)
        status = "Stopped" if stop_requested else "Success"
        update_summary(category, stock_code, status, f"Captured {len(existing_data)}")

# --- WORKER ---
def process_stock_task(task_info):
    category = task_info['category']
    raw_code = str(task_info['stock_code'])
    config = task_info['config']
    
    if stop_requested: return

    is_hk_stock = raw_code.isdigit()
    stock_code = raw_code.zfill(5) if is_hk_stock else raw_code.upper()
    
    file_links, file_data, file_excel = get_filenames(category, stock_code)

    if os.path.exists(file_excel):
        log_message(f"⏭️ {stock_code}: Excel exists. Skipping.")
        update_summary(category, stock_code, "Skipped", "Done")
        return

    with sync_playwright() as p:
        browser = p.firefox.launch(headless=config['headless'])
        context = browser.new_context(viewport={'width': 1920, 'height': 1080})
        
        try:
            # PHASE 1
            page = context.new_page()
            all_links = collect_links_phase(page, category, stock_code, is_hk_stock, config, file_links)
            page.close()
            
            if not all_links:
                log_message(f"❌ {stock_code}: No links found.")
                update_summary(category, stock_code, "Failed", "No Links")
                return

            if stop_requested: return

            # PHASE 2
            scrape_content_phase(context, category, stock_code, all_links, file_data, file_excel)
            
        except Exception as e:
            log_message(f"🔥 Crash {stock_code}: {e}")
        finally:
            browser.close()

# --- MAIN ---
def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ Error: Please create '{CONFIG_FILE}' first.")
        return

    # Create Output Directories
    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    config = load_config()
    max_workers = config.get('max_concurrent', 2)
    categories = config.get('categories', {})

    tasks = []
    for cat, stocks in categories.items():
        if not stocks: continue 
        for stock in stocks:
            tasks.append({'category': cat, 'stock_code': str(stock), 'config': config})

    log_message(f"=== STARTING AASTOCKS 2-PHASE SCRAPER ({max_workers} threads) ===")
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for task in tasks:
            if stop_requested: break
            futures.append(executor.submit(process_stock_task, task))
        
        try:
            for future in concurrent.futures.as_completed(futures):
                if stop_requested: executor.shutdown(wait=False, cancel_futures=True)
                try: future.result()
                except: pass
        except KeyboardInterrupt:
            signal_handler(None, None)

    log_message("=== SESSION END ===")

if __name__ == "__main__":
    main()