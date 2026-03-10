import yaml
import pandas as pd
from datetime import datetime, timedelta
import re
import random
import time
import json
import os
import signal
import concurrent.futures
import threading
import urllib.parse
import traceback
from playwright.sync_api import sync_playwright

# --- GLOBAL SETTINGS ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TMP_DIR = os.path.join(BASE_DIR, "yf_tmp")
OUTPUT_DIR = os.path.join(BASE_DIR, "yf_output")
LOG_FILE = os.path.join(BASE_DIR, "yfinance_progress_log.txt")
SUMMARY_FILE = os.path.join(OUTPUT_DIR, "yfinance_summary_report.csv")
CONFIG_FILE = os.path.join(BASE_DIR, "config_yfinance.yaml")
stop_requested = False

log_lock = threading.Lock()
csv_lock = threading.Lock()

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0",
]

def signal_handler(sig, frame):
    global stop_requested
    if not stop_requested:
        print("\n🛑 STOP REQUESTED! Finishing active tasks... (Press Ctrl+C again to KILL IMMEDIATELY)")
        stop_requested = True
    else:
        print("\n💀 FORCE QUITTING...")
        os._exit(1)

signal.signal(signal.SIGINT, signal_handler)

def load_config():
    with open(CONFIG_FILE, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

def get_time_str():
    return datetime.now().strftime("%H:%M:%S")

def log_message(message):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] {message}"
    with log_lock:
        print(formatted_msg)
        try:
            with open(LOG_FILE, "a", encoding="utf-8") as f:
                f.write(formatted_msg + "\n")
        except: pass

def sanitize_filename(name):
    return re.sub(r'[<>:"/\\|?*]', '', str(name)).strip()

def append_summary_report(category, stock_code, status, details="", current_date=""):
    record = {
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Category": category,
        "Stock Code": stock_code,
        "Status": status,
        "Details": details,
        "Current Search Date": current_date
    }
    df = pd.DataFrame([record])
    with csv_lock:
        header = not os.path.exists(SUMMARY_FILE)
        try: df.to_csv(SUMMARY_FILE, mode='a', header=header, index=False)
        except: pass

def get_resume_date_from_csv(stock_code):
    if not os.path.exists(SUMMARY_FILE): return None
    try:
        with csv_lock:
            df = pd.read_csv(SUMMARY_FILE)
        df['Stock Code'] = df['Stock Code'].astype(str)
        stock_data = df[df['Stock Code'] == str(stock_code)]
        if stock_data.empty: return None
        
        last_entry = stock_data.iloc[-1]
        if 'Current Search Date' in last_entry and pd.notna(last_entry['Current Search Date']) and last_entry['Current Search Date'] != "":
            return str(last_entry['Current Search Date'])
    except: pass
    return None

def get_filenames(category, stock_code):
    s_cat = sanitize_filename(category)
    s_code = sanitize_filename(stock_code)
    jsonl = os.path.join(TMP_DIR, f"yfinance_temp_{s_cat}_{s_code}.jsonl")
    excel_name = f"yfinance_{s_cat}_{s_code}.xlsx"
    excel_path = os.path.join(OUTPUT_DIR, excel_name)
    return jsonl, excel_path, excel_name

def load_existing_data(jsonl_file):
    data = []
    if os.path.exists(jsonl_file):
        try:
            with open(jsonl_file, 'r', encoding="utf-8") as f:
                for line in f:
                    if line.strip(): data.append(json.loads(line))
        except: pass
    return data

def append_to_jsonl(jsonl_file, new_data_list):
    if not new_data_list: return
    try:
        with open(jsonl_file, "a", encoding="utf-8") as f:
            for entry in new_data_list:
                f.write(json.dumps(entry, default=str) + "\n")
    except Exception as e:
        log_message(f"❌ JSON Write Error {jsonl_file}: {e}")

def create_final_excel(data, excel_filepath):
    if not data: return
    df = pd.DataFrame(data)
    df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
    if 'Content' in df.columns:
        df['Content'] = df['Content'].astype(str).str.slice(0, 32000)
    cols = ['Category', 'Stock Code', 'Published Date', 'Title', 'Content', 'Link']
    for c in cols: 
        if c not in df.columns: df[c] = ""
    try:
        df[cols].to_excel(excel_filepath, index=False)
        log_message(f"💾 EXCEL SAVED: {excel_filepath}")
        return True
    except Exception as e:
        log_message(f"❌ Excel Save Failed: {e}")
        return False

def parse_iso_date(date_str):
    if not date_str: return None
    try:
        clean_str = date_str.replace('Z', '+00:00')
        dt = datetime.fromisoformat(clean_str)
        return dt.replace(tzinfo=None) 
    except: pass
    return None

def get_article_details(context, url, fallback_title=""):
    page = None
    published_date = None
    clean_text = ""
    title = fallback_title
    
    try:
        page = context.new_page()
        page.route("**/*.{png,jpg,jpeg,svg,woff,woff2,gif,ico,css}", lambda route: route.abort())
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        
        try:
            consent_btn = page.locator('button[name="agree"], button.accept-all, button[value="agree"]').first
            if consent_btn.count() > 0 and consent_btn.is_visible(timeout=2000):
                consent_btn.click()
                time.sleep(1)
            
            close_login = page.locator('button.close, .close-icon, [aria-label="Close"], .x-icon').first
            if close_login.count() > 0 and close_login.is_visible(timeout=2000):
                close_login.click()
        except: 
            pass 
        
        try:
            time_el = page.locator('time.byline-attr-meta-time, time').first
            if time_el.count() > 0:
                date_str = time_el.get_attribute('datetime')
                published_date = parse_iso_date(date_str)
        except: pass

        try:
            h1_el = page.locator('h1').first
            if h1_el.count() > 0:
                title = h1_el.inner_text().strip()
        except: pass

        try:
            read_more_btn = page.locator('button[aria-label="Story Continues"], button.readmore-button').first
            if read_more_btn.count() > 0 and read_more_btn.is_visible():
                read_more_btn.evaluate("el => el.click()")
                time.sleep(0.5)
        except: pass

        try:
            body_container = page.locator('.bodyItems-wrapper, div[data-testid="article-body"], .caas-body').first
            if body_container.count() > 0:
                paragraphs = body_container.locator('p, h2, h3').all_inner_texts()
                clean_text = "\n\n".join([p.strip() for p in paragraphs if p.strip()])
        except: pass
        
    except Exception as e:
        pass 
    finally:
        if page: 
            if not clean_text:
                log_message(f"      ⚠️ DEBUG: Pausing for 15 seconds so you can inspect this failed page...")
                time.sleep(15)  
            try: page.close()
            except: pass
            
    return title, clean_text, published_date

def process_stock_task(task_info):
    try:
        global stop_requested
        category = task_info['category']
        raw_stock_code = str(task_info['stock_code'])
        config = task_info['config']
        
        # Extracts just 'TSLA' from ''TSLA" OR "Tesla''
        base_ticker = raw_stock_code.split('"')[0].replace("'", "").strip()
        if not base_ticker:
            base_ticker = raw_stock_code 
        
        limit = config.get('limit_items', 10000)
        daily_limit = config.get('daily_limit', 10)
        
        if stop_requested: return

        jsonl_filename, excel_path, excel_name = get_filenames(category, raw_stock_code)
        
        if os.path.exists(excel_path):
            log_message(f"[{get_time_str()}] ⏭️ {base_ticker}: Excel exists. Skipping.")
            append_summary_report(category, raw_stock_code, "Skipped", "Excel Exists")
            return

        existing_data = load_existing_data(jsonl_filename)
        processed_links = set(item['Link'] for item in existing_data)
        items_collected = len(processed_links)

        if items_collected >= limit:
            log_message(f"[{get_time_str()}] ✅ {base_ticker}: Done ({items_collected} items). Generating Excel...")
            create_final_excel(existing_data, excel_path)
            return

        try:
            start_dt = datetime.strptime(config['start_date'], '%Y-%m-%d')
            end_dt = datetime.strptime(config['end_date'], '%Y-%m-%d')
        except ValueError as e:
            log_message(f"❌ DATE FORMAT ERROR in config.yaml: {e}. Please use YYYY-MM-DD format (e.g. 2025-01-01).")
            return
        
        loop_start_dt = end_dt
        
        # --- FIX: SAFE DICTIONARY FETCHING ---
        resume_dates = config.get('resume_dates') or {}
        cat_resume_dates = resume_dates.get(category) or {}
        
        if base_ticker in cat_resume_dates:
            try:
                loop_start_dt = datetime.strptime(str(cat_resume_dates[base_ticker]), '%Y-%m-%d')
                log_message(f"[{get_time_str()}] ⏩ {base_ticker}: Resuming at {loop_start_dt.date()} (from Config)")
            except Exception as e: 
                log_message(f"[{get_time_str()}] ⚠️ Bad resume date format in config for {base_ticker}: {e}")

        csv_resume_dt_str = get_resume_date_from_csv(raw_stock_code)
        if csv_resume_dt_str:
            try:
                csv_resume_dt = datetime.strptime(csv_resume_dt_str, '%Y-%m-%d')
                if csv_resume_dt < loop_start_dt:
                    loop_start_dt = csv_resume_dt - timedelta(days=1)
                    log_message(f"[{get_time_str()}] ⏩ {base_ticker}: Resuming at {loop_start_dt.date()} (from CSV Report)")
            except: pass

        if loop_start_dt < start_dt:
            log_message(f"[{get_time_str()}] ✅ {base_ticker}: Reached target start date ({start_dt.date()}).")
            create_final_excel(existing_data, excel_path)
            return
            
        date_range = pd.date_range(start=start_dt, end=loop_start_dt)
        dates_to_search = [d for d in reversed(date_range)]

        headless = config.get('headless', False) 

        playwright_instance = None
        context = None
        google_page = None

        def start_browser():
            p = sync_playwright().start()
            ua = random.choice(USER_AGENTS)
            user_data_dir = os.path.join(BASE_DIR, "playwright_chrome_profile")
            
            ctx = p.chromium.launch_persistent_context(
                user_data_dir,
                headless=headless, 
                viewport={'width': 1920, 'height': 1080},
                user_agent=ua,
                args=["--disable-blink-features=AutomationControlled"]
            )
            return p, ctx

        try:
            playwright_instance, context = start_browser()
            
            if len(context.pages) > 0:
                google_page = context.pages[0]
            else:
                google_page = context.new_page()
                
            google_page.route("**/*.{png,jpg,jpeg,svg,woff,woff2,gif,ico}", lambda route: route.abort())

            log_message(f"[{get_time_str()}] 🚀 {base_ticker}: Started Google-YFinance Pipeline")

            for current_date in dates_to_search:
                if stop_requested or items_collected >= limit: break
                
                date_log_str = current_date.strftime('%Y-%m-%d')
                previous_date = current_date - timedelta(days=1)
                date_query_max = f"{current_date.month}/{current_date.day}/{current_date.year}"
                date_query_min = f"{previous_date.month}/{previous_date.day}/{previous_date.year}"
                
                query = f'site:finance.yahoo.com/news/ "{raw_stock_code}"'
                encoded_query = urllib.parse.quote(query)
                google_url = f"https://www.google.com/search?q={encoded_query}&tbs=cdr:1,cd_min:{date_query_min},cd_max:{date_query_max}&hl=en&gl=us&lr=lang_en&start=0"
                
                log_message(f"[{get_time_str()}] 🔍 {base_ticker}: Searching Google for {date_log_str} (Limit: {daily_limit}/day)...")
                
                try:
                    google_page.goto(google_url, timeout=30000, wait_until="domcontentloaded")
                    
                    while "sorry" in google_page.title().lower() or google_page.locator('form[action="/errors/"]').count() > 0:
                        log_message(f"[{get_time_str()}] 🚨 CAPTCHA DETECTED! Please manually solve it in the open browser window...")
                        time.sleep(5) 

                    links_data = []
                    link_elements = google_page.locator('#search a[href*="finance.yahoo.com/news/"]').all()
                    
                    for el in link_elements:
                        href = el.get_attribute('href')
                        if not href: continue
                        
                        if "/url?q=" in href:
                            href = href.split("/url?q=")[1].split("&")[0]
                            href = urllib.parse.unquote(href)
                            
                        if not re.match(r'^https:\/\/finance\.yahoo\.com\/news\/', href):
                            continue
                            
                        if href not in processed_links:
                            h3 = el.locator('h3').first
                            google_title = h3.inner_text().strip() if h3.count() > 0 else "Yahoo Finance Article"
                            links_data.append((href, google_title))
                            processed_links.add(href)
                            
                            if len(links_data) >= daily_limit: break

                    if not links_data:
                        log_message(f"   [{get_time_str()}] ⏭️ {base_ticker}: No new links found on {date_log_str}.")
                        append_summary_report(category, raw_stock_code, "No News", current_date=date_log_str)
                        continue

                    log_message(f"   [{get_time_str()}] 📥 {base_ticker}: Extracting {len(links_data)} articles for {date_log_str}...")
                    
                    batch_new_data = []
                    for idx, (url, feed_title) in enumerate(links_data, 1):
                        if stop_requested or items_collected >= limit: break
                        
                        article_title, content, pub_date = get_article_details(context, url, feed_title)
                        
                        if not pub_date: pub_date = current_date 

                        if content:
                            log_message(f"      [{get_time_str()}] ✅ ({idx}/{len(links_data)}): Saved '{article_title[:40]}...' | URL: {url}")
                            item = {
                                'Category': category,
                                'Stock Code': raw_stock_code,
                                'Published Date': pub_date.strftime('%Y-%m-%d %H:%M:%S'),
                                'Title': article_title,
                                'Content': content,
                                'Link': url
                            }
                            batch_new_data.append(item)
                            items_collected += 1
                        else:
                            log_message(f"      [{get_time_str()}] ❌ ({idx}/{len(links_data)}): Skipped '{feed_title[:40]}...' (Empty Content) | URL: {url}")
                            
                        time.sleep(random.uniform(0.5, 1.5))

                    if batch_new_data:
                        append_to_jsonl(jsonl_filename, batch_new_data)
                        existing_data.extend(batch_new_data)
                        append_summary_report(category, raw_stock_code, "In Progress", f"Found {items_collected}", current_date=date_log_str)
                    else:
                        append_summary_report(category, raw_stock_code, "Failed Extraction", f"0/{len(links_data)} parsed", current_date=date_log_str)

                except Exception as e:
                    log_message(f"[{get_time_str()}] ⚠️ {base_ticker}: Error on {date_log_str} - {str(e)[:50]}")
                    time.sleep(5)

            completed_successfully = (items_collected >= limit) or (items_collected > 0) 

            if completed_successfully:
                create_final_excel(existing_data, excel_path)
                append_summary_report(category, raw_stock_code, "Success", f"Finished {items_collected}")
            else:
                status = "Stopped" if stop_requested else "Incomplete"
                append_summary_report(category, raw_stock_code, status, f"Finished {items_collected} (Partial)")

            if context:
                try: context.close()
                except: pass
            playwright_instance.stop()

        except Exception as e:
            log_message(f"[{get_time_str()}] ❌ Crash {base_ticker}: {e}")
            append_summary_report(category, raw_stock_code, "Crash", str(e))
            if context:
                try: context.close()
                except: pass
            if playwright_instance:
                try: playwright_instance.stop()
                except: pass

    except Exception as e:
        error_trace = traceback.format_exc()
        log_message(f"❌ FATAL THREAD CRASH for {task_info.get('stock_code', 'Unknown')}: {e}\n{error_trace}")

def main():
    if not os.path.exists(CONFIG_FILE):
        print(f"❌ Error: Please create '{CONFIG_FILE}' first.")
        return

    os.makedirs(TMP_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    config = load_config()
    max_workers = config.get('max_concurrent', 1)
    
    # --- FIX: SAFE DICTIONARY FETCHING FOR CATEGORIES ---
    categories = config.get('categories') or {}

    tasks = []
    if not categories:
        print("⚠️ WARNING: No valid categories found in your YAML config. Please uncomment or add stocks.")
        return

    for cat, stocks in categories.items():
        if not stocks: continue
        for stock in stocks:
            tasks.append({'category': cat, 'stock_code': str(stock), 'config': config})

    if not tasks:
        print("⚠️ WARNING: Task list is empty. Ensure your stocks are properly formatted in the YAML.")
        return

    log_message(f"=== STARTING YAHOO FINANCE GOOGLE PIPELINE ({max_workers} threads) ===")
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for task in tasks:
            if stop_requested: break
            futures.append(executor.submit(process_stock_task, task))
        
        try:
            for future in concurrent.futures.as_completed(futures):
                if stop_requested: executor.shutdown(wait=False, cancel_futures=True)
                try: 
                    future.result()
                except Exception as e: 
                    log_message(f"❌ Unhandled Thread Error: {e}")
        except KeyboardInterrupt:
            signal_handler(None, None)

    log_message("=== YAHOO FINANCE SESSION END ===")

if __name__ == "__main__":
    main()
