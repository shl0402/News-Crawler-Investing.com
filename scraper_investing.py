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
LOG_FILE = "scraper_progress_log.txt"
SUMMARY_FILE = "scraper_summary_report.csv"
stop_requested = False

# Thread locks
log_lock = threading.Lock()
csv_lock = threading.Lock()

# --- SIGNAL HANDLER (Force Quit Support) ---
def signal_handler(sig, frame):
    global stop_requested
    if not stop_requested:
        print("\n🛑 STOP REQUESTED! Finishing active pages... (Press Ctrl+C again to KILL IMMEDIATELY)")
        stop_requested = True
    else:
        print("\n💀 FORCE QUITTING...")
        os._exit(1) # Force kill everything

signal.signal(signal.SIGINT, signal_handler)

def load_config():
    with open('config.yaml', 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

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

# --- REPORTING ---
def append_summary_report(category, stock_code, status, details=""):
    """Writes to the CSV report."""
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
        try:
            df.to_csv(SUMMARY_FILE, mode='a', header=header, index=False)
        except: pass

# --- DATA MANAGEMENT ---
def get_filenames(category, stock_code, stock_name=""):
    s_cat = sanitize_filename(category)
    s_code = sanitize_filename(stock_code)
    jsonl = f"temp_{s_cat}_{s_code}.jsonl"
    
    if stock_name:
        s_name = sanitize_filename(stock_name)
        excel = f"{s_cat}_{s_code}_{s_name}.xlsx"
    else:
        excel = f"{s_cat}_{s_code}_" 
    
    return jsonl, excel

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

def create_final_excel(data, excel_filename):
    if not data: return
    
    df = pd.DataFrame(data)
    df.drop_duplicates(subset=['Link'], keep='first', inplace=True)
    
    if 'Content' in df.columns:
        df['Content'] = df['Content'].astype(str).str.slice(0, 32000)

    cols = ['Category', 'Stock Code', 'Stock Name', 'Slug', 'Published Date', 'Updated Date', 'Title', 'Content', 'Link']
    for c in cols: 
        if c not in df.columns: df[c] = ""
    
    try:
        df[cols].to_excel(excel_filename, index=False)
        log_message(f"💾 EXCEL SAVED: {excel_filename}")
        return True
    except Exception as e:
        log_message(f"❌ Excel Save Failed: {e}")
        return False

# --- UTILS ---
def resolve_slug(page, stock_code):
    slug = None
    stock_name = None
    try:
        page.goto(f"https://www.investing.com/search/?q={stock_code}", timeout=30000, wait_until="domcontentloaded")
        try:
            page.wait_for_selector('.searchSection', timeout=5000)
            result_link = page.locator('.searchSection .js-inner-all-results-quotes-wrapper a').first
            if result_link.count() > 0:
                href = result_link.get_attribute('href')
                name_text = result_link.locator('.second').inner_text()
                if not name_text: name_text = result_link.locator('.third').inner_text()
                
                if '/equities/' in href:
                    slug = href.split('/equities/')[-1].split('?')[0]
                    stock_name = name_text.strip()
        except: pass
    except: pass
    return slug, stock_name

def parse_date(date_str):
    if not date_str: return None
    date_str = date_str.strip().replace('By ', '').replace('• ', '')
    now = datetime.now()
    try:
        match = re.search(r'(\d{2}/\d{2}/\d{4}, \d{2}:\d{2} [AP]M)', date_str)
        if match: return datetime.strptime(match.group(1), '%m/%d/%Y, %I:%M %p')
    except: pass
    try:
        if '-' in date_str and ':' in date_str:
            return datetime.strptime(date_str, '%Y-%m-%d %H:%M:%S')
    except: pass
    if 'ago' in date_str.lower():
        try:
            val = int(re.search(r'(\d+)', date_str).group(1))
            if 'hour' in date_str.lower(): return now - timedelta(hours=val)
            elif 'minute' in date_str.lower(): return now - timedelta(minutes=val)
            return now
        except: pass
    try:
        clean_date = re.search(r'([A-Z][a-z]{2} \d{1,2}, \d{4})', date_str)
        if clean_date: return datetime.strptime(clean_date.group(1), '%b %d, %Y')
    except: pass
    return None

def get_article_details(context, url, fallback_date):
    page = None
    published_date = fallback_date
    updated_date = None
    clean_text = ""
    try:
        page = context.new_page()
        page.route("**/*.{png,jpg,jpeg,svg,woff,woff2,gif,ico}", lambda route: route.abort())
        page.goto(url, timeout=30000, wait_until="domcontentloaded")
        
        try:
            pub_el = page.get_by_text(re.compile(r"Published \d{2}/\d{2}/\d{4}")).first
            if pub_el.count() > 0:
                p_date = parse_date(pub_el.inner_text().replace("Published", ""))
                if p_date: published_date = p_date
            upd_el = page.get_by_text(re.compile(r"Updated \d{2}/\d{2}/\d{4}")).first
            if upd_el.count() > 0:
                u_date = parse_date(upd_el.inner_text().replace("Updated", ""))
                if u_date: updated_date = u_date
        except: pass

        selector = 'div[class*="article_WYSIWYG"]'
        try: page.wait_for_selector(selector, timeout=5000)
        except: selector = 'div.articlePage'

        page.evaluate(f"""() => {{
            const content = document.querySelector('{selector}');
            if (!content) return;
            const junk = content.querySelectorAll('div[data-test="ad-slot-visible"], div[data-test="contextual-subscription-hook"], div[id="mid-article-hook"]');
            junk.forEach(el => el.remove());
        }}""")

        paragraphs = page.locator(selector).first.locator('p').all_inner_texts()
        clean_text = "\n\n".join([p.strip() for p in paragraphs if p.strip()])
    except: pass
    finally:
        if page: 
            try: page.close()
            except: pass
    return clean_text, published_date, updated_date

# --- WORKER ---
def process_stock_task(task_info):
    global stop_requested
    category = task_info['category']
    stock_code = str(task_info['stock_code'])
    config = task_info['config']
    limit = config['limit_items']
    
    if stop_requested: return

    # 1. SMART RESUME
    jsonl_filename, excel_prefix = get_filenames(category, stock_code)
    
    # Check Final Excel
    for f in os.listdir('.'):
        if f.startswith(excel_prefix) and f.endswith(".xlsx") and stock_code in f:
            log_message(f"⏭️ Skipping {stock_code} (Final Excel exists)")
            append_summary_report(category, stock_code, "Skipped", "Excel Exists")
            return

    # Load JSONL
    existing_data = load_existing_data(jsonl_filename)
    unique_links = set(item['Link'] for item in existing_data)
    items_collected = len(unique_links)
    
    stock_name = "Unknown"
    if existing_data:
        stock_name = existing_data[0].get('Stock Name', 'Unknown')

    if items_collected >= limit:
        log_message(f"✅ {stock_code}: Has {items_collected}/{limit}. Generating Excel...")
        _, final_excel_name = get_filenames(category, stock_code, stock_name)
        create_final_excel(existing_data, final_excel_name)
        append_summary_report(category, stock_code, "Success", f"Finished ({items_collected})")
        return

    if items_collected > 0:
        log_message(f"🔄 Resuming {stock_code}: {items_collected}/{limit}")

    # 2. START SCRAPING
    start_dt = datetime.strptime(config['start_date'], '%Y-%m-%d')
    end_dt = datetime.strptime(config['end_date'], '%Y-%m-%d')
    headless = config.get('headless', False)

    try:
        with sync_playwright() as p:
            browser = p.firefox.launch(headless=headless)
            context = browser.new_context(viewport={'width': 1920, 'height': 1080})
            context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
            page = context.new_page()

            slug = stock_code
            if '-' not in stock_code or stock_code.isdigit():
                found_slug, found_name = resolve_slug(page, stock_code)
                if found_slug: 
                    slug = found_slug
                    stock_name = found_name
                else:
                    log_message(f"⚠️ Slug not found for {stock_code}")
                    append_summary_report(category, stock_code, "Failed", "Slug Not Found")
                    browser.close()
                    return
            
            log_message(f"🚀 Started: {stock_code} ({stock_name})")
            
            page_num = 1
            completed_successfully = False
            page.route("**/*.{png,jpg,jpeg,svg,woff,woff2,gif,ico}", lambda route: route.abort())

            while items_collected < limit and not stop_requested:
                url = f"https://www.investing.com/equities/{slug}-news/{page_num}"
                try:
                    page.goto(url, timeout=45000, wait_until="domcontentloaded")
                    if "Just a moment" in page.title(): page.wait_for_timeout(10000)

                    try: page.wait_for_selector('a[data-test="article-title-link"]', timeout=5000)
                    except: 
                        completed_successfully = True
                        break 
                    
                    links = page.locator('a[data-test="article-title-link"]').all()
                    if not links: 
                        completed_successfully = True
                        break

                    batch_new_data = []

                    for link in links:
                        if items_collected >= limit or stop_requested: break
                        
                        try:
                            href = link.get_attribute('href')
                            full_link = href if href.startswith('http') else "https://www.investing.com" + href
                            
                            if full_link in unique_links: continue

                            container = link.locator('xpath=./ancestor::li | ./ancestor::article').first
                            prelim_date_str = None
                            if container.locator('[data-test="article-publish-date"]').count() > 0:
                                prelim_date_str = container.locator('[data-test="article-publish-date"]').first.get_attribute('datetime')
                            elif container.locator('time').count() > 0:
                                prelim_date_str = container.locator('time').first.get_attribute('datetime')
                            
                            prelim_date = parse_date(prelim_date_str)
                            if not prelim_date: continue

                            if start_dt <= prelim_date <= end_dt:
                                title = link.inner_text().strip()
                                content, pub_date, upd_date = get_article_details(context, full_link, prelim_date)
                                final_date = pub_date if pub_date else prelim_date
                                
                                if start_dt <= final_date <= end_dt:
                                    item = {
                                        'Category': category,
                                        'Stock Code': stock_code,
                                        'Stock Name': stock_name,
                                        'Slug': slug,
                                        'Published Date': final_date.strftime('%Y-%m-%d %H:%M:%S'),
                                        'Updated Date': upd_date.strftime('%Y-%m-%d %H:%M:%S') if upd_date else "",
                                        'Title': title,
                                        'Content': content,
                                        'Link': full_link
                                    }
                                    batch_new_data.append(item)
                                    unique_links.add(full_link)
                                    items_collected += 1
                                    print(f"   + {stock_code}: {items_collected}/{limit}")
                                
                                time.sleep(random.uniform(0.5, 1.5))
                            elif prelim_date < start_dt:
                                completed_successfully = True
                                items_collected = limit + 999 
                                break
                        except: pass
                    
                    if batch_new_data:
                        append_to_jsonl(jsonl_filename, batch_new_data)
                        existing_data.extend(batch_new_data)
                        # --- REPORT UPDATE (THE FIX) ---
                        append_summary_report(category, stock_code, "In Progress", f"Found {items_collected}")
                    
                    page_num += 1
                    
                except Exception as e:
                    log_message(f"⚠️ Page Loop Error {stock_code}: {e}")
                    break
            
            # 3. FINALIZE
            if items_collected >= limit: completed_successfully = True

            if completed_successfully:
                _, final_excel_name = get_filenames(category, stock_code, stock_name)
                create_final_excel(existing_data, final_excel_name)
                append_summary_report(category, stock_code, "Success", f"Found {items_collected}")
            else:
                status = "Stopped" if stop_requested else "Incomplete"
                append_summary_report(category, stock_code, status, f"Found {items_collected} (Partial)")

            browser.close()

    except Exception as e:
        log_message(f"❌ Crash {stock_code}: {e}")
        append_summary_report(category, stock_code, "Crash", str(e))

def main():
    config = load_config()
    max_workers = config.get('max_concurrent', 1)
    categories = config.get('categories', {})

    tasks = []
    for cat, stocks in categories.items():
        for stock in stocks:
            tasks.append({'category': cat, 'stock_code': stock, 'config': config})

    log_message(f"=== STARTING ({max_workers} threads) ===")
    
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