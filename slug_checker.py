import yaml
import pandas as pd
import concurrent.futures
import re
import os
from playwright.sync_api import sync_playwright

# --- CONFIGURATION ---
CONFIG_FILE = 'config.yaml'
REPORT_FILE = 'slug_validation_report.csv'
MAX_THREADS = 5  # Can be higher since we aren't scraping deep content

def load_config():
    with open(CONFIG_FILE, 'r', encoding='utf-8') as file:
        return yaml.safe_load(file)

def resolve_slug(page, stock_code):
    """
    Attempts to resolve a stock code to a valid Investing.com slug.
    Returns: (status, slug, stock_name, error_message)
    """
    try:
        # Navigate to search
        page.goto(f"https://www.investing.com/search/?q={stock_code}", timeout=15000, wait_until="domcontentloaded")
        
        # Check for immediate redirect (sometimes searching a slug redirects to the page)
        if "/equities/" in page.url:
            slug = page.url.split("/equities/")[-1].split("?")[0]
            return "Valid", slug, "Direct Redirect", ""

        # Wait for search results
        try:
            page.wait_for_selector('.searchSection', timeout=5000)
        except:
            return "Invalid", None, None, "No search results found"

        # Grab first result in Equities section
        result_link = page.locator('.searchSection .js-inner-all-results-quotes-wrapper a').first
        
        if result_link.count() > 0:
            href = result_link.get_attribute('href')
            name_text = result_link.locator('.second').inner_text()
            if not name_text: name_text = result_link.locator('.third').inner_text()
            
            if '/equities/' in href:
                slug = href.split('/equities/')[-1].split('?')[0]
                stock_name = name_text.strip()
                return "Valid", slug, stock_name, ""
            else:
                return "Invalid", None, None, f"Top result was not an equity: {href}"
        else:
            return "Invalid", None, None, "Search returned 0 results"

    except Exception as e:
        return "Error", None, None, str(e)

def check_stock_task(task_info):
    category = task_info['category']
    stock_code = str(task_info['stock_code'])
    headless = task_info['headless']
    
    # If code already looks like a slug (has letters/hyphens), verify it directly
    is_direct_slug = '-' in stock_code or not stock_code.isdigit()
    
    with sync_playwright() as p:
        browser = p.firefox.launch(headless=headless)
        page = context = browser.new_context(user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        page = context.new_page()
        
        try:
            if is_direct_slug:
                # Direct check: Does the URL exist?
                response = page.goto(f"https://www.investing.com/equities/{stock_code}", timeout=15000, wait_until="domcontentloaded")
                if response.status == 200 and "404" not in page.title():
                    status, slug, name, err = "Valid", stock_code, stock_code, ""
                else:
                    status, slug, name, err = "Invalid", stock_code, None, "404 Page Not Found"
            else:
                # Search check
                status, slug, name, err = resolve_slug(page, stock_code)
                
            print(f"[{status}] {category} - {stock_code} -> {slug or 'N/A'}")
            return {
                "Category": category,
                "Input Code": stock_code,
                "Status": status,
                "Resolved Slug": slug,
                "Stock Name": name,
                "Error": err
            }
            
        finally:
            browser.close()

def main():
    print("--- 🔍 STARTING SLUG VALIDATION CHECKER ---")
    config = load_config()
    categories = config.get('categories', {})
    headless = config.get('headless', True) # Default to headless for speed

    tasks = []
    for cat, stocks in categories.items():
        for stock in stocks:
            tasks.append({'category': cat, 'stock_code': stock, 'headless': headless})

    results = []
    
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        future_to_stock = {executor.submit(check_stock_task, task): task for task in tasks}
        
        for future in concurrent.futures.as_completed(future_to_stock):
            try:
                data = future.result()
                results.append(data)
            except Exception as e:
                print(f"Wrapper Error: {e}")

    # Save Report
    df = pd.DataFrame(results)
    
    # Sort so 'Invalid' appear at the top for easy fixing
    df.sort_values(by=['Status', 'Category'], ascending=[True, True], inplace=True)
    
    df.to_csv(REPORT_FILE, index=False)
    print(f"\n✅ Validation Complete! Report saved to: {REPORT_FILE}")
    print("\n--- SUMMARY ---")
    print(df['Status'].value_counts())
    
    # Print Invalid ones for immediate view
    invalid_df = df[df['Status'] != 'Valid']
    if not invalid_df.empty:
        print("\n⚠️  ATTENTION REQUIRED FOR THESE STOCKS:")
        print(invalid_df[['Category', 'Input Code', 'Error']].to_string(index=False))
    else:
        print("\n🎉 All codes resolved successfully!")

if __name__ == "__main__":
    main()