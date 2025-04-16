import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

#---Config---
SEARCH_QUERY = "Yamaha MT-07"
EXCEL_FILE = "yamaha_mt07_ads.xlsx"

# Chrome Setup
chrome_options = Options()
# chrome_options.add_argument("--headless")  # DISABLE headless to see browser
chrome_options.add_argument("--start-maximized")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=chrome_options)

# Open OLX
driver.get("https://www.olx.ro")
time.sleep(3)

# Search for the query
search_box = driver.find_element(By.CSS_SELECTOR, 'input[placeholder*="Ce anume cauți?"]')
search_box.send_keys(SEARCH_QUERY)
search_box.send_keys(Keys.RETURN)
time.sleep(4)

# Scroll slowly to ensure all listings are loaded
for _ in range(5):
    driver.execute_script("window.scrollBy(0, 1000);")
    time.sleep(1.5)

# Extract Listings
ads = driver.find_elements(By.CSS_SELECTOR, 'div[data-testid="listing-grid"] > div')
new_ads = []

print(f"Found {len(ads)} ads...")

for i, ad in enumerate(ads):
    try:

        try:
            link_elem = ad.find_element(By.CSS_SELECTOR, 'a')
            url_suffix = link_elem.get_attribute("href")
            url = "https://www.olx.ro" + url_suffix if url_suffix.startswith("/d/") else url_suffix
        except Exception as e:
            print(f"Skipping ad {i+1}: Link element not found")
            continue

        img_elem = link_elem.find_element(By.TAG_NAME, "img")
        title = img_elem.get_attribute("alt").strip()

        # Price (fallback: find <p> with 'lei' or €)
        try:
            price_elem = ad.find_element(By.XPATH, './/p[contains(text(),"€") or contains(text(),"lei")]')
            price = price_elem.text.strip()
        except:
            price = "N/A"

        # Location
        try:
            location_elem = ad.find_element(By.CSS_SELECTOR, 'p[data-testid="location-date"]')
            location = location_elem.text.split(" - ")[0].strip()
        except:
            location = "N/A"

        new_ads.append({
            "Title": title,
            "URL": url,
            "Price": price,
            "Location": location
        })

    except Exception as e:
        print(f"Skipping ad {i+1}: {str(e)}")


if os.path.exists(EXCEL_FILE):
    df_existing = pd.read_excel(EXCEL_FILE)
    existing_urls = set(df_existing["URL"])
else:
    df_existing = pd.DataFrame()
    existing_urls = set()


fresh_ads = [ad for ad in new_ads if ad["URL"] not in existing_urls]

if fresh_ads:
    print(f"{len(fresh_ads)} new listings found:")
    df_new = pd.DataFrame(fresh_ads)
    df_combined = pd.concat([df_existing, df_new], ignore_index=True)
    df_combined.drop_duplicates(subset="URL", inplace=True)
    df_combined.to_excel(EXCEL_FILE, index=False)
    print(f" Updated Excel: {EXCEL_FILE}")
else:
    print("No new listings found.")
