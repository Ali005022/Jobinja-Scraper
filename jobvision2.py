import os
import pickle
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import pandas as pd
import time
import random
from urllib3.exceptions import MaxRetryError
from requests.exceptions import SSLError
import ssl

# تنظیمات SSL
ssl._create_default_https_context = ssl._create_unverified_context

# مسیرهای فایل
output_path = "C:/Users/Asus/Documents/jobvision_data.xlsx"
state_file = "scraper_state.pkl"

# تنظیمات مرورگر برای حل مشکل SSL
chrome_options = Options()
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--ignore-certificate-errors")
chrome_options.add_argument("--allow-insecure-localhost")
chrome_options.add_argument("--disable-web-security")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")

# ایجاد دایرکتوری اگر وجود نداشته باشد
os.makedirs(os.path.dirname(output_path), exist_ok=True)

def init_driver():
    service = Service("C:/Users/ASUS/Desktop/chromedriver-win64/chromedriver.exe")
    return webdriver.Chrome(service=service, options=chrome_options)

def load_state():
    if os.path.exists(state_file):
        with open(state_file, 'rb') as f:
            return pickle.load(f)
    return {
        'current_page': 1,
        'processed_urls': set(),
        'saved_records': 0
    }

def save_state(state):
    with open(state_file, 'wb') as f:
        pickle.dump(state, f)

def init_excel():
    if not os.path.exists(output_path):
        pd.DataFrame(columns=[
            "عنوان شغل", "شرکت", "محل کار", "حقوق", "وضعیت",
            "لینک شغل", "صفحه", "تاریخ استخراج"
        ]).to_excel(output_path, index=False, engine='openpyxl')

def save_to_excel(new_data):
    try:
        # خواندن داده‌های موجود
        try:
            existing_data = pd.read_excel(output_path, engine='openpyxl')
        except:
            existing_data = pd.DataFrame()

        # اضافه کردن داده‌های جدید
        updated_data = pd.concat([existing_data, pd.DataFrame(new_data)], ignore_index=True)
        
        # ذخیره فایل
        updated_data.to_excel(output_path, index=False, engine='openpyxl')
        print(f"ذخیره شد. کل رکوردها: {len(updated_data)}")
        return len(updated_data)
    except Exception as e:
        print(f"خطا در ذخیره فایل: {str(e)}")
        return 0

def random_delay(min=2, max=5):
    time.sleep(random.uniform(min, max))

def safe_extract(element, selector, attribute=None, default="N/A"):
    try:
        elem = element.find_element(By.CSS_SELECTOR, selector)
        if attribute:
            return elem.get_attribute(attribute) or default
        return elem.text or default
    except NoSuchElementException:
        return default

def extract_job_data(job_element, page_num):
    try:
        # استخراج لینک با بررسی کامل
        href = safe_extract(job_element, 'a[class*="mobile-job-card"]', 'href')
        if href != "N/A":
            job_link = f"https://jobvision.ir{href.split('?')[0]}"
        else:
            job_link = "N/A"

        # استخراج حقوق
        salary = safe_extract(job_element, 'span.font-size-12px:not(.text-secondary)')
        if salary == "N/A":
            salary_div = safe_extract(job_element, 'div.d-flex.flex-wrap')
            if 'میلیون' in salary_div or 'تومان' in salary_div:
                salary = salary_div.split('|')[-1].strip()

        return {
            "عنوان شغل": safe_extract(job_element, '.job-card-title'),
            "شرکت": safe_extract(job_element, 'a.text-black.line-height-24'),
            "محل کار": safe_extract(job_element, 'span.text-secondary.pointer-events-none'),
            "حقوق": salary if salary != "N/A" else "توافقی",
            "وضعیت": "فوری" if safe_extract(job_element, '.urgent-tag') != "N/A" else "معمولی",
            "لینک شغل": job_link,
            "صفحه": page_num,
            "تاریخ استخراج": pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        }
    except Exception as e:
        print(f"خطا در استخراج داده: {str(e)}")
        return None

def scrape_page(driver, page_num, state):
    try:
        url = f"https://jobvision.ir/jobs?page={page_num}&sort=0"
        print(f"در حال پردازش صفحه {page_num} - {url}")
        
        driver.get(url)
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'job-card'))
        )
        random_delay(3, 7)
        
        job_cards = driver.find_elements(By.CSS_SELECTOR, 'job-card.col-12.row.cursor.px-0.ng-star-inserted')
        if not job_cards:
            print("صفحه خالی است - احتمالاً پایان نتایج")
            return False
        
        batch_data = []
        for job in job_cards:
            job_data = extract_job_data(job, page_num)
            if job_data:
                batch_data.append(job_data)
        
        if batch_data:
            state['saved_records'] = save_to_excel(batch_data)
            state['current_page'] = page_num + 1
            save_state(state)
        
        return True
    except Exception as e:
        print(f"خطا در پردازش صفحه {page_num}: {str(e)}")
        return False

def main():
    init_excel()
    state = load_state()
    driver = init_driver()
    
    try:
        while state['saved_records'] < 1200:  # حد نصاب رکوردها
            success = scrape_page(driver, state['current_page'], state)
            if not success:
                break
            
            random_delay(5, 10)  # تاخیر بین صفحات
            
    except KeyboardInterrupt:
        print("\nتوقف دستی توسط کاربر...")
    except Exception as e:
        print(f"خطای غیرمنتظره: {str(e)}")
    finally:
        driver.quit()
        print(f"پروسه متوقف شد. آخرین صفحه پردازش شده: {state['current_page'] - 1}")
        print(f"کل رکوردهای ذخیره شده: {state['saved_records']}")

if __name__ == "__main__":
    main()
    
