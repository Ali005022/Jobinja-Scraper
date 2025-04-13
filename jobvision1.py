import os
import pickle
from playwright.sync_api import sync_playwright, Browser, BrowserContext, Page
import pandas as pd
import time
import random
from datetime import datetime
from urllib.parse import urljoin
import logging
from typing import Optional, Dict, List, Tuple, Any

# تنظیمات پایه
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# تنظیمات اسکراپر
class Config:
    OUTPUT_PATH = "jobvision_data.xlsx"
    STATE_FILE = "scraper_state.pkl"
    MAX_RECORDS = 1200
    PAGES_PER_BROWSER = 20
    DELAYS = {
        'page_load': (3, 7),         # تأخیر بارگذاری صفحه
        'between_jobs': (2, 5),     # تأخیر بین پردازش مشاغل
        'between_pages': (5, 12),   # تأخیر بین صفحات
        'new_browser': (10, 20)     # تأخیر برای مرورگر جدید
    }
    USER_AGENTS = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.199 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.199 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:102.0) Gecko/20100101 Firefox/102.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7; rv:102.0) Gecko/20100101 Firefox/102.0",
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.5735.199 Safari/537.36",
        "Mozilla/5.0 (X11; Linux x86_64; rv:102.0) Gecko/20100101 Firefox/102.0"
    ]

class JobVisionScraper:
    def __init__(self):
        self.browser: Optional[Browser] = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None
        self.pages_scraped_in_session: int = 0
        self.init_files()
        self.state = self.load_state()

    def init_files(self) -> None:
        """آماده‌سازی فایل‌های خروجی"""
        if not os.path.exists(Config.OUTPUT_PATH):
            pd.DataFrame(columns=[
                "job_title", "company", "location", "salary", "status",
                "job_link", "page", "extraction_date", "description"
            ]).to_excel(Config.OUTPUT_PATH, index=False)

    def load_state(self) -> Dict[str, Any]:
        """بارگذاری وضعیت قبلی"""
        if os.path.exists(Config.STATE_FILE):
            try:
                with open(Config.STATE_FILE, 'rb') as f:
                    return pickle.load(f)
            except Exception as e:
                logging.error(f"Error loading state: {e}")
        return {'current_page': 1, 'saved_records': 0}

    def save_state(self) -> None:
        """ذخیره وضعیت فعلی"""
        try:
            with open(Config.STATE_FILE, 'wb') as f:
                pickle.dump(self.state, f)
        except Exception as e:
            logging.error(f"Error saving state: {e}")

    def random_delay(self, delay_type: str) -> float:
        """تاخیر تصادفی بر اساس نوع"""
        min_d, max_d = Config.DELAYS[delay_type]
        delay = random.uniform(min_d, max_d)
        time.sleep(delay)
        return delay

    def init_browser(self) -> None:
        """آماده‌سازی مرورگر جدید"""
        self.close_browser()
        
        # انتخاب User-Agent جدید
        current_user_agent = random.choice(Config.USER_AGENTS)
        logging.info(f"استفاده از User-Agent: {current_user_agent}")
        
        playwright = sync_playwright().start()
        self.browser = playwright.chromium.launch(
            headless=False,
            args=[
                "--disable-gpu",
                "--disable-extensions",
                "--disable-blink-features=AutomationControlled",
                "--start-maximized"
            ]
        )
        self.context = self.browser.new_context(
            user_agent=current_user_agent,
            viewport={'width': 1920, 'height': 1080},
            java_script_enabled=True,
            bypass_csp=True
        )
        self.context.route("**/*", lambda route, request: route.abort() if request.resource_type == "image" else route.continue_())
        
        self.page = self.context.new_page()  # مقداردهی صفحه جدید
        self.pages_scraped_in_session = 0
        logging.info("مرورگر جدید راه‌اندازی شد")

    def close_browser(self) -> None:
        """بستن مرورگر فعلی"""
        if self.context:
            self.context.close()
        if self.browser:
            self.browser.close()  # تغییر از stop به close
        self.browser = None
        self.context = None
        self.page = None
        logging.info("مرورگر فعلی بسته شد")

    def scrape_page(self, page_num: int) -> bool:
        """استخراج داده‌های یک صفحه"""
        try:
            if not self.page or self.pages_scraped_in_session >= Config.PAGES_PER_BROWSER:
                self.init_browser()
                if not self.page:
                    raise RuntimeError("Page initialization failed")

            url = f"https://jobvision.ir/jobs?page={page_num}&sort=0"
            logging.info(f"در حال پردازش صفحه {page_num} - {url}")
            
            # تأخیر تصادفی قبل از بارگذاری صفحه
            delay = self.random_delay('page_load')
            logging.debug(f"تاخیر {delay:.1f} ثانیه قبل از بارگذاری صفحه")
            
            self.page.goto(url, timeout=30000)  # type: ignore
            self.page.wait_for_selector('job-card', state='attached', timeout=15000)  # type: ignore
            
            # اسکرول به پایین صفحه با تأخیر تصادفی
            self.page.evaluate("window.scrollTo(0, document.body.scrollHeight);")
            delay = self.random_delay('between_jobs')
            logging.debug(f"تاخیر {delay:.1f} ثانیه پس از اسکرول")
            
            # تأخیر تصادفی پس از بارگذاری صفحه
            delay = self.random_delay('page_load')
            logging.debug(f"تاخیر {delay:.1f} ثانیه پس از بارگذاری صفحه")
            
            job_cards = self.page.query_selector_all('job-card.col-12.row.cursor.px-0.ng-star-inserted')  # type: ignore
            if not job_cards:
                logging.warning(f"صفحه {page_num} خالی است")
                return False
            
            batch_data = []
            for job in job_cards:
                # تأخیر تصادفی بین پردازش هر شغل
                delay = self.random_delay('between_jobs')
                logging.debug(f"تاخیر {delay:.1f} ثانیه بین پردازش مشاغل")
                
                job_data = self.extract_job_data(job, page_num)
                if job_data:
                    batch_data.append(job_data)
            
            if batch_data:
                self.save_data(batch_data)
                self.state['current_page'] = page_num + 1
                self.state['saved_records'] += len(batch_data)
                self.save_state()
                self.pages_scraped_in_session += 1
            
            # تأخیر تصادفی پس از پردازش صفحه
            delay = self.random_delay('between_pages')
            logging.debug(f"تاخیر {delay:.1f} ثانیه پس از پردازش صفحه")
            
            return True
            
        except Exception as e:
            logging.error(f"خطا در پردازش صفحه {page_num}: {str(e)}")
            return False

    def extract_job_data(self, job_element: Any, page_num: int) -> Optional[Dict[str, Any]]:
        """استخراج داده‌های یک شغل"""
        try:
            href = job_element.get_attribute('href')
            job_link = urljoin("https://jobvision.ir", href.split('?')[0]) if href else "N/A"
            
            title_element = job_element.query_selector('.job-card-title')
            title = title_element.inner_text().strip() if title_element else "N/A"
            
            company_element = job_element.query_selector('a.text-black.line-height-24')
            company = company_element.inner_text().strip() if company_element else "N/A"
            
            location_element = job_element.query_selector('span.text-secondary.pointer-events-none')
            location = location_element.inner_text().strip() if location_element else "N/A"
            
            salary_element = job_element.query_selector('span.font-size-12px:not(.text-secondary)')
            salary = salary_element.inner_text().strip() if salary_element else "N/A"
            
            if salary == "N/A":
                salary_div = job_element.query_selector('div.d-flex.flex-wrap')
                if salary_div:
                    salary_text = salary_div.inner_text()
                    if 'میلیون' in salary_text or 'تومان' in salary_text:
                        salary = salary_text.split('|')[-1].strip()
            
            status = "Urgent" if job_element.query_selector('.urgent-tag') else "Normal"
            
            return {
                "job_title": title,
                "company": company,
                "location": location,
                "salary": salary if salary != "N/A" else "Negotiable",
                "status": status,
                "job_link": job_link,
                "page": page_num,
                "extraction_date": datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
        except Exception as e:
            logging.error(f"خطا در استخراج شغل: {str(e)}")
            return None

    def save_data(self, new_data: List[Dict[str, Any]]) -> None:
        """ذخیره داده‌ها در فایل"""
        try:
            existing_data = pd.read_excel(Config.OUTPUT_PATH) if os.path.exists(Config.OUTPUT_PATH) else pd.DataFrame()
            updated_data = pd.concat([existing_data, pd.DataFrame(new_data)], ignore_index=True)
            updated_data.to_excel(Config.OUTPUT_PATH, index=False)
            logging.info(f"داده‌ها ذخیره شدند. کل رکوردها: {len(updated_data)}")
        except Exception as e:
            logging.error(f"خطا در ذخیره داده‌ها: {str(e)}")

    def run(self) -> None:
        """اجرای اصلی اسکراپر"""
        try:
            self.init_browser()
            
            while self.state['saved_records'] < Config.MAX_RECORDS:
                success = self.scrape_page(self.state['current_page'])
                if not success:
                    break
                
                delay = self.random_delay('between_pages')
                logging.info(f"تاخیر {delay:.1f} ثانیه قبل از صفحه بعد...")

                # بررسی تعداد صفحات پردازش شده و باز کردن مرورگر جدید با User-Agent جدید
                if self.pages_scraped_in_session >= Config.PAGES_PER_BROWSER:
                    self.close_browser()
                    delay = self.random_delay('new_browser')
                    logging.info(f"تغییر مرورگر و User-Agent پس از {Config.PAGES_PER_BROWSER} صفحه - تاخیر {delay:.1f} ثانیه")
                    self.init_browser()  # باز کردن مرورگر جدید با User-Agent جدید
            
            logging.info(f"استخراج کامل شد. کل رکوردها: {self.state['saved_records']}")
        except KeyboardInterrupt:
            logging.info("توقف دستی توسط کاربر")
        except Exception as e:
            logging.error(f"خطای غیرمنتظره: {str(e)}")
        finally:
            self.close_browser()
            logging.info(f"آخرین صفحه پردازش شده: {self.state['current_page'] - 1}")

if __name__ == "__main__":
    scraper = JobVisionScraper()
    scraper.run()