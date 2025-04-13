import os
import time
import random
import pickle
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, scrolledtext, Spinbox
from datetime import datetime
from threading import Thread, Event, Lock
from typing import Optional, Dict, List, Any, Hashable, Sequence, Callable, Union
import urllib.parse
import schedule
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (TimeoutException, WebDriverException, 
                                     NoSuchElementException, StaleElementReferenceException)

# Constants
STATUS_FILE = "scraping_status.pkl"
CONFIG_FILE = "scraper_config.json"
MAX_RETRIES = 5
BACKUP_DIR = "backups"
MAX_MATCHES = 5
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0"
]

class JobScraper:
    def __init__(self, gui: Optional['JobScraperGUI'] = None):
        self.driver: Optional[webdriver.Chrome] = None
        self.gui = gui
        self.base_url = "https://jobinja.ir/jobs/latest-job-post-%D8%A7%D8%B3%D8%AA%D8%AE%D8%AF%D8%A7%D9%85%DB%8C-%D8%AC%D8%AF%DB%8C%D8%AF"
        self.url_params = "preferred_before=1743954204&sort_by=published_at_desc"
        self.current_user_agent = random.choice(USER_AGENTS)
        self.paused = Event()
        self.stopped = Event()
        self.new_jobs_paused = Event()
        self.new_jobs_stopped = Event()
        self.pause_lock = Lock()
        
    def extract_job_slug(self, url: str) -> str:
        """
        Extract the job title portion from a Jobinja URL
        Example: 
        Input: "https://jobinja.ir/companies/ceres/jobs/AEXN/استخدام-ویتر-ویترس-رستوران-در-کافه-رستوران-س-ر-س?_ref=16&_t=352e3131362e3134392e3235"
        Output: "استخدام-ویتر-ویترس-رستوران-در-کافه-رستوران-س-ر-س"
        """
        try:
            # Split URL to get the part after '/jobs/'
            parts = url.split('/jobs/')
            if len(parts) > 1:
                # Get the part before query parameters
                slug_part = parts[1].split('?')[0]
                # Further split to get just the title portion (after job ID)
                slug = slug_part.split('/')[1] if '/' in slug_part else slug_part
                return urllib.parse.unquote(slug)  # Decode URL-encoded characters
            return url  # Fallback to full URL if pattern not found
        except Exception:
            return url  # Fallback to full URL on any error

    def is_duplicate(self, new_job: Dict[str, str], 
                    existing_jobs: Sequence[Dict[Hashable, Any]], 
                    check_first_n: int = 5) -> bool:
        """
        Enhanced duplicate check using:
        1. Job title slugs from URLs (primary check)
        2. Title + Company comparison (secondary check)
        3. Full URL comparison (fallback)
        """
        new_slug = self.extract_job_slug(new_job['Link'])
        
        for existing_job in existing_jobs[:check_first_n]:
            existing_slug = self.extract_job_slug(existing_job['Link'])
            
            # 1. Primary check: Compare job title slugs from URLs
            if new_slug == existing_slug:
                return True
                
            # 2. Secondary check: Compare titles and companies (case-insensitive)
            if (new_job['Title'].lower() == existing_job['Title'].lower() and 
                new_job['Company'].lower() == existing_job['Company'].lower()):
                return True
                
            # 3. Fallback: Full URL comparison
            if new_job['Link'] == existing_job['Link']:
                return True
                
        return False

    def initialize_driver(self) -> None:
        """Initialize Chrome WebDriver only when needed"""
        if self.driver is not None:
            return
            
        chrome_options = Options()
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument(f"user-agent={self.current_user_agent}")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        try:
            service = Service(executable_path="C:/Users/ASUS/Desktop/chromedriver-win64/chromedriver.exe")
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.driver.set_page_load_timeout(30)
            self.log("WebDriver initialized")
        except Exception as e:
            self.log(f"Failed to initialize WebDriver: {str(e)}")
            raise

    def get_page_url(self, page_number: int = 1) -> str:
        """Generate URL for specific page number"""
        if page_number == 1:
            return f"{self.base_url}?{self.url_params}"
        return f"{self.base_url}?&page={page_number}&{self.url_params}"

    def random_delay(self, min_seconds: float = 3, max_seconds: float = 7) -> None:
        """Random delay between requests"""
        time.sleep(random.uniform(min_seconds, max_seconds))

    def rotate_user_agent(self) -> None:
        """Rotate to a different user agent"""
        if self.driver is None:
            return
            
        new_ua = random.choice([ua for ua in USER_AGENTS if ua != self.current_user_agent])
        self.current_user_agent = new_ua
        self.driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": new_ua})
        self.log(f"Rotated User Agent to: {new_ua[:50]}...")

    def load_status(self) -> Optional[Dict[str, Any]]:
        """Load previous scraping status"""
        try:
            if os.path.exists(STATUS_FILE):
                with open(STATUS_FILE, "rb") as f:
                    return pickle.load(f)
        except Exception as e:
            self.log(f"Error loading status: {str(e)}")
        return None

    def save_status(self, page_count: int, output_file: str, backup_file: str) -> None:
        """Save current scraping status"""
        try:
            with open(STATUS_FILE, "wb") as f:
                pickle.dump({
                    "page_count": page_count,
                    "output_file": output_file,
                    "backup_file": backup_file
                }, f)
            self.log(f"Status saved: Page {page_count}")
        except Exception as e:
            self.log(f"Error saving status: {str(e)}")

    def go_to_page(self, page_number: int) -> bool:
        """Navigate to specific page"""
        if self.driver is None:
            self.log("WebDriver not initialized")
            return False
        
        for attempt in range(MAX_RETRIES):
            try:
                url = self.get_page_url(page_number)
                self.log(f"Loading page: {url}")
                self.driver.get(url)
                self.random_delay()
                
                WebDriverWait(self.driver, 20).until(
                    lambda d: d.find_elements(By.CSS_SELECTOR, ".o-listView__itemInfo") or 
                             d.find_elements(By.CSS_SELECTOR, ".paginator")
                )
                return True
            except TimeoutException:
                self.log(f"Timeout on page {page_number}, attempt {attempt + 1}")
                if attempt == MAX_RETRIES - 1:
                    return False
                self.random_delay(5, 10)
            except Exception as e:
                self.log(f"Error loading page: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    return False
                self.random_delay(5, 10)
        return False

    def get_max_pages(self) -> int:
        """Get total number of pages available"""
        if self.driver is None:
            self.log("WebDriver not initialized")
            return 1

        for attempt in range(MAX_RETRIES):
            try:
                pagination = WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, ".paginator"))
                )
                page_links = pagination.find_elements(By.CSS_SELECTOR, "li a")
                page_numbers = [int(link.text) for link in page_links if link.text.isdigit()]
                return max(page_numbers) if page_numbers else 1
            except TimeoutException:
                self.log(f"Timeout detecting pages, attempt {attempt + 1}")
                if attempt == MAX_RETRIES - 1:
                    return 1
                self.random_delay(3, 5)
            except Exception as e:
                self.log(f"Error detecting pages: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    return 1
                self.random_delay(3, 5)
        return 1

    def scrape_page(self) -> List[Dict[str, str]]:
        """Scrape job listings from current page"""
        if self.driver is None:
            self.log("WebDriver not initialized")
            return []

        jobs = []
        try:
            job_elements = WebDriverWait(self.driver, 20).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".o-listView__itemInfo"))
            )

            for job in job_elements:
                try:
                    title = job.find_element(By.CSS_SELECTOR, ".c-jobListView__titleLink").text.strip()
                    company = job.find_element(By.XPATH, ".//span[contains(text(), '|')]").text.strip()
                    location = job.find_element(By.XPATH, ".//span[contains(text(), '،')]").text.strip()
                    contract = job.find_element(By.XPATH, ".//span[contains(text(), 'قرارداد')]").text.strip()
                    link = job.find_element(By.CSS_SELECTOR, ".c-jobListView__titleLink").get_attribute("href")

                    jobs.append({
                        'Title': title,
                        'Company': company,
                        'Location': location,
                        'Contract Type': contract,
                        'Link': link,
                    })
                except NoSuchElementException:
                    continue
                except Exception as e:
                    self.log(f"Error parsing job: {str(e)}")
                    continue

        except TimeoutException:
            self.log("Timeout waiting for job listings")
        except Exception as e:
            self.log(f"Error scraping page: {str(e)}")

        return jobs

    def go_to_next_page(self) -> bool:
        """Navigate to next page if available"""
        if self.driver is None:
            self.log("WebDriver not initialized")
            return False

        for attempt in range(MAX_RETRIES):
            try:
                next_btn = WebDriverWait(self.driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a[rel='next']"))
                )
                
                if next_btn is None:
                    self.log("Next button not found")
                    return False
                
                next_btn_class = next_btn.get_attribute("class") or ""
                if "disabled" in next_btn_class:
                    return False
                
                self.driver.execute_script("arguments[0].click();", next_btn)
                self.random_delay(2, 4)
                
                WebDriverWait(self.driver, 20).until(
                    lambda d: d.find_elements(By.CSS_SELECTOR, ".o-listView__itemInfo") or 
                             d.find_elements(By.CSS_SELECTOR, ".paginator")
                )
                return True
                
            except TimeoutException:
                self.log(f"Timeout on next page, attempt {attempt + 1}")
                if attempt == MAX_RETRIES - 1:
                    return False
                self.random_delay(3, 6)
            except Exception as e:
                self.log(f"Error going to next page: {str(e)}")
                if attempt == MAX_RETRIES - 1:
                    return False
                self.random_delay(3, 6)
        return False

    def save_data(self, data: Sequence[Dict[str, Any]], output_file: str, 
                 existing_data: Optional[Sequence[Dict[Hashable, Any]]] = None) -> str:
        """Save data with backup"""
        try:
            combined = list(existing_data) + list(data) if existing_data else list(data)
            df = pd.DataFrame(combined)
            
            os.makedirs(os.path.dirname(output_file), exist_ok=True)
            
            # Main save
            df.to_excel(output_file, index=False, engine='openpyxl')
            self.log(f"Data saved to {output_file}")
            
            # Immediate backup
            backup_path = f"{os.path.splitext(output_file)[0]}_backup.xlsx"
            df.to_excel(backup_path, index=False, engine='openpyxl')
            
            # Timestamped backup
            backup_dir = os.path.join(os.path.dirname(output_file), BACKUP_DIR)
            os.makedirs(backup_dir, exist_ok=True)
            timestamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')
            timestamped_backup = os.path.join(backup_dir, f"backup_{timestamp}.xlsx")
            df.to_excel(timestamped_backup, index=False, engine='openpyxl')
            
            return backup_path
            
        except Exception as e:
            self.log(f"Error saving data: {str(e)}")
            raise

    def scrape_new_jobs(self, reference_file: str, output_file: str,
                       progress_callback: Optional[Callable[[int, int], None]] = None) -> None:
        """Enhanced New Jobs Only mode with proper pause/resume functionality"""
        try:
            # Initialize driver only when needed
            self.initialize_driver()
            
            # Reset pause/stop events
            self.new_jobs_paused.clear()
            self.new_jobs_stopped.clear()
            
            # Load existing data
            existing_df = pd.read_excel(reference_file)
            existing_jobs = existing_df.to_dict('records')
            self.log(f"Loaded {len(existing_jobs)} jobs (checking first 5 for duplicates)")

            # Initialize variables
            new_jobs = []
            current_page = 1
            matches_found = 0
            
            if not self.go_to_page(current_page):
                return
                
            max_pages = self.get_max_pages()
            self.log(f"Total pages to check: {max_pages}")

            while (current_page <= max_pages and 
                  matches_found < MAX_MATCHES and 
                  not self.new_jobs_stopped.is_set()):
                
                # Handle pause state
                if self.new_jobs_paused.is_set():
                    with self.pause_lock:
                        self.save_new_jobs_status(current_page, matches_found, new_jobs, output_file)
                        self.log("New jobs scan paused")
                        while self.new_jobs_paused.is_set() and not self.new_jobs_stopped.is_set():
                            time.sleep(1)
                        if not self.new_jobs_stopped.is_set():
                            self.log("New jobs scan resumed")
                
                if self.new_jobs_stopped.is_set():
                    break
                
                self.log(f"\nChecking page {current_page} of {max_pages}")
                
                # Update progress
                if progress_callback:
                    progress_callback(current_page, max_pages)
                
                # Scroll to load all content
                if self.driver:
                    self.driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                self.random_delay(2, 3)
                
                jobs = self.scrape_page()
                if not jobs:
                    break
                    
                for job in jobs:
                    if self.is_duplicate(job, existing_jobs, 5):
                        matches_found += 1
                        if matches_found >= MAX_MATCHES:
                            break
                    else:
                        new_jobs.append(job)
                
                # Save progress after each page
                if new_jobs:
                    all_jobs = new_jobs + existing_jobs
                    backup_file = self.save_data(all_jobs, output_file)
                    self.save_new_jobs_status(current_page, matches_found, new_jobs, output_file)
                
                if matches_found >= MAX_MATCHES:
                    break
                    
                if not self.go_to_next_page():
                    break
                    
                current_page += 1
                self.random_delay(4, 7)
            
            # Final progress update
            if progress_callback:
                progress_callback(max_pages, max_pages)
            
            if new_jobs:
                self.log(f"\nAdded {len(new_jobs)} new jobs. Total jobs now: {len(existing_jobs) + len(new_jobs)}")
            else:
                self.log("\nNo new jobs found or stopped at duplicates")
                
            # Clear status file when complete
            if os.path.exists(STATUS_FILE):
                os.remove(STATUS_FILE)
                
        except Exception as e:
            self.log(f"Error in new jobs scanning: {str(e)}")
            self.save_new_jobs_status(current_page, matches_found, new_jobs, output_file)
        finally:
            self.new_jobs_stopped.set()
            self.new_jobs_paused.clear()

    def save_new_jobs_status(self, page_count: int, matches_found: int, 
                           new_jobs: List[Dict[str, str]], output_file: str) -> None:
        """Save status specific to New Jobs Only mode"""
        try:
            with open(STATUS_FILE, "wb") as f:
                pickle.dump({
                    "mode": "new_jobs",
                    "page_count": page_count,
                    "matches_found": matches_found,
                    "new_jobs": new_jobs,
                    "output_file": output_file,
                    "timestamp": datetime.now().isoformat()
                }, f)
        except Exception as e:
            self.log(f"Error saving new jobs status: {str(e)}")

    def scrape_all_pages(self, output_file: str, existing_file: Optional[str] = None,
                       progress_callback: Optional[Callable[[int, int], None]] = None) -> None:
        """Scrape all pages with resource management"""
        if not output_file:
            raise ValueError("Output file path not specified")
            
        all_jobs = []
        current_page = 1
        
        try:
            # Initialize driver only when needed
            self.initialize_driver()
            
            if existing_file:
                try:
                    existing_df = pd.read_excel(existing_file)
                    existing_data = existing_df.to_dict('records')
                    self.log(f"Existing jobs loaded: {len(existing_data)}")
                except Exception as e:
                    self.log(f"Error loading existing file: {str(e)}")
                    existing_data = []
            else:
                existing_data = None

            if not self.go_to_page(current_page):
                return
                
            max_pages = self.get_max_pages()
            self.log(f"Total pages to scrape: {max_pages}")
            
            while current_page <= max_pages and not self.stopped.is_set():
                if self.paused.is_set():
                    with self.pause_lock:
                        self.log("Scraping paused")
                        while self.paused.is_set() and not self.stopped.is_set():
                            time.sleep(1)
                        if not self.stopped.is_set():
                            self.log("Scraping resumed")
                
                if self.stopped.is_set():
                    break
                
                # Update progress
                if progress_callback:
                    progress_callback(current_page, max_pages)
                
                self.log(f"\nScraping page {current_page} of {max_pages}")
                
                jobs = self.scrape_page()
                if not jobs:
                    break
                    
                all_jobs.extend(jobs)
                backup_file = self.save_data(all_jobs, output_file, existing_data)
                self.save_status(current_page, output_file, backup_file)
                
                if current_page >= max_pages:
                    break
                    
                if not self.go_to_next_page():
                    break
                    
                current_page += 1
                
                # Rotate user agent every 5 pages
                if current_page % 5 == 0 and self.driver:
                    self.rotate_user_agent()
                    new_max = self.get_max_pages()
                    if new_max > max_pages:
                        self.log(f"Updated max pages to {new_max}")
                        max_pages = new_max
                
                self.random_delay(3, 6)
            
            # Final progress update
            if progress_callback:
                progress_callback(max_pages, max_pages)
            
            self.log(f"\nScraping complete. Total jobs collected: {len(all_jobs)}")
            
        except Exception as e:
            self.log(f"Error in full scraping: {str(e)}")
        finally:
            # Ensure WebDriver is closed after operation
            if hasattr(self, 'driver') and self.driver:
                self.driver.quit()
                self.driver = None

    def pause(self) -> None:
        """Pause main scraping"""
        self.paused.set()
        
    def pause_new_jobs(self) -> None:
        """Pause new jobs scanning"""
        self.new_jobs_paused.set()
        
    def stop(self) -> None:
        """Stop main scraping"""
        self.stopped.set()
        if hasattr(self, 'driver') and self.driver:
            self.driver.quit()
            self.driver = None
            
    def stop_new_jobs(self) -> None:
        """Stop new jobs scanning"""
        self.new_jobs_stopped.set()
        if hasattr(self, 'driver') and self.driver:
            self.driver.quit()
            self.driver = None
            
    def resume(self) -> None:
        """Resume main scraping"""
        self.paused.clear()
        
    def resume_new_jobs(self) -> None:
        """Resume new jobs scanning"""
        self.new_jobs_paused.clear()
        
    def log(self, message: str) -> None:
        """Log message to GUI or console"""
        if self.gui:
            self.gui.log_message(message)
        else:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")

class JobScraperGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Jobinja Scraper")
        self.root.geometry("800x600")
        self.log_text: Optional[scrolledtext.ScrolledText] = None
        self.scraper = JobScraper(self)
        self.running = False
        self.scheduled_job: Optional[schedule.Job] = None
        self.periodic_update_active = False
        
        # Progress tracking
        self.progress_var = tk.DoubleVar()
        self.nj_progress_var = tk.DoubleVar()
        
        self.create_widgets()
        self.load_config()
        
        # Start the schedule checker
        self.check_schedule()

    def check_schedule(self):
        """Check for scheduled jobs"""
        schedule.run_pending()
        self.root.after(1000, self.check_schedule)

    def create_widgets(self) -> None:
        """Create all GUI widgets"""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Create tabs
        self.operations_tab = ttk.Frame(self.notebook)
        self.new_jobs_tab = ttk.Frame(self.notebook)
        self.log_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.operations_tab, text="Operations")
        self.notebook.add(self.new_jobs_tab, text="New Jobs Mode")
        self.notebook.add(self.log_tab, text="Log")
        
        # Populate tabs
        self._create_operations_tab()
        self._create_new_jobs_tab()
        self._create_log_tab()
        
    def _create_operations_tab(self) -> None:
        """Create widgets for operations tab"""
        # Mode Selection
        mode_frame = ttk.LabelFrame(self.operations_tab, text="Operation Mode")
        mode_frame.pack(pady=10, padx=10, fill=tk.X)
        
        self.mode_var = tk.StringVar(value="new")
        
        ttk.Radiobutton(mode_frame, text="Complete New Scrape", 
                       variable=self.mode_var, value="new").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="Continue Previous Scrape", 
                       variable=self.mode_var, value="continue").pack(anchor=tk.W)
        
        # File Selection
        file_frame = ttk.LabelFrame(self.operations_tab, text="File Selection")
        file_frame.pack(pady=10, padx=10, fill=tk.X)
        
        ttk.Button(file_frame, text="Select Input File", 
                  command=self.select_input_file).pack(side=tk.LEFT, padx=5)
        self.input_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.input_file_var, 
                 state='readonly').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        ttk.Button(file_frame, text="Select Output File", 
                  command=self.select_output_file).pack(side=tk.LEFT, padx=5)
        self.output_file_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.output_file_var, 
                 state='readonly').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        # Progress Bar
        ttk.Progressbar(self.operations_tab, variable=self.progress_var, 
                       maximum=100).pack(fill=tk.X, padx=10, pady=5)
        
        # Control Buttons
        control_frame = ttk.Frame(self.operations_tab)
        control_frame.pack(pady=10)
        
        self.start_btn = ttk.Button(control_frame, text="Start", command=self.start_scraping)
        self.start_btn.pack(side=tk.LEFT, padx=5)
        
        self.pause_btn = ttk.Button(control_frame, text="Pause", 
                                   command=self.pause_scraping, state=tk.DISABLED)
        self.pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.resume_btn = ttk.Button(control_frame, text="Resume",
                                   command=self.resume_scraping, state=tk.DISABLED)
        self.resume_btn.pack(side=tk.LEFT, padx=5)
        
        self.stop_btn = ttk.Button(control_frame, text="Stop", 
                                  command=self.stop_scraping, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)
        
    def _create_new_jobs_tab(self) -> None:
        """Create widgets for new jobs tab"""
        # Schedule Options
        schedule_frame = ttk.LabelFrame(self.new_jobs_tab, text="Schedule Options")
        schedule_frame.pack(pady=10, padx=10, fill=tk.X)
        
        self.schedule_var = tk.StringVar(value="immediate")
        
        ttk.Radiobutton(schedule_frame, text="Run Immediately", 
                       variable=self.schedule_var, value="immediate").pack(anchor=tk.W)
        
        ttk.Radiobutton(schedule_frame, text="Schedule Daily At:", 
                       variable=self.schedule_var, value="daily").pack(anchor=tk.W)
        
        self.schedule_time_var = tk.StringVar(value="09:00")
        ttk.Entry(schedule_frame, textvariable=self.schedule_time_var, 
                 width=8).pack(anchor=tk.W, padx=20)
        
        # Periodic Update Option
        ttk.Radiobutton(schedule_frame, text="Periodic Update Every:", 
                       variable=self.schedule_var, value="periodic").pack(anchor=tk.W)
        
        periodic_frame = ttk.Frame(schedule_frame)
        periodic_frame.pack(anchor=tk.W, padx=20)
        
        self.periodic_hours_var = tk.IntVar(value=2)
        Spinbox(periodic_frame, from_=1, to=24, width=2, 
               textvariable=self.periodic_hours_var).pack(side=tk.LEFT)
        ttk.Label(periodic_frame, text="hours").pack(side=tk.LEFT, padx=5)
        
        # File Selection
        nj_file_frame = ttk.LabelFrame(self.new_jobs_tab, text="File Selection")
        nj_file_frame.pack(pady=10, padx=10, fill=tk.X)
        
        ttk.Button(nj_file_frame, text="Select Reference File", 
                  command=self.select_reference_file).pack(side=tk.LEFT, padx=5)
        self.reference_file_var = tk.StringVar()
        ttk.Entry(nj_file_frame, textvariable=self.reference_file_var, 
                 state='readonly').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        ttk.Button(nj_file_frame, text="Select Output File", 
                  command=self.select_new_jobs_output).pack(side=tk.LEFT, padx=5)
        self.new_jobs_output_var = tk.StringVar()
        ttk.Entry(nj_file_frame, textvariable=self.new_jobs_output_var, 
                 state='readonly').pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        
        # Progress Bar
        ttk.Progressbar(self.new_jobs_tab, variable=self.nj_progress_var, 
                       maximum=100).pack(fill=tk.X, padx=10, pady=5)
        
        # New Jobs Controls
        nj_control_frame = ttk.Frame(self.new_jobs_tab)
        nj_control_frame.pack(pady=10)
        
        self.nj_start_btn = ttk.Button(nj_control_frame, text="Start New Jobs Scan", 
                                     command=self.start_new_jobs_scan)
        self.nj_start_btn.pack(side=tk.LEFT, padx=5)
        
        self.nj_pause_btn = ttk.Button(nj_control_frame, text="Pause", 
                                      command=self.pause_new_jobs_scan, state=tk.DISABLED)
        self.nj_pause_btn.pack(side=tk.LEFT, padx=5)
        
        self.nj_resume_btn = ttk.Button(nj_control_frame, text="Resume",
                                       command=self.resume_new_jobs_scan, state=tk.DISABLED)
        self.nj_resume_btn.pack(side=tk.LEFT, padx=5)
        
        self.nj_stop_btn = ttk.Button(nj_control_frame, text="Stop", 
                                     command=self.stop_new_jobs_scan, state=tk.DISABLED)
        self.nj_stop_btn.pack(side=tk.LEFT, padx=5)
        
        # Status
        self.nj_status_var = tk.StringVar(value="Ready")
        ttk.Label(self.new_jobs_tab, textvariable=self.nj_status_var).pack()
        
    def _create_log_tab(self) -> None:
        """Create logging widget"""
        self.log_text = scrolledtext.ScrolledText(self.log_tab, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        self.log_text.config(state=tk.DISABLED)
        
    def log_message(self, message: str) -> None:
        """Add message to log"""
        if self.log_text is None:
            print(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}")
            return
            
        try:
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {message}\n")
            self.log_text.config(state=tk.DISABLED)
            self.log_text.see(tk.END)
        except Exception as e:
            print(f"Error logging message: {str(e)}")
            print(f"Original message: {message}")

    def select_input_file(self) -> None:
        """Select input file dialog"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.input_file_var.set(file_path)
            
    def select_output_file(self) -> None:
        """Select output file dialog"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                               filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.output_file_var.set(file_path)
            
    def select_reference_file(self) -> None:
        """Select reference file dialog"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.reference_file_var.set(file_path)
            
    def select_new_jobs_output(self) -> None:
        """Select new jobs output file dialog"""
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                               filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.new_jobs_output_var.set(file_path)
            
    def start_scraping(self) -> None:
        """Start scraping operation"""
        mode = self.mode_var.get()
        output_file = self.output_file_var.get()
        
        if not output_file:
            messagebox.showerror("Error", "Please select an output file")
            return
            
        if mode == "continue" and not self.input_file_var.get():
            messagebox.showerror("Error", "Please select an input file for continue mode")
            return
            
        self.running = True
        self.progress_var.set(0)
        self.start_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.NORMAL)
        
        Thread(target=self.run_scraping, args=(mode,), daemon=True).start()
        
    def run_scraping(self, mode: str) -> None:
        """Run scraping in background thread"""
        try:
            if mode == "new":
                self.scraper.scrape_all_pages(
                    output_file=self.output_file_var.get(),
                    progress_callback=self.update_progress
                )
            elif mode == "continue":
                self.scraper.scrape_all_pages(
                    output_file=self.output_file_var.get(),
                    existing_file=self.input_file_var.get(),
                    progress_callback=self.update_progress
                )
        except Exception as e:
            self.log_message(f"Error in scraping: {str(e)}")
        finally:
            self.running = False
            self.root.after(0, self.reset_controls)
            
    def update_progress(self, current: int, total: int) -> None:
        """Update progress bar for main scraping"""
        progress = (current / total) * 100 if total > 0 else 0
        self.progress_var.set(progress)
        self.root.update_idletasks()
            
    def pause_scraping(self) -> None:
        """Pause scraping operation"""
        self.scraper.pause()
        self.pause_btn.config(state=tk.DISABLED)
        self.resume_btn.config(state=tk.NORMAL)
        self.log_message("Scraping paused")
        
    def resume_scraping(self) -> None:
        """Resume scraping operation"""
        self.scraper.resume()
        self.resume_btn.config(state=tk.DISABLED)
        self.pause_btn.config(state=tk.NORMAL)
        self.log_message("Scraping resumed")
        
    def stop_scraping(self) -> None:
        """Stop scraping operation"""
        self.scraper.stop()
        self.running = False
        self.reset_controls()
        self.log_message("Scraping stopped")
        
    def reset_controls(self) -> None:
        """Reset control buttons to default state"""
        self.start_btn.config(state=tk.NORMAL)
        self.pause_btn.config(state=tk.DISABLED)
        self.resume_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        
    def start_new_jobs_scan(self) -> None:
        """Start new jobs scan operation"""
        if not self.reference_file_var.get() or not self.new_jobs_output_var.get():
            messagebox.showerror("Error", "Please select both reference and output files")
            return
            
        # Reset progress bar
        self.nj_progress_var.set(0)
        
        schedule_mode = self.schedule_var.get()
        
        if schedule_mode == "immediate":
            Thread(target=self.run_new_jobs_scan, daemon=True).start()
        elif schedule_mode == "daily":
            schedule_time = self.schedule_time_var.get()
            try:
                datetime.strptime(schedule_time, "%H:%M")
                self.scheduled_job = schedule.every().day.at(schedule_time).do(self.run_new_jobs_scan)
                self.nj_status_var.set(f"Scheduled daily at {schedule_time}")
                self.log_message(f"New jobs scan scheduled daily at {schedule_time}")
            except ValueError:
                messagebox.showerror("Error", "Invalid time format. Use HH:MM")
        elif schedule_mode == "periodic":
            hours = self.periodic_hours_var.get()
            if hours < 1 or hours > 24:
                messagebox.showerror("Error", "Please enter hours between 1 and 24")
                return
                
            self.periodic_update_active = True
            self.scheduled_job = schedule.every(hours).hours.do(self.run_new_jobs_scan)
            self.nj_status_var.set(f"Periodic scan every {hours} hours")
            self.log_message(f"Periodic scan scheduled every {hours} hours")
            
            # Run immediately and then periodically
            Thread(target=self.run_new_jobs_scan, daemon=True).start()
            
    def run_new_jobs_scan(self) -> None:
        """Run new jobs scan in background thread"""
        self.nj_start_btn.config(state=tk.DISABLED)
        self.nj_pause_btn.config(state=tk.NORMAL)
        self.nj_stop_btn.config(state=tk.NORMAL)
        self.nj_status_var.set("Scanning for new jobs...")
        
        try:
            # Initialize scraper (which will initialize WebDriver only when needed)
            if not hasattr(self.scraper, 'driver') or not self.scraper.driver:
                self.scraper.initialize_driver()
                
            self.scraper.scrape_new_jobs(
                reference_file=self.reference_file_var.get(),
                output_file=self.new_jobs_output_var.get(),
                progress_callback=self.update_new_jobs_progress
            )
        except Exception as e:
            self.log_message(f"Error in new jobs scan: {str(e)}")
        finally:
            if not self.periodic_update_active:
                self.root.after(0, self.reset_new_jobs_controls)
            self.nj_status_var.set(f"Scan completed at {datetime.now().strftime('%H:%M:%S')}")
            # Ensure WebDriver is closed after operation
            if hasattr(self.scraper, 'driver') and self.scraper.driver:
                self.scraper.driver.quit()
                self.scraper.driver = None

    def update_new_jobs_progress(self, current: int, total: int) -> None:
        """Update progress bar for new jobs mode"""
        progress = (current / total) * 100 if total > 0 else 0
        self.nj_progress_var.set(progress)
        self.root.update_idletasks()
        
    def pause_new_jobs_scan(self) -> None:
        """Pause new jobs scan"""
        self.scraper.pause_new_jobs()
        self.nj_pause_btn.config(state=tk.DISABLED)
        self.nj_resume_btn.config(state=tk.NORMAL)
        self.nj_status_var.set("Scan paused")
        
    def resume_new_jobs_scan(self) -> None:
        """Resume new jobs scan"""
        self.scraper.resume_new_jobs()
        self.nj_resume_btn.config(state=tk.DISABLED)
        self.nj_pause_btn.config(state=tk.NORMAL)
        self.nj_status_var.set("Scan resumed...")
        
    def stop_new_jobs_scan(self) -> None:
        """Stop new jobs scan"""
        self.scraper.stop_new_jobs()
        if self.scheduled_job:
            schedule.cancel_job(self.scheduled_job)
            self.scheduled_job = None
        self.periodic_update_active = False
        self.reset_new_jobs_controls()
        self.nj_status_var.set("Scan stopped")
        
    def reset_new_jobs_controls(self) -> None:
        """Reset new jobs control buttons"""
        self.nj_start_btn.config(state=tk.NORMAL)
        self.nj_pause_btn.config(state=tk.DISABLED)
        self.nj_resume_btn.config(state=tk.DISABLED)
        self.nj_stop_btn.config(state=tk.DISABLED)
        
    def load_config(self) -> None:
        """Load configuration from file"""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r') as f:
                    config = json.load(f)
                    self.input_file_var.set(config.get('input_file', ''))
                    self.output_file_var.set(config.get('output_file', ''))
                    self.reference_file_var.set(config.get('reference_file', ''))
                    self.new_jobs_output_var.set(config.get('new_jobs_output', ''))
                    self.schedule_time_var.set(config.get('schedule_time', '09:00'))
                    self.schedule_var.set(config.get('schedule_mode', 'immediate'))
                    self.periodic_hours_var.set(config.get('periodic_hours', 2))
            except Exception as e:
                self.log_message(f"Error loading config: {str(e)}")
                
    def save_config(self) -> None:
        """Save configuration to file"""
        config = {
            'input_file': self.input_file_var.get(),
            'output_file': self.output_file_var.get(),
            'reference_file': self.reference_file_var.get(),
            'new_jobs_output': self.new_jobs_output_var.get(),
            'schedule_time': self.schedule_time_var.get(),
            'schedule_mode': self.schedule_var.get(),
            'periodic_hours': self.periodic_hours_var.get()
        }
        
        try:
            with open(CONFIG_FILE, 'w') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self.log_message(f"Error saving config: {str(e)}")
            
    def on_closing(self) -> None:
        """Handle window closing event"""
        self.save_config()
        if self.running:
            self.stop_scraping()
        if hasattr(self.scraper, 'driver') and self.scraper.driver:
            self.scraper.driver.quit()
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = JobScraperGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()