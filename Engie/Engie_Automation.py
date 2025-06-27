import re
import time
import os
import sys
import logging
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    InvalidElementStateException
)

# =============================================================================
# Configurations & Constants
# =============================================================================
EXCEL_FILE_PATH = r'C:/Users/nelakve/Documents/Field Engineers/Engine_Site_ID_and_Vendor.xlsx'
ENGIE_LOGIN_URL = 'https://engieimpact.okta.com/'
ENGIE_APP_TITLE = 'ENGIE Impact Platform'
IOP_LOGIN_URL = 'https://iop.vh.vzwnet.com/user/nelakve/sites'
ENGIE_USERNAME = 'veenith.nelakanti@verizonwireless.com'
IOP_USERNAME   = 'nelakve'
IOP_PASSWORD   = 'Vamshin143@'
SHORT_TIMEOUT  = 30
LONG_TIMEOUT   = 90

# Directory for screenshots on error
SCREENSHOT_DIR = os.path.join(os.getcwd(), 'screenshots')
if not os.path.exists(SCREENSHOT_DIR):
    os.makedirs(SCREENSHOT_DIR)

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

# =============================================================================
# Utility Functions
# =============================================================================
def normalize_vendor_name(name: str) -> str:
    """
    Remove non-alphanumeric, lowercase, collapse whitespace.
    """
    if not isinstance(name, str):
        return ''
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned.split())


def take_screenshot(driver, site_id: str, step: str) -> None:
    """
    Save screenshot for debugging.
    """
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    filename = f"error_{site_id}_{step}_{timestamp}.png"
    path = os.path.join(SCREENSHOT_DIR, filename)
    try:
        driver.save_screenshot(path)
        logger.info(f"Screenshot saved: {path}")
    except Exception as e:
        logger.error(f"Failed to save screenshot: {e}")


def load_sites_from_excel(path: str) -> list:
    """
    Read Excel and return list of (site_id, vendor).
    """
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        return [
            (str(r[0]).strip(), str(r[1]).strip() if r[1] else '')
            for r in ws.iter_rows(min_row=2, max_col=2, values_only=True)
            if r[0]
        ]
    except Exception as e:
        logger.error(f"Error loading Excel: {e}")
        sys.exit(1)

# =============================================================================
# Main Automation Class
# =============================================================================
class EngieIOPAutomator:
    def __init__(self, sites):
        self.sites = sites
        self.driver = None
        self.wait_short = None
        self.wait_long = None
        self.engie_handle = None
        self.iop_handle = None

    def setup_driver(self):
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        self.driver = webdriver.Chrome(options=options)
        self.wait_short = WebDriverWait(self.driver, SHORT_TIMEOUT)
        self.wait_long  = WebDriverWait(self.driver, LONG_TIMEOUT)

    def login_engie(self):
        # Load Okta and start login
        self.driver.get(ENGIE_LOGIN_URL)
        self.wait_short.until(
            EC.visibility_of_element_located((By.ID, 'idp-discovery-username'))
        ).send_keys(ENGIE_USERNAME)
        self.wait_short.until(
            EC.element_to_be_clickable((By.ID, 'idp-discovery-submit'))
        ).click()
        input("Complete Okta login and press Enter...")

        # Launch ENGIE Impact Platform
        tile = self.wait_long.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//span[@data-se='app-card-title' and @title='{ENGIE_APP_TITLE}']")
            )
        )
        tile.click()
        self.wait_long.until(EC.number_of_windows_to_be(2))

        # Close original tab, keep only ENGIE
        handles = self.driver.window_handles
        old = handles[0]
        new = handles[-1]
        self.driver.switch_to.window(old)
        self.driver.close()
        self.driver.switch_to.window(new)
        self.engie_handle = new
        logger.info("Switched to ENGIE Impact Platform tab.")

    def login_iop(self):
        # Open and login to IOP
        self.driver.execute_script("window.open('');")
        handles = self.driver.window_handles
        self.iop_handle = handles[-1]
        self.driver.switch_to.window(self.iop_handle)
        self.driver.get(IOP_LOGIN_URL)
        self.wait_short.until(
            EC.visibility_of_element_located((By.ID, 'idToken1'))
        ).send_keys(IOP_USERNAME)
        self.wait_short.until(
            EC.visibility_of_element_located((By.ID, 'idToken2'))
        ).send_keys(IOP_PASSWORD)
        self.wait_short.until(
            EC.element_to_be_clickable((By.ID, 'loginButton_0'))
        ).click()
        logger.info("Logged into IOP.")

    def process_site(self, index, site_id, vendor):
        logger.info(f"{index}/{len(self.sites)} - Processing {site_id} ({vendor})")
        try:
            # Switch to ENGIE details/search page
            self.driver.switch_to.window(self.engie_handle)

            # Perform search on Bill Details page
            search_box = self.wait_long.until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "input.search-box[placeholder='Search']")
                )
            )
            self.driver.execute_script("arguments[0].value=''", search_box)
            search_box.send_keys(site_id)
            self.driver.find_element(By.CSS_SELECTOR, "i.fa-magnifying-glass").click()

            # Extract fields
            self.wait_long.until(
                EC.visibility_of_element_located((By.ID, 'id-uem-bill-details-vendor-name'))
            )
            power_company = self.driver.find_element(By.ID, 'id-uem-bill-details-vendor-name')
            power_company = power_company.text.split('/')[0].strip()

            account_number = self.driver.find_element(By.ID, 'id-uem-bill-details-acct-number').text.strip()
            power_meter = self.driver.find_element(
                By.CSS_SELECTOR, 'td.uem-bill-details-meter-number-widthSet'
            ).text.strip()

            print(f"Site {site_id} -> Company: {power_company}, Account: {account_number}, Meter: {power_meter}")

            # Switch to IOP and enter
            self.driver.switch_to.window(self.iop_handle)
            io_search = self.wait_long.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "input[placeholder='Site/Switch Search']"))
            )
            io_search.clear()
            io_search.send_keys(site_id)
            time.sleep(1)
            self.wait_long.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, 'a.dropdown-item'))
            ).click()

            # Expand Utility Info and fill
            header = self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, "//span[text()='Utility Info']/.."))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true)", header)
            header.click()

            self.wait_short.until(
                EC.element_to_be_clickable((By.XPATH, "//label[text()='Power Company']/following-sibling::input"))
            ).send_keys(power_company)
            self.driver.find_element(
                By.XPATH, "//label[text()='Power Meter']/following-sibling::input"
            ).send_keys(power_meter)
            self.driver.find_element(
                By.XPATH, "//label[text()='Account Number']/following-sibling::input"
            ).send_keys(account_number)

            self.wait_short.until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Save Utility Information']"))
            ).click()
            time.sleep(2)

        except Exception as err:
            logger.error(f"Error on site {site_id}: {err}")
            take_screenshot(self.driver, site_id, 'process_site')

    def run(self):
        try:
            self.setup_driver()
            self.login_engie()
            self.login_iop()
            for idx, (sid, vend) in enumerate(self.sites, 1):
                self.process_site(idx, sid, vend)
        finally:
            if self.driver:
                self.driver.quit()
            logger.info("Automation complete.")

# Entry
if __name__ == '__main__':
    sites = load_sites_from_excel(EXCEL_FILE_PATH)
    auto = EngieIOPAutomator(sites)
    auto.run()
