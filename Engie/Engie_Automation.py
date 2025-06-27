import re
import time
import os
import sys
import logging
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
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
# Excel file containing Site IDs and Vendor Names
EXCEL_FILE_PATH = r'C:/Users/nelakve/Documents/Field Engineers/Engine_Site_ID_and_Vendor.xlsx'

# ENGIE (Okta) URLs and App Title
ENGIE_LOGIN_URL = 'https://engieimpact.okta.com/'
ENGIE_APP_TITLE = 'ENGIE Impact Platform'

# IOP (Integrated Operations Platform) URL
IOP_LOGIN_URL = 'https://iop.vh.vzwnet.com/user/nelakve/sites'

# Credentials (Hardcoded for simplicity)
ENGIE_USERNAME = 'veenith.nelakanti@verizonwireless.com'
IOP_USERNAME = 'nelakve'
IOP_PASSWORD = 'Vamshin143@'

# Timeout settings for slow-loading pages
SHORT_TIMEOUT = 30
LONG_TIMEOUT = 90

# Screenshot directory for errors
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
    Clean up vendor name: remove non-alphanumeric, lowercase, collapse spaces.
    """
    if not isinstance(name, str):
        return ''
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned.split())


def take_screenshot(driver, site_id: str, step: str) -> None:
    """
    Capture screenshot on error for debugging.
    """
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    filename = f"error_{site_id}_{step}_{timestamp}.png"
    filepath = os.path.join(SCREENSHOT_DIR, filename)
    try:
        driver.save_screenshot(filepath)
        logger.info(f"Saved screenshot: {filepath}")
    except Exception as e:
        logger.error(f"Failed to save screenshot {filepath}: {e}")


def load_sites_from_excel(path: str) -> list:
    """
    Read the Excel file and return list of tuples: (site_id, vendor_name).
    """
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        sites = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            site_id, vendor = row
            if site_id:
                sites.append((str(site_id).strip(), str(vendor).strip() if vendor else ''))
        return sites
    except FileNotFoundError:
        logger.error(f"Excel file not found at: {path}")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Error reading Excel file: {e}")
        sys.exit(1)

# =============================================================================
# Main Automation Class
# =============================================================================
class EngieIOPAutomator:
    def __init__(self, sites: list):
        self.sites = sites
        self.driver = None
        self.wait_short = None
        self.wait_long = None
        self.engie_handle = None
        self.iop_handle = None

    def setup_driver(self):
        logger.info("Launching Chrome WebDriver...")
        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        self.driver = webdriver.Chrome(options=options)
        self.wait_short = WebDriverWait(self.driver, SHORT_TIMEOUT)
        self.wait_long = WebDriverWait(self.driver, LONG_TIMEOUT)

    def login_engie(self):
        logger.info("Navigating to ENGIE Okta login...")
        self.driver.get(ENGIE_LOGIN_URL)
        self.wait_short.until(
            EC.visibility_of_element_located((By.ID, 'idp-discovery-username'))
        ).send_keys(ENGIE_USERNAME)
        self.wait_short.until(
            EC.element_to_be_clickable((By.ID, 'idp-discovery-submit'))
        ).click()
        input("Please complete Okta sign-in and press Enter to continue...")
        logger.info("Selecting ENGIE Impact Platform from Okta apps...")
        engie_tile = self.wait_long.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//span[@data-se='app-card-title' and @title='{ENGIE_APP_TITLE}']")
            )
        )
        engie_tile.click()
        self.wait_long.until(EC.number_of_windows_to_be(2))
        self.engie_handle = self.driver.window_handles[-1]
        self.driver.switch_to.window(self.engie_handle)
        logger.info("Switched to ENGIE Impact Platform tab.")

    def login_iop(self):
        logger.info("Opening new tab for IOP login...")
        self.driver.execute_script("window.open('');")
        self.iop_handle = self.driver.window_handles[-1]
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

    def process_site(self, index: int, site_id: str, vendor: str):
        vendor_norm = normalize_vendor_name(vendor)
        logger.info(f"Processing site {index}: {site_id} / Vendor: {vendor}")
        try:
            self.driver.switch_to.window(self.engie_handle)
            self.driver.refresh()
            overlay_xpath = "//div[contains(@class, 'ui-widget-overlay')]"
            self.wait_long.until(EC.invisibility_of_element_located((By.XPATH, overlay_xpath)))
            search_input_xpath = "//input[contains(@class,'search-box') and @placeholder='Search']"
            search_btn_xpath = "//button[contains(@class,'search-btn-enabled')]"
            search_input = self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, search_input_xpath))
            )
            self.driver.execute_script("arguments[0].value=''", search_input)
            for c in site_id:
                search_input.send_keys(c)
                time.sleep(0.05)
            self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, search_btn_xpath))
            ).click()
            bill_rows_xpath = (
                "//table[contains(@id,'BillResultsGrid')]//tr[.//a[contains(@id,'VendorName')]]"
            )
            self.wait_long.until(
                EC.presence_of_element_located((By.XPATH, bill_rows_xpath))
            )
            bill_rows = self.driver.find_elements(By.XPATH, bill_rows_xpath)
            extracted_data = {}
            for row in bill_rows:
                try:
                    txt = row.find_element(By.XPATH, ".//a[contains(@id,'VendorName')]").text
                    if vendor_norm in normalize_vendor_name(txt):
                        before_handles = set(self.driver.window_handles)
                        view_btn = row.find_element(By.XPATH, ".//a[text()='View...']")
                        self.driver.execute_script("arguments[0].click()", view_btn)
                        self.wait_long.until(EC.new_window_is_opened(before_handles))
                        new_window = (set(self.driver.window_handles) - before_handles).pop()
                        self.driver.switch_to.window(new_window)
                        iframe_xpath = (
                            "//iframe[contains(@id,'iframe') or contains(@name,'iframe') or @title='content']"
                        )
                        self.wait_long.until(
                            EC.frame_to_be_available_and_switch_to_it((By.XPATH, iframe_xpath))
                        )
                        extracted_data['power_company'] = (
                            self.driver.find_element(By.ID, 'id-uem-bill-details-vendor-name')
                            .text.split('/')[0].strip()
                        )
                        extracted_data['account_number'] = (
                            self.driver.find_element(By.ID, 'id-uem-bill-details-acct-number')
                            .text.strip()
                        )
                        extracted_data['power_meter'] = (
                            self.driver.find_element(
                                By.XPATH,
                                "//td[contains(@class,'uem-bill-details-meter-number-widthSet')]")
                            .text.strip()
                        )
                        self.driver.switch_to.default_content()
                        self.driver.close()
                        self.driver.switch_to.window(self.engie_handle)
                        break
                except Exception:
                    continue
            if not extracted_data:
                logger.warning(f"No matching vendor row found for {vendor} on ENGIE.")
                return
            logger.info(f"Extracted data: {extracted_data}")
            self.driver.switch_to.window(self.iop_handle)
            search_io_input_xpath = "//input[@placeholder='Site/Switch Search']"
            search_io = self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, search_io_input_xpath))
            )
            search_io.clear()
            search_io.send_keys(site_id)
            time.sleep(1)
            dropdown_xpath = "//a[@class='dropdown-item']"
            self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, dropdown_xpath))
            ).click()
            util_section_xpath = "//span[text()='Utility Info']/.."
            util_header = self.wait_long.until(
                EC.element_to_be_clickable((By.XPATH, util_section_xpath))
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true)", util_header)
            time.sleep(1)
            util_header.click()
            field_map = {
                'power_company': "//label[text()='Power Company']/following-sibling::input",
                'power_meter': "//label[text()='Power Meter']/following-sibling::input",
                'account_number': "//label[text()='Account Number']/following-sibling::input"
            }
            for key, xpath in field_map.items():
                inp = self.wait_short.until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                inp.clear()
                inp.send_keys(extracted_data[key])
                logger.info(f"Entered {key}: {extracted_data[key]}")
            save_btn = self.wait_short.until(
                EC.element_to_be_clickable((By.XPATH, "//button[text()='Save Utility Information']"))
            )
            save_btn.click()
            logger.info(f"Successfully saved data for site {site_id} in IOP.")
            time.sleep(2)
        except Exception as e:
            logger.error(f"Error processing site {site_id}: {e}")
            take_screenshot(self.driver, site_id, 'processing')

    def run(self):
        try:
            self.setup_driver()
            self.login_engie()
            self.login_iop()
            for idx, (site_id, vendor) in enumerate(self.sites, start=1):
                self.process_site(idx, site_id, vendor)
        except Exception as e:
            logger.critical(f"Automation failed unexpectedly: {e}")
        finally:
            if self.driver:
                time.sleep(1)
                self.driver.quit()
                logger.info("WebDriver closed. Automation complete.")

if __name__ == '__main__':
    logger.info("Loading site list from Excel...")
    site_list = load_sites_from_excel(EXCEL_FILE_PATH)
    automator = EngieIOPAutomator(site_list)
    automator.run()
Hi
