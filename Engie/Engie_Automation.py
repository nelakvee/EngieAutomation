"""
ENGIE → IOP Utility-Info Sync
────────────────────────────
• Reads Site IDs & expected Vendor names from an Excel file
• Logs into ENGIE Impact (via Okta), extracts Power Company, Account Number,
  and Power Meter from each site’s bill
• Logs into IOP, opens the site, updates the “Utility Info” section, and saves
──────────────────────────────────────────────────────────────────────────────
Prereqs:
    pip install selenium openpyxl webdriver-manager

Chrome must be installed.  The matching chromedriver will be downloaded
automatically the first time the script runs (courtesy of webdriver-manager).
"""

import re
import os
import sys
import time
import logging
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, SessionNotCreatedException
)

# =============================================================================
# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# =============================================================================
EXCEL_FILE_PATH  = r'C:/Users/nelakve/Documents/Field Engineers/Engine_Site_ID_and_Vendor.xlsx'

ENGIE_LOGIN_URL  = 'https://engieimpact.okta.com/'
ENGIE_APP_TITLE  = 'ENGIE Impact Platform'

IOP_LOGIN_URL    = 'https://iop.vh.vzwnet.com/user/nelakve/sites'

ENGIE_USERNAME   = 'veenith.nelakanti@verizonwireless.com'
IOP_USERNAME     = 'nelakve'
IOP_PASSWORD     = 'Vamshin143@'

SHORT_TIMEOUT    = 30
LONG_TIMEOUT     = 90

# Where screenshots are saved upon error
SCREENSHOT_DIR   = os.path.join(os.getcwd(), "screenshots")
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
logging.basicConfig(
    level   = logging.INFO,
    format  = '%(asctime)s [%(levelname)s] %(message)s',
    datefmt = '%Y-%m-%d %H:%M:%S',
    force   = True          # override any prior logging config
)
logger = logging.getLogger(__name__)

# =============================================================================
# ── HELPER FUNCTIONS ──────────────────────────────────────────────────────────
# =============================================================================
def normalize_vendor_name(name: str) -> str:
    """Lower-case, strip special chars, collapse whitespace."""
    if not isinstance(name, str):
        return ''
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned.split())

def take_screenshot(driver, site_id: str, step: str):
    """Capture screenshot for debugging."""
    ts = time.strftime("%Y%m%d_%H%M%S")
    path = os.path.join(SCREENSHOT_DIR, f"{site_id}_{step}_{ts}.png")
    try:
        driver.save_screenshot(path)
        logger.info(f"Saved screenshot → {path}")
    except Exception as e:
        logger.error(f"Could not save screenshot: {e}")

def load_sites_from_excel(path: str):
    """Return list[(site_id, vendor)] from Excel."""
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        sites = []
        for site_id, vendor in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            if site_id:
                sites.append((str(site_id).strip(), str(vendor or '').strip()))
        return sites
    except FileNotFoundError:
        logger.critical(f"Excel file not found → {path}")
        sys.exit(1)
    except Exception as exc:
        logger.critical(f"Failed reading Excel: {exc}")
        sys.exit(1)

# =============================================================================
# ── MAIN AUTOMATION CLASS ─────────────────────────────────────────────────────
# =============================================================================
class EngieIOPAutomator:
    def __init__(self, sites):
        self.sites        = sites
        self.driver       = None
        self.wait_short   = None
        self.wait_long    = None
        self.engie_handle = None
        self.iop_handle   = None

    # ──────────────────────────────────────────────────────────────────────────
    def setup_driver(self):
        logger.info("Launching Chrome …")
        from selenium.webdriver.chrome.service import Service
        from webdriver_manager.chrome import ChromeDriverManager

        options = webdriver.ChromeOptions()
        options.add_argument('--start-maximized')
        # options.add_argument('--headless=new')   # ← uncomment for headless

        try:
            self.driver = webdriver.Chrome(
                service = Service(ChromeDriverManager().install()),
                options = options
            )
        except SessionNotCreatedException as e:
            logger.critical(f"Chrome / chromedriver mismatch: {e}")
            sys.exit(1)

        self.wait_short = WebDriverWait(self.driver, SHORT_TIMEOUT)
        self.wait_long  = WebDriverWait(self.driver, LONG_TIMEOUT)

    # ──────────────────────────────────────────────────────────────────────────
    def login_engie(self):
        logger.info("Opening Okta login for ENGIE …")
        self.driver.get(ENGIE_LOGIN_URL)

        # Username on first Okta screen
        self.wait_short.until(
            EC.visibility_of_element_located((By.ID, 'idp-discovery-username'))
        ).send_keys(ENGIE_USERNAME)
        self.driver.find_element(By.ID, 'idp-discovery-submit').click()

        # Pause for password + MFA
        input("\n🔑  Complete Okta sign-in (password + MFA) then press <Enter> here…")

        # Choose ENGIE Impact Platform app
        logger.info("Selecting ENGIE Impact Platform app …")
        self.wait_long.until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//span[@data-se='app-card-title' and @title='{ENGIE_APP_TITLE}']"))
        ).click()

        # New tab opens
        self.wait_long.until(EC.number_of_windows_to_be(2))
        self.engie_handle = self.driver.window_handles[-1]
        self.driver.switch_to.window(self.engie_handle)
        logger.info("ENGIE tab ready ✅")

    # ──────────────────────────────────────────────────────────────────────────
    def login_iop(self):
        logger.info("Opening IOP in a separate tab …")
        self.driver.execute_script("window.open('');")
        self.iop_handle = self.driver.window_handles[-1]
        self.driver.switch_to.window(self.iop_handle)
        self.driver.get(IOP_LOGIN_URL)

        self.wait_short.until(EC.visibility_of_element_located((By.ID, 'idToken1'))).send_keys(IOP_USERNAME)
        self.driver.find_element(By.ID, 'idToken2').send_keys(IOP_PASSWORD)
        self.driver.find_element(By.ID, 'loginButton_0').click()
        logger.info("Logged into IOP ✅")

    # ──────────────────────────────────────────────────────────────────────────
    def process_site(self, idx, site_id, vendor):
        logger.info(f"\n── Site {idx}/{len(self.sites)}  (ID={site_id}) ──")
        vendor_norm = normalize_vendor_name(vendor)

        try:
            # ===== ENGIE SEARCH =================================================
            self.driver.switch_to.window(self.engie_handle)

            # Wait for overlay gone
            try:
                self.wait_long.until(EC.invisibility_of_element_located(
                    (By.XPATH, "//div[contains(@class,'ui-widget-overlay')]")))
            except TimeoutException:
                pass

            # Search bar and button (present on both dashboard & bill pages)
            search_box = self.wait_long.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[contains(@class,'search-box') and @placeholder='Search']")))
            self.driver.execute_script("arguments[0].value='';", search_box)
            search_box.send_keys(site_id)

            # Click enabled magnifier
            self.wait_long.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[contains(@class,'search-btn-enabled')]"))).click()

            logger.info("Waiting for bill list …")
            rows_xpath = "//table[contains(@id,'BillResultsGrid')]//tr[.//a[contains(@id,'VendorName')]]"
            self.wait_long.until(EC.presence_of_element_located((By.XPATH, rows_xpath)))
            rows = self.driver.find_elements(By.XPATH, rows_xpath)

            # Pick matching vendor
            data = {}
            for row in rows:
                if vendor_norm and vendor_norm not in normalize_vendor_name(
                        row.find_element(By.XPATH, ".//a[contains(@id,'VendorName')]").text):
                    continue

                logger.info("Opening bill details …")
                before_tabs = set(self.driver.window_handles)
                row.find_element(By.XPATH, ".//a[normalize-space()='View...']").click()
                self.wait_long.until(EC.new_window_is_opened(before_tabs))
                new_tab = (set(self.driver.window_handles) - before_tabs).pop()

                # Close old ENGIE tab, switch to new details tab
                self.driver.switch_to.window(self.engie_handle)
                self.driver.close()
                self.driver.switch_to.window(new_tab)
                self.engie_handle = new_tab

                # Wait for vendor name span (may be inside an iframe)
                try:
                    self.wait_long.until(EC.presence_of_element_located(
                        (By.ID, "id-uem-bill-details-vendor-name")))
                except TimeoutException:
                    # Try first iframe
                    frames = self.driver.find_elements(By.TAG_NAME, "iframe")
                    if frames:
                        self.driver.switch_to.frame(frames[0])
                        self.wait_long.until(EC.presence_of_element_located(
                            (By.ID, "id-uem-bill-details-vendor-name")))

                # Extract values
                data['power_company']  = self.driver.find_element(
                    By.ID, "id-uem-bill-details-vendor-name").text.split('/')[0].strip()
                data['account_number'] = self.driver.find_element(
                    By.ID, "id-uem-bill-details-acct-number").text.strip()
                data['power_meter']    = self.driver.find_element(
                    By.XPATH, "//td[contains(@class,'uem-bill-details-meter-number-widthSet')]").text.strip()

                # Reset any frame
                self.driver.switch_to.default_content()
                break

            if not data:
                logger.warning("No matching vendor row found; skipping site.")
                return

            logger.info(f"Extracted → {data}")
            print(f"Extracted  {site_id}:  {data}")

            # ===== UPDATE IOP ===================================================
            self.driver.switch_to.window(self.iop_handle)
            search = self.wait_long.until(EC.element_to_be_clickable(
                (By.XPATH, "//input[@placeholder='Site/Switch Search']")))
            search.clear()
            search.send_keys(site_id)
            time.sleep(1)
            self.wait_short.until(EC.element_to_be_clickable(
                (By.XPATH, "//a[@class='dropdown-item']"))).click()

            util_hdr = self.wait_long.until(EC.element_to_be_clickable(
                (By.XPATH, "//span[normalize-space()='Utility Info']/..")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", util_hdr)
            time.sleep(1)
            util_hdr.click()

            mapping = {
                'Power Company': ("//label[.='Power Company']/following-sibling::input", data['power_company']),
                'Power Meter'  : ("//label[.='Power Meter']/following-sibling::input",   data['power_meter']),
                'Account Number':("//label[.='Account Number']/following-sibling::input",data['account_number'])
            }

            for label, (xpath, val) in mapping.items():
                elem = self.wait_short.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                elem.clear()
                elem.send_keys(val)
                logger.info(f"🖊  {label}: {val}")

            self.wait_short.until(EC.element_to_be_clickable(
                (By.XPATH, "//button[normalize-space()='Save Utility Information']"))).click()
            logger.info("Utility info saved ✅")
            time.sleep(2)

        except Exception as exc:
            logger.error(f"❌  Error processing {site_id}: {exc}")
            take_screenshot(self.driver, site_id, "error")

    # ──────────────────────────────────────────────────────────────────────────
    def run(self):
        try:
            self.setup_driver()
            self.login_engie()
            self.login_iop()

            for idx, (site_id, vendor) in enumerate(self.sites, 1):
                self.process_site(idx, site_id, vendor)

        finally:
            if self.driver:
                self.driver.quit()
                logger.info("Browser closed.")

# =============================================================================
# ── MAIN ──────────────────────────────────────────────────────────────────────
# =============================================================================
if __name__ == "__main__":
    print("=== ENGIE ➜ IOP Automation starting ===")
    logger.info("Reading Excel …")
    site_list = load_sites_from_excel(EXCEL_FILE_PATH)
    print(f"Sites loaded: {len(site_list)}")
    if not site_list:
        print("No sites found. Check the spreadsheet path / contents.")
        sys.exit(0)

    EngieIOPAutomator(site_list).run()
