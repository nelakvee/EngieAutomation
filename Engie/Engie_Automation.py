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
from selenium.common.exceptions import TimeoutException, NoSuchElementException

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

# Credentials (Hardcoded for simplicity - consider using a secure method in production)
ENGIE_USERNAME = 'veenith.nelakanti@verizonwireless.com'
IOP_USERNAME   = 'nelakve'
IOP_PASSWORD   = 'Vamshin143@'

# Timeout settings for slow-loading pages
SHORT_TIMEOUT = 30
LONG_TIMEOUT  = 90

# Screenshot directory for errors
SCREENSHOT_DIR = os.path.join(os.getcwd(), 'screenshots')
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

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
    """Clean up vendor name: remove non-alphanumeric characters, lowercase, collapse spaces."""
    if not isinstance(name, str):
        return ''
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned.split())

def take_screenshot(driver, site_id: str, step: str) -> None:
    """Capture screenshot on error for debugging."""
    timestamp = time.strftime('%Y%m%d_%H%M%S')
    filename = f"error_{site_id}_{step}_{timestamp}.png"
    filepath = os.path.join(SCREENSHOT_DIR, filename)
    try:
        driver.save_screenshot(filepath)
        logger.info(f"Saved screenshot: {filepath}")
    except Exception as e:
        logger.error(f"Failed to save screenshot {filename}: {e}")

def load_sites_from_excel(path: str) -> list:
    """Read the Excel file and return list of tuples: (site_id, vendor_name)."""
    try:
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        sites = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            site_id, vendor = row
            if site_id:
                site_id = str(site_id).strip()
                vendor = str(vendor).strip() if vendor else ''
                sites.append((site_id, vendor))
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
        # (Optional: add headless or other options if needed)
        self.driver = webdriver.Chrome(options=options)
        self.wait_short = WebDriverWait(self.driver, SHORT_TIMEOUT)
        self.wait_long  = WebDriverWait(self.driver, LONG_TIMEOUT)

    def login_engie(self):
        logger.info("Navigating to ENGIE Okta login page...")
        self.driver.get(ENGIE_LOGIN_URL)
        # Enter username and submit (Okta will handle password and MFA externally)
        self.wait_short.until(EC.visibility_of_element_located((By.ID, 'idp-discovery-username'))).send_keys(ENGIE_USERNAME)
        self.wait_short.until(EC.element_to_be_clickable((By.ID, 'idp-discovery-submit'))).click()
        # Pause for Okta MFA (user must complete manually)
        input("Please complete Okta sign-in (MFA) and press Enter here to continue...")
        # Select the ENGIE Impact Platform app after Okta login
        logger.info("Selecting 'ENGIE Impact Platform' app from Okta dashboard...")
        engie_app_tile = self.wait_long.until(
            EC.element_to_be_clickable((By.XPATH, f"//span[@data-se='app-card-title' and @title='{ENGIE_APP_TITLE}']"))
        )
        engie_app_tile.click()
        # Wait for ENGIE platform to open in a new tab
        self.wait_long.until(EC.number_of_windows_to_be(2))
        self.engie_handle = self.driver.window_handles[-1]
        self.driver.switch_to.window(self.engie_handle)
        logger.info("ENGIE Impact Platform opened in a new tab and activated.")

    def login_iop(self):
        logger.info("Opening IOP site in a new browser tab...")
        self.driver.execute_script("window.open('');")  # open a blank new tab
        self.iop_handle = self.driver.window_handles[-1]
        self.driver.switch_to.window(self.iop_handle)
        self.driver.get(IOP_LOGIN_URL)
        # Enter IOP credentials and log in
        self.wait_short.until(EC.visibility_of_element_located((By.ID, 'idToken1'))).send_keys(IOP_USERNAME)
        self.wait_short.until(EC.visibility_of_element_located((By.ID, 'idToken2'))).send_keys(IOP_PASSWORD)
        self.wait_short.until(EC.element_to_be_clickable((By.ID, 'loginButton_0'))).click()
        logger.info("Logged into IOP successfully.")

    def process_site(self, index: int, site_id: str, vendor: str):
        """Process a single site: extract data from ENGIE and update IOP."""
        vendor_norm = normalize_vendor_name(vendor)
        logger.info(f"\n--- Processing Site {index}: SiteID={site_id}, Expected Vendor='{vendor}' ---")
        try:
            # Ensure we're on the ENGIE tab
            self.driver.switch_to.window(self.engie_handle)
            # Refresh the ENGIE page to ensure a clean state (especially for the first search or if still showing old data)
            try:
                self.driver.refresh()
            except Exception as e:
                logger.debug(f"Refresh not applicable or failed (possibly first load): {e}")
            # Wait for any loading overlay to disappear (ENGIE pages often show an overlay during loading)
            overlay_xpath = "//div[contains(@class, 'ui-widget-overlay')]"
            self.wait_long.until(EC.invisibility_of_element_located((By.XPATH, overlay_xpath)))
            # Find the search input box and search button on the ENGIE page
            search_input_xpath = "//input[contains(@class,'search-box') and @placeholder='Search']"
            search_btn_xpath   = "//button[contains(@class,'search-btn-enabled')]"
            search_input = self.wait_long.until(EC.element_to_be_clickable((By.XPATH, search_input_xpath)))
            # Clear any previous text in the search box and enter the new Site ID
            self.driver.execute_script("arguments[0].value = '';", search_input)  # clear via JS to avoid readonly issues
            for ch in site_id:
                search_input.send_keys(ch)
                time.sleep(0.05)  # slight delay to mimic typing and allow binding updates
            # Wait for the search button to become enabled, then click it
            self.wait_long.until(EC.element_to_be_clickable((By.XPATH, search_btn_xpath))).click()
            logger.info(f"Initiated search on ENGIE for Site ID {site_id}. Waiting for results...")
            # Wait for at least one result row to appear in the Bill Results grid (which contains VendorName links)
            results_xpath = "//table[contains(@id,'BillResultsGrid')]//tr[.//a[contains(@id,'VendorName')]]"
            self.wait_long.until(EC.presence_of_element_located((By.XPATH, results_xpath)))
            bill_rows = self.driver.find_elements(By.XPATH, results_xpath)
            if not bill_rows:
                logger.warning(f"No bill entries found for Site ID {site_id} on ENGIE.")
            extracted_data = {}
            # Search through the result rows for the given vendor name (normalized comparison for robustness)
            for row in bill_rows:
                try:
                    vendor_text = row.find_element(By.XPATH, ".//a[contains(@id,'VendorName')]").text
                except NoSuchElementException:
                    continue  # skip this row if no vendor link (shouldn't happen due to XPath filtering)
                if vendor_norm and vendor_norm in normalize_vendor_name(vendor_text):
                    logger.info(f"Found matching vendor '{vendor_text}' for Site ID {site_id}, opening bill details...")
                    # Click the "View..." link in this row to open the bill details
                    before_handles = set(self.driver.window_handles)
                    view_link = row.find_element(By.XPATH, ".//a[text()='View...']")
                    self.driver.execute_script("arguments[0].click();", view_link)
                    # Wait for a new window/tab to open
                    self.wait_long.until(EC.new_window_is_opened(before_handles))
                    new_tab_handle = (set(self.driver.window_handles) - before_handles).pop()
                    # Close the previous ENGIE tab (no longer needed) to save memory
                    try:
                        self.driver.switch_to.window(self.engie_handle)
                        self.driver.close()
                        logger.info("Closed previous ENGIE tab after opening new bill details tab.")
                    except Exception as e:
                        logger.warning(f"Could not close previous ENGIE tab: {e}")
                    # Switch focus to the new bill details tab
                    self.driver.switch_to.window(new_tab_handle)
                    self.engie_handle = new_tab_handle  # update engie_handle to the new tab for next iteration
                    logger.info("Switched to new bill details tab. Extracting data...")
                    # Wait for the bill details content to load (identify by vendor name element presence)
                    try:
                        self.wait_long.until(EC.presence_of_element_located((By.ID, 'id-uem-bill-details-vendor-name')))
                    except TimeoutException:
                        logger.warning("Bill details content not fully loaded within timeout; checking for iframe...")
                        # If content might be in an iframe, attempt to switch
                        frames = self.driver.find_elements(By.TAG_NAME, 'iframe')
                        if frames:
                            self.driver.switch_to.frame(frames[0])
                            self.wait_long.until(EC.presence_of_element_located((By.ID, 'id-uem-bill-details-vendor-name')))
                    # Extract the required fields from the bill details page
                    try:
                        # Vendor name might include a slash and code (e.g., "VendorName/12345"), take the part before '/'
                        full_vendor_text = self.driver.find_element(By.ID, 'id-uem-bill-details-vendor-name').text
                        extracted_data['power_company'] = full_vendor_text.split('/')[0].strip()
                    except NoSuchElementException:
                        extracted_data['power_company'] = ''
                    try:
                        extracted_data['account_number'] = self.driver.find_element(By.ID, 'id-uem-bill-details-acct-number').text.strip()
                    except NoSuchElementException:
                        extracted_data['account_number'] = ''
                    try:
                        extracted_data['power_meter'] = self.driver.find_element(By.XPATH, "//td[contains(@class,'uem-bill-details-meter-number-widthSet')]").text.strip()
                    except NoSuchElementException:
                        extracted_data['power_meter'] = ''
                    # If an iframe was used for extraction, switch back to main document
                    try:
                        self.driver.switch_to.default_content()
                    except Exception:
                        pass
                    # Log and print the extracted data for verification
                    logger.info(f"Extracted Data for Site {site_id} -> Power Company: {extracted_data['power_company']}, "
                                f"Account Number: {extracted_data['account_number']}, Power Meter: {extracted_data['power_meter']}")
                    print(f"Extracted -> Power Company: {extracted_data['power_company']}, "
                          f"Account Number: {extracted_data['account_number']}, Power Meter: {extracted_data['power_meter']}")
                    break  # exit loop after finding the matching vendor and extracting data
            # End of loop over bill_rows
            if not extracted_data:
                # No data found for this site (vendor mismatch or no results)
                logger.warning(f"Could not extract data for Site ID {site_id} (vendor '{vendor}' not found in results). Skipping update for this site.")
                return  # skip updating IOP for this site
            # Switch to IOP tab to update the site’s Utility Info
            self.driver.switch_to.window(self.iop_handle)
            # Search for the site in IOP (using the Site/Switch search bar)
            search_field = self.wait_long.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Site/Switch Search']")))
            search_field.clear()
            search_field.send_keys(site_id)
            time.sleep(1)  # small pause for the dropdown to appear with results
            dropdown_option = self.wait_long.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='dropdown-item']")))
            dropdown_option.click()
            logger.info(f"Opened site {site_id} in IOP interface.")
            # Scroll to and expand the "Utility Info" section
            utility_header = self.wait_long.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Utility Info']/..")))
            self.driver.execute_script("arguments[0].scrollIntoView(true);", utility_header)
            time.sleep(1)
            utility_header.click()  # expand section if collapsible
            # Fill in the fields with extracted ENGIE data
            field_mapping = {
                'Power Company': ("//label[text()='Power Company']/following-sibling::input", extracted_data['power_company']),
                'Power Meter':   ("//label[text()='Power Meter']/following-sibling::input",   extracted_data['power_meter']),
                'Account Number':("//label[text()='Account Number']/following-sibling::input", extracted_data['account_number'])
            }
            for field_name, (xpath, value) in field_mapping.items():
                input_elem = self.wait_short.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                input_elem.clear()
                input_elem.send_keys(value)
                logger.info(f"Entered {field_name}: {value}")
            # Click the Save button to save the Utility Information
            save_button = self.wait_short.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Save Utility Information']")))
            save_button.click()
            logger.info(f"Utility information saved for Site ID {site_id} in IOP.")
            time.sleep(2)  # wait a moment to ensure save action completes (could be replaced with explicit wait for a success message if available)
        except Exception as e:
            logger.error(f"Error processing Site ID {site_id}: {e}")
            take_screenshot(self.driver, site_id, 'processing')  # take screenshot for debugging
            # Note: continuing to next site even if this one failed

    def run(self):
        """Run the automation for all sites."""
        try:
            self.setup_driver()
            self.login_engie()
            self.login_iop()
            for idx, (site_id, vendor) in enumerate(self.sites, start=1):
                self.process_site(idx, site_id, vendor)
        except Exception as e:
            logger.critical(f"Automation terminated due to unexpected error: {e}")
        finally:
            if self.driver:
                self.driver.quit()
                logger.info("Browser closed. Automation complete.")
