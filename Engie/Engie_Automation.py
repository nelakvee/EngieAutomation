import re
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, InvalidElementStateException

# --- Helper Functions ---
def normalize_vendor_name(name):
    """
    Cleans vendor names for comparison: removes special chars, extra spaces, converts to lowercase.
    """
    if not isinstance(name, str):
        return ""
    cleaned = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower().strip()
    return ' '.join(cleaned.split())


def wait_for_page_load(driver, timeout=90):
    """Waits until the document.readyState is complete."""
    WebDriverWait(driver, timeout).until(
        lambda d: d.execute_script('return document.readyState') == 'complete'
    )

# --- Configuration ---
EXCEL_FILE_PATH = r'C:\Users\nelakve\Documents\Field Engineers\Engine_Site_ID_and_Vendor.xlsx'
ENGIE_LOGIN_URL = 'https://engieimpact.okta.com/'
ENGIE_USERNAME = 'veenith.nelakanti@verizonwireless.com'
IOP_LOGIN_URL = 'https://iop.vh.vzwnet.com/user/nelakve/sites'
IOP_USERNAME = 'nelakve'
IOP_PASSWORD = 'Vamshin143@'
SCREENSHOT_TEMPLATE = 'error_{site_id}.png'


def run_automation():
    # --- 1. Load Excel ---
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        sites = []
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            sid, vend = row
            if sid:
                sites.append({'site_id': str(sid).strip(), 'vendor_name': str(vend or '').strip()})
        if not sites:
            print('No sites found. Exiting.')
            return
    except Exception as e:
        print(f"Failed to read Excel: {e}")
        return

    # --- 2. Setup WebDriver ---
    driver = webdriver.Chrome()
    driver.maximize_window()
    long_wait = WebDriverWait(driver, 90)
    short_wait = WebDriverWait(driver, 30)

    try:
        # --- 3. Login to ENGIE (manual MFA) ---
        driver.get(ENGIE_LOGIN_URL)
        wait_for_page_load(driver)
        short_wait.until(EC.element_to_be_clickable((By.ID, 'idp-discovery-username'))).send_keys(ENGIE_USERNAME)
        short_wait.until(EC.element_to_be_clickable((By.ID, 'idp-discovery-submit'))).click()
        input('>>> Complete ENGIE sign-in, then press Enter to continue...')

        platform_btn = long_wait.until(EC.element_to_be_clickable((By.XPATH,
            "//span[@data-se='app-card-title' and @title='ENGIE Impact Platform']")))
        platform_btn.click()
        long_wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[-1])
        engie_tab = driver.current_window_handle
        wait_for_page_load(driver)

        # --- 4. Login to IOP ---
        driver.execute_script('window.open();')
        driver.switch_to.window(driver.window_handles[-1])
        iop_tab = driver.current_window_handle
        driver.get(IOP_LOGIN_URL)
        wait_for_page_load(driver)
        short_wait.until(EC.presence_of_element_located((By.ID, 'idToken1'))).send_keys(IOP_USERNAME)
        short_wait.until(EC.presence_of_element_located((By.ID, 'idToken2'))).send_keys(IOP_PASSWORD)
        short_wait.until(EC.element_to_be_clickable((By.ID, 'loginButton_0'))).click()
        wait_for_page_load(driver)

        # --- 5. Process Each Site ---
        for idx, entry in enumerate(sites, 1):
            sid = entry['site_id']
            vendor = entry['vendor_name']
            norm_vendor = normalize_vendor_name(vendor)
            print(f"\n--- Processing {idx}/{len(sites)}: {sid} ---")

            try:
                # 5a. Search in ENGIE
                driver.switch_to.window(engie_tab)
                driver.refresh()
                wait_for_page_load(driver)
                long_wait.until(EC.invisibility_of_element_located((By.XPATH, "//div[contains(@class,'ui-widget-overlay')]")))

                search_input = long_wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//input[contains(@class,'search-box') and @placeholder='Search']")))
                driver.execute_script('arguments[0].value="";', search_input)
                for ch in sid:
                    search_input.send_keys(ch)
                    time.sleep(0.05)

                search_btn = long_wait.until(EC.element_to_be_clickable((By.XPATH,
                    "//button[contains(@class,'search-btn-enabled')]")))
                driver.execute_script('arguments[0].click();', search_btn)
                wait_for_page_load(driver)

                rows = long_wait.until(EC.presence_of_all_elements_located((By.XPATH,
                    "//table[contains(@id,'BillResultsGrid')]//tr[.//a[contains(@id,'VendorName')]]")))

                # find matching vendor and extract
                found = False
                for row in rows:
                    try:
                        vtext = row.find_element(By.XPATH, ".//a[contains(@id,'VendorName')]").text
                        if norm_vendor in normalize_vendor_name(vtext):
                            # scroll + click
                            view_btn = row.find_element(By.XPATH, ".//a[text()='View...']")
                            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", view_btn)
                            long_wait.until(EC.element_to_be_clickable((By.XPATH, ".//a[text()='View...']")))
                            driver.execute_script('arguments[0].click();', view_btn)

                            long_wait.until(EC.new_window_is_opened)
                            new_win = [w for w in driver.window_handles if w != engie_tab and w != iop_tab][0]
                            driver.switch_to.window(new_win)
                            wait_for_page_load(driver)

                            vendor_el = long_wait.until(EC.visibility_of_element_located((By.ID, 'id-uem-bill-details-vendor-name')))
                            acct_el   = short_wait.until(EC.visibility_of_element_located((By.ID, 'id-uem-bill-details-acct-number')))
                            meter_el  = driver.find_element(By.XPATH,
                                "//td[contains(@class,'uem-bill-details-meter-number-widthSet') and contains(@class,'wrapword')]")

                            power_company  = vendor_el.text.split('/')[0].strip()
                            account_number = acct_el.text.strip()
                            power_meter    = meter_el.text.strip()
                            print(f"Extracted: {power_company}, {account_number}, {power_meter}")

                            driver.close()
                            driver.switch_to.window(engie_tab)
                            found = True
                            break
                    except Exception:
                        continue

                if not found:
                    print(f"No matching bill for vendor '{vendor}'. Skipping.")
                    continue

                # 5b. Enter into IOP
                driver.switch_to.window(iop_tab)
                wait_for_page_load(driver)
                # wait for search field
                search_iop = long_wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@placeholder='Site/Switch Search']")))
                search_iop.clear()
                search_iop.send_keys(sid)
                time.sleep(1)
                long_wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='dropdown-item']"))).click()
                wait_for_page_load(driver)

                # expand Utility Info
                util_hdr = long_wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Utility Info']/..")))
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', util_hdr)
                long_wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Utility Info']/.."))).click()

                # fill fields
                pc_input = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//label[text()='Power Company']/following-sibling::input")))
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', pc_input)
                pc_input.clear()
                pc_input.send_keys(power_company)

                pm_input = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//label[text()='Power Meter']/following-sibling::input")))
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', pm_input)
                pm_input.clear()
                pm_input.send_keys(power_meter)

                an_input = short_wait.until(EC.element_to_be_clickable((By.XPATH, "//label[text()='Account Number']/following-sibling::input")))
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', an_input)
                an_input.clear()
                an_input.send_keys(account_number)

                # save
                save_btn = long_wait.until(EC.element_to_be_clickable((By.XPATH, "//button[text()='Save Utility Information']")))
                driver.execute_script('arguments[0].scrollIntoView({block:"center"});', save_btn)
                driver.execute_script('arguments[0].click();', save_btn)
                wait_for_page_load(driver)
                print(f"Data saved for {sid}.")

            except Exception as err:
                print(f"Error on site {sid}: {err}")
                try:
                    driver.save_screenshot(SCREENSHOT_TEMPLATE.format(site_id=sid))
                except:
                    pass
                driver.switch_to.window(engie_tab)
                continue

    finally:
        print('Automation completed. Closing browser.')
        driver.quit()

if __name__ == '__main__':
    run_automation()
1
