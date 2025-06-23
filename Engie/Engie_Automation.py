import re
import time
import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException, \
    InvalidElementStateException


# --- Helper Function to Normalize Vendor Names ---
# This function cleans up vendor names for easier comparison.
def normalize_vendor_name(name):
    """
    Removes special characters and extra spaces from a name, and converts to lowercase.
    This helps in comparing vendor names that might have slight variations.
    Example: 'National Grid - New York/371376' becomes 'national grid new york'.
    """
    if not isinstance(name, str):
        return ""
    # Keep letters and numbers for a slightly more specific match if needed later
    cleaned_name = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned_name.split())


# --- Configuration ---
# Store credentials and file paths here for easy access.
# IMPORTANT: For security, avoid hardcoding credentials. Consider environment variables or a secure config file.
EXCEL_FILE_PATH = r'C:\Users\nelakve\Documents\Field Engineers\Engine_Site_ID_and_Vendor.xlsx'
ENGIE_LOGIN_URL = 'https://engieimpact.okta.com/'
ENGIE_USERNAME = 'veenith.nelakanti@verizonwireless.com'
IOP_LOGIN_URL = "https://iop.vh.vzwnet.com/user/nelakve/sites"
IOP_USERNAME = "nelakve"
IOP_PASSWORD = "Vamshin143@"  # Note: Storing passwords in plain text is not secure.
SCREENSHOT_FILE = "automation_error_screenshot.png"


# --- Main Automation Function ---
def run_automation():
    """
    Main function to orchestrate the entire automation process.
    """
    # --- 1. Excel Setup ---
    try:
        wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
        ws = wb.active
        # Read all sites into a list of dictionaries for processing
        sites_to_process = []
        # Start from row 2 to skip header
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):
            site_id_val, vendor_name_val = row
            if site_id_val:  # Process only if a Site ID exists
                site_id = str(site_id_val).strip()
                vendor_name = str(vendor_name_val).strip() if vendor_name_val else "Not Found"
                sites_to_process.append({"site_id": site_id, "vendor_name": vendor_name})

        if not sites_to_process:
            print("❌ No sites found to process in the Excel file. Exiting.")
            return

    except FileNotFoundError:
        print(f"❌ FATAL ERROR: The Excel file was not found at '{EXCEL_FILE_PATH}'.")
        return
    except Exception as e:
        print(f"❌ FATAL ERROR: Could not read the Excel file. Reason: {e}")
        return

    # --- 2. WebDriver Setup ---
    # The driver is set up once and used for the entire session.
    driver = webdriver.Chrome()
    long_wait = WebDriverWait(driver, 90)  # Increased wait time for very slow pages
    short_wait = WebDriverWait(driver, 30)
    driver.maximize_window()

    try:
        # --- 3. Initial Login to ENGIE ---
        print("--- Starting ENGIE Platform Automation ---")
        driver.get(ENGIE_LOGIN_URL)
        short_wait.until(EC.presence_of_element_located((By.ID, 'idp-discovery-username'))).send_keys(ENGIE_USERNAME)
        short_wait.until(EC.element_to_be_clickable((By.ID, 'idp-discovery-submit'))).click()

        print("\n>>> ACTION REQUIRED: Please complete the sign-in process for ENGIE in the browser.")
        input(">>> After you have successfully signed in, press Enter here to continue...")

        print("Waiting for the ENGIE application dashboard to load...")
        platform_button = long_wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[@data-se='app-card-title' and @title='ENGIE Impact Platform']"))
        )
        platform_button.click()
        print("✅ Successfully clicked the ENGIE Impact Platform button.")

        print("Switching to the new platform tab...")
        long_wait.until(EC.number_of_windows_to_be(2))
        driver.switch_to.window(driver.window_handles[1])
        engie_main_tab_handle = driver.current_window_handle

        # --- 4. ONE-TIME LOGIN TO IOP ---
        print("\n--- Performing one-time login to IOP Website ---")
        driver.execute_script("window.open('');")
        driver.switch_to.window(driver.window_handles[-1])
        iop_tab_handle = driver.current_window_handle  # Store the IOP tab handle

        driver.get(IOP_LOGIN_URL)
        print(f"Navigated to: {IOP_LOGIN_URL}")

        print("Waiting for IOP login fields...")
        short_wait.until(EC.presence_of_element_located((By.ID, 'idToken1'))).send_keys(IOP_USERNAME)
        short_wait.until(EC.presence_of_element_located((By.ID, 'idToken2'))).send_keys(IOP_PASSWORD)
        short_wait.until(EC.element_to_be_clickable((By.ID, "loginButton_0"))).click()
        print("✅ IOP login submitted. This tab will be reused.")

        # --- 5. Loop Through Each Site from Excel ---
        for i, site_data in enumerate(sites_to_process):
            site_id = site_data["site_id"]
            vendor_name = site_data["vendor_name"]
            normalized_vendor_from_excel = normalize_vendor_name(vendor_name)

            print("\n" + "=" * 50)
            print(f"Processing Site {i + 1}/{len(sites_to_process)}: Site ID '{site_id}', Vendor '{vendor_name}'")
            print("=" * 50)

            try:
                # --- 5a. ENGIE SEARCH ---
                driver.switch_to.window(engie_main_tab_handle)
                print("Refreshing ENGIE page for a clean search...")
                driver.refresh()

                loading_overlay_xpath = "//div[contains(@class, 'ui-widget-overlay')]"
                print("Waiting for ENGIE page to become fully interactive...")
                long_wait.until(EC.invisibility_of_element_located((By.XPATH, loading_overlay_xpath)))
                print("ENGIE page is now interactive.")

                search_field_xpath = "//input[@placeholder='Search' and contains(@class, 'search-box')]"

                max_retries = 4
                for attempt in range(max_retries):
                    try:
                        print(f"Attempting to interact with ENGIE search box (Attempt {attempt + 1}/{max_retries})...")
                        search_field = long_wait.until(EC.element_to_be_clickable((By.XPATH, search_field_xpath)))
                        driver.execute_script("arguments[0].value = '';", search_field)
                        for char in site_id:
                            search_field.send_keys(char)
                            time.sleep(0.05)
                        print(f"✅ Successfully entered Site ID '{site_id}' into ENGIE search.")
                        break
                    except (InvalidElementStateException, StaleElementReferenceException) as e:
                        print(
                            f"  - Encountered a temporary page instability ({type(e).__name__}). Retrying in 3 seconds...")
                        time.sleep(3)
                        if attempt == max_retries - 1: raise Exception(
                            "Failed to interact with ENGIE search box after multiple retries.")

                print("Waiting for ENGIE search button to become enabled...")
                enabled_search_button_xpath = "//button[contains(@class, 'search-btn-enabled')]"
                search_button = long_wait.until(EC.element_to_be_clickable((By.XPATH, enabled_search_button_xpath)))
                driver.execute_script("arguments[0].click();", search_button)
                print("✅ Search button clicked on ENGIE.")

                # --- 5b. Find Bill on ENGIE and Extract Data ---
                print("Searching for bill with a matching Vendor Name...")
                bill_grid_xpath = "//table[contains(@id, 'BillResultsGrid')]"
                long_wait.until(EC.presence_of_element_located((By.XPATH, bill_grid_xpath)))
                bill_rows_xpath = f"{bill_grid_xpath}//tr[.//a[contains(@id, 'VendorName')]]"
                long_wait.until(EC.presence_of_element_located((By.XPATH, bill_rows_xpath)))
                bill_rows = driver.find_elements(By.XPATH, bill_rows_xpath)

                power_company = ""
                account_number = ""
                power_meter = ""
                found_engie_match = False

                for row in bill_rows:
                    try:
                        vendor_name_element = row.find_element(By.XPATH, ".//a[contains(@id, 'VendorName')]")
                        if normalized_vendor_from_excel in normalize_vendor_name(vendor_name_element.text):
                            print(f"✅ Match found on ENGIE platform!")

                            # Get the current window handles BEFORE clicking
                            initial_windows = set(driver.window_handles)

                            view_button = row.find_element(By.XPATH, ".//a[text()='View...']")
                            driver.execute_script("arguments[0].click();", view_button)

                            # --- FINAL FIX: Most robust new tab and data extraction logic ---
                            print("Waiting for the new bill details tab to open...")
                            long_wait.until(EC.new_window_is_opened(initial_windows))

                            new_window_handle = (set(driver.window_handles) - initial_windows).pop()
                            driver.switch_to.window(new_window_handle)
                            print("✅ Switched to bill details tab.")

                            try:
                                print("Waiting for the bill details content frame to be available...")
                                content_frame_xpath = "//iframe[contains(@id, 'iframe') or contains(@name, 'iframe') or @title='content']"
                                long_wait.until(
                                    EC.frame_to_be_available_and_switch_to_it((By.XPATH, content_frame_xpath)))
                                print("✅ Successfully switched into the content frame.")

                                print("Waiting for content inside the frame to load...")
                                key_element_id = 'id-uem-bill-details-vendor-name'
                                long_wait.until(EC.visibility_of_element_located((By.ID, key_element_id)))
                                print("✅ Content inside frame has loaded. Extracting data...")

                                power_company_span = driver.find_element(By.ID, 'id-uem-bill-details-vendor-name')
                                power_company = power_company_span.text.split('/')[0].strip()

                                account_number_span = driver.find_element(By.ID, 'id-uem-bill-details-acct-number')
                                account_number = account_number_span.text.strip()

                                meter_number_td = driver.find_element(By.XPATH,
                                                                      "//td[@class='uem-bill-details-meter-number-widthSet wrapword']")
                                power_meter = meter_number_td.text.strip()

                                print("\n--- Extracted Data from ENGIE (Proof) ---")
                                print(
                                    f"  - Power Company: {power_company}\n  - Account Number: {account_number}\n  - Power Meter: {power_meter}")

                            finally:
                                driver.switch_to.default_content()
                                print("Switched back to the main page content.")

                            print("\nClosing the bill details tab.")
                            driver.close()
                            found_engie_match = True
                            break
                    except (NoSuchElementException, StaleElementReferenceException):
                        continue

                if not found_engie_match:
                    print(f"❌ No matching bill found on ENGIE for vendor '{vendor_name}'. Skipping to next site.")
                    continue

                # --- 5c. Re-use IOP Tab and Enter Data ---
                print("\n--- Switching to existing IOP Tab for data entry ---")
                driver.switch_to.window(iop_tab_handle)

                print(f"Waiting for IOP search field and entering Site ID: {site_id}")
                iop_search_input = long_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//input[contains(@placeholder, 'Site/Switch Search')]")))
                iop_search_input.clear()
                iop_search_input.send_keys(site_id)
                time.sleep(1)

                print("Waiting for search result dropdown on IOP...")
                dropdown_result_xpath = "//a[@class='dropdown-item']"
                dropdown_result = long_wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_result_xpath)))
                print("Clicking on the site from the dropdown...")
                dropdown_result.click()

                print("Waiting for 'fuze Utility Info' section to load...")
                utility_info_xpath = "//div[contains(@class, 'tmET65dnMPIvRdTilZcA') and .//span[contains(text(), 'Utility Info')]]"
                utility_info_header = long_wait.until(EC.presence_of_element_located((By.XPATH, utility_info_xpath)))

                print("Scrolling to and expanding 'fuze Utility Info' section...")
                driver.execute_script("arguments[0].scrollIntoView(true);", utility_info_header)
                time.sleep(1)
                utility_info_header.click()

                print("Entering extracted data into IOP fields...")
                power_company_input = short_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//label[text()='Power Company']/following-sibling::input")))
                power_company_input.clear()
                power_company_input.send_keys(power_company)
                print(f"  - Entered Power Company: {power_company}")

                power_meter_input = short_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//label[text()='Power Meter']/following-sibling::input")))
                power_meter_input.clear()
                power_meter_input.send_keys(power_meter)
                print(f"  - Entered Power Meter: {power_meter}")

                account_number_input = short_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//label[text()='Account Number']/following-sibling::input")))
                account_number_input.clear()
                account_number_input.send_keys(account_number)
                print(f"  - Entered Account Number: {account_number}")

                print("Data entry complete.")

                save_button = short_wait.until(
                    EC.element_to_be_clickable((By.XPATH, "//button[text()='Save Utility Information']")))
                # save_button.click() # UNCOMMENT THIS LINE TO ACTUALLY SAVE THE DATA
                print("✅✅✅ Successfully processed and entered data for Site ID '{site_id}'.")
                print("NOTE: The 'Save' button click is currently COMMENTED OUT for safety.")
                print("IOP tab will remain open for the next site.")

            except Exception as e:
                print(f"❌ An error occurred while processing Site ID '{site_id}': {type(e).__name__} - {e}")
                print("Saving a screenshot for debugging and moving to the next site...")
                try:
                    driver.save_screenshot(f"error_{site_id}.png")
                except Exception as screenshot_e:
                    print(f"Could not save screenshot: {screenshot_e}")

                if engie_main_tab_handle in driver.window_handles:
                    driver.switch_to.window(engie_main_tab_handle)
                else:
                    print("Main ENGIE tab was closed. Cannot continue. Please restart.")
                    raise Exception("Main ENGIE tab was closed, recovery not possible.")
                continue

    except Exception as e:
        print(f"\n❌ A FATAL, UNEXPECTED ERROR occurred: {type(e).__name__} - {e}")
        print(f"A screenshot of the page at the time of the error has been saved to '{SCREENSHOT_FILE}'.")
        driver.save_screenshot(SCREENSHOT_FILE)

    finally:
        print("\n--- Automation Complete ---")
        input("Press Enter to close the script and the browser...")
        driver.quit()


# --- Run the Script ---
if __name__ == "__main__":
    run_automation()
