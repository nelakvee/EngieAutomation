import re
import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, StaleElementReferenceException

import config


# --- Helper Function ---
def normalize_vendor_name(name):
    """Cleans up vendor names for reliable comparison."""
    if not isinstance(name, str):
        return ""
    cleaned_name = re.sub(r'[^a-zA-Z0-9\s]', '', name).lower()
    return ' '.join(cleaned_name.split())


# --- Locator Dictionary ---
# Centralizing locators makes the script easier to maintain if the website UI changes.
class EngieLocators:
    USERNAME_INPUT = (By.ID, 'idp-discovery-username')
    SUBMIT_BUTTON = (By.ID, 'idp-discovery-submit')
    PLATFORM_BUTTON = (By.XPATH, "//span[@data-se='app-card-title' and @title='ENGIE Impact Platform']")
    SEARCH_INPUT = (By.XPATH, "//input")
    ENABLED_SEARCH_BUTTON = (By.XPATH, "//button[contains(@class, 'search-btn-enabled')]")
    LOADING_OVERLAY = (By.XPATH, "//div[contains(@class, 'ui-widget-overlay')]")
    BILL_RESULTS_GRID = (By.ID, 'BillResultsGrid')
    # This XPath finds rows that contain a vendor name link, making it robust.
    BILL_ROWS = (By.XPATH, ".//tr[.//a[contains(@id, 'VendorName')]]")
    VENDOR_NAME_LINK = (By.XPATH, ".//a[contains(@id, 'VendorName')]")
    VIEW_BUTTON = (By.XPATH, ".//a[text()='View...']")
    # This is the key to the iframe solution: a reliable locator for the frame itself.
    CONTENT_IFRAME = (By.XPATH, "//iframe[@title='content']")
    # Locators for elements *inside* the iframe
    POWER_COMPANY_SPAN = (By.ID, 'id-uem-bill-details-vendor-name')
    ACCOUNT_NUMBER_SPAN = (By.ID, 'id-uem-bill-details-acct-number')
    POWER_METER_CELL = (By.XPATH, "//td")


def login_to_engie(driver, username):
    """Handles the semi-automated login process for ENGIE."""
    print("--- Starting ENGIE Platform Login ---")
    wait = WebDriverWait(driver, config.SHORT_WAIT_TIME)

    driver.get(config.ENGIE_LOGIN_URL)
    wait.until(EC.visibility_of_element_located(EngieLocators.USERNAME_INPUT)).send_keys(username)
    wait.until(EC.element_to_be_clickable(EngieLocators.SUBMIT_BUTTON)).click()

    print("\n>>> ACTION REQUIRED: Please complete the multi-factor authentication for ENGIE in the browser.")
    input(">>> After you have successfully signed in, press Enter in this console to continue...")

    print("Waiting for the ENGIE application dashboard to load...")
    long_wait = WebDriverWait(driver, config.LONG_WAIT_TIME)
    platform_button = long_wait.until(EC.element_to_be_clickable(EngieLocators.PLATFORM_BUTTON))

    initial_windows = set(driver.window_handles)
    platform_button.click()
    print("✅ Clicked the ENGIE Impact Platform button.")

    print("Waiting for new platform tab and switching to it...")
    long_wait.until(EC.new_window_is_opened(initial_windows))
    new_window_handle = (set(driver.window_handles) - initial_windows).pop()
    driver.switch_to.window(new_window_handle)

    print("✅ Switched to ENGIE main dashboard tab.")
    return driver.current_window_handle


def extract_bill_data(driver, site_id, vendor_name_from_excel):
    """
    Searches for a site on ENGIE, finds the correct bill, and extracts data.
    Returns a dictionary with the data or None if not found.
    """
    engie_main_tab = driver.current_window_handle
    long_wait = WebDriverWait(driver, config.LONG_WAIT_TIME)

    # --- 1. Search for the Site ID ---
    print("   [ENGIE] Refreshing page to ensure a clean search state.")
    driver.refresh()

    # Wait for any loading overlays to disappear before proceeding. This is crucial for stability.
    print("   [ENGIE] Waiting for page to become fully interactive...")
    long_wait.until(EC.invisibility_of_element_located(EngieLocators.LOADING_OVERLAY))

    print(f"   [ENGIE] Entering Site ID '{site_id}' into search box...")
    search_field = long_wait.until(EC.element_to_be_clickable(EngieLocators.SEARCH_INPUT))
    search_field.clear()
    search_field.send_keys(site_id)

    print("   [ENGIE] Waiting for search button to become enabled...")
    search_button = long_wait.until(EC.element_to_be_clickable(EngieLocators.ENABLED_SEARCH_BUTTON))
    search_button.click()
    print("   [ENGIE] Search initiated.")

    # --- 2. Find Matching Bill and Click 'View' ---
    print(f"   [ENGIE] Searching for a bill matching vendor: '{vendor_name_from_excel}'...")
    normalized_vendor_from_excel = normalize_vendor_name(vendor_name_from_excel)

    try:
        long_wait.until(EC.presence_of_element_located(EngieLocators.BILL_RESULTS_GRID))
        bill_rows = long_wait.until(EC.presence_of_all_elements_located(EngieLocators.BILL_ROWS))
    except TimeoutException:
        print(f"   [ENGIE] ❌ No bill results found for Site ID '{site_id}'.")
        return None

    found_match = False
    for row in bill_rows:
        try:
            vendor_name_element = row.find_element(*EngieLocators.VENDOR_NAME_LINK)
            engie_vendor_text = vendor_name_element.text

            # Use normalized names for a more reliable comparison
            if normalized_vendor_from_excel in normalize_vendor_name(engie_vendor_text):
                print(f"   [ENGIE] ✅ Match found for vendor '{engie_vendor_text}'. Clicking 'View...'.")
                initial_windows = set(driver.window_handles)
                view_button = row.find_element(*EngieLocators.VIEW_BUTTON)
                view_button.click()

                # --- 3. Robustly Handle New Tab and Iframe ---
                print("   [ENGIE] Waiting for bill details tab to open...")
                long_wait.until(EC.new_window_is_opened(initial_windows))
                new_window_handle = (set(driver.window_handles) - initial_windows).pop()
                driver.switch_to.window(new_window_handle)
                print("   [ENGIE] ✅ Switched to bill details tab.")

                try:
                    # THE CORE IFRAME SOLUTION: A multi-stage wait protocol.
                    # Step A: Wait for the iframe element to be PRESENT in the DOM.
                    print("   [ENGIE] Waiting for content iframe to be present in the DOM...")
                    long_wait.until(EC.presence_of_element_located(EngieLocators.CONTENT_IFRAME))

                    # Step B: Now that it's present, wait for it to be fully available and switch to it.
                    print("   [ENGIE] Switching into the content iframe...")
                    long_wait.until(EC.frame_to_be_available_and_switch_to_it(EngieLocators.CONTENT_IFRAME))

                    # Step C: Wait for a known key element *inside* the frame to be visible.
                    print("   [ENGIE] Waiting for content inside iframe to load...")
                    long_wait.until(EC.visibility_of_element_located(EngieLocators.POWER_COMPANY_SPAN))
                    print("   [ENGIE] ✅ Iframe content loaded. Extracting data...")

                    # --- 4. Extract Data ---
                    power_company_raw = driver.find_element(*EngieLocators.POWER_COMPANY_SPAN).text
                    power_company = power_company_raw.split('/').strip()

                    account_number = driver.find_element(*EngieLocators.ACCOUNT_NUMBER_SPAN).text.strip()
                    power_meter = driver.find_element(*EngieLocators.POWER_METER_CELL).text.strip()

                    utility_data = {
                        "power_company": power_company,
                        "account_number": account_number,
                        "power_meter": power_meter
                    }

                    print("\n   --- Extracted Data from ENGIE ---")
                    print(f"     - Power Company: {utility_data['power_company']}")
                    print(f"     - Account Number: {utility_data['account_number']}")
                    print(f"     - Power Meter: {utility_data['power_meter']}\n")

                    found_match = True
                    return utility_data

                finally:
                    # --- 5. Cleanup ---
                    # Always ensure we clean up the state, regardless of success or failure inside the tab.
                    print("   [ENGIE] Closing bill details tab and returning to main dashboard.")
                    driver.close()
                    driver.switch_to.window(engie_main_tab)
                    # Break the loop once the correct vendor has been found and processed.
                    if found_match:
                        break

        except (NoSuchElementException, StaleElementReferenceException) as e:
            # This can happen if the page re-renders while iterating. Simply continue to the next row.
            print(f"   [ENGIE] Encountered a minor DOM issue while checking a row: {e}. Trying next row.")
            continue

    if not found_match:
        print(f"   [ENGIE] ❌ No bill rows found with a vendor name matching '{vendor_name_from_excel}'.")

    return None