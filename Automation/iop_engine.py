import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config


# --- Locator Dictionary ---
class IopLocators:
    USERNAME_INPUT = (By.ID, 'idToken1')
    PASSWORD_INPUT = (By.ID, 'idToken2')
    LOGIN_BUTTON = (By.ID, 'loginButton_0')
    SITE_SEARCH_INPUT = (By.XPATH, "//input")
    SEARCH_DROPDOWN_RESULT = (By.XPATH, "//a[@class='dropdown-item']")
    # This locator is more specific and robust for finding the Utility Info section.
    UTILITY_INFO_HEADER = (By.XPATH, "//div[contains(@class, 'card-header') and.//span[text()='Utility Info']]")
    # Using following-sibling is a powerful and stable way to find inputs associated with labels.
    POWER_COMPANY_INPUT = (By.XPATH, "//label[text()='Power Company']/following-sibling::input")
    POWER_METER_INPUT = (By.XPATH, "//label[text()='Power Meter']/following-sibling::input")
    ACCOUNT_NUMBER_INPUT = (By.XPATH, "//label[text()='Account Number']/following-sibling::input")
    SAVE_UTILITY_BUTTON = (By.XPATH, "//button")


def login_to_iop(driver, username, password):
    """Handles the fully automated login for the IOP platform in a new tab."""
    print("--- Starting IOP Platform Login ---")
    wait = WebDriverWait(driver, config.SHORT_WAIT_TIME)

    # Open IOP in a new, dedicated tab
    driver.execute_script("window.open('');")
    iop_tab_handle = driver.window_handles[-1]
    driver.switch_to.window(iop_tab_handle)

    driver.get(config.IOP_LOGIN_URL)
    print(f"   [IOP] Navigated to: {config.IOP_LOGIN_URL}")

    print("   [IOP] Entering credentials...")
    wait.until(EC.visibility_of_element_located(IopLocators.USERNAME_INPUT)).send_keys(username)
    wait.until(EC.visibility_of_element_located(IopLocators.PASSWORD_INPUT)).send_keys(password)
    wait.until(EC.element_to_be_clickable(IopLocators.LOGIN_BUTTON)).click()

    # Wait for the search input to appear, which confirms successful login and page load.
    print("   [IOP] Waiting for IOP dashboard to load after login...")
    long_wait = WebDriverWait(driver, config.LONG_WAIT_TIME)
    long_wait.until(EC.visibility_of_element_located(IopLocators.SITE_SEARCH_INPUT))
    print("   [IOP] ✅ Login successful. IOP tab is ready.")

    return iop_tab_handle


def update_iop_record(driver, iop_tab_handle, site_id, utility_data):
    """
    Switches to the IOP tab, finds the site, and enters the utility data.
    """
    print("\n   --- Updating IOP Record ---")
    driver.switch_to.window(iop_tab_handle)
    long_wait = WebDriverWait(driver, config.LONG_WAIT_TIME)
    short_wait = WebDriverWait(driver, config.SHORT_WAIT_TIME)

    # --- 1. Search for Site ---
    print(f"   [IOP] Searching for Site ID: {site_id}")
    search_input = long_wait.until(EC.element_to_be_clickable(IopLocators.SITE_SEARCH_INPUT))
    search_input.clear()
    search_input.send_keys(site_id)
    time.sleep(1)  # A small static pause can help dropdowns appear reliably.

    print("   [IOP] Waiting for search result dropdown...")
    dropdown_result = long_wait.until(EC.element_to_be_clickable(IopLocators.SEARCH_DROPDOWN_RESULT))
    dropdown_result.click()
    print("   [IOP] ✅ Navigated to site details page.")

    # --- 2. Expand Utility Info Section ---
    print("   [IOP] Waiting for 'Utility Info' section to load...")
    utility_info_header = long_wait.until(EC.presence_of_element_located(IopLocators.UTILITY_INFO_HEADER))

    print("   [IOP] Scrolling to and expanding 'Utility Info' section...")
    # Use JavaScript for reliable scrolling and clicking, which can bypass some UI obstructions.
    driver.execute_script("arguments.scrollIntoView({block: 'center'});", utility_info_header)
    time.sleep(1)  # Allow for scroll animation to complete.
    driver.execute_script("arguments.click();", utility_info_header)

    # --- 3. Enter Data into Fields ---
    print("   [IOP] Entering extracted data into form fields...")

    power_company_input = short_wait.until(EC.visibility_of_element_located(IopLocators.POWER_COMPANY_INPUT))
    power_company_input.clear()
    power_company_input.send_keys(utility_data["power_company"])
    print(f"     - Entered Power Company: {utility_data['power_company']}")

    power_meter_input = short_wait.until(EC.visibility_of_element_located(IopLocators.POWER_METER_INPUT))
    power_meter_input.clear()
    power_meter_input.send_keys(utility_data["power_meter"])
    print(f"     - Entered Power Meter: {utility_data['power_meter']}")

    account_number_input = short_wait.until(EC.visibility_of_element_located(IopLocators.ACCOUNT_NUMBER_INPUT))
    account_number_input.clear()
    account_number_input.send_keys(utility_data["account_number"])
    print(f"     - Entered Account Number: {utility_data['account_number']}")

    # --- 4. Save Changes (Action is commented out for safety) ---
    save_button = short_wait.until(EC.element_to_be_clickable(IopLocators.SAVE_UTILITY_BUTTON))

    # UNCOMMENT THE FOLLOWING LINE TO ENABLE SAVING DATA TO IOP
    # save_button.click()

    print("\n   [IOP] ✅ Data entry complete for this site.")
    print("   [IOP] NOTE: The 'Save Utility Information' button click is currently COMMENTED OUT for safety.")