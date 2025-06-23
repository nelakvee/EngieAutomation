# page_objects/engie_page.py
"""
Page Object for the ENGIE Impact Platform.
Encapsulates all element locators and actions for ENGIE pages.
"""
import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from config import WEBDRIVER_TIMEOUT


class EngiePage:
    """Models the interactions on the ENGIE platform."""

    class _Locators:
        """Holds all locators for the ENGIE pages."""
        # Login Page (Assuming standard login fields)
        USERNAME_INPUT = (By.ID, "username")
        PASSWORD_INPUT = (By.ID, "password")
        LOGIN_BUTTON = (By.ID, "loginButton")

        # Dashboard Page
        SEARCH_INPUT = (By.CSS_SELECTOR, "input.search-box")
        SEARCH_BUTTON = (By.CSS_SELECTOR, "button.search-btn-enabled")

        # Search Results (Assuming a link to the bill)
        VIEW_BILL_LINK = (By.XPATH, "//a[contains(text(), 'View')]")  # Placeholder, needs refinement

        # Bill Details Page
        BILL_DETAILS_IFRAME = (By.ID, "id-bill-details-content-frame")  # This is a predicted ID, may need adjustment
        VENDOR_NAME_SPAN = (By.ID, "id-uem-bill-details-vendor-name")
        ACCOUNT_NUMBER_SPAN = (By.ID, "id-uem-bill-details-acct-number")
        METER_NUMBER_TD = (By.CSS_SELECTOR, "td.uem-bill-details-meter-number-widthSet")

    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, WEBDRIVER_TIMEOUT)

    def login(self, url, username, password):
        """Navigates to the URL and performs login."""
        logging.info("Navigating to ENGIE login page.")
        self.driver.get(url)
        # NOTE: The actual login locators and flow may differ.
        # This is a placeholder for the user's existing working login logic.
        # self.wait.until(EC.visibility_of_element_located(self._Locators.USERNAME_INPUT)).send_keys(username)
        # self.wait.until(EC.visibility_of_element_located(self._Locators.PASSWORD_INPUT)).send_keys(password)
        # self.wait.until(EC.element_to_be_clickable(self._Locators.LOGIN_BUTTON)).click()
        logging.info("ENGIE login process initiated (placeholder). Assuming manual login or existing session for now.")

    def search_for_site(self, site_id: str):
        """Searches for a given Site ID on the ENGIE dashboard."""
        logging.info(f"Searching for Site ID '{site_id}' on ENGIE.")
        search_box = self.wait.until(EC.visibility_of_element_located(self._Locators.SEARCH_INPUT))
        search_box.clear()
        search_box.send_keys(site_id)

        search_button = self.wait.until(EC.element_to_be_clickable(self._Locators.SEARCH_BUTTON))
        search_button.click()
        logging.info("Search button clicked.")

    def go_to_latest_bill(self):
        """Finds and clicks the link to the latest bill, handling the new tab."""
        logging.info("Waiting for bill search results and navigating to details.")
        original_window = self.driver.current_window_handle

        # This locator is a placeholder and may need to be more specific
        # based on the actual vendor name matching logic required.
        view_bill_button = self.wait.until(EC.element_to_be_clickable(self._Locators.VIEW_BILL_LINK))
        view_bill_button.click()

        self.wait.until(EC.number_of_windows_to_be(2))

        # Switch to the new tab
        for window_handle in self.driver.window_handles:
            if window_handle != original_window:
                self.driver.switch_to.window(window_handle)
                logging.info("Switched to new bill details tab.")
                break

        # Close the old, redundant tab
        self.driver.switch_to.window(original_window)
        self.driver.close()
        self.driver.switch_to.window(self.driver.window_handles)
        logging.info("Closed original ENGIE tab to conserve memory.")

    def extract_bill_details(self) -> dict:
        """
        Extracts vendor name, account number, and meter number from the bill details page.
        This method handles switching into and out of the content iframe.
        """
        logging.info("Extracting bill details from the content iframe.")
        try:
            # THE CRITICAL STEP: Wait for the iframe and switch to it.
            # The locator for the iframe needs to be confirmed from the actual page source.
            # A more robust locator might be needed, e.g., (By.CSS_SELECTOR, "iframe")
            self.wait.until(EC.frame_to_be_available_and_switch_to_it(self._Locators.BILL_DETAILS_IFRAME))
            logging.info("Successfully switched into the bill details iframe.")

            # Now, find elements within the iframe
            vendor_name = self.wait.until(EC.visibility_of_element_located(self._Locators.VENDOR_NAME_SPAN)).text
            account_number = self.wait.until(EC.visibility_of_element_located(self._Locators.ACCOUNT_NUMBER_SPAN)).text
            power_meter = self.wait.until(EC.visibility_of_element_located(self._Locators.METER_NUMBER_TD)).text

            extracted_data = {
                "power_company": vendor_name.strip(),
                "account_number": account_number.strip(),
                "power_meter": power_meter.strip()
            }

            logging.info(f"Successfully extracted data: {extracted_data}")
            return extracted_data

        finally:
            # CRITICAL CLEANUP: Always switch back to the main document context.
            self.driver.switch_to.default_content()
            logging.info("Switched back to the main page content.")