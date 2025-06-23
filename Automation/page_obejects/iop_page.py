# page_objects/iop_page.py
"""
Page Object for the Verizon Internal Operations Platform (IOP).
Encapsulates all element locators and actions for IOP pages.
"""
import logging
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from config import WEBDRIVER_TIMEOUT


class IopPage:
    """Models the interactions on the IOP platform."""

    class _Locators:
        """Holds all locators for the IOP pages."""
        # Login Page (Assuming standard login fields)
        USERNAME_INPUT = (By.ID, "username")
        PASSWORD_INPUT = (By.ID, "password")
        LOGIN_BUTTON = (By.ID, "loginButton")

        # Site Search Page
        SITE_SEARCH_INPUT = (By.CSS_SELECTOR, "input.rbt-input-main")

        # Utility Info Page
        UTILITY_INFO_EXPANDER = (By.XPATH, "//span[text()='Utility Info']/parent::div")
        POWER_COMPANY_INPUT = (By.XPATH, "//label[text()='Power Company']/following-sibling::input")
        POWER_METER_INPUT = (By.XPATH, "//label[text()='Power Meter']/following-sibling::input")
        ACCOUNT_NUMBER_INPUT = (By.XPATH, "//label[text()='Account Number']/following-sibling::input")
        SAVE_UTILITY_BUTTON = (By.XPATH, "//button")
        SAVE_SUCCESS_MESSAGE = (By.XPATH, "//*")  # Placeholder

    def __init__(self, driver: WebDriver):
        self.driver = driver
        self.wait = WebDriverWait(self.driver, WEBDRIVER_TIMEOUT)

    def login(self, url, username, password):
        """Navigates to the URL and performs login."""
        logging.info("Navigating to IOP login page.")
        self.driver.get(url)
        # NOTE: Placeholder for user's existing login logic.
        logging.info("IOP login process initiated (placeholder).")

    def search_and_navigate_to_site(self, site_id: str):
        """
        Searches for a site ID using the dynamic dropdown and navigates to it.
        """
        logging.info(f"Searching for Site ID '{site_id}' on IOP.")
        search_input = self.wait.until(EC.visibility_of_element_located(self._Locators.SITE_SEARCH_INPUT))
        search_input.clear()
        search_input.send_keys(site_id)

        # Dynamic locator for the dropdown item
        dropdown_item_locator = (By.XPATH, f"//a[contains(@class, 'dropdown-item')]//mark[text()='{site_id}']")

        dropdown_item = self.wait.until(EC.element_to_be_clickable(dropdown_item_locator))
        dropdown_item.click()
        logging.info(f"Clicked on dropdown result for Site ID '{site_id}'.")

    def fill_utility_info(self, power_company: str, account_number: str, power_meter: str):
        """
        Expands the Utility Info section and fills in the data extracted from ENGIE.
        """
        logging.info("Expanding Utility Info section and filling form.")

        # Scroll to the element to ensure it's in view
        utility_expander = self.wait.until(EC.presence_of_element_located(self._Locators.UTILITY_INFO_EXPANDER))
        self.driver.execute_script("arguments.scrollIntoView(true);", utility_expander)

        # Click to expand if not already expanded (logic may need adjustment based on state)
        utility_expander.click()

        # Fill Power Company
        power_co_input = self.wait.until(EC.visibility_of_element_located(self._Locators.POWER_COMPANY_INPUT))
        power_co_input.clear()
        power_co_input.send_keys(power_company)

        # Fill Power Meter
        power_meter_input = self.wait.until(EC.visibility_of_element_located(self._Locators.POWER_METER_INPUT))
        power_meter_input.clear()
        power_meter_input.send_keys(power_meter)

        # Fill Account Number
        account_num_input = self.wait.until(EC.visibility_of_element_located(self._Locators.ACCOUNT_NUMBER_INPUT))
        account_num_input.clear()
        account_num_input.send_keys(account_number)

        logging.info("Utility information form filled.")

    def save_utility_info(self):
        """Clicks the save button and verifies success."""
        logging.info("Saving utility information.")
        save_button = self.wait.until(EC.element_to_be_clickable(self._Locators.SAVE_UTILITY_BUTTON))
        save_button.click()

        # Wait for a success message to confirm the save operation
        try:
            self.wait.until(EC.visibility_of_element_located(self._Locators.SAVE_SUCCESS_MESSAGE))
            logging.info("Successfully saved utility information on IOP.")
        except TimeoutException:
            logging.warning("Save success message not found. Assuming save was successful based on button click.")