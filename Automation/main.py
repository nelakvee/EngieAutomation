# main.py
"""
Main orchestration script for the ENGIE-to-IOP data transfer automation.
"""
import csv
import logging
from datetime import datetime
from selenium.common.exceptions import WebDriverException

import config
from utils.driver_setup import get_webdriver
from utils.logger_setup import setup_logger
from page_objects.engie_page import EngiePage
from page_objects.iop_page import IopPage


def take_screenshot(driver, site_id, stage):
    """Saves a timestamped screenshot on failure."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{config.SCREENSHOT_DIR}failure_{stage}_{site_id}_{timestamp}.png"
    try:
        driver.save_screenshot(filename)
        logging.info(f"Screenshot saved to {filename}")
    except WebDriverException as e:
        logging.error(f"Failed to take screenshot: {e}")


def process_site_id(driver, site_id: str):
    """
    Executes the full workflow for a single Site ID.
    1. Logs into ENGIE and extracts bill data.
    2. Logs into IOP and updates utility info.
    """
    # --- Step 1: ENGIE Workflow ---
    engie = EngiePage(driver)
    engie.login(config.ENGIE_URL, config.ENGIE_USERNAME, config.ENGIE_PASSWORD)
    engie.search_for_site(site_id)
    engie.go_to_latest_bill()
    bill_data = engie.extract_bill_details()

    # Print extracted data to screen as requested
    print("\n--- Extracted Data from ENGIE ---")
    print(f"  Power Company: {bill_data['power_company']}")
    print(f"  Account Number: {bill_data['account_number']}")
    print(f"  Power Meter: {bill_data['power_meter']}")
    print("---------------------------------\n")
    logging.info(f"Data extraction from ENGIE for Site ID {site_id} complete.")

    # --- Step 2: IOP Workflow ---
    iop = IopPage(driver)
    iop.login(config.IOP_URL, config.IOP_USERNAME, config.IOP_PASSWORD)
    iop.search_and_navigate_to_site(site_id)
    iop.fill_utility_info(
        power_company=bill_data['power_company'],
        account_number=bill_data['account_number'],
        power_meter=bill_data['power_meter']
    )
    iop.save_utility_info()
    logging.info(f"IOP update for Site ID {site_id} complete.")


def main():
    """
    Main function to initialize and run the automation.
    """
    setup_logger()
    logging.info("Automation script started.")

    # Read Site IDs from a CSV file
    try:
        with open(config.INPUT_FILE_PATH, mode='r', encoding='utf-8') as infile:
            reader = csv.reader(infile)
            next(reader)  # Skip header row
            site_ids = [row for row in reader]
    except FileNotFoundError:
        logging.critical(f"Input file not found at {config.INPUT_FILE_PATH}. Exiting.")
        return

    logging.info(f"Found {len(site_ids)} Site IDs to process.")

    driver = None
    for site_id in site_ids:
        try:
            # Initialize a new driver for each iteration for maximum stability,
            # or initialize once outside the loop for speed.
            # For this highly complex, multi-site workflow, a fresh session is safer.
            if driver:
                driver.quit()

            driver = get_webdriver()
            logging.info(f"--- Starting processing for Site ID: {site_id} ---")
            process_site_id(driver, site_id)
            logging.info(f"--- Successfully completed processing for Site ID: {site_id} ---")

        except WebDriverException as e:
            logging.error(f"A critical error occurred while processing Site ID {site_id}: {e}")
            if driver:
                # Determine stage for better screenshot naming
                stage = "IOP" if "iop" in driver.current_url else "ENGIE"
                take_screenshot(driver, site_id, stage)
            logging.warning(f"Skipping Site ID {site_id} due to error. Moving to next.")
            continue  # Move to the next site_id
        except Exception as e:
            logging.critical(f"An unexpected non-Selenium error occurred for Site ID {site_id}: {e}")
            if driver:
                stage = "IOP" if "iop" in driver.current_url else "ENGIE"
                take_screenshot(driver, site_id, stage)
            logging.warning(f"Skipping Site ID {site_id}. Moving to next.")
            continue

    if driver:
        driver.quit()

    logging.info("Automation script finished.")


if __name__ == "__main__":
    main()