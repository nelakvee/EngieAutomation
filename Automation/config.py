# config.py
"""
Centralized configuration for the automation script.
Stores URLs, credentials, and file paths.
"""

# ENGIE Platform Configuration
ENGIE_URL = "https://platformprd.engieimpact.com/sites/Verizon/SitePages/Dashboard.aspx"
ENGIE_USERNAME = "veenith.nelakanti@verizonwireless.com"  # <-- IMPORTANT: Replace with actual username
ENGIE_PASSWORD = "Vamshin143@"  # <-- IMPORTANT: Replace with actual password

# IOP Platform Configuration
IOP_URL = "https://iop.vh.vzwnet.com/user/nelakve/sites"
IOP_USERNAME = "nelakve"  # <-- IMPORTANT: Replace with actual username
IOP_PASSWORD = "Vamshin143@"  # <-- IMPORTANT: Replace with actual password

# File and Directory Paths
INPUT_FILE_PATH = "C:/Users/nelakve/Documents/Field Engineers/Engine_Site_ID_and_Vendor.xlsx"
LOG_FILE_PATH = "output/logs/automation_run.log"
SCREENSHOT_DIR = "output/screenshots/"

# WebDriver and Wait Configuration
WEBDRIVER_TIMEOUT = 20  # Maximum time in seconds for explicit waits