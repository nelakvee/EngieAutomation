# utils/driver_setup.py
"""
Utility for setting up and configuring the Selenium WebDriver.
"""
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def get_webdriver():
    """
    Configures and returns a Chrome WebDriver instance.
    - Sets page load strategy to 'eager' for faster navigation.
    - Installs and manages the correct chromedriver binary.
    """
    chrome_options = Options()
    # 'eager' returns control after HTML is parsed, not waiting for all images/CSS.
    # This speeds up navigation but requires robust explicit waits for interactions.
    chrome_options.page_load_strategy = 'eager'
    chrome_options.add_argument("--start-maximized")
    # Uncomment the line below to run in headless mode (no visible browser window)
    # chrome_options.add_argument("--headless")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver