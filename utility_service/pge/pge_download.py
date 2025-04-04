import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Define ChromeDriver path
chrome_driver_path = r"path/to/your/input/folder/chromedriver.exe"
service = Service(chrome_driver_path)

# Define login credentials
USERNAME = "username"
PASSWORD = "password"

# Define download directory (Commented out for now)
# download_directory = r"path/to/your/input/folder"

# Set Chrome options
chrome_options = webdriver.ChromeOptions()
# Commented out download directory settings
# chrome_options.add_experimental_option("prefs", {
#     "download.default_directory": download_directory,  # Set the download folder
#     "download.prompt_for_download": False,  # Disable the download prompt
#     "plugins.always_open_pdf_externally": True  # Prevent PDFs from opening in browser
# })

# Start Selenium WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 10)  # Set explicit wait

def login():
    """Logs into PG&E if logged out."""
    print("Re-logging in...")
    driver.get("https://m.pge.com/#login")
    time.sleep(3)

    # Check if username field is present (confirming we're at the login page)
    try:
        username_field = wait.until(EC.presence_of_element_located((By.ID, "usernameField")))
        password_field = driver.find_element(By.ID, "passwordField")
        login_button = driver.find_element(By.ID, "home_login_submit")

        username_field.send_keys(USERNAME)
        password_field.send_keys(PASSWORD)
        login_button.click()
        time.sleep(5)  # Ensure login is complete
        print("Successfully logged back in.")
    except Exception as e:
        print(f"Error logging in: {e}")

def is_logged_out():
    """Checks if the script is logged out by looking for the login form."""
    try:
        # If the username field is present, we're logged out
        driver.find_element(By.ID, "usernameField")
        print("Detected logout.")
        return True
    except:
        return False  # No login form found = still logged in

def handle_popup():
    """Detects and clicks 'Cancel' on the contact info popup if it appears."""
    try:
        popup = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.CLASS_NAME, "modal-dialog"))
        )
        print("Contact info popup detected.")

        # Wait for the cancel button and click it
        cancel_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "mfaBusPopupCancelbtn"))
        )
        cancel_button.click()
        print("Clicked 'Cancel' on the popup.")

        time.sleep(3)  # Allow UI to process the dismissal
    except:
        print("No contact info popup detected.")

# **Step 1: Initial Login**
login()

# **Step 2: Navigate to Each Account Page & Download PDF**
account_numbers = [
    "1051232522-5", "1947775177-0", "7656100844-2", "1425713638-3", "8176683824-7",
    "3937199847-5", "8478299735-1", "9655656259-1", "7447349142-6", "9010017104-1",
    "0270611640-3", "8947324014-4", "7947324078-0", "1739465846-6", "6447538128-7",
    "1906082438-0", "4783150844-6", "3228944784-2", "2697323371-6", "2531108473-2",
    "1467300186-4", "2822775121-2", "1417723039-9", "0937554826-1", "0406127383-7",
    "2176684208-8", "9739460119-5", "2020219221-5", "1852724251-7"
]

for account_no in account_numbers:
    account_url = f"https://m.pge.com/#myaccount/dashboard/summary/{account_no}"
    print(f"Navigating to: {account_url}")
    driver.get(account_url)

    time.sleep(3)  # Allow time for the page to load

    try:
        # **Step 2.1: Check if We Got Logged Out**
        if is_logged_out():
            print(f"Logged out detected on account: {account_no}. Re-logging in...")
            login()  # Re-login
            driver.get(account_url)  # Retry navigating to the account page
            time.sleep(3)  # Allow page to load

        # **Step 2.2: Handle Contact Info Popup if It Appears**
        handle_popup()

        # **Step 2.3: Click "View Current Bill (PDF)"**
        time.sleep(3)  # Mimic human behavior

        pdf_button = wait.until(EC.element_to_be_clickable((By.ID, "utag-view-pay-view-bill-pdf")))
        pdf_button.click()
        print(f"Clicked 'View Current Bill (PDF)' for account: {account_no}")

        # **Step 2.4: Wait for the Download to Complete**
        time.sleep(10)  # Adjust as needed

    except Exception as e:
        print(f"Error processing account {account_no}: {e}")

# Close the browser
driver.quit()
print(f"Completed PDF downloads!")
