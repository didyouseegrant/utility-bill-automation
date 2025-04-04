import os
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains

# Define ChromeDriver path
chrome_driver_path = r"path/to/your/input/folder/chromedriver.exe"
service = Service(chrome_driver_path)

# Define login credentials (Replace these with your actual credentials)
USERNAME = "email"
PASSWORD = "password"

# Define download directory
download_directory = r"path/to/your/input/folder"


# Set Chrome options
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": download_directory,  # Set the download folder
    "download.prompt_for_download": False,  # Disable download prompt
    "plugins.always_open_pdf_externally": True  # Prevent PDFs from opening in browser
})

# Start Selenium WebDriver
driver = webdriver.Chrome(service=service, options=chrome_options)

### **Step 1: Open Login Page & Log In**
driver.get("https://mmwd.billonline.com/login.aspx")

# Allow page to load
time.sleep(3)

# Locate username and password fields and enter credentials
username_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbUsername")  # Adjust if necessary
password_field = driver.find_element(By.ID, "ContentPlaceHolder1_tbPassword")  # Adjust if necessary
login_button = driver.find_element(By.ID, "ContentPlaceHolder1_btnLogin")  # Adjust if necessary

username_field.send_keys(USERNAME)
password_field.send_keys(PASSWORD)
login_button.click()

# Wait for login to process
time.sleep(5)

### **Step 2: Navigate to Invoice History Page**
driver.get("https://mmwd.billonline.com/Accounts/InvoiceHistory.aspx")
time.sleep(3)  # Allow time for the page to load


# Locate all rows containing invoices (excluding header row)
invoice_rows = driver.find_elements(By.XPATH, "//table[@id='ContentPlaceHolder1_GridView1']/tbody/tr[not(@class='gridviewheader')]")


### **Step 3: Download PDFs for February 2025**
for row in invoice_rows:
    try:
        # Extract invoice date (4th column)
        invoice_date_element = row.find_element(By.XPATH, "./td[4]")
        invoice_date = invoice_date_element.text.strip()

        # Extract customer number (2nd column)
        customer_no_element = row.find_element(By.XPATH, "./td[2]")
        customer_no = customer_no_element.text.strip()

        print(f"Checking row - Invoice Date: {invoice_date}, Customer No: {customer_no}")

        # Check if the invoice date is in February 2025
        if invoice_date.startswith("2/") and invoice_date.endswith("/2025"):
            print(f"Match found! Processing invoice: {invoice_date}, Customer: {customer_no}")

	    # Store the handle of the original tab before opening the PDF
            original_window = driver.current_window_handle


            # Locate "View" link (1st column)
            view_link = row.find_element(By.XPATH, ".//a[contains(text(), 'View')]")
            ##### pdf_url = view_link.get_attribute("href")

            # Open the PDF in a new tab
            view_link.click()  

            # The browser automatically switches to the new tab

            # Allow time for PDF to fully load
            time.sleep(5)

 	    # Find all open tabs and switch to the newly opened one (last in list)
            new_tab = [handle for handle in driver.window_handles if handle != original_window][0]
            driver.switch_to.window(new_tab)

            # Trigger "Save As" dialog using Ctrl+S
            ActionChains(driver).key_down(Keys.CONTROL).send_keys("s").key_up(Keys.CONTROL).perform()
            time.sleep(5)

            # Simulate typing the new file path and hitting 'Enter'
            new_file_path = os.path.join(download_directory, f"{customer_no}.pdf")
            ActionChains(driver).send_keys(new_file_path).send_keys(Keys.ENTER).perform()

            # Allow time for download to complete
            time.sleep(5)

            # Close the new tab
            driver.close()


    except Exception as e:
        print(f"Error processing invoice row: {e}")

# Close the browser
driver.quit()
print(f"Invoice download completed!")
