from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from dotenv import dotenv_values
from scraper import run_scraper  # Import the scraper logic
import time

# Load environment variables
env_vars = dotenv_values(".env")

# Extract multiple users dynamically from .env
accounts = []
for key, value in env_vars.items():
    if key.startswith("USERNAME_ACCOUNT"):
        account_number = key[len("USERNAME_ACCOUNT"):]
        password_key = f"PASSWORD_ACCOUNT{account_number}"
        if password_key in env_vars:
            accounts.append({
                "username": value,
                "password": env_vars[password_key],
                "account_number": account_number
            })

# Function to initialize ChromeDriver
def setup_driver(download_dir):
    chrome_options = Options()
    
    # Preferences for automatic file downloads
    prefs = {
        "download.default_directory": download_dir,  # Set the download directory
        "download.prompt_for_download": False,       # Disable prompt
        "download.directory_upgrade": True,          # Allow directory upgrades
        "safebrowsing.enabled": True,                # Allow safe downloads
        "plugins.always_open_pdf_externally": True   # If PDFs are downloaded
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--start-maximized")
    #return webdriver.Chrome(options=chrome_options)

    print(f"Download directory set to: {download_dir}")
    driver = webdriver.Chrome(options=chrome_options)
    return driver

# Login function
def login(driver, username, password):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.keys import Keys

    wait = WebDriverWait(driver, 20)
    print(f"Logging in for user: {username}")

    driver.get("https://www.capitaliq.spglobal.com/web/login?ignoreIDMContext=1#/")

    # Enter username
    username_field = wait.until(EC.presence_of_element_located((By.ID, "input28")))
    username_field.send_keys(username)
    next_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@type='submit' and @value='Next']")))
    next_button.click()

    # Enter password
    password_field = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
    password_field.send_keys(password)
    password_field.send_keys(Keys.RETURN)

    print("Login successful!")

# Main function to loop through accounts
def main():
    DOWNLOAD_DIR = "/Users/minjun/Desktop/FYP"   # Path to the desired folder
    for account in accounts:
        print(f"\nProcessing account {account['account_number']}")
        driver = setup_driver(DOWNLOAD_DIR)
        try:
            # Login
            login(driver, account["username"], account["password"])
            # Call the scraper function
            run_scraper(driver, DOWNLOAD_DIR, "4156490")  # Pass company ID and download folder
        except Exception as e:
            print(f"Error for account {account['account_number']}: {e}")
        finally:
            driver.quit()
            time.sleep(2)  # Delay before next login

if __name__ == "__main__":
    main()