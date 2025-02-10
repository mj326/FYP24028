import os
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

def run_scraper(driver, download_dir, company_id):
    """
    Navigates to the Filings & Reports page, scrapes version IDs, downloads PDFs, and updates the CSV file.
    """
    wait = WebDriverWait(driver, 20)

    try:
        # Step 1: Navigate to Filings & Reports
        filings_url = f"https://www.capitaliq.spglobal.com/web/client?auth=inherit#company/documents?id={company_id}"
        driver.get(filings_url)
        print("Navigated to Filings & Reports.")
        time.sleep(3)  # Allow page to load

        # Step 2: Extract document links and version IDs
        # Initialize a list to store all report links and version IDs
        all_report_links = []
        all_version_ids = []

        while True:  # Loop through all pages
            print("Extracting report links from the current page...")
            try:
                # Wait for report links to load
                report_links = WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'docviewer?mid=')]"))
                )

                # Extract links and version IDs
                for link in report_links:
                    href = link.get_attribute("href")
                    if "mid=" in href:
                        version_id = href.split("mid=")[-1].split("&")[0]
                        all_version_ids.append(version_id)
                        all_report_links.append(link.text)
                    else:
                        print("No version ID found for this link.")
            except Exception as e:
                print(f"Error extracting links: {e}")
                break

            try:
                next_button_span = driver.find_element(By.XPATH, "//div[contains(@id, '_grid_table_next_page')]/span")
                next_button_classes = next_button_span.get_attribute("class")

                # Check if the "Next" span class includes "ui-state-disabled"
                if "ui-state-disabled" in next_button_classes:
                    break

                # Otherwise, wait for the "Next" button to be clickable and click
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//div[contains(@id, '_grid_table_next_page') and not(contains(@class, 'ui-state-disabled'))]"))
                ).click()

                # Wait for the next page's data to load
                WebDriverWait(driver, 10).until(
                    EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@href, 'docviewer?mid=')]"))
                )
                
            except Exception as e:
                print(f"Error navigating to the next page: {e}")
                break

        # Print the total number of version IDs extracted
        print(f"Total Version IDs Extracted: {len(all_version_ids)}")

        # Step 3: Download Excel file
        latest_file = None

        try:
            # Scroll to and click on dropdown toggle
            dropdown_toggle = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='dropdown-toggle' and @data-id='3233133301']")))
            driver.execute_script("arguments[0].scrollIntoView(true);", dropdown_toggle)  # Scroll to ensure visibility
            time.sleep(1)  # Allow UI to stabilize
            driver.execute_script("arguments[0].click();", dropdown_toggle)  

            # Wait for the dropdown menu to render and click the export button
            export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[@class='hui-toolbutton' and @data-id='32']")))
            driver.execute_script("arguments[0].click();", export_button)

            # Wait for the Excel file to download
            latest_file = wait_for_latest_excel(download_dir)

        except Exception as e:
            print(f"An error occurred during the export process: {e}")

        # Step 4: Convert Excel to CSV and Manipulate Data
        csv_file_path = os.path.join(download_dir, f"{company_id}.csv")  # Use company_id in the filename

        # Read the Excel file
        excel_data = pd.read_excel(latest_file, skiprows=14)  # Adjust skiprows if needed

        # Drop the 'Abstract' column
        if "Abstract" in excel_data.columns:  # Check if the column exists
            excel_data = excel_data.drop(columns=["Abstract"])
        else:
            print("'Abstract' column not found. Skipping removal.")

        # Filter rows containing "Annual Report" or relevant report type
        relevant_rows = excel_data[excel_data.apply(lambda row: row.astype(str).str.contains("Annual Report", case=False).any(), axis=1)].copy()

        # Check if the number of relevant rows matches the number of Version_IDs
        if len(relevant_rows) != len(all_version_ids):
            raise ValueError(
                f"Mismatch detected: Number of relevant rows ({len(relevant_rows)}) does not match the number of Version_IDs ({len(all_version_ids)})."
            )
        
        relevant_rows["Filing Date"] = pd.to_datetime(relevant_rows["Filing Date"]).dt.strftime('%d/%m/%Y')
        relevant_rows["Filing Date"] = relevant_rows["Filing Date"].astype(str)

        relevant_rows["Event Date"] = pd.to_datetime(relevant_rows["Event Date"]).dt.strftime('%d/%m/%Y')
        relevant_rows["Event Date"] = relevant_rows["Event Date"].astype(str)

        # Map Version_IDs to relevant rows
        relevant_rows.loc[:, "Version_ID"] = all_version_ids

        # Add the 'Path' column with placeholder values
        relevant_rows.loc[:, "Path"] = " "  # Placeholder for S3 paths
        # Add the 'Country' column 
        relevant_rows.loc[:, "Country"] = "Malaysia" 

        # Save the cleaned and processed data
        relevant_rows.to_csv(csv_file_path, index=False)
        print(f"Processed file saved to: {csv_file_path}")


    except Exception as e:
        print(f"An error occurred during scraping: {e}")

def wait_for_latest_excel(download_dir, timeout=60):
    start_time = time.time()
    latest_file = None

    while time.time() - start_time < timeout:
        excel_files = [f for f in os.listdir(download_dir) if f.endswith(('.xlsx', '.xls'))]
        if excel_files:
            latest_file = max([os.path.join(download_dir, f) for f in excel_files], key=os.path.getmtime)
            if os.path.exists(latest_file):
                return latest_file
        time.sleep(1)

    raise Exception("File download timed out. No Excel file found.")

def wait_for_latest_excel(download_dir, timeout=60):
    start_time = time.time()
    latest_file = None

    while time.time() - start_time < timeout:
        excel_files = [f for f in os.listdir(download_dir) if f.endswith(('.xlsx', '.xls'))]
        if excel_files:
            latest_file = max([os.path.join(download_dir, f) for f in excel_files], key=os.path.getmtime)
            if os.path.exists(latest_file):
                return latest_file
        time.sleep(1)

    raise Exception("File download timed out. No Excel file found.")
