import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import read_settings, read_am_codes, loginToEdata


import time
import sys

def process_am_code(driver, am_code, date):
    """
    Processes a single AM code on the website
    
    Args:
        driver: Selenium WebDriver instance
        am_code: 6-digit AM code to process
    """
    try:
        print(f"Processing AM code: {am_code}")
        wait = WebDriverWait(driver, 10)
        # Navigate to the service search page# Wait for the link to be clickable
        ypiretiseis_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='worker.fin.list.aspx?type=service']"))
        )
        ypiretiseis_link.click()

        # Wait for the page to load
        time.sleep(2)        
        # Define wait for this function
        
        # Insert am into the input field
        am_input = wait.until(
            EC.presence_of_element_located((By.XPATH, "//span[@id='ctl00_ContentMain_dxpanelCriteria_lblRegistryNo']/../..//input[@type='text']"))
        )
        am_input.clear()
        am_input.send_keys(am_code)

        # Click the Αναζήτηση (Search) button
        search_button = driver.find_element(By.ID, "ctl00_ContentMain_dxpanelCriteria_btnSearch")
        search_button.click()

        # Wait for results to load
        time.sleep(2)
        
        
        # Wait for the search results and click the first row link
        result_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//tr[@id='ctl00_ContentMain_dxgridResults_DXDataRow0']//a"))
        )
        result_link.click()


        # Wait for the page to load
        time.sleep(2)
        
        # Wait for the results table to load
        time.sleep(2)

        # Find all data rows (excluding the summary row which has "ΣΥΝΟΛΟ ΥΠΗΡΕΤΗΣΕΩΝ")
        data_rows = driver.find_elements(By.XPATH, "//tr[contains(@id, 'ctl00_ContentMain_dxgridResults_DXDataRow')]//a[contains(@href, 'worker.singleservice.edit.aspx')]")

        # Click the last row's edit link
        if data_rows:
            last_row_link = data_rows[-1]  # Get the last element
            last_row_link.click()
            time.sleep(2)
            print("Clicked on the last service entry")
        else:
            print("No data rows found")
        
        # Find the date input field and change the date
        date_input = wait.until(
            EC.presence_of_element_located((By.ID, "ctl00_ContentMain_dxdateDutyStopDate_I"))
        )

        # Clear the existing date and enter the new one
        date_input.clear()
        date_input.send_keys(date)

        # Optional: trigger the onchange event
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", date_input)
        button = driver.find_element(By.ID, "ctl00_ContentMain_lnkbtnSave")
        button.click()
        time.sleep(8)  # Small delay between requests
        print(f"Completed processing: {am_code}")
        
    except Exception as e:
        print(f"Error processing {am_code}: {e}")

def main():
    """Main execution function"""
    
    # Configuration
    excel_file = "am_codes.xlsx"  # Change to your file path
    settings_file = "settings.xlsx"  # Change to your settings file path
    username, password, date = read_settings(settings_file)
    # Read AM codes from Excel
    am_codes = read_am_codes(excel_file)
    
    if not am_codes:
        print("No AM codes found. Exiting.")
        return
    
    # Initialize the web driver
    # For Chrome (make sure you have Chrome installed)
    driver = webdriver.Chrome()
    
    # Alternatively for Firefox:
    # driver = webdriver.Firefox()
    
    loginToEdata(driver, username, password)
    try:
        # Process each AM code
        for am_code in am_codes:
            process_am_code(driver, am_code, date)
        
        print(f"\nSuccessfully processed all {len(am_codes)} AM codes!")
        
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        # Close the browser
        driver.quit()

if __name__ == "__main__":
    main()