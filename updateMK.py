import pandas as pd
import logging
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from utils import read_settings, read_mk_am_codes, loginToEdata

import time
import sys

logger = logging.getLogger(__name__)
def mk_to_value(mk):
    """
    Maps MK code to the corresponding dropdown value.
    This function needs to be implemented based on the actual mapping.
    
    Args:
        mk: MK code to map
    
    Returns:
        Corresponding value for the dropdown
    """
    mk_mapping = {
        1: "30",
        2: "31",
        3: "33",
        4: "35",
        5: "36",
        6: "37",
        7: "39",
        8: "40",
        9: "42",
        10: "45",
        11: "46",
        12: "47",
        13: "48",
        14: "49",
        15: "51",
        16: "52",
        17: "53",
        18: "54",
        19: "55",
        
    }
    return mk_mapping.get(mk, "")

def process_am_code(driver, am_code, mk):
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
        basic_elements_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='worker.fin.list.aspx?type=worker']"))
        )
        basic_elements_link.click()

        # Wait for the page to load
        time.sleep(2)        
    
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

        # Wait for the results table to load
        time.sleep(2)

        mk_value_for_dropdown =  mk_to_value(mk)  # Implement this function to map am_code to dropdown value
        print(f"MK: {mk}, MK value for dropdown: {mk_value_for_dropdown}")
        wait = WebDriverWait(driver, 10)
        dropdown_element = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ContentMain_dxtabWorker_ddlSalaryCategory")))

        dropdown = Select(dropdown_element)
        dropdown.select_by_value(mk_value_for_dropdown)

        # Verify the selection was made
        selected_option = dropdown.first_selected_option
        selected_mk_value = selected_option.text.split(" ")[1]

        print(f"Selected: {selected_mk_value} (Value: {selected_option.text})")
        # Verify the selection was made
        selected_option = dropdown.first_selected_option
        print(f"Selected: {selected_option.text} (Value: {selected_option.get_attribute('value')})")
        # Wait for the page to load
        time.sleep(2)        

        # Save the changes by clicking the save button
        button = driver.find_element(By.ID, "ctl00_ContentMain_lnkbtnSave")
        
        if(selected_mk_value != str(mk)):
            logger.error('AM %s: MK value not updated correctly. Expected %s but got %s', am_code, mk, selected_mk_value)
        else:
            logger.info('AM %s: MK value updated correctly. Expected %s and got %s', am_code, mk, selected_mk_value)
            # button.click()
        time.sleep(4)  # Small delay between requests
        print(f"Completed processing: {am_code}")
        
    except Exception as e:
        print(f"Error processing {am_code}: {e}")

def main():
    """Main execution function"""
    logging.basicConfig(
        format='%(asctime)s %(levelname)-8s %(message)s',
        filename='eDataMKUpdate.log', level=logging.INFO,
        datefmt='%Y-%m-%d %H:%M:%S'
        )
    logger.info('Started')
    # Configuration
    excel_file = "mk_am_codes.xlsx"  # Change to your file path
    settings_file = "settings.xlsx"  # Change to your settings file path
    username, password, date = read_settings(settings_file)
    
    # Initialize the web driver
    # For Chrome (make sure you have Chrome installed)
    driver = webdriver.Chrome()
    
    # Alternatively for Firefox:
    # driver = webdriver.Firefox()
    
    loginToEdata(driver, username, password)

    try:
        # Process each AM code
        for am_code, mk in read_mk_am_codes(excel_file):
            process_am_code(driver, am_code, mk)
        
        print(f"\nSuccessfully processed all AM codes!")
        
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        # Close the browser
        driver.quit()

if __name__ == "__main__":
    main()