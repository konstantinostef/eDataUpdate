import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import sys

def read_settings(settings_file_path):
    """
    Reads settings from a text file
    
    Args:
        settings_file_path: Path to the settings file
    """
    # Read the Excel file
    
    df = pd.read_excel(settings_file_path)
    #print(df)
    username = df.iloc[0, 1].strip()
    password = df.iloc[1, 1].strip()
    #print(f"Username: {username}, Password: {password}")
    date = df.iloc[2, 1].strip()
    #print(f"Date: {date}")
    #sys.exit()
    return username, password, date

def read_am_codes(file_path):
    """
    Reads 6-digit AM codes from an Excel file
    
    Args:
        file_path: Path to the .xlsx file
    
    Returns:
        List of AM codes
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Assuming AM codes are in the first column
        # Adjust column name as needed
        am_codes = df.iloc[:, 0].tolist()
        
        # Convert to string and ensure 6 digits (pad with zeros if needed)
        am_codes = [str(code).zfill(6) for code in am_codes]
        
        print(f"Successfully read {len(am_codes)} AM codes")
        return am_codes
    
    except Exception as e:
        print(f"Error reading file: {e}")
  
def connectAndPressServices(driver, username, password):
    try:
        #try to connect to the website
        driver.get("https://edcportal.minedu.gov.gr/")
        
        # 2. Find input field and enter AM code
        #input_field = driver.find_element(By.ID, "am_code_input")
        # Wait for page to load and find the Σύνδεση button
        wait = WebDriverWait(driver, 10)
        
        # The button is inside an <a> tag with specific href
        syndesi_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='https://edcapps.minedu.gov.gr/edc']"))
        )
        syndesi_link.click()
        time.sleep(2)
        
        # Switch to the new tab (the SSO login page)
        driver.switch_to.window(driver.window_handles[-1])

        # Now wait for the login form to load
        wait = WebDriverWait(driver, 10)
        wait.until(EC.presence_of_element_located((By.ID, "username")))


        username_field = wait.until(
            EC.presence_of_element_located((By.ID, "username"))
        )
        username_field.clear()
        username_field.send_keys(username)  # Replace with actual username

        # Enter password
        password_field = driver.find_element(By.ID, "password")
        password_field.clear()
        password_field.send_keys(password) 
        # Click the login button
        login_button = driver.find_element(By.ID, "loginButton")
        login_button.click()

        # Wait for login to complete (adjust time as needed)
        time.sleep(3)

    except Exception as e:
        print(f"Error connecting to the website: {e}")
        driver.quit()
    return
    

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
    
    connectAndPressServices(driver, username, password)
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