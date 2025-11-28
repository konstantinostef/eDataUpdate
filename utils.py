import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import logging
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
    date = "1" #df.iloc[2, 1].strip()
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

def read_mk_am_codes(file_path):
    """
    Reads 6-digit AM codes and MK values from an Excel file
    
    Args:
        file_path: Path to the .xlsx file
    
    Returns:
        List of AM and MK codes
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Assuming AM codes are in the first column
        # Adjust column name as needed
        am_codes = df.iloc[:, 0].tolist()
        mk_values = df.iloc[:, 1].tolist()
        # Convert to string and ensure 6 digits (pad with zeros if needed)
        # am_codes = [str(code).zfill(6) for code in am_codes]
        
        # print(f"Successfully read {len(am_codes)} AM codes")
        return zip(am_codes, mk_values)
    
    
    except Exception as e:
        print(f"Error reading file: {e}")

    
def loginToEdata(driver, username, password):
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
        username_field.send_keys(username)

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
        logging.error(e)
        print(f"Error connecting to the website: {e}")
        driver.quit()
    return
    