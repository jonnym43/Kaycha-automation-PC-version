from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from datetime import datetime, timedelta
import os
from config import confidential_info
from config.confidential_info import operating_file_path_instance
from config.confidential_info import website_instance

start_time = time.time()


def get_date_position(date):
    # Find the first day of the month and its weekday (0=Monday, 6=Sunday)
    first_day_of_month = date.replace(day=1)
    first_day_weekday = first_day_of_month.weekday()  # Monday is 0 and Sunday is 6

    # Adjust the weekday to start from Sunday
    first_day_weekday = (first_day_weekday + 1) % 7

    # Calculate the position in the date picker grid
    row = (date.day + first_day_weekday - 1) // 7
    col = (date.day + first_day_weekday - 1) % 7

    return row, col


def select_date(driver, date, date_field_id, is_end_date=False):
    try:
        # Get the position of the date in the date picker
        row, col = get_date_position(date)
        print(f"Date position for {date}: row {row}, col {col}")

        # Construct the XPath for the date and click it
        date_xpath = f"//td[contains(@class, 'available') and @data-title='r{row}c{col}']"
        date_to_select = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, date_xpath))
        )
        date_to_select.click()

        # Verify the click by checking the class change
        expected_class = "active end-date in-range available" if is_end_date else "active start-date available"
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.XPATH, f"//td[contains(@class, '{expected_class}') and @data-title='r{row}c{col}']"))
        )
        # time.sleep(2)
        print(f"Clicked date: {date}")

    except Exception as e:
        print(f"Error in select_date: {e}")


driver = webdriver.Firefox()
driver.get(website_instance.kaycha_website)

try:
    login_username = WebDriverWait(driver, 20).until(
        EC.presence_of_element_located((By.ID, "loginform-username"))
    )
    login_username.send_keys(confidential_info.username_instance.username)
    login_username.send_keys(Keys.RETURN)

    # Add a wait of 2 seconds here
    time.sleep(2)

    for attempt in range(2):  # Try to enter password up to 2 times
        try:
            login_password_box = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "loginform-password"))
            )
            login_password_box.clear()  # Clear any existing text in the password field
            login_password_box.send_keys(confidential_info.password_instance.password)

            login_button = WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable((By.ID, "login"))
            )
            login_button.click()

            # Wait for the company_id select element to pop up
            company_id_select = WebDriverWait(driver, 20).until(
                EC.presence_of_element_located((By.ID, "loginform-company_id"))
            )

            # Perform another login button click
            login_button.click()  # Use the same login_button variable
            
            break  # If no exception was raised, break out of the loop

        except Exception as e:
            print(f"Attempt {attempt + 1} - Error entering password: {e}")
            if attempt == 1:  # If this was the second attempt
                raise  # Re-raise the last exception
            
except Exception as e:
    print(f"Error in login process: {e}")
    
time.sleep(2)

nav_bar_coa_history = WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.CLASS_NAME, 'coa-history-refresh'))
    )
nav_bar_coa_history.click()

date_field = WebDriverWait(driver, 20).until(
    EC.element_to_be_clickable((By.ID, 'w0'))
)
date_field.click()
# print("Opened calendar widget") #debug statement, not needed to run script

# Define the start and end dates
today = datetime.now().date()
endtime = today - timedelta(days=0)  # More recent date
starttime = today - timedelta(days=3)  # Older date

# Select the start date
select_date(driver, starttime, 'w0', is_end_date=False)

# Select the end date
select_date(driver, endtime, 'w0', is_end_date=True)

# Wait and then click the apply button
try:
        apply_button = WebDriverWait(driver, 5).until(
            # Ensure this is the correct class for the apply button
            EC.element_to_be_clickable(
                (By.CLASS_NAME, "applyBtn btn btn-sm btn-primary"))
        )
        apply_button.click()
        # print("Clicked Apply button") #debug statement, not needed to run script
except Exception as e:
    print(f"Error clicking Apply button: {e}")

# Wait and then click the search button
    try:
        find_search_button = WebDriverWait(driver, 5).until(
            # Ensure this is the correct class for the search button
            EC.presence_of_element_located((By.ID, 'team-form-submit'))
        )
        find_search_button.click()
        time.sleep(1)
        # print("Clicked Search button")
    except Exception as e:
        print(f"Error clicking Search button: {e}")

# Wait and then click the select all button
    try:
        find_select_all_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'select_all'))
        )
        find_select_all_button.click()
    except Exception as e:
        print(f"Error clicking Select All button: {e}")

# Wait and then click the export button
    try:
        find_export_button = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'export-new-coa-btn'))
                                                             )
        find_export_button.click()
    except Exception as e:
        print(f"Error: {e}")

# Wait and then click the export button on pop-up
    try:
        find_export_button_popup = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, 'export-format-save'))
        )
        find_export_button_popup.click()
    except Exception as e:
        print(f"Error: {e}")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        time.sleep(6)
        driver.close()


    elapsed_time = round(time.time() - start_time, 1)
    print(f"Time elapsed until the end of main.py: {elapsed_time} seconds")

    os.system(f'python "{operating_file_path_instance.work_with_export}"')


    print("Chained script execution completed.")