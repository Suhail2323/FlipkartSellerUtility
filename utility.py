
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Load GST numbers from an Excel file
def load_gst_numbers_from_excel(excel_file, sheet_name, column_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    return df[column_name]

# Verify GST numbers on Flipkart Seller platform
def verify_gst_numbers(gst_numbers, excel_file, sheet_name, result_column_name):
    driver = webdriver.Chrome()

    # Open the Flipkart Seller platform login page
    driver.get("https://seller.flipkart.com/index.html#signUp/accountCreation/new")

    # You may need to perform any login/authentication steps if required

    results = []

    # Iterate through GST numbers
    for gst_number in gst_numbers:
        if not pd.isna(gst_number):
            try:
                # Locate the input field for GST number and enter the GST number
                input_element = driver.find_element(By.NAME, "gst")
                input_element.clear()
                input_element.send_keys(gst_number)
                input_element.send_keys(Keys.RETURN)

                # Wait for the verification result element to appear on the page
                result_element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                    (By.NAME, 'gst')))

                # Check the verification result to determine if GST exists or not
                verification_status = result_element.text
                if "does not exist" in verification_status.lower():
                    print(f"GST Number: {gst_number} does not exist.")
                    results.append("Does Not Exist")
                else:
                    print(f"GST Number: {gst_number} exists. Verification Status: {verification_status}")
                    results.append("Exists")
            except:
                print(f"GST Number: {gst_number}, Verification Failed")
                results.append("Verification Failed")

            # Add a short delay between verifications to avoid rate limiting or bans
            time.sleep(5)
        else:
            print("Skipping invalid or missing GST Number")
            results.append("Invalid or Missing GST Number")

    driver.quit()

    # Add the verification results to the existing Excel file
    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df[result_column_name] = results
    df.to_excel(excel_file, sheet_name=sheet_name, index=False)

if __name__ == "__main__":
    excel_file = r"C:\Users\91981\Downloads\excelrecord\Book1.xlsx"  # Replace with the path to your Excel file
    sheet_name = "Sheet1"  # Replace with the sheet name
    column_name = "GST Number"  # Replace with the column name containing GST numbers
    result_column_name = "Result"  # Choose a column name for the verification results

    gst_numbers = load_gst_numbers_from_excel(excel_file, sheet_name, column_name)
    verify_gst_numbers(gst_numbers, excel_file, sheet_name, result_column_name)

