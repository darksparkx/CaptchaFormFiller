"""
This is an automated form filling script with basic captcha solving capabilities

you will need to install tesseract to make this work.

Download this script and edit it as shown in the steps below.

Steps:
1. add the url of the website where the form is located
2. add a file called "data.xlsx" with your form data in the same folder as this script
3. Change the coordinates on line 155 so that only the captcha is visible. 
    you can check this using the screenshot.png that is saved in the folder where this script is located.
4. Once the screenshot contains only the captcha, check the 
    extracted_numbers (as the captcha in my case only had numbers) make sure the captcha is correct. 
    you can edit the splicing of the return value on line 54 to ensure the correct values are returned.
5. if the captcha is returning correct most of the time, you can now use selenium to target different part of the form
    and fill it .An example has been provided in the send_form_data() function.
6. After sending the form data, check for errors and return the error if there is any. 
    an example has been provided in the check_errors() function.
7. If there are no errors, access the returned values, put them in an object and save it into the results list.
    This will be saved at the end into an Excel spreadsheet called results.xlsx.

you can use this library to make try/except blocks and catch errors:
selenium.common.exceptions

"""

import traceback
from selenium import webdriver
from PIL import Image
import pytesseract
import pandas as pd


# Functions
def crop_image(image_path, crop_coordinates):
    # Open the image using Pillow
    image = Image.open(image_path)

    # Crop the image using the defined coordinates
    cropped_image = image.crop(crop_coordinates)

    # Save the cropped image
    cropped_image.save(image_path)


def extract_numbers(image_path):
    # Open the image using Pillow
    image = Image.open(image_path)

    # Perform OCR using Tesseract
    extracted_text = pytesseract.image_to_string(image)

    return extracted_text[-8:-1]  # slicing string to get only the numbers


def send_form_data(form_data):
    """
    Here you can send the form data to populate each field.
    select the element by:
    select_element = driver.find_element(By.ID, "ELEMENT_ID")
    select_element.send_keys(form_data["KEY"])
    
    For example:
    
    # Input Captcha
    captcha = driver.find_element(By.ID, "captcha_txtCaptcha")
    captcha.send_keys(form_data["Captcha"])
    
    
    After filing the fields, use:
    button = driver.find_element(By.ID, "SUBMIT_BUTTON_ID")
    button.click()
    
    
    """
    

def check_errors():
    """
    Here you can check any errors that might've occured, like a wrong captcha.
    
    If there are no errors, return false.
    
    For example:
    
    verification = driver.find_element(By.ID,"CAPTCHA_VERIFICATION_ERROR_ID")  # this ID can usually be found under or around the captcha box 
    
    if verification and verification.text:
        return 'captcha_error'
    else:
        return False

    """
    


def get_results(results_list):
    """
    Here you can finally get the results and add it to the 'return' list which will be saved later.
    
    This part can vary from website to website depending on what is returned.
    
    For example: ( in this example, a table is returned on the same page)
    
    result_table = driver.find_element(By.ID, "RESULT_TABLE_ID")

    # Get all rows
    rows = result_table.find_elements(By.TAG_NAME, "tr")

    # Extract the table data
    cells = rows[1].find_elements(By.TAG_NAME, "td")  # rows[0] are the headers
    row_data = { POPULATE ROW DATA}

    # Append to results list
    results_list.append(row_data)

    """
    return results_list


if __name__ == "__main__":

    # variables
    website_url = "ENTER URL HERE"
    ss_path = "screenshot.png"
    file_path = "data.xlsx"
    result_path = "results.xlsx"

    # Preparing the browser
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--disable-cookies")
    chrome_options.add_argument("start-maximized")
    driver = webdriver.Chrome(options=chrome_options)

    # Get xls data as a list
    df = pd.read_excel(file_path)
    data = df.to_dict(orient='records')

    # create a results list
    results = []

    while data:
        try:
            # get single object
            curr_data = data[0]

            # Open url in a browser
            driver.get(website_url)

            # Take a screenshot
            driver.save_screenshot(ss_path)

            # Crop the image to get only the captcha
            coordinates = (650, 600, 900, 700)  # coordinates = left, top, right, bottom
            crop_image(ss_path, coordinates)

            # Extract captcha numbers from the image
            extracted_numbers = extract_numbers(ss_path)

            # Concatenate data with verification code
            curr_data["Captcha"] = extracted_numbers[:-1]  # splicing to remove the '/n'

            # Submit form
            send_form_data(curr_data)

            # Check for errors
            error = check_errors()

            # if there are no errors, get the results
            if not error:
                results = get_results(results)

            # remove obj from data if captcha was correct IE, there is an error in the data.
            # this part can be edited as you see fit
            if error != "captcha_error":
                data.pop(0)
            
        except Exception as e:
            print("An error occurred:", e)
            traceback.print_exc()

    # Change results into a pandas dataframe and write it
    results = pd.DataFrame(results)
    
    with pd.ExcelWriter(result_path, engine='xlsxwriter') as writer:
        results.to_excel(writer, index=False, sheet_name='Sheet1')
