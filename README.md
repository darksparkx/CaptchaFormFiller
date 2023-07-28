# CaptchaFormFiller
Fill a form with captcha in it, for webscrapping purposes

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
