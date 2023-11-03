from RPA.Browser.Selenium import Selenium
from time import sleep
from RPA.HTTP import HTTP
import openpyxl

# Initialize Selenium and HTTP objects
browser_lib = Selenium()
http = HTTP()

url = "https://rpachallenge.com/"
def open_the_website(url):
    browser_lib.open_available_browser(url)

def download_excel():
    excel_link = browser_lib.get_element_attribute('//a[text()=" Download Excel "]','href')
    http.download(url = excel_link)     

def input_data_from_spreadsheet():
    # Load the Excel file with the data
    workbook = openpyxl.load_workbook("challenge.xlsx")  # Change "data.xlsx" to the path of your Excel file
    sheet = workbook.active

    # Loop through rows and input data into form fields
    for row in sheet.iter_rows(min_row=2, values_only=True):
        cleaned_row = [cell for cell in row if cell is not None]

        print(cleaned_row)
        if len(cleaned_row)==7:
            first_name, last_name, company_name, role_in_the_company,address,email,phone_number = cleaned_row
            print(first_name, last_name,company_name,role_in_the_company,address,email,phone_number)
            browser_lib.input_text('//label[text()="Address"]//..//input', address)
            browser_lib.input_text('//label[text()="Email"]//..//input', email)
            browser_lib.input_text('//label[text()="Phone Number"]//..//input', phone_number)
            browser_lib.input_text('//label[text()="Role in Company"]//..//input',role_in_the_company)
            browser_lib.input_text('//label[text()="Company Name"]//..//input',company_name)
            browser_lib.input_text('//label[text()="Last Name"]//..//input',last_name)
            browser_lib.input_text('//label[text()="First Name"]//..//input',first_name)
            browser_lib.click_element('//input[@type="submit"]')
        else:
            break

open_the_website(url)     
download_excel() 
sleep(15)
input_data_from_spreadsheet()
