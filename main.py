# Original work Copyright (c) 2020 [Hao Nguyen]

import openpyxl
from bs4 import BeautifulSoup
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Getting things ready!!

# Check if excel file exists, if not create one
file_name = 'sheridan_job_board_fall_2020_v1.xlsx'
FileExistence = Path(file_name)
if FileExistence.is_file():
    wb = openpyxl.load_workbook(file_name)
else:
    wb = openpyxl.Workbook()  # Create a blank workbook.
sheet = wb.active

# from webdriver_manager.chrome import ChromeDriverManager
# driver = webdriver.Chrome(ChromeDriverManager().install())
browser = webdriver.Chrome()
browser.get('https://jobs.sheridancollege.ca/Student/Default.asp#Content')
time.sleep(1)  # Sleep to wait for the browser to load the content

# Fill in login info.
# Account
userElem = browser.find_element_by_id('user')
userElem.send_keys("")  # Put your studentID in ""
# Password
passwordElem = browser.find_element_by_id('password')
passwordElem.send_keys("")  # Put your password in ""
htmlElem = browser.find_element_by_tag_name('html')
passwordElem.send_keys(Keys.ENTER)

# Navigate to Job board tab
linkElem = browser.find_element_by_link_text('Job Postings')
linkElem.click()

# Set headers for table
header_list = [
    "Job Number",
    "Employer",
    "Job Title",
    "Status",
    "Job Type",
    "Job Sub-Type",
    "City",
    "Positions",
    "Months",
    "Posting Date",
    "Closing Date",
    "Salary",
    "Compensation",
    "Description"
]

count = 0  # Number of moves (cell to cell)
countX = 1  # Column
countY = 2  # Row


def create_headers(sheet, headers):
    column = 1
    for header in headers:
        sheet.cell(row=1, column=column).value = header
        column += 1


# Get job board table content
def collect_job_postings():
    soup = BeautifulSoup(browser.page_source, 'lxml')  # load page content
    my_table = soup.find('table', {'class': 'table table-bordered'})
    find_td = my_table.find_all('td')  # scan through table, and store all 'td' elements into an array
    my_table.find_all_next()
    for i in find_td:
        global count
        global countX
        global countY
        # copy salary & compensation
        if count % 11 == 0:  # execute if first cell
            sheet.cell(row=countY, column=countX).value = job_id = i.text.replace('\n', '').strip()
            locator = browser.find_element_by_link_text(job_id)
            locator.click()
            soup = BeautifulSoup(browser.page_source, 'lxml')  # update buffered link to current link
            my_salary = soup.find('div', {'id': 'jd9'})
            my_compensation = soup.find('div', {'id': 'jd10'})
            my_description = soup.find('div', {'id': 'jobd1'})
            if my_salary is not None:
                find_salary = my_salary.find_all('p')[0].text.replace('\n', '').strip()
                sheet.cell(row=countY, column=12).value = find_salary
            if my_compensation is not None:
                find_compensation = my_compensation.find_all('p')[0].text.replace('\n', '').strip()
                sheet.cell(row=countY, column=13).value = find_compensation
            if my_description is not None:
                find_description = my_description.find_all('p')[0].text.replace('\n', '').strip()
                sheet.cell(row=countY, column=14).value = find_description
            browser.back()

        elif count % 11 != 0:  # execute if not first cell
            sheet.cell(row=countY, column=countX).value = i.text.replace('\n', '').strip()
            # Check if the last cell
            if count % 11 == 10:
                countY += 1
                countX = 0

        # Increment anyways
        countX += 1
        count += 1


# Create headers for the table
create_headers(sheet, header_list)

# first page
collect_job_postings()

# later pages.
for page in range(0, 11):
    # Go to the next page
    linkElem = browser.find_element_by_link_text('Next')
    linkElem.click()
    collect_job_postings()

# Write (save) the excel file into the drive
wb.save(file_name)
