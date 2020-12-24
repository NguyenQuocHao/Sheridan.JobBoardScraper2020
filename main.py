# Original work Copyright (c) 2020 [Hao Nguyen]

import openpyxl
from bs4 import BeautifulSoup
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time

# Getting things ready!!

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

# Check if excel file exists, if not create one
FileExistence = Path('sheridan_job_board_fall_2020.xlsx')
if FileExistence.is_file():
    wb = openpyxl.load_workbook('sheridan_job_board_fall_2020.xlsx')
else:
    wb = openpyxl.Workbook()  # Create a blank workbook.
sheet = wb.active

# Set headers for table
sheet.cell(row=1, column=1).value = "Job Number"
sheet.cell(row=1, column=2).value = "Employer"
sheet.cell(row=1, column=3).value = "Job Title"
sheet.cell(row=1, column=4).value = "Status"
sheet.cell(row=1, column=5).value = "Job Type"
sheet.cell(row=1, column=6).value = "Job Sub-Type"
sheet.cell(row=1, column=7).value = "City"
sheet.cell(row=1, column=8).value = "Positions"
sheet.cell(row=1, column=9).value = "Months"
sheet.cell(row=1, column=10).value = "Posting Date"
sheet.cell(row=1, column=11).value = "Closing Date"
sheet.cell(row=1, column=12).value = "Salary"
sheet.cell(row=1, column=13).value = "Compensation"
sheet.cell(row=1, column=14).value = "Description"

# Get job board table content
soup = BeautifulSoup(browser.page_source, 'lxml')  # load page content
My_table = soup.find('table', {'class': 'table table-bordered'})
findTd = My_table.findAll('td')  # scan through table, and store all 'td' elems into an array

count = 0  # Number of moves (cell to cell)
countX = 1  # Column
countY = 2  # Row

def collect_job_postings():
    for i in findTd:
        global count
        global countX
        global countY
        # copy salary & compensation
        if count % 11 == 0:  # execute if first cell
             sheet.cell(row=countY, column=countX).value = job_id = i.text.replace('\n', '').strip()
             locator = browser.find_element_by_link_text(job_id)
             locator.click()
             soup = BeautifulSoup(browser.page_source, 'lxml')  # update buffered link to current link
             My_Salary = soup.find('div', {'id': 'jd9'})
             My_Compensation = soup.find('div', {'id': 'jd10'})
             My_Description = soup.find('div', {'id': 'jobd1'})
             if My_Salary is not None:
                 findSalary = My_Salary.findAll('p')[0].text.replace('\n', '').strip()
                 sheet.cell(row=countY, column=12).value = findSalary
             if My_Compensation is not None:
                 findCompensation = My_Compensation.findAll('p')[0].text.replace('\n', '').strip()
                 sheet.cell(row=countY, column=13).value = findCompensation
             if My_Description is not None:
                 findDescription = My_Description.findAll('p')[0].text.replace('\n', '').strip()
                 sheet.cell(row=countY, column=14).value = findDescription
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

# first page
collect_job_postings()
# for i in findTd:
#      # copy salary & compensation
#      if count % 11 == 0:  # execute if first cell
#          sheet.cell(row=countY, column=countX).value = job_id = i.text.replace('\n', '').strip()
#          locator = browser.find_element_by_link_text(job_id)
#          locator.click()
#          soup = BeautifulSoup(browser.page_source, 'lxml')  # update buffered link to current link
#          My_Salary = soup.find('div', {'id': 'jd9'})
#          My_Compensation = soup.find('div', {'id': 'jd10'})
#          My_Description = soup.find('div', {'id': 'jobd1'})
#          if My_Salary is not None:
#              findSalary = My_Salary.findAll('p')[0].text.replace('\n', '').strip()
#              sheet.cell(row=countY, column=12).value = findSalary
#          if My_Compensation is not None:
#              findCompensation = My_Compensation.findAll('p')[0].text.replace('\n', '').strip()
#              sheet.cell(row=countY, column=13).value = findCompensation
#          if My_Description is not None:
#              findDescription = My_Description.findAll('p')[0].text.replace('\n', '').strip()
#              sheet.cell(row=countY, column=14).value = findDescription
#          browser.back()
#
#      elif count % 11 != 0:  # execute if not first cell
#          sheet.cell(row=countY, column=countX).value = i.text.replace('\n', '').strip()
#          # Check if the last cell
#          if count % 11 == 10:
#              countY += 1
#              countX = 0
#
#      # Increment anyways
#      countX += 1
#      count += 1

# later pages.
# for page in range(0, 17):
#     linkElem = browser.find_element_by_link_text('Next')
#     linkElem.click()
#     soup = BeautifulSoup(browser.page_source, 'lxml')  # update buffered link to current link
#     My_table = soup.find('table', {'class': 'table table-bordered'})
#     findTd = My_table.findAll('td')
#     for i in findTd:
#         # copy table data into excel file
#         if count % 11 != 0:
#             sheet.cell(row=countY, column=countX).value = i.text.replace('\n', '').strip()
#             countX += 1
#         else:
#             countY += 1
#             countX = 1
#             sheet.cell(row=countY, column=countX).value = i.text.replace('\n', '').strip()
#             countX += 1
#
#         # copy salary & compensation
#         if count % 11 == 0:
#             job_id = i.text.replace('\n', '').strip()
#             print('                ' + job_id)
#             locator = browser.find_element_by_link_text(job_id)
#             locator.click()
#             soup = BeautifulSoup(browser.page_source, 'lxml')
#             My_Salary = soup.find('div', {'id': 'jd9'})
#             My_Compensation = soup.find('div', {'id': 'jd10'})
#             My_Description = soup.find('div', {'id': 'jobd1'})
#             if My_Salary is not None:
#                 findSalary = My_Salary.findAll('p')[0].text.replace('\n', '').strip()
#                 print("sal :" + findSalary)
#                 sheet.cell(row=countY, column=12).value = findSalary
#             if My_Compensation is not None:
#                 findCompensation = My_Compensation.findAll('p')[0].text.replace('\n', '').strip()
#                 print("pen :" + findCompensation)
#                 sheet.cell(row=countY, column=13).value = findCompensation
#             # Get Description
#             if My_Description is not None:
#                 findDescription = My_Description.findAll('p')[0].text  # .replace('\n', '').strip()
#                 print("des :" + findDescription)
#                 sheet.cell(row=countY, column=14).value = findDescription
#             browser.back()
#         count += 1

# Write (save) the excel file into the drive
wb.save('sheridan_job_board_fall_2020.xlsx')


