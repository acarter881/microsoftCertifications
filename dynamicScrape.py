from helium import *
from bs4 import BeautifulSoup
import pandas as pd
import time

# This is the URL that the script goes to first
url = 'https://docs.microsoft.com/en-us/learn/certifications/browse'

# Start Chrome
browser = start_chrome(url, headless=True)

# Create an empty list for the names of the exams and certifications
rows = []

# Set the number to increment the results by
number = 0

# Sleep to give time for the browser to load
time.sleep(2)

# Main portion of the script
while True:
    # Create an instance of the BeautifulSoup class
    soup = BeautifulSoup(browser.page_source, 'html.parser')

    # If there are no results, break out of the while loop
    if soup.find('h2', {'class': 'title is-2 margin-bottom-xs'}) is not None:
        kill_browser()
        break

    # Find all of the cards on the page
    cards = soup.find_all('article', {'class': 'card'})

    # Get the relevant data from each card
    for card in cards:
        certOrExam = card.find('p', {'class': 'card-content-super-title'}).text
        title = card.find('a', {'class': 'card-content-title'}).text.strip()
    
        tags = ''

        for tag in card.find_all('li', {'class': 'tag is-small'}):
            tags += tag.text + ', '

        rows.append((certOrExam, title, tags.strip()[:-1]))

    number += 30

    url = 'https://docs.microsoft.com/en-us/learn/certifications/browse/?skip=' + str(number)

    kill_browser()

    time.sleep(1)

    browser = start_chrome(url, headless=True)

    time.sleep(2)

    soup = BeautifulSoup(browser.page_source, 'html.parser')

# Create a pandas DataFrame
df = pd.DataFrame(data=rows, columns=['Certification or Exam?', 'Name', 'Tags'])

# Write the pandas DataFrame to an Excel workbook
df.to_excel(excel_writer=r'C:\Users\Alex\Desktop\Test.xlsx', sheet_name='Microsoft Certifications', index=False, freeze_panes=(1,0))

# Let the user know the script is finished
print('The script is complete.')
