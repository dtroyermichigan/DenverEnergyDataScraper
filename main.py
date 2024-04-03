#Imports
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
import openpyxl
import random
import time

#FOR THIS TO WORK, ALL OF THE ABOVE MUST BE INSTALLED AND MUST INSTALL CHROMEDRIVER.


# Set variables
mBuildingDict = {}
mCurrentID = ""
mBuildingID = 0

#Range Changers
mRangeInitial, mRangeEnd = 5601,5650



# Local Directory: MUST CHANGE IF RUNNING ON OTHER COMPUTERS!!!
s = Service(r'C:\Users\dtroy\Downloads\chromedriver-win64\chromedriver-win64\chromedriver.exe')

# Set up the Chrome tab
driver = webdriver.Chrome(service=s)

#This loops through all building IDS in the variable range.
for mBuildingID in range(mRangeInitial, mRangeEnd):
    #formats the ID to four digits. This shouldn't be an issue, more as a failsafe.
    formattedID = str(mBuildingID).zfill(4)

    #This pulls up the webpage.
    driver.get(f'https://lookup.energizedenver.org/?building_id={formattedID}')

    # Wait for the elements to be loaded, also to ensure we are not timed out.
    time.sleep(random.randint(1,10))

    #Give 5 seconds to do the calculation. This **shouldn't** be neccessary, again as a failsafe.
    wait = WebDriverWait(driver, 5)


    #If this is a vaild id
    try:
        #Find the Display Data: Set it into a list by splitting each of the entries
        display_data_element = wait.until(EC.presence_of_element_located((By.ID, "displayData")))
        display_data_text = display_data_element.text
        display_data_text = display_data_text.split('\n')
        #formatted like Building ID: xxxx , Address: xxxx etc.
        #For each item in the displaydata, it will loop and add it to the dictionary.
        for item in display_data_text:
            #if a new building ID, it will add it to the dictionary.
            if "Building ID:" in item:
                mCurrentID = item.split(":")[1].strip()
                mBuildingDict[mCurrentID] = {}
            #if entry is already apart of the ID, try to add to the ID. If the value doesn't exist, continue.
            else:
                try:
                    key, value = item.split(":")
                    mBuildingDict[mCurrentID][key.strip()] = value.strip()
                except ValueError:
                    # Handle cases where splitting by ":" is not possible (if item doesn't contain a ":")
                    continue

        #This is similar to the above, but for the table that's included.
        targets_table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.min-w-full")))
        rows = targets_table.find_elements(By.TAG_NAME, "tr")[1:]  # Skip header row

        # Loop through the rows of table to grab the EUI values
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            target_name = cells[0].text
            target_year = cells[1].text
            target_eui = cells[2].text

            key, value = target_year, target_eui
            mBuildingDict[mCurrentID][key.strip()] = value.strip()

    #Not a valid ID
    except:
        continue

#Quit chrome
driver.quit()

#Make the excel
df = pd.DataFrame.from_dict(mBuildingDict, orient='index')
df.to_excel('building_data.xlsx')



