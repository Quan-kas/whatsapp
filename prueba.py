# -*- coding: utf-8 -*-
"""
Created on Wed Jul  7 11:02:03 2021

@author: kevin
"""

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time
import os
import sys

# Load the chrome driver
driver = webdriver.Chrome()
count = 0
mainMsg = ""

# Open WhatsApp URL in chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
excel_data = pandas.read_excel('Customer bulk email data.xlsx', sheet_name='Customers')
image_filename="images.png"

# Iterate excel rows till to finish
for column in excel_data['Name'].tolist():
    # Assign customized message
    message = excel_data['Image'][0]

    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(str(excel_data['Contact'][count]))
    count = count + 1

    # Wait for 2 seconds to search contact number
    time.sleep(2)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel sheet
        filepath = message.replace('{customer_name}', column)
        person_title.send_keys(Keys.ENTER)
        attachment_box = driver.find_element_by_xpath('//div[@title = "Adjuntar"]')
        attachment_box.click()
        imgvid_box = driver.find_element_by_xpath("//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")
        imgvid_box.send_keys(filepath)
        time.sleep(3)
        send_button = driver.find_element_by_xpath('//span[@data-icon="send-light"]')
        send_button.click()

# Close chrome browser
time.sleep(10)
driver.quit()
