# Program to send bulk customized messages/videos through the WhatsApp web application
# Code by Bharath G in collaboration with Suhail Ahmed

import os
from subprocess import getoutput
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time


def sendmsg(contact, message, fromwhere):
    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    time.sleep(0.35)
    person_title = wait.until(
        lambda driver: driver.find_element_by_xpath(search_box))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(contact)

    # Wait for 4 seconds to search contact number
    time.sleep(0.35)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath(
            '//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel the sheet
        time.sleep(0.1)
        person_title.send_keys(Keys.ENTER)
        actions = ActionChains(driver)
        # Whether you want the message to be copied from the clipboard or the Spreadsheet
        if fromwhere=='1':
            actions.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL)
        else:
            if '{customer_car}' in message:
                message = message.replace('{customer_car}', car)
            if '{customer_name}' in message:
                message = message.replace('{customer_name}', column)
            actions.send_keys(message)
        actions.send_keys(Keys.ENTER)
        time.sleep(0.2)
        print("Sending message to ", column)
        print(count+1, " messages sent.")
        actions.perform()


def sendvid(contact, path, if_text):
    # Locate search box through x_path
    person_title = wait.until(
        lambda driver: driver.find_element_by_xpath('//*[@id="side"]/div[1]/div/label/div/div[2]'))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(contact)

    # Wait for 4 seconds to search contact number
    time.sleep(0.25)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath(
            '//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        time.sleep(0.35)
        person_title.send_keys(Keys.ENTER)
        # Find and click the attach button
        attachment_box = driver.find_element_by_xpath(
            '//div[@title = "Attach"]')
        attachment_box.click()

        # Find and click the video option
        image_box = driver.find_element_by_xpath(
            '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(path)

        # Wait until the video loads
        # time.sleep(load)
        actions = ActionChains(driver)
        #If you want text to be sent with the video. Only takes from clipboard
        # Hit the send button
        send_button = wait.until(
            lambda driver: driver.find_element_by_xpath('//span[@data-icon="send"]'))

        if if_text=='1':
            actions.key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL)
        time.sleep(1)
#        send_button.click()

        actions.send_keys(Keys.ENTER)
        print("Video succesfully sent to ", column)
        print(count+1, " videos sent.")
        actions.perform()

# Load the chrome driver
print("Opening Chrome...\n")
driver = webdriver.Chrome(executable_path='/usr/local/bin/chromedriver')
count = 0

# Open WhatsApp URL in chrome browser
print("Going to WhatsApp web...\n")
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
print("Reading excel data...\n")
excel_data = pandas.read_excel('Customer.xlsx', sheet_name='Customers')

# User choses between sending a text message or a video message
print('\n\nv: Video/Photo \nt: Text message')
msgtyp=''
while not msgtyp:
    msgtyp = input("Enter the kind of message you want to send(v/t): ")
fromwhere=''
if_text=''
correct='0'
if msgtyp.lower() == 't':
    while not fromwhere:
        fromwhere = input("Enter 0 to send Message from the Spreadsheet\n      1 to send Message from Clipboard\n:")
    if fromwhere==1:
        input("The text sent will be copied from the clipboard.\nPress Enter when the text is copied: ")
    # Iterate excel rows till to finish
    for column in excel_data['Name'].tolist():
        # Stop once it reaches first nan value
        if excel_data['Name'].isnull()[count]:
            print("Job completed.")
            break
        # Assign customized message
        message = excel_data['Message'][0]
#        car = excel_data['Car'][count]
        contact = str(int(excel_data['Contact'][count]))
        sendmsg(contact, message, fromwhere)
        count = count + 1
    time.sleep(3)
if msgtyp.lower() == 'v':
    # Take input for path to video
    print("The video must be in the same folder as the program.")
    print("\nThese are the videos/photos in the folder: \n\n")
    files=getoutput("ls -1 | grep -E '.mp4|.3gp|.jpg|.png|.jpeg'")
    print(files)
    filelist=files.splitlines()
    if len(filelist)==1:
        correct='1'
        filename=filelist[0]
        path=''
        path+= os.path.dirname(os.path.realpath(__file__))+'/'
        path+=filename
        input('{} \nPress Enter if the path is correct...'.format(path))
#    else:
#        no=0
#        for file in filelist:
#            filelist[n] = '{}. '.format(no) + filelist[n] 
#            no = no + 1
#        print(filelist)
    while correct=='0':
        filename = input("\nEnter the filename: ")
        path=''
        path+= os.path.dirname(os.path.realpath(__file__))+'/'
        path+=filename
        if filename not in filelist:   
            print("Please enter the correct file name")
            print(files)
            continue
        correct=input('{} \nPress Enter if the path is correct'.format(path))
        print("Enter 0 to change filename: ")
    # load = int(input("Enter a time(s) load: "))
    print("The video or photo will be loaded from:\n{}\n".format(path))
    if_text = input("\n\nEnter 1 if you want to send text with the video\n      0 if no\n :")
    if if_text=='1':
        input("The text sent will be copied from the clipboard.\nPress Enter when the text is copied...")
    # Iterate excel rows till to finish
    for column in excel_data['Name'].tolist():
        # Stop once it reaches first nan value

        if excel_data['Name'].isnull()[count]:
            print("Job completed.")
            break
        # Assign customized message
        contact = str(int(excel_data['Contact'][count]))
        sendvid(contact, path, if_text)
        count = count + 1
        time.sleep(2)
    # This time is dependent on the size of the file that you are sharing
    input(
        "Press Enter when all videos have been uploaded...")

# Close chrome browser
driver.quit()
