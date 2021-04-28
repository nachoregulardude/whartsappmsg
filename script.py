# Program to send bulk customized message through WhatsApp web application
# initial code by @inforkgodara, tweaks by Bharath G

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time

def sendmsg(contact,message):
    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(contact)
    
    # Wait for 2 seconds to search contact number
    time.sleep(2)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel sheet
        message = message.replace('{customer_name}', column)
        message = message.replace('{customer_car}', car)
        person_title.send_keys(Keys.ENTER)
        actions = ActionChains(driver)
        actions.send_keys(message)
        actions.send_keys(Keys.ENTER)
        time.sleep(0.5)
        print("Message succesfully sent to ", column)
        print(count, " messages sent.")
        actions.perform()

def sendvid(contact, path):
    # Locate search box through x_path
    search_box = '//*[@id="side"]/div[1]/div/label/div/div[2]'
    person_title = wait.until(lambda driver:driver.find_element_by_xpath(search_box))

    # Clear search box if any contact number is written in it
    person_title.clear()

    # Send contact number in search box
    person_title.send_keys(contact)
    
    # Wait for 2 seconds to search contact number
    time.sleep(2)

    try:
        # Load error message in case unavailability of contact number
        element = driver.find_element_by_xpath('//*[@id="pane-side"]/div[1]/div/span')
    except NoSuchElementException:
        # Format the message from excel sheet
        person_title.send_keys(Keys.ENTER)
        # Find and click the attach button
        attachment_box = driver.find_element_by_xpath('//div[@title = "Attach"]')
        attachment_box.click()

        #Find and click the video option
        image_box = driver.find_element_by_xpath(
            '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')
        image_box.send_keys(path)

        # Wait until the video loads
        time.sleep(3)

        # Hit the send button
        send_button = driver.find_element_by_xpath('//span[@data-icon="send"]')
        send_button.click()

        print("Message succesfully sent to ", column)
        print(count, " messages sent.")
    
# Load the chrome driver
driver = webdriver.Chrome(executable_path='/usr/local/bin/chromedriver')
count = 0

# Open WhatsApp URL in chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
excel_data = pandas.read_excel('Customer.xlsx', sheet_name='Customers')

# User choses between sending a text message or a video message
msgtyp = input("Enter the kind of message you want to send(v/t): ")

if msgtyp == 't':
    # Iterate excel rows till to finish
    for column in excel_data['Name'].tolist():
        #Stop once it reaches first nan value
        if excel_data['Name'].isnull()[count]:
                print("Job completed.")
                break
        # Assign customized message
        message = excel_data['Message'][0]
        car = excel_data['Car'][count]
        contact = str(int(excel_data['Contact'][count]))
        sendmsg(contact,message)
        count = count + 1

if msgtyp == 'v':
    # Take input for path to video
    path = input("Enter complete path to the video: ")

    # Iterate excel rows till to finish
    for column in excel_data['Name'].tolist():
        #Stop once it reaches first nan value
        if excel_data['Name'].isnull()[count]:
                print("Job completed.")
                break
        # Assign customized message
        contact = str(int(excel_data['Contact'][count]))
        sendvid(contact,path)
        count = count + 1
    # This time is dependent on the size of the file that you are sharing
    print("Leaving the page open for an additional 30 seconds so that files upload.")
    time.sleep(30) 

# Close chrome browser
driver.quit()
