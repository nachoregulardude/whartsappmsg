# Program to send bulk customized message through WhatsApp web application
# initial code by @inforkgodara, tweaks by Bharath G

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import pandas
import time

# Load the chrome driver
driver = webdriver.Chrome(executable_path='/usr/local/bin/chromedriver')
count = 0

# Open WhatsApp URL in chrome browser
driver.get("https://web.whatsapp.com/")
wait = WebDriverWait(driver, 20)

# Read data from excel
excel_data = pandas.read_excel('Customer.xlsx', sheet_name='Customers')

# Iterate excel rows till to finish
for column in excel_data['Name'].tolist():
    # Assign customized message
    message = excel_data['Message'][0]
    car = excel_data['Car'][count]
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
        message = message.replace('{customer_name}', column)
        message = message.replace('{customer_car}', car)
        person_title.send_keys(Keys.ENTER)
        actions = ActionChains(driver)
        actions.send_keys(message)
        actions.send_keys(Keys.ENTER)
        print("Message succesfully sent to ", column)
        print(count, " messages sent.")
        actions.perform()

# Close chrome browser
driver.quit()