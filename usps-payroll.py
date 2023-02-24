from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import imaplib
import email
import app_password
import re
import creds
import maskpass
import time
import pyperclip as pc

chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_experimental_option("useAutomationExtension", False)
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_argument("--incognito")

username = creds.username
password = maskpass.advpass('Enter your password:\n', '*')

# Copies the user's password into the clipboard to paste later on in the script
pc.copy(password)

login_url = 'https://liteblue.usps.gov/wps/myportal'

s = Service('/Users/jemiller/chromedriver.exe')
driver = webdriver.Chrome(service = s, options = chrome_options)

driver.maximize_window()


driver.get(login_url)

# Generic 'button.button-primary' button click function
def primary_button():
	driver.find_element(By.CLASS_NAME, 'button.button-primary').send_keys(Keys.RETURN)

# time.sleep(2)

# 'Sign in' button click
# sign_in_btn = driver.find_element(By.XPATH, "//button[text()='Sign in']")
# sign_in_btn.click()
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.XPATH, "//button[text()='Sign in']"))
	)
	element.click()
except:
	driver.quit()

# Employee ID to be populated
# driver.find_element(By.ID, 'input27').send_keys(username)
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'input27'))
	)
	element.send_keys(username)
except:
	driver.quit()

# 'Next' button click
primary_button()

# Password to be populated
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'input59'))
	)
	element.send_keys(password)
except:
	driver.quit()

# 'Verify' button click
primary_button()

# Select how user would like to receive the one time code
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.LINK_TEXT, 'Select'))
	)
	element.click()
except:
	driver.quit()

# 'Send me an email' button click
# send_me = driver.find_element(By.CLASS_NAME, 'button.button-primary')
# send_me.click()
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.CLASS_NAME, 'button.button-primary'))
	)
	element.click()
except:
	driver.quit()

# Enter code manually link instead of link provided in email
# enter_code = driver.find_element(By.CLASS_NAME, 'button-link.enter-auth-code-instead-link')
# enter_code.click()
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.CLASS_NAME, 'button-link.enter-auth-code-instead-link'))
	)
	element.click()
except:
	driver.quit()

time.sleep(5)

########################
# Gmail code extractor #

imap_url = 'imap.gmail.com'

my_mail = imaplib.IMAP4_SSL(imap_url)

my_mail.login(app_password.user, app_password.password)

my_mail.select('Inbox')

key = 'FROM'
value = 'noreply@litebluemail.usps.gov'

_, data = my_mail.search(None, key, value)

mail_id_list = data[0].split()

msgs = []
for num in mail_id_list:
	typ, data = my_mail.fetch(num, '(RFC822)')
	msgs.append(data)

for msg in msgs[::-1]:
	for response_part in msg:
		if type(response_part) is tuple:
			my_msg = email.message_from_bytes((response_part[1]))
			# print('____________________________________________')
			# print('subj: ', my_msg['subject'])
			# print('from: ', my_msg['from'])
			# print('Body:')
			for part in my_msg.walk():
				# print(part.get_content_type())
				if part.get_content_type() == 'text/plain':
					pattern = re.compile(r'\d{6}')
					matches = pattern.findall(part.get_payload())
					#print(part.get_payload())
one_time_code = matches[0]

########################

time.sleep(2)

# Inputs a unique one time code sent to user's Gmail Inbox
# driver.find_element(By.ID, 'input137').send_keys(one_time_code)
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'input137'))
	)
	element.send_keys(one_time_code)
except:
	driver.quit()

primary_button()

# Clicks the 'ePayroll' link at the bottom of landing page
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.LINK_TEXT, 'ePayroll'))
	)
	element.click()
except:
	driver.quit()

# Clicks on the 'Enter Application' button
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'enter-button'))
	)
	element.click()
except:
	driver.quit()

time.sleep(5)

# Inputs eight digit username
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.NAME, 'j_username'))
	)
	element.send_keys(username)
except:
	driver.quit()

# time.sleep(2)

# Inputs password provided by user
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.NAME, 'j_password'))
	)
	element.send_keys(password)
except:
	driver.quit()

time.sleep(2)

# Make duplicate tab in the browser... try this ... THIS WORKS! WHEN I ENTER PASSWORD FROM pyperclip module

# Clicks on the 'Login' button for 'ePayroll'
# login_btn = driver.find_element(By.XPATH, "//button[text()='Login']")
# login_btn.click()

# time.sleep(5)

# User has to hit back in the browser instead of 'driver.back()'

# This where the user pastes in the password saved in the clipboard from the pyperclip module

# The following does not work
# try:
# 	element = WebDriverWait(driver, 30).until(
# 		EC.presence_of_element_located((By.CLASS_NAME, 'list-group-item.list-group-item-action.col-md-6'))
# 	)
# 	element.click()
# except:
# 	driver.quit()
# 	print('Element not located...')


# driver.quit()