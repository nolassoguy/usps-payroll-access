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
# This adds the Duplicate Tab Extension in Google Chrome
chrome_options.add_extension("extension_1_5_1_0.crx")
# The following code makes Chrome go into 'incognito' mode
# chrome_options.add_argument("--incognito") - not needed as extension will not be added

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

# 'Sign in' button click
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.XPATH, "//button[text()='Sign in']"))
	)
	element.click()
except:
	driver.quit()

# Employee ID to be populated
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.NAME, 'identifier'))
	)
	element.send_keys(username)
except:
	driver.quit()

# 'Next' button click
primary_button()

# Password to be populated
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.CLASS_NAME, 'password-with-toggle')) # the ID recently changed - switched to use CLASS_NAME instead
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
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.CLASS_NAME, 'button.button-primary'))
	)
	element.click()
except:
	driver.quit()

# Program clicks to send one-time verification code to email registered on LiteBlue
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.CLASS_NAME, 'button-link.enter-auth-code-instead-link'))
	)
	element.click()
except:
	driver.quit()

time.sleep(2)

# Gmail code extractor #
# START - Gmail code extractor snippet #

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
                       
# END - Gmail code extractor snippet #

time.sleep(2)

# Program enters the unique one-time-code sent to user's Gmail Inbox
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'input138')) # this number recently changed to 138 from 137 - need to find a better way to access this textfield with a CLASS_NAME perhaps?
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

time.sleep(2)

# Inputs eight digit username
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.NAME, 'j_username'))
	)
	element.send_keys(username)
except:
	driver.quit()

# Inputs password provided by user
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.NAME, 'j_password'))
	)
	element.send_keys(password)
except:
	driver.quit()

time.sleep(20)

# User makes a duplicate tab in the browser using keys: 'SHIFT' + 'ALT + d'

try:
	driver.switch_to.window(driver.window_handles[1])
except IndexError as e:
	print(f'IndexError: {e}')
	driver.quit()

time.sleep(2)

# User pastes in the password provided via the pyperclip module: 'COMMAND' + 'v'

# User mouse clicks on the 'Login' button

# The following gives user the most recent net pay to first paycheck of 2023

# TRYING A WHILE LOOP BELOW
count = 0
while count < 6: # <--- This number needs to be changed Monday 13 March 2023 to '7' 
	try:
		elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'list-group-item.list-group-item-action.col-md-6')))
		
		elements[count].click()
		
		#time.sleep(1)
		
		net_pay = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'net-pay')))
		
		#time.sleep(1)
		
		print(net_pay.text)
		
		next_statement = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.LINK_TEXT, 'Statements')))

		next_statement.click()

		count += 1

		if count == 5: # <--- This number needs to be changed Monday 13 March 2023 to '6' 
			break

	except: # need to add exception handling or do I use 'finally' here instead of 'except?
		driver.quit()

time.sleep(2)

print('That\'s all folks.')

driver.quit()

# try:
# 	element = WebDriverWait(driver, 10).until(
# 		EC.presence_of_all_elements_located((By.CLASS_NAME, 'list-group-item.list-group-item-action.col-md-6'))
# 	)
# 	element[0].click()
# 	element = WebDriverWait(driver, 10).until(
# 		EC.presence_of_element_located((By.ID, 'net-pay'))
# 	)
# 	print(element.text)
# 	time.sleep(5)
# 	element = WebDriverWait(driver, 10).until(
# 		EC.presence_of_element_located((By.ID, 'logout-link'))
# 	)
# 	element.click()
# except:
# 	driver.quit()
# 	print('Element not located...')

# time.sleep(30)

# driver.quit()