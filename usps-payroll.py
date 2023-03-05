from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from time import sleep
import imaplib
import email, app_password, re, creds, maskpass
import pyperclip as pc

chrome_options = Options()
chrome_options.add_experimental_option('detach', True)
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation'])
# This temporarily adds the 'Duplicate Tab Shortcut' extension in Google Chrome
chrome_options.add_extension('extension_1_5_1_0.crx')
# The following code makes Chrome go into 'incognito' mode if ever needed
# chrome_options.add_argument('--incognito')

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
		EC.presence_of_element_located((By.CLASS_NAME, 'password-with-toggle')) # the ID recently changed - switched to use CLASS_NAME instead - more fool proof
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

sleep(2)

# - START - Gmail code extractor snippet #
# Code taken and modified from the YouTube video 'AMT2 - Extracting Emails from your Gmail Inbox using python' 
# (https://youtu.be/K21BSZPFIjQ)

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

# - END - Gmail code extractor snippet #

sleep(2)

# Program enters the unique one-time-code sent to user's Gmail Inbox
try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'input138')) # This number recently changed to input138 from input137 - need to find a better way to access this text field with a CLASS_NAME perhaps?
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

sleep(2)

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

sleep(15)

# IMPORTANT: User makes a duplicate tab in the browser using the following keys within 15 seconds: 'SHIFT' + 'ALT + d'
# This bypasses 'j_security' for some reason and allows user access to ePayroll

try:
	driver.switch_to.window(driver.window_handles[1])
except:
	driver.quit() #TODO need an exception here for when there is no 2nd tab

sleep(2)

# User pastes in the password provided via the 'pyperclip' module: 'COMMAND' + 'v'

# User mouse clicks on or sends RETURN key for the 'Login' button

# Selenium finds all paychecks displayed on page for the year 2023 (TODO: make variables and ask for user input if they would like to change the year - this can be changed with '/22' or '/21' for example)
num_of_paychecks = len(driver.find_elements(By.PARTIAL_LINK_TEXT, '/23')) #TODO: put in a Try/Except block like the others

# While loop goes to each paycheck, finds the net pay and prints it in the CLI
count = 0

while count < num_of_paychecks + 1:
	try:
		elements = WebDriverWait(driver, 10).until(EC.presence_of_all_elements_located((By.CLASS_NAME, 'list-group-item.list-group-item-action.col-md-6'))
		)
		elements[count].click()
		
		sleep(2)

		# This expands the 'Leave & Retirement' accordian menu 
		try:
			element = WebDriverWait(driver, 10).until(
				EC.presence_of_all_elements_located((By.CLASS_NAME, 'btn.btn-link.w-100.text-left'))
			)
			element[2].click()
		except:
			driver.quit()

		sleep(2)
		
		pay_date = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'pay-date')))
		net_pay = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'net-pay')))
		al_balance = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'annual-leave-earned-available')))
		sl_balance = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'sick-leave-current-balance')))
		xday_balance = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, 'other-leave-amount-4')))
		
		sleep(2)
		
		print(pay_date.text + ': ' + net_pay.text.replace('$','').replace(',',''))
		print(f'Available AL Balance: {al_balance.text}')
		print(f'Current SL Balance: {sl_balance.text}')
		print(f'Current X-Days Balance: {xday_balance.text}\n')
		
		next_statement = WebDriverWait(driver,10).until(EC.presence_of_element_located((By.LINK_TEXT, 'Statements')))

		next_statement.click()

		count += 1

		if count == num_of_paychecks:
			break
	except:
		driver.quit()

print('That\'s all folks.')

sleep(10)

try:
	element = WebDriverWait(driver, 10).until(
		EC.presence_of_element_located((By.ID, 'logout-link'))
	)
	element.click()
except:
	driver.quit()

sleep(10)

driver.quit()