from selenium import webdriver
import pickle
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = Options()
options.add_experimental_option("excludeSwitches", ["enable-logging"])

browser = webdriver.Chrome(options=options)
browser.get("https://www.venkel.com/")

browser.maximize_window()

_user_Email = "israelmensah92@gmail.com"
_user_Password = "perekeme123"


button_element = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')))

# find login btn and click
_login = browser.find_element(
    By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')
_login.click()

# find input fields
_input_email = browser.find_element(By.ID, 'email')
_input_password = browser.find_element(By.ID, 'pass')
_login2 = browser.find_element(By.ID, 'send2')

# to login as a user
_input_email.send_keys(_user_Email)
_input_password.send_keys(_user_Password)
_login2.click()
time.sleep(25) #time to solve captcha    

cookies = browser.get_cookies()
pickle.dump(cookies, open("cookies.pkl","wb"))

print("loaded cookies")

browser.close()