from selenium import webdriver
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

options = Options()
options.add_experimental_option('detach', True)
options.add_experimental_option("excludeSwitches", ["enable-logging"])

browser = webdriver.Chrome(options=options)
browser.get("https://www.venkel.com/")
browser.maximize_window()    # to maximise window because of login button


_user_Email = "israelmensah92@gmail.com"
_user_Password = "perekeme123"


""""
        excellll
"""

excel_file = "Venkel Scraper - Input Sample.xlsx"
df_excel = pd.read_excel(excel_file)
df_excel_query = df_excel["Query"]
# print(df_excel_query)

listt = df_excel_query.tolist()
# print(listt)

""""
        excellll
"""

button_element = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')))

_login = browser.find_element(
    By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')
_login.click()
# debug  print("login clicked" )


_input_email = browser.find_element(By.ID, 'email')
_input_password = browser.find_element(By.ID, 'pass')
_login2 = browser.find_element(By.ID, 'send2')
# debug  print(_input_id)
# debug  print(_input_password)

# to login as a user
_input_email.send_keys(_user_Email)
_input_password.send_keys(_user_Password)
_login2.click()
# debug print(_login2)

get_url = browser.current_url
# updated_url = get_url[:23]
# updated_url = "https://www.venkel.com/catalogsearch/result/?q={0}".format(item)

# get_url = "https://www.venkel.com/C0201X6S160-104KNP"
# print(get_url)
# browser.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS)
# https://www.venkel.com/catalogsearch/result/?q=C0201X6S160-104KNP

for item in listt:
    updated_url = "https://www.venkel.com/catalogsearch/result/?q={0}".format(item)
    browser.get(updated_url)


    
