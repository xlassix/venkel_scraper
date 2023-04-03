from selenium import webdriver
import pickle
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import random

options = Options()
options.add_experimental_option('detach', True)
options.add_experimental_option("excludeSwitches", ["enable-logging"])

browser = webdriver.Chrome(options=options)
browser.get("https://www.venkel.com/")
cookies = pickle.load(open("cookies.pkl", "rb"))

for cookie in cookies:
    try:
        browser.add_cookie(cookie)
    except Exception as e:
        print(e)

browser.maximize_window()    # to maximise window because of login button

# hard coded by user
_user_Email = "israelmensah92@gmail.com"
_user_Password = "perekeme123"

min_time= "" or 15
max_time= "" or 20

""""
      read query from excel input file
"""
excel_file = "Venkel Scraper - Input Sample.xlsx"
df_excel = pd.read_excel(excel_file)
df_excel_query = df_excel["Query"]

listt = df_excel_query.tolist()
# print(listt)
""""
        excellll
"""

button_element = WebDriverWait(browser, 50).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')))

# Columns
part_numbercol = []
descriptioncol = []
qty_availcol = []
reel_sizecol = []
moqcol = []
standard_leadcol = []

def main():
    for item in listt:
        updated_url = "https://www.venkel.com/{0}".format(item)
        browser.get(updated_url)
        time.sleep(random.randint(min_time,max_time))
        btitle = browser.title
        if "404 Not Found" in btitle:
            part_numbercol.append("Part not found")
            descriptioncol.append("Error")
            qty_availcol.append("Error")
            reel_sizecol.append("Error")
            moqcol.append("Error")
            standard_leadcol.append("Error")
            print("404 NOt Found")
        else:

            button_element2 = WebDriverWait(browser, 200).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'product-collateral')))  # you should only use classname, using xpath will cause errors
        
            try:
                part_number = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[1]/td').text 
                description = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[2]/td').text     
                qty_avail = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[3]/td').text.replace(",", "") 
                reel_size = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[4]/td').text.replace(",", "") 
                moq = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[5]/td').text 
                standard_lead = browser.find_element(
                    By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[6]/td').text[:2] 

                part_numbercol.append(part_number)
                descriptioncol.append(description)
                qty_availcol.append(qty_avail)
                reel_sizecol.append(reel_size)
                moqcol.append(moq)
                standard_leadcol.append(standard_lead)
                #     print(f'{part_number} \n {description}\n {qty_avail}\n {reel_size}\n {moq}\n {standard_lead}')
            except Exception as e:
                (e, "Main Error")



# check if user is logged in
try:
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
    time.sleep(25)#time to solve captcha
    main()
except:
    main()

""""
      write data into excel output file
"""
df = pd.DataFrame({"part_number": part_numbercol, "description": descriptioncol, " qty_avail": qty_availcol,
                  "reel_size": reel_sizecol, "moq":   moqcol, "standard_lead": standard_leadcol})
df.to_excel("Venkel.xlsx", index=False)
print(df)
""""
       excel file
"""
browser.close()