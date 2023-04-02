from selenium import webdriver
import pickle
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

options = Options()
options.add_experimental_option('detach', True)
options.add_experimental_option("excludeSwitches", ["enable-logging"])

browser = webdriver.Chrome(options=options)
browser.get("https://www.venkel.com/")
cookies = pickle.load(open("cookies.pkl", "rb"))
# browser.delete_all_cookies()
for cookie in cookies:
    try:
        browser.add_cookie(cookie)
    except Exception as e:
        print(e)
browser.maximize_window()    # to maximise window because of login button

# hard coded by user
_user_Email = "israelmensah92@gmail.com"
_user_Password = "perekeme123"


""""
      read query from excel input file
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

# click on login on home page
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
time.sleep(15)

part_numbercol=[]
descriptioncol=[]
qty_availcol=[]
reel_sizecol=[]
moqcol=[]
standard_leadcol=[]


for item in listt:
    updated_url = "https://www.venkel.com/{0}".format(item)
    browser.get(updated_url)
    print(updated_url)
    button_element2 = WebDriverWait(browser, 200).until(
    EC.presence_of_element_located((By.CLASS_NAME, 'product-collateral')))  # you should only use classname, using xpath will cause errors
    
    try:
        part_number = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[1]/td').text
        description = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[2]/td').text
        qty_avail = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[3]/td').text
        reel_size = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[4]/td').text
        moq = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[5]/td').text
        standard_lead = browser.find_element(
                By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[6]/td').text
        
        part_numbercol.append(part_number)         
        descriptioncol.append(description)        
        qty_availcol.append(qty_avail)        
        reel_sizecol.append(reel_size)     
        moqcol.append(moq)     
        standard_leadcol.append(standard_lead)
        #     print(f'{part_number} \n {description}\n {qty_avail}\n {reel_size}\n {moq}\n {standard_lead}')
    except Exception as e:
        print(e, "Main Error")

df = pd.DataFrame({"part_number": part_numbercol, "description": descriptioncol, " qty_avail": qty_availcol, "reel_size": reel_sizecol, "moq":   moqcol, "standard_lead": standard_leadcol })
df.to_excel("Venkel.xlsx", index=False)
print(df)





# browser.get("https://www.venkel.com/C0201X6S160-104KNP")

# button_element2 = WebDriverWait(browser, 200).until(
#     EC.presence_of_element_located((By.CLASS_NAME, 'product-collateral')))  # you should only use classname, using xpath will cause errors

# part_numbercol=[]
# descriptioncol=[]
# qty_availcol=[]
# reel_sizecol=[]
# moqcol=[]
# standard_leadcol=[]
# try:
# #     table1=  browser.find_element(
# #         By.CLASS_NAME, 'table-bordered')
# #     print(table1)

# #     rows= table1.find_elements(By.TAG_NAME,"tr")
# #     print(rows)

# #     for row in rows:
# #         detail.append(row.find_element(By.XPATH,"./td").text)
# #     print(detail)    

#     part_number = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[1]/td').text
#     description = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[2]/td').text
#     qty_avail = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[3]/td').text
#     reel_size = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[4]/td').text
#     moq = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[5]/td').text
#     standard_lead = browser.find_element(
#         By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[6]/td').text
    
#     part_numbercol.append(part_number)         
#     descriptioncol.append(description)        
#     qty_availcol.append(qty_avail)        
#     reel_sizecol.append(reel_size)     
#     moqcol.append(moq)     
#     standard_leadcol.append(standard_lead)         

# #     print(f'{part_number} \n {description}\n {qty_avail}\n {reel_size}\n {moq}\n {standard_lead}')
# except Exception as e:
#     print(e, "Main Error")

# df = pd.DataFrame({"part_number": part_numbercol, "description": descriptioncol, " qty_avail": qty_availcol, "reel_size": reel_sizecol, "moq":   moqcol, "standard_lead": standard_leadcol })
# df.to_excel("Venkel.xlsx", index=False)
# print(df)