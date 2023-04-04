from selenium import webdriver
from os import listdir, path, makedirs
import pickle
from datetime import datetime
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


# basic input folders
_dir = "input"
_output_dir = "output"


makedirs(_output_dir, exist_ok=True)

# for brave
# brave_path = '/usr/bin/brave'
# options.binary_location = brave_path
# options.add_argument('--remote-debugging-port=9224')  # NOT 9222


# hard coded by user
_user_Email = "israelmensah92@gmail.com"
_user_Password = "perekeme123"


def loadCookiesToBrowserOrLogin(browser, username, password):

    try:
        cookies = pickle.load(open("cookies.pkl", "rb"))

        print(cookies[0]['expiry']-time.time())

        if (cookies[0]['expiry']-time.time() < 0):
            return login(username, password, browser)

        for cookie in cookies:
            try:
                browser.add_cookie(cookie)
            except Exception as e:
                print(e)
        browser.maximize_window()
    except:
        return login(username, password, browser)


def login(_user_Email, _user_Password, browser):
    browser.maximize_window()

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
    time.sleep(25)  # time to solve captcha

    cookies = browser.get_cookies()
    pickle.dump(cookies, open("cookies.pkl", "wb"))

    print("saved cookies")


def getExcels(path: str) -> [str]:
    """This function returns the list of excel in path

    Args:
        path (str)

    Returns:
        [str]: list of excel(.xlsx) in path
    """
    return (list(filter(lambda elem: elem.endswith(".csv") or elem.endswith(".xlsx"), listdir(path))))


min_time = "" or 15
max_time = "" or 25


# button_element = WebDriverWait(browser, 50).until(
#     EC.presence_of_element_located((By.XPATH, '//*[@id="header"]/div[1]/div/div[2]/div[2]/div[1]/div/div/ul/li[5]/a')))
_columns = [
    "Query",	"Qty",	"Description",	"Schematic Reference",	"Internal Part Number",	"Venkel Part Number",	"_Description",	"Quantity Available",	"Reel Size",	"Full Reel Unit Price",
    "Full Reel Ext Price",	"MOQ",	"Standard Lead Time",	"URL",	"Current Date",	"Current Time"
]


def main():
    # initialise Browser
    browser = webdriver.Chrome(options=options)
    browser.get("https://www.venkel.com/")

    # loadCookies if it exist or still valid
    loadCookiesToBrowserOrLogin(browser, _user_Email, _user_Password)

    for excel in (getExcels(_dir)):
        print('\n\n\n')
        result_df = pd.DataFrame(columns=_columns)
        timestamp = datetime.now()
        raw_data = pd.read_excel(path.join(_dir, excel)) if excel.endswith(
            '.xlsx') else pd.read_csv(path.join(_dir, excel))
        present_columns = set(raw_data.columns).intersection(
            ['Internal Part Number', 'Description', 'Manufacturer', 'Query', 'Qty'])
        print(raw_data)
        if ("Query" in present_columns):
            for index, row in enumerate(raw_data.to_dict(orient='records')):
                print("currently at index: {} \nData\t:{}".format(index, row))
                updated_url = "https://www.venkel.com/{0}".format(row["Query"])
                browser.get(updated_url)
                time.sleep(random.randint(min_time, max_time))
                browser_title = browser.title
                if "404 Not Found" in browser_title:
                    print("404 Not Found")
                else:
                    button_element2 = WebDriverWait(browser, 200).until(
                        EC.presence_of_element_located((By.CLASS_NAME, 'product-collateral')))  # you should only use classname, using xpath will cause errors

                    try:
                        row["Venkel Part Number"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[1]/td').text
                        row["_Description"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[2]/td').text
                        row["Quantity Available"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[3]/td').text.replace(",", "")
                        row["Reel Size"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[4]/td').text.replace(",", "")
                        row["MOQ"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[5]/td').text
                        row["Standard Lead Time"] = browser.find_element(
                            By.XPATH, '//*[@id="product_addtocart_form"]/div[5]/table/tbody/tr[6]/td').text[:2]
                        row["URL"] = browser.current_url
                        row["Current Date"] = timestamp.date()
                        row["Current Time"] = timestamp.time()
                    except Exception as e:
                        (e, "Main Error")
                result_df = result_df.append(
                    row, ignore_index=True, sort=False)
        else:
            print("could not find `Query` in {}".format(excel))
        result_df[_columns].to_excel(
            path.join(_output_dir, str(timestamp)+"_"+(excel if excel.endswith(".xlsx") else excel+".xlsx")), index=False)

    browser.close()


# check if user is logged in
try:
    time.sleep(5)  # time to solve captcha
    main()
except:
    main()
