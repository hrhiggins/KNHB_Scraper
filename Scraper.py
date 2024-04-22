from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook

# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N7/results'
driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(5)

cookies_popup = driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]')
if cookies_popup:
    driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]').click()
    driver.implicitly_wait(5)

shadow = Shadow(driver)
z = shadow.chrome_driver.get('https://www.knhb.nl/match-center#/competitions/N7/results')
element = shadow.find_element("match-center")
shadow.set_implicit_wait(3)
text = element.text
text = text.splitlines()

text_odd = text[1::2]
text_even = text[0::2]

team_away = []
team_home = []
pool = []
score = []


def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


for s in text_odd:
    if "H1" in s:
        team_home.append(s)
    elif len(s) == 1:
        pool.append(s)

for s in text_even:
    if "H1" in s:
        team_away.append(s)
    elif has_numbers(s) and "-" in s and len(s) < 10:
        score.append(s)

pd_array = pd.DataFrame(data=[team_home, score, team_away, pool]).T

print(pd_array)
pd_array.to_excel("C:\H1Results\h1_results.xlsx", sheet_name='Sheet1', startcol=1)

driver.close()
