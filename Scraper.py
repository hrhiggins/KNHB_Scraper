from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
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
shadow.set_implicit_wait(5)
text = element.text
text = text.splitlines()
driver.close()

# months in dutch
months = ["januari", 'februari', 'maart', 'april', 'mei', 'juni', 'juli', 'augustus', 'september', 'oktober',
          'november', 'december']
# removing dates
for i in text:
    for k in months:
        if k in i:
            text.remove(i)

text_odd = text[1::2]
text_even = text[0::2]

team_away = []
team_home = []
pool = []
score = []
split_scores = []


def has_numbers(i):
    return any(char.isdigit() for char in i)


for s in text_even:
    if "H1" in s:
        team_home.append(s)
    elif len(s) == 1:
        pool.append(s)

for s in text_odd:
    if "H1" in s:
        team_away.append(s)
    elif has_numbers(s) and "-" in s and len(s) < 6:
        score.append(s)

# split the scores into home score and away score
for s in score:
    splitscore = s.replace('-', ' ').split()
    split_scores.append(splitscore)

split_scores = np.array(split_scores).flatten()
home_score = split_scores[0::2]
home_score = pd.to_numeric(home_score)
away_score = split_scores[1::2]
away_score = pd.to_numeric(away_score)

new_results = pd.DataFrame(data=[team_home, home_score, away_score, team_away, pool]).T

# location of excel file with results
file_path = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\H1\1K_results_2324.xlsx"

# read the current scores off the excel file
old_results = pd.read_excel(file_path, sheet_name='All Results')
# create a data frame of the old and new results
all_results = pd.concat([new_results, old_results])
# remove duplicate results
all_results = all_results.drop_duplicates()
# write all the results to the excel file
all_results.to_excel(file_path, sheet_name='All Results', index=False)

print("results uploaded")

##### NEXT: IF POOL COLUMN = A THEN ADD THAT ROW TO 'POOL A' RESULTS SHEET
