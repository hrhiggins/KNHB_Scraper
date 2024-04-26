from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
import openpyxl

# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N8/results'
driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(5)

cookies_popup = driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]')
if cookies_popup:
    driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]').click()
    driver.implicitly_wait(5)

shadow = Shadow(driver)
z = shadow.chrome_driver.get('https://www.knhb.nl/match-center#/competitions/N8/results')
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
    if "D1" in s:
        team_home.append(s)
    elif len(s) == 1:
        pool.append(s)

for s in text_odd:
    if "D1" in s:
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
new_results = new_results.rename(columns={0: 'Home Team', 1: 'Home Score', 2: 'Away Score', 3: 'Away Team', 4: 'Pool'})

# location of excel file with results
file_pathlittle = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\D1\D1_1K_results_2324.xlsx"
file_pathbig = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\D1\D1_1K_results_2324.xlsx"

# read the current scores off the excel file
try:
    old_results = pd.read_excel(file_pathlittle, sheet_name='All Results')
    file_path = file_pathlittle
except IOError:
    old_results = pd.read_excel(file_pathbig, sheet_name='All Results')
    file_path = file_pathbig

# create a data frame of the old and new results
all_results = pd.concat([new_results, old_results])
# remove duplicate results
all_results = all_results.drop_duplicates()
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    all_results.to_excel(writer, sheet_name='All Results', index=False)

for team in all_results['Home Team'].unique():
    team_df = pd.concat([all_results[all_results['Home Team'] == team].reset_index(drop=True),
                         all_results[all_results['Away Team'] == team].reset_index(drop=True)], axis=1)
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        team_df.to_excel(writer, sheet_name=team, index=False)

print("results uploaded")