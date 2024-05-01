from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import numpy as np
import openpyxl

# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N7/results'
driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(5)

cookies_popup = driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]')
if cookies_popup:
    driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]').click()
    driver.implicitly_wait(10)

shadow = Shadow(driver)
z = shadow.chrome_driver.get('https://www.knhb.nl/match-center#/competitions/N7/results')
element = shadow.find_element("match-center")
shadow.set_implicit_wait(10)
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
winner = []
goal_difference = []


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
goal_difference = pd.to_numeric(goal_difference)

new_results = pd.DataFrame(data=[team_home, home_score, away_score, team_away, goal_difference, winner, pool]).T
new_results = new_results.rename(columns={0: 'Home Team', 1: 'Home Score', 2: 'Away Score', 3: 'Away Team',
                                          4: 'Goal Difference', 5: 'Winner', 6: 'Pool'})

new_results['Goal Difference'] = new_results['Home Score'] - new_results['Away Score']

new_results.loc[new_results['Goal Difference'] < 0, 'Winner'] = 'Away'
new_results.loc[new_results['Goal Difference'] == 0, 'Winner'] = 'Draw'
new_results.loc[new_results['Goal Difference'] > 0, 'Winner'] = 'Home'

# location of excel file with results
file_path_little = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\H1\H1_1K_results_2324.xlsx"
file_path_big = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\H1\H1_1K_results_2324.xlsx"

# read the current scores off the excel file
try:
    old_results = pd.read_excel(file_path_little, sheet_name='All Results')
    file_path = file_path_little
except IOError:
    old_results = pd.read_excel(file_path_big, sheet_name='All Results')
    file_path = file_path_big

# create a data frame of the old and new results
all_results = pd.concat([new_results, old_results])
# remove duplicate results
all_results = all_results.drop_duplicates()
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    all_results.to_excel(writer, sheet_name='All Results', index=False)

pool_A = pd.DataFrame()
pool_A = pool_A.rename(columns={0: 'Team', 1: 'Points', 2: 'Games Played', 3: 'Goals For',
                                4: 'Goals Against', 5: 'Goal Difference'})
pool_B = pd.DataFrame()
pool_B = pool_B.rename(columns={0: 'Team', 1: 'Points', 2: 'Games Played', 3: 'Goals For',
                                4: 'Goals Against', 5: 'Goal Difference'})
pool_C = pd.DataFrame()
pool_C = pool_C.rename(columns={0: 'Team', 1: 'Points', 2: 'Games Played', 3: 'Goals For', 4: 'Goals Against',
                                5: 'Goal Difference'})
pool_D = pd.DataFrame()
pool_D = pool_D.rename(columns={0: 'Team', 1: 'Points', 2: 'Games Played', 3: 'Goals For', 4: 'Goals Against',
                                5: 'Goal Difference'})

for team in all_results['Home Team'].unique():
    team_df = pd.concat([all_results[all_results['Home Team'] == team].reset_index(drop=True),
                         all_results[all_results['Away Team'] == team].reset_index(drop=True)], axis=1)
    # renaming all the columns of the data frame
    team_df = team_df.set_axis(['Home Team', 'Home Score', 'Away Score', 'Away Team', 'Goal Difference',
                                'Winner', 'Pool', 'home team', 'home score', 'away score', 'away team',
                                'goal difference', 'winner', 'pool'], axis=1)
    total_goals_for = sum(team_df['Home Score']) + sum(team_df['away score'])
    total_goals_against = sum(team_df['Away Score']) + sum(team_df['home score'])
    total_goal_difference = sum(team_df['Goal Difference']) - sum(team_df['goal difference'])
    total_games_played = team_df['Home Team'].shape[0] + team_df['away team'].shape[0]
    wins = (team_df[team_df['Winner'] == 'Home'].shape[0]) + (team_df[team_df['winner'] == 'Away'].shape[0])
    draws = (team_df[team_df['Winner'] == 'Draw'].shape[0]) + (team_df[team_df['winner'] == 'Draw'].shape[0])
    losses = (team_df[team_df['Winner'] == 'Away'].shape[0]) + (team_df[team_df['winner'] == 'Home'].shape[0])
    points = (wins*3)+(draws*1)
    team_result = pd.DataFrame(data=[team, points, total_games_played, total_goals_for,
                                     total_goals_against, total_goal_difference])

    try:
        pool = team_df.at[0, 'Pool']
    except KeyError:
        pool = team_df.at[0, 'pool']

    if pool == 'A':
        pool_A = pd.concat([pool_A, team_result])
        pool_result = pool_A
    elif pool == 'B':
        pool_B = pd.concat([pool_B, team_result])
        pool_result = pool_B
    elif pool == 'C':
        pool_C = pd.concat([pool_C, team_result])
        pool_result = pool_C
    else:
        pool_D = pd.concat([pool_D, team_result])
        pool_result = pool_D

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        pool_result.to_excel(writer, sheet_name=pool, index=False)

    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        team_df.to_excel(writer, sheet_name=team, index=False)

print("results uploaded")
