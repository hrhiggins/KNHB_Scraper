import pandas as pd
import numpy as np
from openpyxl import load_workbook
import re
import itertools

# location of excel file with results
file_pathlittle = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\H1\H1_1K_results_2324.xlsx"
file_pathbig = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\H1\H1_1K_results_2324.xlsx"

# read the current scores off the excel file
try:
    old_results = pd.read_excel(file_pathlittle, sheet_name='All Results')
    file_path = file_pathlittle
except IOError:
    old_results = pd.read_excel(file_pathbig, sheet_name='All Results')
    file_path = file_pathbig

# create a data frame of the old and new results
all_results = old_results
# remove duplicate results
all_results = all_results.drop_duplicates()
with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
    all_results.to_excel(writer, sheet_name='All Results', index=False)

for team in all_results['Home Team'].unique():
    team_results_df = pd.concat([all_results[all_results['Home Team'] == team].reset_index(drop=True),
                         all_results[all_results['Away Team'] == team].reset_index(drop=True)], axis=1)
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        team_results_df.to_excel(writer, sheet_name=team, index=False)

print("results uploaded")

# write league table to each pool sheet

