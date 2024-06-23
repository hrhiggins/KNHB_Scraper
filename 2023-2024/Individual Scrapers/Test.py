from General_Scraper import general_scraper
from Only_Excel_DataParsing import only_excel_dataparsing
import datetime

# create data time string
now = str(datetime.datetime.now())[:19]
now = now.replace(":", "_")

url = 'https://www.knhb.nl/match-center#/competitions/N7/results'
file_path_little = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\2023-2024\H1\Test.xlsx"
file_path_big = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\2023-2024\H1\Test.xlsx"
dst_dir_little = (r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\2023-2024\H1\Previous\Test_"
                  + str(now) + ".xlsx")
dst_dir_big = (r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\2023-2024\H1\Previous\Test_"
               + str(now) + ".xlsx")
club_location_little = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\Club_locations.xlsx"
club_location_big = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\Club_locations.xlsx"
even = "True"

scraper = general_scraper(url, file_path_little, file_path_big, dst_dir_little, dst_dir_big, club_location_little,
                          club_location_big, even)
# scraper = only_excel_dataparsing(file_path_little, file_path_big, dst_dir_little, dst_dir_big)
