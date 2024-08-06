from General_Scraper import general_scraper
import datetime

# create data time string
now = str(datetime.datetime.now())[:19]
now = now.replace(":", "_")

url = 'https://www.knhb.nl/match-center#/competitions/N2/results'
file_path_little = r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\2024-2025\D1\D1_HK_results_2425.xlsx"
file_path_big = r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\2024-2025\D1\D1_HK_results_2425.xlsx"
dst_dir_little = (r"C:\Users\Harry\OneDrive\Hockey\Results and Analysis\2024-2025\D1\Previous\D1_HK_results_2425_"
                  + str(now) + ".xlsx")
dst_dir_big = (r"C:\Users\Harry Higgins\OneDrive\Hockey\Results and Analysis\2024-2025\D1\Previous\D1_HK_results_2425_"
               + str(now) + ".xlsx")
even = "True"
scraper = general_scraper(url, file_path_little, file_path_big, dst_dir_little, dst_dir_big, even)
