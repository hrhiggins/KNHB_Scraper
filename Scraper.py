import requests
from bs4 import BeautifulSoup

# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N7/results'

# request.get function fetches the raw html
page = requests.get(url)

# create beautifulsoup object
soup = BeautifulSoup(page.content)

# find all table cells (with tag 'td')
cells = soup.findALL('span')
print(cells)


