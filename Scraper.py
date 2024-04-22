import requests
from bs4 import BeautifulSoup

# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N7/results'

# request.get function fetches the raw html
page = requests.get(url)

# create beautifulsoup object
soup = BeautifulSoup(page.text, "html.parser")

# badges = soup.body.find('div', attrs={'class': 'badges'})
#for span in badges.span.find_all('span', recursive=False):
 #   print

span_tags = soup.find_all('span')
for span in span_tags:
    print(span.text)

#li_tags = soup.find_all('li')
#for li in li_tags:
    #print(li.text)


