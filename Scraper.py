import time
import requests
from bs4 import BeautifulSoup
from pyshadow.main import Shadow
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys


# page to access as a string
url = 'https://www.knhb.nl/match-center#/competitions/N7/results'
driver = webdriver.Chrome()
driver.get(url)
driver.implicitly_wait(3)

driver.find_element(By.XPATH, '//*[@id="bcSubmitConsentToAll"]').click()
driver.implicitly_wait(3)

#root1 = driver.find_element(By.CSS_SELECTOR, ".tcol12 > match-center:nth-child(2)")
#shadow_root = root1.shadow_root
#shadow_content = shadow_root.find_element(By.TAG_NAME, "span")

#shadow = Shadow(driver)
#element = shadow.find_element()
#element.click()
#driver.implicitly_wait(3)

shadow_host1 = driver.find_element(By.CSS_SELECTOR, '.tcol12 > match-center:nth-child(2)')
shadow_root1 = driver.execute_script('return arguments[0].shadowRoot', shadow_host1)
shadow_content = shadow_host1.text
print (shadow_content)

# create beautifulsoup object
#soup = BeautifulSoup(driver.page_source, "html.parser")

#span_tags = soup.find_all(By.TAG_NAME, 'span')

# find and print all text in the span tags
#for span in span_tags:
 # print(span.text)

#driver.quit()

