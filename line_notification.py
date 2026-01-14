from ast import If, While
from encodings import utf_8
import numbers
from re import search

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from bs4 import BeautifulSoup
import time
import requests
import datetime
import os
import pygame

chrome_options = Options() 
chrome_options.add_argument('--headless')  # 啟動Headless 無頭
chrome_options.add_argument('--disable-gpu') #關閉GPU 避免某些系統或是網頁出錯E


URL = "https://course.ncku.edu.tw/index.php?c=qry_all"
course = input("Enter Lesson number: ")

PATH = "C:/Users/User/Downloads/chromedriver_win322/chromedriver.exe"

driver = webdriver.Chrome(PATH,chrome_options = chrome_options)
driver.get(URL)

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')

department = driver.find_element(By.XPATH, "//*[contains(text(), '%s')]" %course[0:2])
department.click()

time.sleep(0.7)

html = driver.page_source
soup = BeautifulSoup(html, 'html.parser')
courses = soup.find_all("td", align = "center")
coursesname = soup.find_all("span", class_ = "course_name")
number = soup.find_all("div", class_ = "dept_seq")

def index_containing_substring(the_list, substring):
    for i, s in enumerate(the_list):
        if substring in s:
            return i
    return -1

j = index_containing_substring(number, course)

def lineNotifyMessage(token, msg):

    headers = {
        "Authorization": "Bearer " + token, 
        "Content-Type" : "application/x-www-form-urlencoded"
    }

    payload = {'message': msg }
    r = requests.post("https://notify-api.line.me/api/notify", headers = headers, params = payload)
    return r.status_code

c = j*2 + 1
i = 0

while i == 0:
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    courses = soup.find_all("td",align="center")

    try:
        print(coursesname[j].text)
        print(courses[c].text)
        print(courses[c].text[len(courses[c].text)-1])
    except IndexError:
        driver = webdriver.Chrome(PATH,chrome_options = chrome_options)
        driver.get(URL)

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        department = driver.find_element(By.XPATH, "//*[contains(text(), '%s')]" %course[0:2])
        department.click()

        time.sleep(0.7)
        continue

    driver.refresh()
    time.sleep(0.7)
    os.system('cls')
    """
    i = i + 1
    if i == 100:
        driver = webdriver.Chrome(PATH,chrome_options = chrome_options)
        driver.get(URL)

        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        department = driver.find_element(By.XPATH, "//*[contains(text(), '%s')]" %course[0:2])
        department.click()

        time.sleep(0.7)
        i = 0
    """
    if courses[c].text[len(courses[c].text)-1] != "滿":
        token = 'linebot代碼'
        message = coursesname[j-1].text + "搶課啦"
        lineNotifyMessage(token, message)
        i = -1
        """
        file=r'C:/Users/User/Music/Music/ .mp3'
        pygame.mixer.init()
        print("播放音樂")
        track = pygame.mixer.music.load(file)
        pygame.mixer.music.play()
        time.sleep(15)
        pygame.mixer.music.stop()
        """
#driver.back()
