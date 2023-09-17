import emoji
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from datetime import datetime
from time import sleep
import random

ID = 'qaxsd530'
PW = 'asdt5230!'

keyword_list = ['실험', '피험자', '연구', '사례', '참가']
title_list = []

service = Service(executable_path='chromedriver_win32/chromedriver')
webdriver_options = webdriver.ChromeOptions()
# webdriver_options.add_argument('--headless') #내부 창을 띄울 수 없으므로 설정
webdriver_options.add_argument('--no-sandbox')
webdriver_options.add_argument('--disable-dev-shm-usage')
webdriver_options.add_argument("user-agent={Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.5845.111 Safari/537.36}")

driver = webdriver.Chrome(options=webdriver_options)
driver.implicitly_wait(1)

# 로그인 페이지에서 로그인하기
url = 'https://everytime.kr/login'
driver.get(url)

driver.find_element(By.NAME, 'userid').send_keys(ID)
driver.find_element(By.NAME, 'password').send_keys(PW)
driver.find_element(By.XPATH, '//*[@id="container"]/form/p[3]/input').click()
driver.implicitly_wait(5)

results = [] # 글, 댓글을 저장하는 리스트
cnt = 0 # 첫 페이지
max_cnt = 0 # 마지막 페이지

while True:
    print('Page ' + str(cnt))

    if cnt > max_cnt:
        break

    cnt = cnt + 1
    driver.get('https://everytime.kr/{$ 학교 고유의 자유게시판 번호}/p/' + str(cnt))
    driver.implicitly_wait(50)

    posts = driver.find_elements_by_css_selector('article > a.article')
    links = [post.get_attribute('href') for post in posts]

    for link in links:
        driver.implicitly_wait(50)

        # 1~6 초 사이 랜덤 n초 만큼 sleep
        # 에타는 크롤링 봇을 강력히 금지해서 sleep 필수
        # 크롤링 임시 계정은 영구정지

        sleep(random.randrange(1, 7))
        driver.get(link)

        commnets = driver.find_elements_by_css_selector('p.large')

        for comment in commnets:
            results.append(comment.text)

# 리스트를 하나의 문자열로 변환
long_results = ' '.join([str(result) for result in results])
print(long_results)

from wordcloud import WordCloud
import numpy as np
import matplotlib.pyplot as plt
import PIL.Image
wc = WordCloud(font_path='/content/malgun.ttf').generate(long_results)

plt.imshow(wc)
plt.axis("off") # 이미지 축 제거
plt.show() # 워드클라우드 이미지 띄우기

