from selenium import webdriver
from PIL import Image
import time

# Chrome 드라이버 경로 설정
driver_path = '/chromedriver_win32/chromedriver.exe'

# 웹 드라이버 초기화
driver = webdriver.Chrome(executable_path=driver_path)

# 순회할 URL 목록
url = 'https://suis.seoultech.ac.kr/nxui/index.html'
urls = [
    '1',
    '2',
    '3',
    # 추가 URL 입력
]

for i, number in enumerate(urls):
    # 웹 페이지 로드
    driver.get(url + '#' + number)
    
    # 페이지 로드 완료 대기
    time.sleep(2)
    
    # 캡처화면 저장
    screenshot = driver.save_screenshot(f'screenshot_{i+1}.png')
    print(f'Screenshot saved: {screenshot}')

# 웹 드라이버 종료
driver.quit()
