from selenium import webdriver
import pyautogui
import pyperclip
from PIL import Image
from selenium.common.exceptions import NoAlertPresentException

import os
import numpy as np
import time
import cv2

pyautogui.FAILSAFE = True
path = os.path.dirname(os.path.realpath(__file__)) # 파일 읽는 상대 경로

class Mouse:
    def __init__(self, loc):
        self.location = loc

    # @staticmethod
    def error_img(self, file_path, img_loc):    # 이미지 에러검증
        if img_loc:
            print(f"{file_path} 이미지가 발견되었습니다: {img_loc}")
            return True
        else:
            print(f"{file_path} 이미지를 찾을 수 없습니다.")
            return False
    
    def find_img(self, file_path, time=5):    # 이미지 찾기
        np_img = np.fromfile(os.path.join(path, file_path), np.uint8)
        img = cv2.imdecode(np_img, cv2.IMREAD_COLOR)
        img_loc = pyautogui.locateOnScreen(img, minSearchTime=time, confidence=0.8)
    
        if self.error_img(file_path, img_loc):
            self.location = img_loc
        else:
            self.location = (0, 0)
        return

    def click_img(self, file_path, time=5):    # 이미지 클릭
        self.find_img(file_path, time)
        pyautogui.click(self.location[0], self.location[1], button='left')
        return

    def draw_info(self):    # 페이지가 로드 되고 작업
        
        try: # 팝업 창 예외 처리
            load_loc = self.find_img('load.png', 4) # load 글귀
            rec_off_loc = self.find_img('record_off.png') # record off

            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left', confidence=0.8)
            pyautogui.moveTo(load_loc[0], load_loc[1])
            pyautogui.click(load_loc[0], load_loc[1], button='left', confidence=0.8)

        except TypeError:
            print("팝업 윈도우가 열렸습니다.")
            print(f"driver.window_handles : {driver.window_handles}")
            main_window_handle = driver.window_handles[0] # 현재 윈도우와 팝업 윈도우의 핸들을 가져옵니다.
            popup_window_handle = driver.window_handles[2]
            driver.switch_to.window(popup_window_handle) # 팝업 윈도우로 전환

            driver.close() # 팝업 윈도우 닫기

            driver.switch_to.window(main_window_handle) # 메인 웹페이지로 복귀
            pyautogui.press('f12') # 개발자 모드 다시 키고

            load_loc = self.find_img('load.png', 1) # load 글귀
            rec_off_loc = self.find_img('record_off.png') # record off
            print(f"전 {rec_off_loc}")

            print("여기까지 되는지1")
            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left', confidence=0.8)
            print("여기까지 되는지2")
            pyautogui.moveTo(load_loc[0], load_loc[1])
            print("여기까지 되는지3")
            pyautogui.click(load_loc[0], load_loc[1], button='left', confidence=0.8)
            print(f"전 {rec_off_loc}")

        finally:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(2)
            print(f"후 {rec_off_loc}")
            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left', confidence=0.8)

            return pyperclip.paste()

# 웹드라이버 설정
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(options=options)
myMouse = Mouse((0, 0)) # 클릭 클래스

pyautogui.press('f12')  # 'f12' 키를 누름 -> 개발자 모드

# 네트워크 체크 초기 설정
myMouse.click_img('network.png') # network 탭
myMouse.click_img('disable_cache.png') # disable_cache 탭
myMouse.click_img('setting.png') # setting 탭
myMouse.click_img('popup.png') # popup 탭

# 웹사이트 접속
target_sites = {'서울과학기술대학교(국문)':'https://www.seoultech.ac.kr/index.jsp'}
for site in target_sites.values():
    print(site)
    before_window_cnt = len(driver.window_handles) # 현재 윈도우 핸들 개수 저장
    driver.get(site)

    time.sleep(1)

    after_window_cnt = len(driver.window_handles) # 대기 이후 윈도우 핸글 개수 저장
    
    # # 팝업 윈도우가 열렸는지 확인
    # while after_window_cnt > before_window_cnt:
    #     print(f"{site} 사이트에서 열린 팝업의 수 / 초기 : {before_window_cnt} -> 현재 : {after_window_cnt}")
    #     close_popup()

    # driver.maximize_window()
    print(myMouse.draw_info())

# 웹드라이버 종료
driver.quit()

# x, y = pyautogui.position()
# pyautogui.moveTo(x, y, duration=1)  # duration은 이동하는 데 걸리는 시간(초)
# pyautogui.click(x, y, button='left')  # button은 'left', 'middle', 'right' 중 하나

# pyautogui.write('Hello, world!')
# pyautogui.press('enter')  # 'enter' 키를 누름
# pyautogui.hotkey('ctrl', 'c')  # 'ctrl' + 'c' 조합키를 누름

# screenshot = pyautogui.screenshot()
# screenshot.save('screenshot.png')

# location = pyautogui.locateOnScreen('image.png')

# Exception has occurred: TypeError
# 'NoneType' object is not subscriptable