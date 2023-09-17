from selenium import webdriver
import pyautogui
import pyperclip

import os
import numpy as np
import time
import cv2

import tkinter.messagebox as msgbox
from tkinter import * # __all__

from Gui import *

class Auto_Web_Driver:
    def __init__(self, loc, path):
        self.location = loc
        self.path = path

        options = webdriver.ChromeOptions()
        self.driver = webdriver.Chrome(options=options)

    def error_img(self, file_path, img_loc):    # 이미지 에러검증
        if img_loc:
            print(f"{file_path} 이미지가 발견되었습니다: {img_loc}")
            return True
        else:
            print(f"{file_path} 이미지를 찾을 수 없습니다.")
            return False
    
    def find_img(self, file_path, time=1):    # 이미지 찾기
        np_img = np.fromfile(os.path.join(self.path, 'assets', 'img', file_path), np.uint8)
        img = cv2.imdecode(np_img, cv2.IMREAD_COLOR)
        img_loc = pyautogui.locateOnScreen(img, minSearchTime=time, confidence=0.8)
    
        if self.error_img(file_path, img_loc):
            return img_loc
        else:
            return (0, 0)
        
    def no_wait_find_img(self, file_path):    # 기다려서 이미지 찾기
        start_time = time.time()
        elapsed_time = time.time() - start_time
        timeout = 6
        np_img = np.fromfile(os.path.join(self.path, 'assets', 'img', file_path), np.uint8)
        img = cv2.imdecode(np_img, cv2.IMREAD_COLOR)
        while True:
            img_loc = pyautogui.locateOnScreen(img, confidence=0.8)
            if img_loc:
                break

            elapsed_time = time.time() - start_time  # 경과 시간 계산
            if elapsed_time > timeout:
                print("시간 제한에 도달했습니다.")
                break

            time.sleep(0.1)

        print(f"로드바 이미지 감지 여부 : {elapsed_time}")

        if self.error_img(file_path, img_loc):
            return img_loc
        else:
            return (0, 0)

    def click_img(self, file_path, time = 2):    # 이미지 클릭
        self.location = self.find_img(file_path, time)
        pyautogui.click(self.location[0], self.location[1], button='left')
        return

    def draw_info(self, wait_time):    # 페이지가 로드 되고 작업
        try: # 팝업 창 예외 처리
            load_loc = self.find_img('load.png', wait_time) # load 글귀
            rec_off_loc = self.find_img('record_off.png') # record off

            if load_loc == (0, 0) or rec_off_loc == (0, 0):
                raise TypeError()

            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left')
            pyautogui.moveTo(load_loc[0], load_loc[1])
            pyautogui.click(load_loc[0], load_loc[1], button='left')

        except TypeError:
            print("팝업 윈도우가 열렸습니다.")
            print(f"driver.window_handles : {self.driver.window_handles}")
            main_window_handle = self.driver.window_handles[0] # 현재 윈도우와 팝업 윈도우의 핸들을 가져옵니다.
            popup_window_handle = self.driver.window_handles[1]
            self.driver.switch_to.window(popup_window_handle) # 팝업 윈도우로 전환

            self.driver.close() # 팝업 윈도우 닫기

            self.driver.switch_to.window(main_window_handle) # 메인 웹페이지로 복귀
            pyautogui.press('f12') # 개발자 모드 다시 키고

            load_loc = self.find_img('load.png', wait_time) # load 글귀
            rec_off_loc = self.find_img('record_off.png') # record off

            if load_loc == (0, 0) or rec_off_loc == (0, 0) or load_loc == None or rec_off_loc == None:
                msgbox.showinfo('정보', '이미지 찾기를 실패하였습니다.')
            
            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left')
            pyautogui.moveTo(load_loc[0], load_loc[1])
            pyautogui.click(load_loc[0], load_loc[1], button='left')

        finally:
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'c')
            time.sleep(1)
            pyautogui.click(rec_off_loc[0], rec_off_loc[1], button='left')
            pyautogui.moveTo(10, 10)

            return pyperclip.paste()