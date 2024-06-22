# -*- coding: utf-8 -*-
# x, y = pyautogui.position()
# pyautogui.moveTo(x, y, duration=1)  # duration은 이동하는 데 걸리는 시간(초)
# pyautogui.click(x, y, button='left')  # button은 'left', 'middle', 'right' 중 하나

# pyautogui.write('Hello, world!')
# pyautogui.press('enter')  # 'enter' 키를 누름
# pyautogui.hotkey('ctrl', 'c')  # 'ctrl' + 'c' 조합키를 누름

# screenshot = pyautogui.screenshot()
# screenshot.save('screenshot.png')

# location = pyautogui.locateOnScreen('image.png')

from selenium import webdriver
import pyautogui
import pyperclip

import os
import json
import numpy as np
import time
import cv2
from datetime import datetime

import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
from collections import defaultdict

from openpyxl import Workbook

# 네트워크탭, 캐쉬 비활성화, 셋팅, 팝업창
# load 감지, record 감지

# 모니터 작은거, 설정 탭, 어뎁터 옵션 변경, 와이파이, 이더넷, 사용, 사용 안함

# 사이트 목록, 측정 횟수, 각 이미지 파일 전체 저장 위치

class Gui:
    def __init__(self, path):
        self.root = Tk()
        self.root.title("Auto Performace Tester")
        self.path = path
        self.option_name = 'default.txt'
        self.option_path = os.path.join(self.path, 'assets', 'set', self.option_name)
        self.draw_ui()
    
    def draw_ui(self): # gui 기본 프레임 그리기
        # 환경설정 프레임
        self.action_frame = Frame(self.root)
        self.action_frame.pack(fill="x", padx=5, pady=5) # 간격 띄우기
    
        # 측정반복횟수
        self.frame1 = LabelFrame(self.action_frame, text="측정")
        self.frame1.pack(fill="x", padx=5, pady=5, ipady=8)
    
        self.frame1_lbl1 = Label(self.frame1, text="반복 횟수 : ")
        self.frame1_lbl1.pack(side="left", padx=6, pady=2)
    
        self.frame1_spb1_value = IntVar()
        self.frame1_spb1 = Spinbox(self.frame1, from_=1, to=4, textvariable=self.frame1_spb1_value)
        self.frame1_spb1.pack(side="left", padx=10, pady=2) # 높이 변경
    
        # 환경변수옵션
        self.frame2 = LabelFrame(self.action_frame, text="로드")
        self.frame2.pack(fill="x", padx=5, pady=6, ipady=8)
    
        self.frame2_lbl1 = Label(self.frame1, text="대기 시간 : ")
        self.frame2_lbl1.pack(side="left", padx=6, pady=2)
    
        self.frame2_spb1_value = IntVar()
        self.frame2_spb1 = Spinbox(self.frame1, from_=1, to=4, textvariable=self.frame2_spb1_value)
        self.frame2_spb1.pack(side="left", padx=10, pady=2) # 높이 변경
    
        # 현재파일위치
        self.frame3 = LabelFrame(self.action_frame, text="파일 설정")
        self.frame3.pack(fill="x", padx=5, pady=5, ipady=8)

        self.frame3_lbl1 = Label(self.frame3, text="파일 이름 : ")
        self.frame3_lbl1.pack(side="left", padx=6, pady=2)

        self.frame3_ety1 = Entry(self.frame3)
        self.frame3_ety1.pack(side="right", fill="x", expand=True, padx=6, pady=2)
    
        # 저장 경로 프레임
        self.path_frame = LabelFrame(self.root, text="저장 경로")
        self.path_frame.pack(fill="x", padx=5, pady=5, ipady=5)
    
        self.txt_dest_path = Entry(self.path_frame)
        self.txt_dest_path.pack(side="left", fill="x", expand=True, padx=6, pady=5, ipady=4) # 높이 변경
    
        self.btn_dest_path = Button(self.path_frame, text="찾아보기", width=10, command=self.browse_dest_path)
        self.btn_dest_path.pack(side="right", padx=5, pady=5)
    
        # 실행 프레임
        self.frame_run = Frame(self.root)
        self.frame_run.pack(fill="x", padx=5, pady=5)
    
        self.btn_close = Button(self.frame_run, padx=5, pady=5, text="닫기", width=10, command=self.on_exit)
        self.btn_close.pack(side="right", padx=5, pady=5)
    
        self.btn_start = Button(self.frame_run, padx=5, pady=5, text="시작", width=10, command=self.start)
        self.btn_start.pack(side="right", padx=5, pady=5)

        self.load_last_date()

        self.root.resizable(False, False)
        self.root.mainloop()

    def make_data(self):
        data = dict()
        data['test_times'] = self.frame1_spb1.get()
        data['wait_time'] = self.frame2_spb1.get()
        data['name'] = self.frame3_ety1.get()
        data['path'] = self.txt_dest_path.get()
        return data

    def on_exit(self): # gui 창 닫기
        data = self.make_data()
        print("data", data)
        self.root.destroy()
        self.save_last_date(data)
        return
    
    def browse_dest_path(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected == "": # 사용자가 취소를 누를 때
            return
        self.txt_dest_path.delete(0, END)
        self.txt_dest_path.insert(0, folder_selected)

    def save_last_date(self, data):
        if os.path.exists(self.option_path):
            with open(self.option_path, 'w') as f:
                f.write(json.dumps(data))
        else:
            with open(self.option_path, mode="w+") as f:
                f.write(json.dumps(data))
        print(f"File path: {os.path.abspath(self.option_path)}")

    def load_last_date(self):
        try:
            if os.path.exists(self.option_path):
                with open(self.option_path, 'r') as f:
                    last_data = json.loads(f.read())
        except json.JSONDecodeError:
            last_data = {}
        if not last_data:
            last_data = {'test_times' : 1, 'wait_time' : 6, 'name' : '', 'path' : ''}

        # 로드 데이터 gui에 뿌리기
        if last_data['test_times'] != '':
            self.frame1_spb1.delete(0, END)  # 현재 값을 지웁니다.
            self.frame1_spb1.insert(0, last_data['test_times'])     # 새 값을 삽입합니다.
        if last_data['wait_time'] != '':
            self.frame2_spb1_value.set(last_data['wait_time'])
        if last_data['name'] is not None: self.frame3_ety1.insert('0', last_data['name'])
        if last_data['path'] is not None: self.txt_dest_path.insert('0', last_data['path'])
        
    def start(self):
        data = self.make_data()
        self.save_last_date(data)
        run(data, self.path)
        return

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

def run(data, path):
    # 웹드라이버 설정
    site_path = os.path.join(path, 'data', 'set', 'site_list.txt') # 저장된 사이트 목록 이름

    target_sites = {}
    # 사이트 목록 불러오고 저장하기
    try:
        if os.path.exists(site_path):
            with open(site_path, 'r') as f:
                for line in f:
                    if ":" in line:
                        name, site = line.strip().split(":", 1)
                        target_sites[name.strip()] = site.strip()
    except Exception as e:
        print(f"parsing err : {e}")

    if not target_sites:
        target_sites = {
            '서울과학기술대학교(국문)':'https://www.seoultech.ac.kr/index.jsp',
            '서울과학기술대학교(영문)':'https://en.seoultech.ac.kr/',
            '카이스트(국문)':'https://www.kaist.ac.kr/kr/',
            '카이스트(영문)':'https://www.kaist.ac.kr/en//',
            '서울시립대학교(국문)':'https://www.uos.ac.kr/main.do?epTicket=LOG',
            '서울시립대학교(영문)':'https://www.uos.ac.kr/en/main.do',
            '서울대학교(국문)':'https://www.snu.ac.kr/',
            '서울대학교(영문)':'https://en.snu.ac.kr/',
        }

    try:
        with open(site_path, 'w' if os.path.exists(site_path) else 'w+') as f:
            for name, site in target_sites.items():
                f.write(f"{name}:{site}\n")
        print(f"파일 경로: {os.path.abspath(site_path)}")
    except Exception as e:
        print(f"파일 저장 중 오류 발생: {e}")
        
    print(f"Sile path: {os.path.abspath(site_path)}")

    result = test(data, target_sites, path)

    print(result)

    file_name = os.path.join(data['name'], datetime.today().strftime('_%y%m%d')) # 저장된 파일 이름
    result = "\n".join(result) # 성능 측정 결과

    # Writing Korean text to a file
    if os.path.exists(file_name):
        with open(file_name, mode='w', encoding='utf-8') as f:
            f.write(result)
    else:
        with open(file_name, mode="w+", encoding='utf-8') as f:
            f.write(result)

    gui.root.destroy()

def test(data, target_sites, path):
    # 웹드라이버 설정
    auto_web_driver = Auto_Web_Driver((0, 0), path) # 클릭 클래스

    pyautogui.press('f12')  # 'f12' 키를 누름 -> 개발자 모드

    # 네트워크 체크 초기 설정
    auto_web_driver.click_img('network.png') # network 탭
    auto_web_driver.click_img('disable_cache.png') # disable_cache 탭
    auto_web_driver.click_img('setting.png') # setting 탭
    auto_web_driver.click_img('popup.png') # popup 탭

    # 웹사이트 성능 측정
    performance_data = [[] for _ in range(len(target_sites))]
    test_times = data['test_times'] # 측정 횟수
    wait_time = data['wait_time'] # 로딩 대기 시간

    for i, site in enumerate(target_sites.keys()):
        performance_data[i].append(site)
        print(f"-----{site}-----")

        for _ in range(int(test_times)):
            auto_web_driver.driver.get(target_sites[site])

            time.sleep(1)
            performance_data[i].append(auto_web_driver.draw_info(wait_time))

        performance_data[i] = "\n".join(performance_data[i])

    auto_web_driver.driver.quit() # 웹드라이버 종료
    result = "\n".join(performance_data) # 성능 측정 결과
    return result

if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    path = os.path.dirname(os.path.realpath(__file__)) # 파일 읽는 상대 경로
    path = os.path.dirname(path)
    gui = Gui(path)