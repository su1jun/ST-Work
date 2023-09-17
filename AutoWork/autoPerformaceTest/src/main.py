import pyautogui

import os
import json
import time
from datetime import datetime

from Auto_Web_Driver import *
import Gui

def run(data, path):
    # 웹드라이버 설정
    site_path = os.path.join(path, 'assets', 'set', 'site_list.txt') # 저장된 사이트 목록 이름

    target_sites = {}
    # 사이트 목록 불러오고 저장하기
    try:
        if os.path.exists(site_path):
            with open(site_path, 'r') as f:
                target_sites = json.loads(f.read())
    except json.JSONDecodeError:
        pass

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

    if os.path.exists(site_path):
        with open(site_path, 'w') as f:
            f.write(json.dumps(target_sites))
    else:
        with open(site_path, mode="w+") as f:
            f.write(json.dumps(target_sites))
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

def test(data, target_sites):
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