import os
import json

from tkinter import * # __all__
from tkinter import filedialog

from macro import *
from main import *

class MainInterface:
    def __init__(self, path):
        self.root = Tk()
        self.root.title("Auto Macro")
        self.path = path
        self.option_name = 'default.txt'
        self.option_path = os.path.join(self.path, 'data', 'setting', self.option_name)
        
        self.drawUi()
    
    def drawUi(self): # gui 기본 프레임 그리기
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
    
        self.frame2_lbl1 = Label(self.frame2, text="대기 시간 : ")
        self.frame2_lbl1.pack(side="left", padx=6, pady=2)
    
        self.frame2_spb1_value = IntVar()
        self.frame2_spb1 = Spinbox(self.frame2, from_=1, to=4, textvariable=self.frame2_spb1_value)
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
    
        self.btn_dest_path = Button(self.path_frame, text="찾아보기", width=10, command=self.browseDestPath)
        self.btn_dest_path.pack(side="right", padx=5, pady=5)
    
        # 실행 프레임
        self.frame_run = Frame(self.root)
        self.frame_run.pack(fill="x", padx=5, pady=5)
    
        self.btn_close = Button(self.frame_run, padx=5, pady=5, text="닫기", width=10, command=self.onExit)
        self.btn_close.pack(side="right", padx=5, pady=5)
    
        self.btn_start = Button(self.frame_run, padx=5, pady=5, text="시작", width=10, command=self.onStart)
        self.btn_start.pack(side="right", padx=5, pady=5)

        self.loadDate()

        self.root.resizable(False, False)
        self.root.mainloop()

    def loadDate(self):
        try:
            if os.path.exists(self.option_path):
                with open(self.option_path, 'r') as f:
                    last_data = json.loads(f.read())
        except Exception as e:
            print(f"last_data_parsing : {e}")
            last_data = {'test_times' : 1, 'wait_time' : 6, 'name' : '', 'path' : ''}

        # 로드 데이터 gui에 뿌리기
        if last_data['test_times'] != '':
            self.frame1_spb1.delete(0, END)  # 현재 값을 지웁니다.
            self.frame1_spb1.insert(0, last_data['test_times'])     # 새 값을 삽입합니다.
        if last_data['wait_time'] != '':
            self.frame2_spb1_value.set(last_data['wait_time'])
        if last_data['file_name'] is not None: self.frame3_ety1.insert('0', last_data['file_name'])
        if last_data['save_path'] is not None: self.txt_dest_path.insert('0', last_data['save_path'])
        
    def browseDestPath(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected == "": # 사용자가 취소를 누를 때
            return
        self.txt_dest_path.delete(0, END)
        self.txt_dest_path.insert(0, folder_selected)

    def saveTestData(self, data):
        if os.path.exists(self.option_path):
            with open(self.option_path, 'w') as f:
                f.write(json.dumps(data))
        else:
            with open(self.option_path, mode="w+") as f:
                f.write(json.dumps(data))
        print(f"File path: {os.path.abspath(self.option_path)}")
    
    def makeOptionData(self):
        data = dict()
        data['test_times'] = self.frame1_spb1.get()
        data['wait_time'] = self.frame2_spb1.get()
        data['file_name'] = self.frame3_ety1.get()
        data['save_path'] = self.txt_dest_path.get()
        return data
    
    def onExit(self): # gui 창 닫기
        data = self.makeOptionData()
        print("data", data)
        self.root.destroy()
        self.saveTestData(data)
    
    def onStart(self): # 작업 시작
        data = self.makeOptionData()
        macro = ChromeDriver((0, 0), data, self.path)
        performance_data = macro.run()
        self.saveTestData(data)
    
if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    path = os.path.dirname(os.path.dirname(os.path.realpath(__file__))) # 파일 읽는 상대 경로
    main_interface = MainInterface(path)