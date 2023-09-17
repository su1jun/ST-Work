import os
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import * # __all__
from tkinter import filedialog
from turtle import left

root = Tk()
root.title("Window Macro")

# 파일 추가
def add_file():
    files = filedialog.askopenfilenames(title="이미지 파일을 선택하세요", \
        filetypes=(("PNG 파일", "*.png"), ("모든 파일", "*.*")), \
        initialdir=r"C:\Users\Nadocoding\Desktop\PythonWorkspace")
        # 최초에 사용자가 지정한 경로를 보여줌
    
    # 사용자가 선택한 파일 목록
    for file in files:
        list_file.insert(END, file)

# 선택 삭제
def del_file():
    #print(list_file.curselection())
    for index in reversed(list_file.curselection()):
        list_file.delete(index)

# 저장 경로 (폴더)
def browse_dest_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "": # 사용자가 취소를 누를 때
        print("폴더 선택 취소")
        return
    #print(folder_selected)
    txt_dest_path.delete(0, END)
    txt_dest_path.insert(0, folder_selected)

# 이미지 통합
def merge_image():
    # print("가로넓이 : ", cmb_width.get())
    # print("간격 : ", cmb_space.get())
    # print("포맷 : ", cmb_format.get())

    try:
        # 가로넓이
        img_width = cmb_width.get()
        if img_width == "원본유지":
            img_width = -1 # -1 일때는 원본 기준으로
        else:
            img_width = int(img_width)

        # 간격
        img_space = cmb_space.get()
        if img_space == "좁게":
            img_space = 30
        elif img_space == "보통":
            img_space = 60
        elif img_space == "넓게":
            img_space = 90
        else: # 없음
            img_space = 0

        # 포맷
        img_format = cmb_format.get().lower() # PNG, JPG, BMP 값을 받아와서 소문자로 변경

        #####################################

        images = [Image.open(x) for x in list_file.get(0, END)]    

        # 이미지 사이즈 리스트에 넣어서 하나씩 처리
        image_sizes = [] # [(width1, height1), (width2, height2), ...]
        if img_width > -1:
            # width  값 변경
            image_sizes = [(int(img_width), int(img_width * x.size[1] / x.size[0])) for x in images]
        else:
            # 원본 사이즈 사용
            image_sizes = [(x.size[0], x.size[1]) for x in images]

        # 계산식
        # 100 * 60 이미지가 있음. -> width 를 80 으로 줄이면 height 는?
        # (원본 width) : (원본 height) = (변경 width) : (변경 height)
        #     100      :     60       =      80      :     ?
        #      x       :     y        =      x'      :     y'
        #    xy' = x'y
        #    y' = x'y / x -> 이 식을 적용
        #  100:60=80:48

        # 우리 코드에 대입하려면?
        # x = width = size[0]
        # y = height = size[1]
        # x' = img_width # 이 값으로 변경 해야 함
        # y' = x'y / x = img_width * size[1] / size[0]

        widths, heights = zip(*(image_sizes))

        # 최대 넓이, 전체 높이 구해옴
        max_width, total_height = max(widths), sum(heights)
        
        # 스케치북 준비
        if img_space > 0: # 이미지 간격 옵션 적용
            total_height += (img_space * (len(images) - 1))

        result_img = Image.new("RGB", (max_width, total_height), (255, 255, 255)) # 배경 흰색
        y_offset = 0 # y 위치

        for idx, img in enumerate(images):
            # width 가 원본유지가 아닐 때에는 이미지 크기 조정
            if img_width > -1:
                img = img.resize(image_sizes[idx])

            result_img.paste(img, (0, y_offset))
            y_offset += (img.size[1] + img_space) # height 값 + 사용자가 지정한 간격

            progress = (idx + 1) / len(images) * 100 # 실제 percent 정보를 계산
            p_var.set(progress)
            progress_bar.update()

        # 포맷 옵션 처리
        file_name = "nado_photo." + img_format
        dest_path = os.path.join(txt_dest_path.get(), file_name)
        result_img.save(dest_path)
        msgbox.showinfo("알림", "작업이 완료되었습니다.")
    except Exception as err: # 예외처리
        msgbox.showerror("에러", err)

# 시작
def start():
    # 각 옵션들 값을 확인
    # print("가로넓이 : ", cmb_width.get())
    # print("간격 : ", cmb_space.get())
    # print("포맷 : ", cmb_format.get())

    # 파일 목록 확인
    if list_file.size() == 0:
        msgbox.showwarning("경고", "이미지 파일을 추가하세요")
        return

    # 저장 경로 확인
    if len(txt_dest_path.get()) == 0:
        msgbox.showwarning("경고", "저장 경로를 선택하세요")
        return

    # 이미지 통합 작업
    merge_image()

# 액션 프레임 (마우스, 키보드, 기타 액션)
action_frame = Frame(root)
action_frame.pack(fill="x", padx=5, pady=5) # 간격 띄우기

# 마우스 영역
mouse_frame = LabelFrame(action_frame, text="마우스 액션")
mouse_frame.pack(fill="x", padx=5, pady=5, ipady=8)

mouse_act1 = Button(mouse_frame, text="이동", width=15)
mouse_act1.pack(side="left", padx=8, pady=2)

mouse_act2 = Button(mouse_frame, text="클릭", width=15)
mouse_act2.pack(side="left", padx=8, pady=2)

mouse_act3 = Button(mouse_frame, text="휠 이동", width=15)
mouse_act3.pack(side="left", padx=8, pady=2)

mouse_act4 = Button(mouse_frame, text="드레그", width=15)
mouse_act4.pack(side="left", padx=8, pady=2)

# 키보드 영역
keybd_frame = LabelFrame(action_frame, text="키보드 액션")
keybd_frame.pack(fill="x", padx=5, pady=5, ipady=8)

keybd_act1 = Button(keybd_frame, text="단일 키", width=15)
keybd_act1.pack(side="left", padx=8, pady=2)

keybd_act2 = Button(keybd_frame, text="단축 키", width=15)
keybd_act2.pack(side="left", padx=8, pady=2)

keybd_actt3 = Button(keybd_frame, text="키 다운", width=15)
keybd_actt3.pack(side="left", padx=8, pady=2)

keybd_act4 = Button(keybd_frame, text="키 업", width=15)
keybd_act4.pack(side="left", padx=8, pady=2)

# 기타 액션 영역
else_frame = LabelFrame(action_frame, text="기타 액션")
else_frame.pack(fill="x", padx=5, pady=5, ipady=8)

else_act1 = Button(else_frame, text="IF", width=15)
else_act1.pack(side="left", padx=8, pady=2)

else_act2 = Button(else_frame, text="While", width=15)
else_act2.pack(side="left", padx=8, pady=2)

else_actt3 = Button(else_frame, text="반복", width=15)
else_actt3.pack(side="left", padx=8, pady=2)

else_act4 = Button(else_frame, text="대기", width=15)
else_act4.pack(side="left", padx=8, pady=2)


# btn_add_file = Button(action_frame, padx=5, pady=5, width=12, text="파일추가", command=add_file)
# btn_add_file.pack(side="left")

# btn_del_file = Button(action_frame, padx=5, pady=5, width=12, text="선택삭제", command=del_file)
# btn_del_file.pack(side="right")

# # 옵션 프레임
# frame_option = LabelFrame(root, text="옵션")
# frame_option.pack(padx=5, pady=5, ipady=5)

# # 1. 가로 넓이 옵션
# # 가로 넓이 레이블
# lbl_width = Label(frame_option, text="가로넓이", width=8)
# lbl_width.pack(side="left", padx=5, pady=5)

# # 가로 넓이 콤보
# opt_width = ["원본유지", "1024", "800", "640"]
# cmb_width = ttk.Combobox(frame_option, state="readonly", values=opt_width, width=10)
# cmb_width.current(0)
# cmb_width.pack(side="left", padx=5, pady=5)

# # 2. 간격 옵션
# # 간격 옵션 레이블
# lbl_space = Label(frame_option, text="간격", width=8)
# lbl_space.pack(side="left", padx=5, pady=5)

# # 간격 옵션 콤보
# opt_space = ["없음", "좁게", "보통", "넓게"]
# cmb_space = ttk.Combobox(frame_option, state="readonly", values=opt_space, width=10)
# cmb_space.current(0)
# cmb_space.pack(side="left", padx=5, pady=5)

# # 3. 파일 포맷 옵션
# # 파일 포맷 옵션 레이블
# lbl_format = Label(frame_option, text="포맷", width=8)
# lbl_format.pack(side="left", padx=5, pady=5)

# # 파일 포맷 옵션 콤보
# opt_format = ["PNG", "JPG", "BMP"]
# cmb_format = ttk.Combobox(frame_option, state="readonly", values=opt_format, width=10)
# cmb_format.current(0)
# cmb_format.pack(side="left", padx=5, pady=5)

# 리스트 프레임
list_frame = Frame(root)
list_frame.pack(fill="both", padx=5, pady=5)

scrollbar = Scrollbar(list_frame)
scrollbar.pack(side="right", fill="y")

list_file = Listbox(list_frame, selectmode="extended", height=15, yscrollcommand=scrollbar.set)
list_file.pack(side="left", fill="both", expand=True)
scrollbar.config(command=list_file.yview)

# 저장 경로 프레임
path_frame = LabelFrame(root, text="저장경로")
path_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_dest_path = Entry(path_frame)
txt_dest_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4) # 높이 변경

btn_dest_path = Button(path_frame, text="찾아보기", width=10, command=browse_dest_path)
btn_dest_path.pack(side="right", padx=5, pady=5)

# 진행 상황 Progress Bar
frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x", padx=5, pady=5, ipady=5)

p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x", padx=5, pady=5)

# 실행 프레임
frame_run = Frame(root)
frame_run.pack(fill="x", padx=5, pady=5)

btn_close = Button(frame_run, padx=5, pady=5, text="닫기", width=12, command=root.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_start = Button(frame_run, padx=5, pady=5, text="시작", width=12, command=start)
btn_start.pack(side="right", padx=5, pady=5)

root.resizable(False, False)
root.mainloop()