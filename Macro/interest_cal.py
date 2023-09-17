from tkinter import *
import pyautogui

##################################################<< 창 기본설정 (제목, 크기) >>////////////////////
size = pyautogui.size()
monitor_width = size[0]
monitor_height = size[1]

main_witdh = 640
main_height = 480

root = Tk()
root.title("간단 복리 계산기")
root.geometry(str(main_witdh)+"x"+str(main_height)+"+"+str(int((monitor_width-main_witdh)/2))+"+"+str(int((monitor_height-main_height)/2)))
root.resizable(0, 0)

##################################################<< 함수 설정 >>////////////////////
def calc():
    pass
    # st_mo = int(First_e.get())
    # times = int(Sencon_e.get())
    # goal_it = int(Third_e.get()) / 100

    # Last_e.insert(0, str(st_mo * pow((1 + goal_it / 1),1 * times)))



##################################################<< 기능 설정 >>////////////////////
#| Header
title_label = Label(root, text="간단 복리 계산기", width=40, height=2, relief="solid", bd=5, font=('맑은 고딕',18,'bold'))
title_label.grid(row=0, column=0, columnspan=3, sticky=EW)

#| Frm1
lab11 = Label(root, text="투자 원금", font=('맑은 고딕',12,'bold'))
ent1 = Entry(root, width=30, font=('맑은 고딕',12,'bold'))
lab12 = Label(root, text="원", font=('맑은 고딕',12,'bold'))
lab11.grid(row=1, column=0, pady=25)
ent1.grid(row=1, column=1, pady=25)
lab12.grid(row=1, column=2, pady=25)
# First_e.insert(0, "투자 원금을 입력하시오")

#| Frm2
lab21 = Label(root, text="계산 기간", font=('맑은 고딕',12,'bold'))
ent2 = Entry(root, width=30, font=('맑은 고딕',12,'bold'))
lab22 = Label(root, text="일", font=('맑은 고딕',12,'bold'))
lab21.grid(row=2, column=0, pady=25)
ent2.grid(row=2, column=1, pady=25)
lab22.grid(row=2, column=2, pady=25)

#| Frm2
lab31 = Label(root, text="목표 수익", font=('맑은 고딕',12,'bold'))
ent3 = Entry(root, width=30, font=('맑은 고딕',12,'bold'))
lab32 = Label(root, text="%", font=('맑은 고딕',12,'bold'))
lab31.grid(row=3, column=0, pady=25)
ent3.grid(row=3, column=1, pady=25)
lab32.grid(row=3, column=2, pady=25)

#| Footer


root.mainloop()