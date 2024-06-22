import autopygui as pag
import time

# 단축키 설정
start_stop_shortcut = 'f2'
video_start = 'space'

# 녹화할 비디오의 시간 설정 (예: 10초)
# video_time = [1584, 1391, 1725]
video_time = [2, 2, 2]

# 경로 설정
button_image_path = ['button_image1.png','button_image2.png','button_image3.png']
button_close_path = 'button_close.png'

# 무한 루프로 계속 실행
for i in range(1, 4):
    button_location = pag.locateOnScreen(button_image_path[i])
    pag.scroll(-1000)
    if button_location is not None:
        pag.click(button_location)
        pag.hotkey(video_start)
        time.sleep(10)
        pag.hotkey(start_stop_shortcut)
        time.sleep(video_time[i])
        pag.hotkey(start_stop_shortcut)
        pag.hotkey(button_close_path)
        time.sleep(10)