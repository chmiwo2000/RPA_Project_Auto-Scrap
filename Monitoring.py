import threading
import tkinter.messagebox as msgbox
import tkinter.filedialog as filedialog
from tkinter import *
import cv2
import pyautogui as pg
import numpy as np
import keyboard
import os

def record():
    resolution = (1920, 1080)
    codec = cv2.VideoWriter_fourcc(*'XVID')
    filename = f'{file_name_entry.get()}.avi'
    location = save_dir_entry.get()
    fps = 10.0
    out = cv2.VideoWriter(location+'/'+filename, codec, fps, resolution)

    while True:
        img = pg.screenshot()
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        out.write(frame)

        if cv2.waitKey(1) == ord('q'):
            break

    out.release()
    cv2.destroyAllWindows()

record()


recording = False

def record_startstop():
    global recording
    recording = not recording

    if save_dir_entry.get() == '':
        msgbox.showerror('경고','저장 경로를 설정해주세요.')
        recording = not recording
    elif file_name_entry.get() == '':
        msgbox.showerror('경고', '파일 이름을 저장해주세요.')

    elif recording:
        msgbox.showinfo('알림', '녹화를 시작합니다.')
        btn_stop.config(state='active')
        btn_start.config(state='disabled')
        thread = threading.Thread(target=record)
        thread.setDaemon(True)
        thread.start()

    elif recording == False:
        keyboard.press_and_release('q')
        btn_start.config(state='active')
        btn_stop.config(state='disabled')
        msgbox.showinfo('알림', '녹화가 완료되었습니다.')



def save_dir():
    folder_selected = filedialog.askdirectory()
    if folder_selected == '':
        return
    save_dir_entry.delete(0, END)
    save_dir_entry.insert(0, folder_selected)


# 기본적인 화면 세팅
root = Tk() # Tk()는 인터페이스 창을 만들어주는 역할
root.title('화면 녹화')
root.resizable(False, False)

# 녹화 시작, 중지 버튼 프레임
btn_frame = Frame(root)
btn_frame.pack(fill='x', pady=5, padx=5) # pack는 비유하자면 '그리기' 기능이라고 보면 된다.
# 녹화 시작, 중지 버튼
btn_start = Button(btn_frame, text='녹화 시작', width=10, command=record_startstop)
btn_start.pack(side='left', padx=5, pady=5, expand=True)
btn_stop = Button(btn_frame, text='녹화 중지', width=10, command=record_startstop)
btn_stop.pack(side='right', padx=5, pady=5, expand=True)

# 저장 경로 프레임
save_dir_frame = LabelFrame(root, text='저장 경로')
save_dir_frame.pack(padx=5, pady=5, fill='x')
# 저장경로 찾아보기 버튼
btn_save_dir = Button(save_dir_frame, text='찾아보기...', width=10, command=save_dir)
btn_save_dir.pack(padx=5, pady=5, side='left')
# 저장경로 엔트리
save_dir_entry = Entry(save_dir_frame)
save_dir_entry.pack(side='right', fill='x', padx=5, pady=5)


# 파일명 프레임
file_name_frame = Frame(root)
file_name_frame.pack(padx=5, pady=5, fill='x')
# 파일명 입력창과 설명칸
file_name_label = Label(file_name_frame, text='저장할 파일명 : ')
file_name_label.pack(side='left')
file_name_entry = Entry(file_name_frame)
file_name_entry.pack(padx=5, pady=5, fill='x')


root.mainloop() # 화면이 계속 돌아가게 해주는 기능

