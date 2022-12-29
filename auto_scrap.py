# 관리자 권한으로 Pycharm실행 필요(윈도우에서 직접 실행 필요)

# 필요한 라이브러리 import
import xlwings as xw
import pandas as pd
import pyautogui as pg
import subprocess
import pyperclip
import sys
import os
import webbrowser
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from win32comext.shell import shell # pywin32 라이브러리 설치 필요



# 관리자 권한으로 실행하는 코드이나, 적용이 잘 안됨
print(sys.argv[-1])
if sys.argv[-1] != 'asadmin':
    script = os.path.abspath(sys.argv[0])
    params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
    shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
    sys.exit()

# 스크랩 마스터 실행
subprocess.call(["C:\Program Files (x86)\Dahami\ScrapMaster5\SM5UpdateProgram.exe"])
pg.sleep(10)
pg.click(717, 531)


# 검색 시작일자 및 종료일자 지정(당일 기준 하루 전 까지 세팅)
from datetime import datetime, timedelta
end = datetime.now().date()
start = end - timedelta(days=1)
start, end = str(start), str(end)


pg.doubleClick(396, 26)
pg.click(145, 62)
pg.click(89, 205)
pg.drag(-70,0, 1, button='left')
pg.hotkey('backspace')
pyperclip.copy(start)
pg.hotkey('ctrl', 'v')
pg.hotkey('tab', interval=0.5)
pyperclip.copy(end)
pg.hotkey('ctrl', 'v')

pg.click(158, 360)
pg.hotkey('ctrl', 'a')
pg.hotkey('backspace')

search_list = ['양재 하나로', '농협유통']


pyperclip.copy(search_list[0])
pg.hotkey('ctrl', 'v', inteval=1)
pg.doubleClick(119, 537)

pg.position()
pg.click(462, 167, interval=1)
pg.click(299, 997, interval=1)
pg.click(596, 997, interval=1)
pg.click(1234, 565, interval=1)
pg.hotkey('enter', interval=0.5)


# 작업할 엑셀 파일 경로 지정
work_path = f"C:/Users/DELL9020/Desktop/검색기사_{end}.xlsx"

# 필요한 워크시트 적용
wb = xw.Book(work_path)
sheet = wb.sheets('기사 목록')
print(sheet.name)

# 필요한 범위 선택 및 DF타입으로 변환
_li = sheet.range('A1').expand('table').value
_df = pd.DataFrame(_li)

# 매체명을 기준으로 중복 값 제거
_df.drop_duplicates([1], keep='first', inplace=True)

# 접속이 필요한 URL주소만 추출
url_li = list(_df[7])
url_li[1:]

# chrome으로 연결해서 url접속 확인
chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"
webbrowser.get(chrome_path).open(url_li[1])

# url접속 후, 웹페이지를 전체 캡쳐

# 옵션 지정
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")

# 크롬으로 오픈 후, 해당 파일 저장
for i in range(1, 10):
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get(url_li[i])
    width = driver.execute_script("return document.body.scrollWidth")
    height = driver.execute_script("return document.body.scrollHeight")
    driver.set_window_size(width, height)
    driver.save_screenshot(f"C:/Users/DELL9020/Desktop/test/test{i}.png")