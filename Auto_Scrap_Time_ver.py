import schedule
import time
import cv2
import pyautogui as pg
import numpy as np

# 프로그램이 어떻게 돌아가다가 오류가 발생하는지 알 수 있는 화면 녹화 코드
def record():
    resolution = (1920, 1080)
    codec = cv2.VideoWriter_fourcc(*'XVID')
    filename = 'verify.avi'
    location = 'C:/Users/DELL9020/Desktop/화면 녹화 테스트'
    fps = 10.0
    out = cv2.VideoWriter(location+'/'+filename, codec, fps, resolution)

    while True:
        img = pg.screenshot()
        frame = np.array(img)
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        out.write(frame)

        if cv2.waitKey(1) == ord('q'):
            out.release()
            cv2.destroyAllWindows()
            break


# 관리자 권한으로 Pycharm실행 필요(윈도우에서 직접 실행 필요)

def auto_scrap():
    # 필요한 라이브러리 import
    import xlwings as xw
    import pandas as pd
    import pyautogui as pg
    import subprocess
    import pyperclip
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from webdriver_manager.chrome import ChromeDriverManager
    from win32comext.shell import shell # pywin32 라이브러리 설치 필요

# 필요한 몇 가지 자체 함수 정의
    def confirm_image(img):
        if pg.locateOnScreen(f'{img}.png', confidence=0.8) is None:
            return False
        else:
            return True


    def image_click(name, sec, conf=0.8):
        for i in range(1, sec, 1):
            pg.sleep(1)
            print(f'이미지 검색 중 {i}번째 시도 입니다.')
            if confirm_image(name):
                loc = pg.locateOnScreen(f'{name}.png', confidence=f'{conf}')
                click_loc = pg.center(loc)
                pg.moveTo(click_loc)
                pg.click()
                break
            elif i == sec:
                print(f'{name} 이미지를 찾지 못했습니다. 프로그램을 종료합니다.')
                exit()
            else:
                pass


    # 관리자 권한으로 실행하는 코드이나, 적용이 잘 안됨
    # print(sys.argv[-1])
    # if sys.argv[-1] != 'asadmin':
    #     script = os.path.abspath(sys.argv[0])
    #     params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
    #     shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
    #     sys.exit()

    # 스크랩 마스터 실행
    subprocess.call(["C:\Program Files (x86)\Dahami\ScrapMaster5\SM5UpdateProgram.exe"])
    image_click('login', 10)
    pg.click(717, 531)
    pg.sleep(10)

    # 공지사항 업데이트 여부 확인 후, Ture면 창 닫고 False면 진행
    if confirm_image('notice') == True:
        image_click('close', 5)
    else:
        pass

    # 검색 시작일자 및 종료일자 지정(당일 기준 하루 전 까지 세팅)
    from datetime import datetime, timedelta
    end = datetime.now().date()
    start = end - timedelta(days=1)
    start, end = str(start), str(end)

    # 검색일자 지정
    pg.doubleClick(396, 26)
    pg.position()
    pg.click(145, 62)
    pg.click(207, 171)
    pg.click(89, 205)
    pg.drag(-70,0, 1, button='left')
    pg.hotkey('backspace')
    pyperclip.copy(start)
    pg.hotkey('ctrl', 'v')
    pg.hotkey('tab', interval=0.5)
    pyperclip.copy(end)
    pg.hotkey('ctrl', 'v')


    # 검색어 타이핑 공간 마련
    search_list = ['농협유통', '양재 하나로']

    for n in range(0, len(search_list)):
        pg.click(396, 26)
        pg.click(106, 376)
        pg.click(158, 360)
        pg.hotkey('ctrl', 'a')
        pg.hotkey('backspace')
        pyperclip.copy(search_list[n])
        pg.hotkey('ctrl', 'v', inteval=1)
        pg.doubleClick(119, 537)
        pg.sleep(10)

    # 엑셀 저장
        pg.click(462, 167, interval=1)
        pg.click(299, 997, interval=1)
        pg.click(596, 997, interval=1)
        pyperclip.copy(f'검색기사_{end}_{search_list[n]}')
        pg.hotkey('ctrl', 'v')
        pg.click(1234, 565, interval=1)
        pg.hotkey('enter', interval=0.5)
        pg.sleep(10)

    # 작업할 엑셀 파일 경로 지정
        work_path = f"C:/Users/DELL9020/Desktop/검색기사_{end}_{search_list[n]}.xlsx"

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
        url_li = list(filter(None, url_li))
        url_li[1:]

    # chrome으로 연결해서 url접속 확인
        chrome_path = "C:/Program Files/Google/Chrome/Application/chrome.exe %s"

    # url접속 후, 웹페이지를 전체 캡쳐

    # 옵션 지정
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--headless")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")

    # 크롬으로 오픈 후, 해당 파일 저장
        for i in range(1, len(url_li)):  # len(url_li) <- 테스트 끝나면 다시 돌려놓기
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), chrome_options=chrome_options)
            # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
            driver.get(url_li[i])
            width = driver.execute_script("return document.body.scrollWidth")
            height = driver.execute_script("return document.body.scrollHeight")
            driver.set_window_size(width, height)
            driver.save_screenshot(f"C:/Users/DELL9020/Desktop/test/{search_list[n]}test{i}.png")
            print(i)

        pg.sleep(2)


schedule.every().day.at("06:30:00").do(auto_scrap)
#schedule.every().day.at("17:40:10").do(record)
schedule.every().day.at("06:50:00").do(exit)


while True:
    schedule.run_pending()
    time.sleep(1)
