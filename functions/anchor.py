
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import WebDriverException,NoSuchElementException 
from webdriver_manager.chrome import ChromeDriverManager
from urllib.error import URLError, HTTPError
import datetime,time, os

options = Options()
#로그숨김
options.add_experimental_option('detach',True) #브라우저 닫힘 방지
options.add_experimental_option('excludeSwitches',['enable-logging']) #usb 오류 
options.add_argument("User-Agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/106.0.0.0 Safari/537.36")

# 페이지 로드가 끝나야지 아래 코드 실행
options.page_load_strategy = 'normal'
#options.page_load_strategy = 'none'
options.add_argument('headless')
service = Service(ChromeDriverManager(path="DRIVER").install())
driver = webdriver.Chrome(service=service,options=options)

#해상도 설정
#driver.maximize_window()
#driver.set_window_size(1920,1200)
driver.set_window_size(360,800)
#driver.set_window_size(1023,850)
wb = load_workbook('files\\anchorList.xlsx')
ws = wb['all']
start = time.time()

for i in range(2, ws.max_row+1) :
    no =  ws['A'+str(i)].value
    anchorId = ws['B'+str(i)].value
    global_url = ws['C'+str(i)].value
    try: 
        driver.get(global_url)
        time.sleep(2)
        
        # 타겟 경로 구함
        folder_name = "anchor_Screenshot"
        current_directory = os.getcwd()
        new_folder_path = os.path.join(current_directory,folder_name)
        if not os.path.exists(new_folder_path):
            os.makedirs(new_folder_path)
        else:
            pass
        #스샷 저장
        driver.save_screenshot(folder_name+'\\'+anchorId+".png")
        time.sleep(2) 
        #driver.execute_script("window.scrollTo(0,22145)")
       
    except NoSuchElementException  :
        print('찾은 엘레먼트가 없음')
    except WebDriverException:
        print('Driver ERROR')
    except HTTPError as e : 
        print("ERROR"+str(e.code))
    except URLError as e :
        print("reason"+str(e.reason)) 
    
driver.close()
wb.save('Result_AnchorID'+'.xlsx')
sec = time.time()-start # 종료
times = str(datetime.timedelta(seconds=sec))
short = times.split(".")[0] #초단위 

print('완료시간',f"{short} sec")
