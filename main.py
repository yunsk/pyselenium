################################
# python -m unittest -v main.py 입력하여 실행
################################

# unittest 모듈 불러오기
import unittest
import HtmlTestRunner
# webdriver 모듈을 selenium 패키지에서 불러오기
from selenium import webdriver
# Options 클래스를 selenium.webdriver.chrome.options 모듈에서 불러오기
from selenium.webdriver.chrome.options import Options
# 엑셀 사용을 위한 라이브러리
from openpyxl import load_workbook, Workbook
# By 사용을 위한 라이브러리
from selenium.webdriver.common.by import By
# 키입력을 위한 라이브러리
from selenium.webdriver.common.keys import Keys

import functions.anchor, functions.compare, functions.kvCTA
# 커스텀 모듈인 functions 불러오기

#import functions.function1_copyApplied, functions.function2_highlightsPageReport, functions.function3_accessoriesPageReport, functions.function4_disclaimerReport, functions.function5_taggingReport,functions.function6_anchorListReport
class MainTest(unittest.TestCase):
    
    # 테스트를 위한 설정
    def setUp(self) -> None:
        options = webdriver.ChromeOptions()
        options.add_experimental_option("detach", True)
        options.add_experimental_option("excludeSwitches", ["enable-logging"])
        
        # 창 열지 않고 실행 원할 시 활성화
        # options.add_argument("headless")

        self.driver = webdriver.Chrome(options=options)

    
    # 테스트 종료 후
    def tearDown(self) -> None:
        self.driver.quit()


    # 앵커 스크린샷 테스트 실행
    @unittest.skip("This is a skipped test.")
    def test_anchor_test(self):
        anchorTestfucntion = functions.anchor.MainFunction(self.driver)
        excelName = 'files\\anchorList.xlsx'
        sheetname = 'all'
        windowsize = 360,800
        #1920,1200
        #1023,850
        anchorTestfucntion.anchorTest(excelName,sheetname,windowsize)

    # Compare 기본 동작성 테스트 실행 
    @unittest.skip("This is a skipped test.")
    def test_compare_test(self):
        compareTestfunction = functions.compare.MainFunction(self.driver)
        excelName = 'files\\compare_vari_format.xlsx'
        windowsize = 1920,1080
        compareTestfunction.CompareTest(excelName, windowsize)
    
    # KV CTA 존재 유무 테스트 실행 
    def test_kvCTA_test(self):
        anchor = functions.kvCTA.MainFunction(self.driver)
        excelName = 'files\\(GRO)KV_CTA_D3.xlsx'
        anchor.kvCTATest(excelName)



if __name__ == "__main__":
    reportFolder= "TestResult"
    unittest.main(testRunner=HtmlTestRunner.HTMLTestRunner(output=reportFolder))