# Chapter 부분만 추출 진행하고, 홈페이지에 기재된 데이터와 엑셀 상에 결합하여 오름차순 정리
import os

import time
import pandas as pd

from selenium.webdriver.common.by import By # 태그 존재 여부 확인 가능
from selenium.webdriver.support.ui import WebDriverWait   # 해당 태그를 기다림
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException    # 태그가 없는 예외 처리

from bs4 import BeautifulSoup
# selenium은 가상 브라우저를 활용하여 크롤링을 진행
from selenium import webdriver

# 드롭 다운 메뉴값을 선택할 수 있는 라이브러리
from selenium.webdriver.support.ui import Select

# 엑셀 함수 작업 자동화
import openpyxl as op
import requests

# 드라이브 생성, chrome driver 93.0.4577.63
driverpath='C:\datapare\chromedriver.exe'
driver = webdriver.Chrome(driverpath)

# 페이지 주소
page_address='https://www.customs.gov.sa/en/customsTariffSearch'
driver.implicitly_wait(3)
driver.get(page_address)

# 홈페이지에 데이터가 추출되어 있으므로 가장 큰 chapter 데이터 추출 시행 후 엑셀 내 오름차순 정렬
# chapter 열기
chapr=[1,7,17,19,29,33,45,48,52,56,60,75,80,84,86,99,102,107,111,113,117]
for i in chapr:
    driver.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/main/div[2]/div[2]/div/div[5]/div/div[2]/div[3]/div/div/ul/li['+str(i)+']').click()

# 열 지정하기
col1='HS Code'
col2='Item'
col_list=[col1, col2]
# 최종 데이터프레임 생성하기
df_ta=pd.DataFrame(columns=col_list)

# 임시 배열 생성
df_tem={col1:'', col2:''}
for i in range(1, 120):
    if i in chapr:
        continue
    else:
        # hscode 담기
        hs=driver.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/main/div[2]/div[2]/div/div[5]/div/div[2]/div[3]/div/div/ul/li['+str(i)+']/div/div/div/div/span[2]').text
        # items 담기
        it=driver.find_element_by_xpath('/html/body/div/div/div[2]/div[2]/main/div[2]/div[2]/div/div[5]/div/div[2]/div[3]/div/div/ul/li['+str(i)+']/div/div/div/div/span[4]').text
        # hscode와 items를 결합한 배열을 최종 데이터프레임에 결합하기
        df_tem[col1]=hs
        df_tem[col2]=it
        df_ta=df_ta.append(df_tem, ignore_index=True)

# 생성할 excel 경로
excel_path="C:\\datapare\\SA\\CHAP.xlsx"
# xlsxwriter 엔진으로 pandas writer 객체 만들기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
# dataframe을 xlsx에 쓰기
df_ta.to_excel(writer, index=False)
# pandas excel writer을 닫고, 엑셀 파일을 출력
writer.save()
