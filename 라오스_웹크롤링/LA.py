import os # 파일 생성 시에 필요한 모듈

import time # 시간 데이터를 처리하기 위해 사용되는 모듈
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
driverpath='C:\datapre\chromedriver.exe'
driver = webdriver.Chrome(driverpath)

# 페이지 주소
page_address='https://www.laotradeportal.gov.la/index.php?r=tradeInfo/listAll'
driver.implicitly_wait(3)
driver.get(page_address)

time.sleep(2)
# dropbox 선택
driver.find_element_by_xpath('//*[@id="bs-example-navbar-collapse-1"]/div[2]/div/a').click()

# dropbox 내 영어 선택
driver.find_element_by_xpath('//*[@id="bs-example-navbar-collapse-1"]/div[2]/div/div/div/ul/li[1]/a').click()

# hscode 선택
driver.find_element_by_xpath('//*[@id="CommoditySearchForm_searchType_0"]').click()

# 모든 목록을 열기
for i in range(1, 98):
    if i==77:
        continue
    else:
        driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div[1]/div/ul/li['+str(i)+']/div').click()
        time.sleep(5)
        print(i)
print("success")

# 1 ~ 97까지의 개수에 대한 배열을 만들기, 검색에 사용할 예정
farr=[]
for i in range(1, 98):
    if i==77:
        farr.append(0)
        time.sleep(3)
        print(i)
    else:
        flen=driver.find_elements_by_xpath('/html/body/div/div[2]/div[2]/div[1]/div/ul/li['+str(i)+']/ul/li')
        farr.append(len(flen))
        time.sleep(3)
        print(str(i)+" : "+str(len(flen)))
print("end")

# 빠른 작업을 위해 위 두 과정을 거치지 않고 검색에 사용할 예정인 배열을 바로 생성
farr=[6, 10, 8, 10, 9, 4, 14, 14, 10, 8, 9, 14, 2, 2, 21, 5, 4, 6, 5, 9, 6, 9, 9, 3, 29, 21, 16, 51, 42, 6, 5, 15, 7, 7, 7, 6, 7, 26, 26, 17,
11, 5, 21, 4, 2, 7, 22, 11, 7, 13, 12, 10, 8, 16, 9, 5, 11, 11, 6, 17, 17, 10, 6, 6, 3, 4, 15, 14, 19, 18, 29, 26, 16, 8, 16, 0, 4, 6, 4,
13, 15, 11, 86, 46, 9, 16, 5, 8, 32, 14, 7, 7, 6, 6, 20, 6]

# 검색어의 형태를 지정, 검색할 수 있게 만들어주는 함수
def chk(l, k):
    if k < 10:
        if l < 10:
            fname='0'+str(l)+'0'+str(k)
        else:
            fname=str(l)+'0'+str(k)
    else:
        if l < 10:
            fname='0'+str(l)+str(k)
        else:
            fname=str(l)+str(k)
    return fname

# 파일 생성 시에 요구되는 경로을 만들어주는 함수
def chap(l):
    if l < 10:
        mname='0'+str(l)
    else:
        mname=str(l)
    return mname

# 빈 데이터 프레임 생성
df=pd.DataFrame(columns=['Country Group', 'Tariff Rate', 'Unit'])

# 데이터 추출, 시간 문제로 완전하게 작동시키지 못하여 수정 필요한 부분
for l in range(1, 98):
    if l == 77:
        continue
    for k in range(1, farr[l - 1] + 1):
        fname = chk(l, k)
        element = driver.find_element_by_name('CommoditySearchForm[searchValue]')
        # 존재하지 않는 페이지에 대한 처리
        if fname == '0503' or fname == '0509' or fname == '1519' or fname == '1402' or fname == '1403' or fname == '2527' or fname == '2838' or fname == '2851' or fname == '2851' or fname == '2851' or fname == '4108' or fname == '4109' or fname == '4110' or fname == '4111' or fname == '4204' or fname == '4815' or fname == '5304' or fname == '6503' or fname == '7012' or fname == '7414' or fname == '7416' or fname == '7417' or fname == '7803' or fname == '7805' or fname == '7906' or fname == '8004' or fname == '8005' or fname == '8006' or fname == '8485' or fname == '8520' or fname == '8524' or fname == '9009' or fname == '9203' or fname == '9204' or fname == '9501' or fname == '9502':
            farr[l - 1] = farr[l - 1] + 1
            continue
        element.send_keys(fname)
        # 검색 버튼 누르기
        driver.find_element_by_xpath('//*[@id="commodity-search-form"]/div[3]/input').click()
        # 리스트의 항목 개수
        elements = driver.find_elements_by_css_selector('#yw0 > li')

        # 리스트 내의 항목에 대한 데이터 추출
        for i in range(0, len(elements)):
            time.sleep(2)
            driver.find_element_by_xpath('/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/div').click()
            time.sleep(2)

            lists = driver.find_elements_by_xpath('/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/ul/li')
            if not len(lists) == 1:
                # 또다른 분류가 존재할 경우
                for j in range(0, len(lists)):
                    driver.find_element_by_xpath(
                        '/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/span/span/a').click()
                    time.sleep(2)
                    html = driver.page_source
                    bsoup = BeautifulSoup(html, 'html.parser')

                    # 모든 table 찾기
                    table = bsoup.find_all('table')

                    # 테이블이 존재하지 않는 경우에 대한 예외처리
                    if not table:
                        time.sleep(2)
                        print(str(l) + ' ' + str(k) + ' ' + str(i) + ' ' + str(j) + ' ' + "no tables")
                        driver.back()
                        if (j + 1) < len(lists):
                            driver.find_element_by_xpath(
                                '/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/div').click()
                            time.sleep(2)
                    else:
                        # 그 중 추출할 table 찾기
                        ta_lst = pd.read_html(str(table))
                        ta_df = ta_lst[0]
                        ta_df = ta_df.drop(columns=['Group Description', 'Activity', 'Valid From', 'Valid To'])

                        # table 합치기
                        df = pd.concat([df, ta_df], ignore_index=True)
                        time.sleep(2)
                        print(str(l) + ' ' + str(k) + ' ' + str(i) + ' ' + str(j))

                        # 뒤로 가기
                        driver.back()
                        if (j + 1) < len(lists):
                            driver.find_element_by_xpath(
                                '/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/div').click()
                            time.sleep(2)
            else:
                # 하나만 존재할 경우
                driver.find_element_by_xpath(
                    '/html/body/div/div[2]/div[2]/div/ul/li[' + str(i + 1) + ']/ul/li/span/span/a').click()
                time.sleep(2)
                # 데이터 추출
                html = driver.page_source
                bsoup = BeautifulSoup(html, 'html.parser')

                # 모든 table 찾기
                table = bsoup.find_all('table')

                # 테이블이 존재하지 않는 경우에 대한 예외처리
                if not table:
                    time.sleep(2)
                    print(str(l) + ' ' + str(k) + ' ' + str(i) + ' ' + "no tables")
                    # 뒤로 가기
                    driver.back()
                else:
                    # 그 중 추출할 table 찾기
                    ta_lst = pd.read_html(str(table))
                    ta_df = ta_lst[0]
                    ta_df = ta_df.drop(columns=['Group Description', 'Activity', 'Valid From', 'Valid To'])

                    # table 합치기
                    df = pd.concat([df, ta_df], ignore_index=True)
                    time.sleep(2)
                    print(str(l) + ' ' + str(k) + ' ' + str(i))

                    # 뒤로 가기
                    driver.back()

        time.sleep(3)
        print(str(l) + ' ' + str(k))

excel_path="C:\datapre\LA\LA.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()