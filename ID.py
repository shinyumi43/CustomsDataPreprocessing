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

# 항상 최신버전의 chromedriver를 자동으로 사용할 수 있도록 하는 라이브러리
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

# 버전 업데이트 필요없이 항상 최신 버전의 구글 드라이브를 사용할 수 있도록 함
driver = webdriver.Chrome(ChromeDriverManager().install())
driver.implicitly_wait(3)

# 페이지 주소
page_address='https://insw.go.id/intr'
driver.implicitly_wait(3)
driver.get(page_address)

# 최초 화면을 벗어나는 용도로 활용, 전부 중간 화면에서 검색 진행
element=driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div/div[1]/div[1]/div/div/div[3]/div/div/div/input')
# 최초 화면에서 검색할 내용 입력
element.send_keys('0101')

# 최초 화면에서 검색 버튼 클릭
driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div/div[1]/div[1]/div/div/div[3]/div/div/div/div/span').click()

# 엑셀 내의 hs code를 읽어와서 검색 구현
exc_df = pd.read_excel('C:\\datapre\\ID\\hscode.xlsx')

# 엑셀 내의 칼럼명
exc_col1='HS Code'
exc_col2='Description of Goods'

# 기본 관세율 테이블
col1='HS Code'
col2='Description of Goods'
col3='Uraian Barang'
col4='Nama'
col5='MFN (Most Favored Nation)'
col_list=[col1, col2, col3, col4, col5]

# FTA 협정세율 테이블
c1='Tarif Nama'
c2='Tarif Preferensi'
c_lists=[c1, c2]

# 최종 데이터프레임 생성
ta_df=pd.DataFrame(columns=[col1, col2, col3, col4, col5, c1, c2])

# 엑셀 내의 전체 hs code를 탐색
for i in range(0, len(exc_df)):
    # 엑셀 내에서 hs code 반환
    exc_txt = str(exc_df[exc_col1][i])
    # hs code가 없거나 최종 hs code가 아닌 경우에 품목 부분만 엑셀에서 읽어오고, 나머지 부분은 ''으로 처리 진행
    # 8528.69.10은 유일하게 관세율 정보가 존재하지 않아서 예외적으로 처리를 진행, 관세율 부분을 공란으로 지정
    if (pd.isna(exc_txt)) or (len(exc_txt) != 10) or (exc_txt=='8528.69.10'):
        # 임시 데이터프레임 생성
        df = pd.DataFrame(columns=[col1, col2, col3, col4, col5, c1, c2])
        # 임시 배열 생성
        df_forcat = {col1: '', col2: '', col3: '', col4: '', col5: '', c1: '', c2: ''}
        # 엑셀에서 품목 부분을 읽어오기
        if exc_txt == 'nan':
            df_forcat[col1] = ""
        else:
            df_forcat[col1] = exc_txt
        df_forcat[col2] = str(exc_df[exc_col2][i])
        # 데이터프레임에 결합
        df = df.append(df_forcat, ignore_index=True)

        # 최종 데이터프레임에 결합
        ta_df = pd.concat([ta_df, df], ignore_index=True)

        # 데이터 결합이 완료됨을 출력
        print(exc_txt)
    # 최종 hs code가 있는 경우에 웹크롤링 진행
    else:
        # 검색 형태로 변환
        txt = str(exc_txt[0:4]) + str(exc_txt[5:7]) + str(exc_txt[8:])

        # 중간 화면에서 검색할 내용 입력
        elems = driver.find_element_by_xpath('/html/body/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div/input')
        elems.send_keys(txt)
        # 중간 화면에서 검색 버튼 클릭
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[1]/div[1]/div/div/span').click()
        driver.implicitly_wait(3)
        # 중간 화면에서 이전 검색 내용 지우기
        elems.clear()
        # 중간 화면에서 hs code 입력 후, detail 버튼 클릭
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[1]/div/div/table/tbody/tr/td[1]/button').click()
        driver.implicitly_wait(3)

        # 일반 관세율을 담을 데이터프레임 생성
        df1 = pd.DataFrame(columns=col_list)
        # 원문 품목명
        otxt = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[1]/div/div/table/tbody/tr/td[2]/p[1]').text
        # 영문 품목명
        etxt = driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[1]/div/div[1]/div/div/table/tbody/tr/td[2]/p[2]').text

        # 일반 관세율 가져오기
        for j in range(1, 10):
            df_tem = {col1: '', col2: '', col3: '', col4: '', col5: ''}
            nama = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div[2]/div/div/div/div[' + str(
                    j + 1) + ']/p[1]').text
            mfn = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[1]/div[2]/div/div/div/div[' + str(
                    j + 1) + ']/p[3]').text
            # 첫 번째일 경우, hs code와 품목명을 담기
            if j == 1:
                df_tem[col1] = exc_txt
                df_tem[col2] = etxt
                df_tem[col3] = otxt
            else:
                df_tem[col1] = ""
                df_tem[col2] = ""
                df_tem[col3] = ""
            df_tem[col4] = nama
            df_tem[col5] = mfn
            df1 = df1.append(df_tem, ignore_index=True)

        # FTA 협정세율 가져오기
        # FTA 협정세율을 확인하기 위해서는 스크롤을 내려서 선택이 필요
        prev_height = driver.execute_script("return document.body.scrollHeight")
        # 웹페이지 맨 아래까지 무한 스크롤
        while True:
            # 스크롤을 화면 가장 아래로 내린다
            driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")

            # 페이지 로딩 대기
            time.sleep(2)

            # 현재 문서 높이를 가져와서 저장
            curr_height = driver.execute_script("return document.body.scrollHeight")

            if (curr_height == prev_height):
                break
            else:
                prev_height = browser.execute_script("return document.body.scrollHeight")

        # FTA 협정세율 선택
        driver.find_element_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[3]/div[1]/div').click()
        driver.implicitly_wait(3)

        # FTA 협정세율을 담을 데이터프레임 생성
        df2 = pd.DataFrame(columns=c_lists)

        # FTA 협정세율 컬럼의 개수를 반환하기 위해 길이 측정
        hlists = driver.find_elements_by_xpath(
            '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[3]/div[2]/div/div/div')

        for k in range(0, len(hlists)):
            # 임시 배열 생성
            df_mk = {c1: '', c2: ''}
            # 관세율 칼럼명 삽입
            df_mk[c1] = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[3]/div[2]/div/div/div[' + str(
                    k + 1) + ']/ul/li').text
            # 컬럼별로 존재하는 개수를 세고, 가장 끝 값만 반환
            lists = driver.find_elements_by_xpath(
                '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[3]/div[2]/div/div/div[' + str(
                    k + 1) + ']/div')
            # 관세율 칼럼명에 따른 관세율 정보 삽입
            df_mk[c2] = driver.find_element_by_xpath(
                '/html/body/div[1]/div/div/div[2]/div/div/div/div/div[2]/div[2]/div/div[3]/div[2]/div/div/div[' + str(
                    k + 1) + ']/div[' + str(len(lists)) + ']/ul/li').text
            df2 = df2.append(df_mk, ignore_index=True)

        # 기본 관세율과 FTA 협정세율이 담긴 관세율 테이블을 결합
        ta = pd.concat([df1, df2], axis=1)
        ta = ta.fillna('')

        # 최종 데이터프레임에 결합
        ta_df = pd.concat([ta_df, ta], ignore_index=True)

        # 데이터 결합이 완료됨을 출력
        print(exc_txt)

# 엑셀에 출력
excel_path="C:\\datapre\\ID\\ID.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
ta_df.to_excel(writer, index=False)
writer.save()