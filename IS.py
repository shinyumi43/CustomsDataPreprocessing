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

# 드라이브 생성, chrome driver 93.0.4577.63, 개별 설치가 필요
driverpath='C:\datapare\chromedriver.exe'
driver = webdriver.Chrome(driverpath)

# 페이지 주소
page_address='https://shaarolami-query.customs.mof.gov.il/CustomspilotWeb/en/CustomsBook/Import/CustomsTaarifEntry'
driver.implicitly_wait(3)
driver.get(page_address)

# 열을 지정
col1='Hs code'
col2='Items'
col3='Agreement Name'
col4='Customs rate'
col5='Tax rate'
col_list=[col1, col2, col3, col4, col5]


# 하나의 hs에 들어가서 table 추출해오는 과정
def tabu():
    # 페이지 불러오기
    html = driver.page_source
    bsoup = BeautifulSoup(html, 'html.parser')

    # 모든 table 찾기
    table = bsoup.find_all('table')

    # 테이블 두 가지 받아오기
    if not table:
        return False, None
    else:
        ta_lst = pd.read_html(str(table))
        cu_df = ta_lst[0]
        pu_df = ta_lst[1]

    # 필요없는 테이블에 대한 예외처리
    if cu_df.columns[1] != 'Agreement Name':
        # 테이블 추출 진행하지 않고, return 0 반환
        return False, None
    else:
        # 테이블 추출 진행
        # custom tax 추출
        cu_df = cu_df.drop(
            columns=['Unnamed: 0', 'Customs tariff within quota', 'Quota number', 'Statistical measurement unit',
                     'Valid from', 'Valid until'])
        # Tax Free에 대해 0%로 변경
        for i in range(0, len(cu_df)):
            if cu_df[col4][i] == 'Tax Free':
                cu_df[col4][i] = "0%"

        # purchase tax 추출
        pu_df = pu_df.drop(
            columns=['Unnamed: 0', 'Measurement unit', 'Optional tax addition', 'Valid from', 'Valid until'])
        # Tax Free에 대해 0%로 변경
        if pu_df[col5][0] == 'Tax Free':
            pu_df[col5][0] = "0%"

        # custom tax와 purchase tax 결합
        ta_df = pd.concat([cu_df, pu_df], axis=1)
        ta_df = ta_df.fillna('')

    return True, ta_df


# 순차적 접근과 데이터 추출 결합(chrome 업데이트로 인한 코드 재수정)
# 재귀적으로 구현하고자 했으나 불가피하게 직관적으로 다중 반복문을 활용
# 너무 많은 반복문으로 인해 제대로 동작하지 않을 수 있어서
# Section과 Chapter는 마우스로 직접 클릭하고(시간 단축), Chapter 내 내용부터 자동
# 첫 번째 ul부터 접근하기 위해 축약
driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

# 지속적으로 넣어줄 문자열
ulli = '#customsItemClassificationTreeContainer'
li = ' > ul > li'

# 4단위 기준의 목록에 순차적으로 접근
flists = driver.find_elements_by_css_selector(ulli + li)

# 최종 데이터프레임 생성
ta_df = pd.DataFrame(columns=col_list)

# 순차적인 접근을 구현, 우선적으로 뽑아낸 뒤에 엑셀 상에서 정제
for i in range(0, len(flists)):
    # 계속해서 바뀌는 임시 데이터공간이므로 내부에 선언
    ta = pd.DataFrame(columns=[col1, col2])

    # 임시적으로 데이터를 저장할 공간 생성
    df_tem = {col1: '', col2: ''}

    # 첫 번째 ul에서 hs code 선택 및 데이터 추출
    df_tem[col1] = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
        i + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text
    df_tem[col2] = driver.find_element_by_xpath(
        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/a/div/div/div[2]').text

    # col1과 col2로 이루어진 데이터프레임 결합
    ta = ta.append(df_tem, ignore_index=True)

    # hs code 클릭
    driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
        i + 1) + ']/a/div/div/div[1]/div[2]/span[1]').click()
    time.sleep(3)

    # 데이터 추출
    tab_bool, df = tabu()
    time.sleep(3)
    if tab_bool:
        ta = pd.concat([ta, df], axis=1)
        time.sleep(3)

    # 최종 데이터프레임에 결합
    ta_df = pd.concat([ta_df, ta], ignore_index=True)

    # 뒤로 가기
    driver.back()
    time.sleep(3)

    # 접근 확인
    print(driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
        i + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text)

    # 다시 축약
    driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

    # 첫 번째 ul에서 두 번째 ul이 있는지 화살표 클릭
    driver.find_element_by_xpath(
        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/i').click()
    time.sleep(3)

    # 6단위 기준의 목록에 순차적으로 접근
    slists = driver.find_elements_by_css_selector(ulli + li + li)

    # 두 번째 ul 존재 여부 판단
    if len(slists) != 0:
        for j in range(0, len(slists)):
            # 계속해서 바뀌는 임시 데이터공간이므로 내부에 선언
            ta = pd.DataFrame(columns=[col1, col2])

            # 임시적으로 데이터를 저장할 공간
            df_tem = {col1: '', col2: ''}
            df_tem[col1] = driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                    j + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text
            df_tem[col2] = driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                    j + 1) + ']/a/div/div/div[2]').text
            ta = ta.append(df_tem, ignore_index=True)

            # hs code 클릭
            driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                    j + 1) + ']/a/div/div/div[1]/div[2]/span[1]').click()
            time.sleep(3)

            # 데이터 추출
            tab_bool, df = tabu()
            time.sleep(3)
            if tab_bool:
                ta = pd.concat([ta, df], axis=1)
                time.sleep(3)

            # 최종 데이터프레임에 결합
            ta_df = pd.concat([ta_df, ta], ignore_index=True)

            # 뒤로 가기
            driver.back()
            time.sleep(3)

            # 접근 확인
            print(driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                    j + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text)

            # 다시 축약
            driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

            # 첫 번째 ul에서 두 번째 ul이 있는지 화살표 클릭
            driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/i').click()
            time.sleep(3)

            # 두 번째 ul에서 세 번째 ul이 있는지 화살표 클릭
            driver.find_element_by_xpath(
                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                    j + 1) + ']/i').click()
            time.sleep(3)

            # 세 번째 ul 내의 항목 개수를 측정
            tlists = driver.find_elements_by_css_selector(ulli + li + li + li)

            # 세 번째 ul 존재 여부 판단
            if len(tlists) != 0:
                for k in range(0, len(tlists)):
                    # 계속해서 바뀌는 임시 데이터공간이므로 내부에 선언
                    ta = pd.DataFrame(columns=[col1, col2])

                    # 임시적으로 데이터를 저장할 공간
                    df_tem = {col1: '', col2: ''}

                    # hs code, items로 이루어진 데이터 추출
                    df_tem[col1] = driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/ul/li[' + str(k + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text
                    df_tem[col2] = driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/ul/li[' + str(k + 1) + ']/a/div/div/div[2]').text
                    ta = ta.append(df_tem, ignore_index=True)

                    # hs code 클릭
                    driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/ul/li[' + str(k + 1) + ']/a/div/div/div[1]/div[2]/span[1]').click()
                    time.sleep(3)

                    # 데이터 추출
                    tab_bool, df = tabu()
                    time.sleep(3)
                    if tab_bool:
                        ta = pd.concat([ta, df], axis=1)
                        time.sleep(3)

                    # 최종 데이터프레임에 결합
                    ta_df = pd.concat([ta_df, ta], ignore_index=True)

                    # 뒤로 가기
                    driver.back()
                    time.sleep(3)

                    # 접근 확인
                    print(driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/ul/li[' + str(k + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text)

                    # 다시 축약
                    driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

                    # 첫 번째 ul에서 두 번째 ul이 있는지 화살표 클릭
                    driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/i').click()
                    time.sleep(3)

                    # 두 번째 ul에서 세 번째 ul이 있는지 화살표 클릭
                    driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/i').click()
                    time.sleep(3)

                    # 세 번째 ul에서 네 번째 ul이 있는지 화살표 클릭
                    driver.find_element_by_xpath(
                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                            j + 1) + ']/ul/li[' + str(k + 1) + ']/i').click()
                    time.sleep(3)

                    # 네 번째 ul 내의 항목 개수를 측정
                    olists = driver.find_elements_by_css_selector(ulli + li + li + li + li)

                    if len(olists) != 0:
                        for m in range(0, len(olists)):
                            # 계속해서 바뀌는 임시 데이터공간이므로 내부에 선언
                            ta = pd.DataFrame(columns=[col1, col2])

                            # 임시적으로 데이터를 저장할 공간
                            df_tem = {col1: '', col2: ''}

                            # hs code, items로 이루어진 데이터 추출
                            df_tem[col1] = driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                    m + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text
                            df_tem[col2] = driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                    m + 1) + ']/a/div/div/div[2]').text
                            ta = ta.append(df_tem, ignore_index=True)

                            # hs code 클릭
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                    m + 1) + ']/a/div/div/div[1]/div[2]/span[1]').click()
                            time.sleep(3)

                            # 데이터 추출
                            tab_bool, df = tabu()
                            time.sleep(3)
                            if tab_bool:
                                ta = pd.concat([ta, df], axis=1)
                                time.sleep(3)

                            # 최종 데이터프레임에 결합
                            ta_df = pd.concat([ta_df, ta], ignore_index=True)

                            # 뒤로 가기
                            driver.back()
                            time.sleep(3)

                            # 접근 확인
                            print(driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                    m + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text)

                            # 다시 축약
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

                            # 첫 번째 ul에서 두 번째 ul이 있는지 화살표 클릭
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/i').click()
                            time.sleep(3)

                            # 두 번째 ul에서 세 번째 ul이 있는지 화살표 클릭
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/i').click()
                            time.sleep(3)

                            # 세 번째 ul에서 네 번째 ul이 있는지 화살표 클릭
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/i').click()
                            time.sleep(3)

                            # 네 번째 ul에서 다섯 번째 ul이 있는지 화살표 클릭
                            driver.find_element_by_xpath(
                                '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                    i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                    m + 1) + ']/i').click()
                            time.sleep(3)

                            # 다섯 번째 ul 내의 항목 개수를 측정
                            ilists = driver.find_elements_by_css_selector(ulli + li + li + li + li + li)

                            if len(ilists) != 0:
                                for n in range(0, len(ilists)):
                                    # 계속해서 바뀌는 임시 데이터공간이므로 내부에 선언
                                    ta = pd.DataFrame(columns=[col1, col2])

                                    # 임시적으로 데이터를 저장할 공간
                                    df_tem = {col1: '', col2: ''}

                                    # hs code, items로 이루어진 데이터 추출
                                    df_tem[col1] = driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/ul/li[' + str(
                                            n + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text
                                    df_tem[col2] = driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/ul/li[' + str(
                                            n + 1) + ']/a/div/div/div[2]').text
                                    ta = ta.append(df_tem, ignore_index=True)

                                    # hs code 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/ul/li[' + str(
                                            n + 1) + ']/a/div/div/div[1]/div[2]/span[1]').click()
                                    time.sleep(3)

                                    # 데이터 추출
                                    tab_bool, df = tabu()
                                    time.sleep(3)
                                    if tab_bool:
                                        ta = pd.concat([ta, df], axis=1)
                                        time.sleep(3)

                                    # 최종 데이터프레임에 결합
                                    ta_df = pd.concat([ta_df, ta], ignore_index=True)

                                    # 뒤로 가기
                                    driver.back()
                                    time.sleep(3)

                                    # 접근 확인
                                    print(driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/ul/li[' + str(
                                            n + 1) + ']/a/div/div/div[1]/div[2]/span[1]').text)

                                    # 다시 축약
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[2]/div/div[1]/span').click()

                                    # 첫 번째 ul에서 두 번째 ul이 있는지 화살표 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/i').click()
                                    time.sleep(3)

                                    # 두 번째 ul에서 세 번째 ul이 있는지 화살표 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/i').click()
                                    time.sleep(3)

                                    # 세 번째 ul에서 네 번째 ul이 있는지 화살표 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/i').click()
                                    time.sleep(3)

                                    # 네 번째 ul에서 다섯 번째 ul이 있는지 화살표 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/i').click()
                                    time.sleep(3)

                                    # 다섯 번째 ul에서 여섯 번째 ul 있는지 화살표 클릭
                                    driver.find_element_by_xpath(
                                        '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                            i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(
                                            k + 1) + ']/ul/li[' + str(m + 1) + ']/ul/li[' + str(n + 1) + ']/i').click()
                                    time.sleep(3)

                                    # 여섯 번째 ul 내의 항목의 개수 측정
                                    xlists = driver.find_elements_by_css_selector(ulli + li + li + li + li + li + li)
                                    time.sleep(3)

                                    if len(xlists) != 0:
                                        print("여섯 번째 ul이 존재")

                                # 다섯 번째 ul에 대해 전부 수행했으므로, 다섯 번째 ul을 닫기
                                driver.find_element_by_xpath(
                                    '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                        i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/ul/li[' + str(
                                        m + 1) + ']/i').click()
                                time.sleep(3)
                        # 네 번째 ul에 대해 전부 수행했으므로, 네 번째 ul을 닫기
                        driver.find_element_by_xpath(
                            '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(
                                i + 1) + ']/ul/li[' + str(j + 1) + ']/ul/li[' + str(k + 1) + ']/i').click()
                        time.sleep(3)
                # 세 번째 ul에 대해 전부 수행했으므로, 세 번째 ul을 닫기
                driver.find_element_by_xpath(
                    '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/ul/li[' + str(
                        j + 1) + ']/i').click()
                time.sleep(3)
        # 두 번째 ul에 대해 전부 수행했으므로, 두 번째 ul를 닫기
        driver.find_element_by_xpath(
            '/html/body/div[2]/div[1]/div[3]/div[3]/div[2]/div[3]/ul/li[' + str(i + 1) + ']/i').click()
        time.sleep(3)

# NaN 값 처리
ta_df = ta_df.fillna('')
ta_df

# 엑셀에 작성하기
excel_path="C:\\datapare\\IS\\13.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
ta_df.to_excel(writer, index=False)
writer.save()