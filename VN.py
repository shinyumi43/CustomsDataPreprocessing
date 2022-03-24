import os
import time
import pandas as pd

# 엑셀 함수 작업 자동화
import openpyxl as op
import requests

# 엑셀 파일 읽어오기
exc_df = pd.read_excel('C:\\datapre\\VN\\01_BIEU_THUE_XNK_2022.xlsx', sheet_name='BIEU THUE 2022')

# 불필요한 열 제거 전에 길이 측정
exc_df_len = len(exc_df.columns)

# 컬럼명 변환 전에 불필요한 열 제거
exc_df = exc_df.drop(['Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 86'], axis=1)

# 컬럼명 변환
for i in range(3, exc_df_len - 1):
    col = exc_df['Unnamed: ' + str(i)][1]
    exc_df = exc_df.rename(columns={'Unnamed: ' + str(i): col})

# 결측값 제거
exc_df = exc_df.fillna('')

# 불필요한 행 제거
exc_df = exc_df.drop([0, 1, 2, 3, 4], axis=0)

# 엑셀 상에 숨기기로 된 열을 제거
del_col_01 = 'Văn bản'
del_col_02 = 'Ngày hiệu lực'
exc_df = exc_df.drop([del_col_01, del_col_02], axis=1)

# 엑셀에 작성 및 추출
excel_path = "C:\\datapre\\VN\\1.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
exc_df.to_excel(writer, index=False)
writer.save()