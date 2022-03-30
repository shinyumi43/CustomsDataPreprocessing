import pandas as pd
import numpy as np
import glob
import sys
import xlrd

# 최종 데이터프레임 생성
all_data = pd.DataFrame()

# 첫 번째 파일을 기준
excel_path = "C:\\datapre\\JP\\01.xlsx"
# 엑셀 파일 읽어오기
data = pd.read_excel(excel_path, sheet_name="Sheet1")

# 칼럼명 확인
col_list = []
for i in range(0, len(data.columns)):
    col_list.append(data.columns[i])

# 임시 데이터프레임 생성
df = pd.DataFrame(data, columns=col_list)
# 최종 데이터프레임에 병합
all_data = all_data.append(df, ignore_index=True)

# 폴더 경로 지정
for f in glob.glob('C:\\datapre\\JP\\excel\\*.xlsx'):
    data = pd.read_excel(f)

    # 필요한 컬럼 입력
    df = pd.DataFrame(data, columns=col_list)
    # 불필요한 행 제거
    df = df.drop(index=0, axis=0)
    # 최종 데이터프레임에 병합
    all_data = all_data.append(df, ignore_index=True)

    # 데이터 갯수 확인
    print(all_data.shape)

    # 데이터 잘 들어오는지 확인
    all_data.head()

# 불필요한 열을 제거
all_data=all_data.drop(columns=col_list[0], axis=1)

# 엑셀에 작성하기
excel_path = "C:\\datapre\\JP\\excel\\JP.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
all_data.to_excel(writer, index=False)
writer.save()


