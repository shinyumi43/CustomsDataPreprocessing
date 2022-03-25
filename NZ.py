# pdf 내의 데이터 추출을 위한 라이브러리
import pandas as pd
import numpy as np
# 벡터, 행렬과 같은 수치연산을 수행하는 라이브러리
import tabula
# pdf 변환 시에 사용하는 라이브러리
import PyPDF2

# 열을 string으로 변환
col2str = {'dtype': str}
kwargs = {'output_format': 'dataframe',
          'pandas_options': col2str,
          'stream': True}

# i ~ xxi까지 바꾸어주면서 설정 변경
# 읽을 pdf 경로
pdf_path="C:\\datapare\\NZ\\i.pdf"
# 생성할 excel 경로
excel_path="C:\\datapare\\NZ\\i.xlsx"

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()

# 홀수와 짝수 페이지 구성이 다름
# 홀수 페이지 측정
top=(44.03*72)/25.4
left=(19.93*72)/25.4
width=(180*72)/25.4
height=(240.3*72)/25.4

col1=(25*72)/25.4+left
col2=(12.5*72)/25.4+col1
col3=(12.5*72)/25.4+col2
col4=(90*72)/25.4+col3
col5=(20*72)/25.4+col4
col6=(20*72)/25.4+col5

y1=top
x1=left
y2=top+height
x2=left+width

odd_cols=(col1, col2, col3, col4, col5, col6)
odd_area=(y1, x1, y2, x2)

# 짝수 페이지 측정
top=(43.38*72)/25.4
left=(10*72)/25.4
width=(180*72)/25.4
height=(240.29*72)/25.4

col1=(25*72)/25.4+left
col2=(12.5*72)/25.4+col1
col3=(12.5*72)/25.4+col2
col4=(90*72)/25.4+col3
col5=(20*72)/25.4+col4
col6=(20*72)/25.4+col5

y1=top
x1=left
y2=top+height
x2=left+width

even_cols=(col1, col2, col3, col4, col5, col6)
even_area=(y1, x1, y2, x2)

col1='Number'
col2='Code'
col3='Unit'
col4='Goods'
col5='Normal Tariff'
col6='*Preferential Tariff'
col_list=[col1, col2, col3, col4, col5, col6]

def tabu(page_p):
    # pdf 파일 읽기, 홀수 페이지
    if int(page_p) % 2 != 0:
        ta_df = tabula.read_pdf(pdf_path, columns=odd_cols, area=odd_area, pages=page_p, **kwargs)
    else:
        ta_df = tabula.read_pdf(pdf_path, columns=even_cols, area=even_area, pages=page_p, **kwargs)

    # i ~ xxi까지 각 파일마다 테이블이 존재하지 않는 페이지 입력 및 실행
    # 테이블이 존재하지 않는 페이지
    # i
    del_temp=[1,2,6,7,8,17,18,41,42,47,48]
    # ii
    # del_temp=[1, 2, 5, 6, 12, 13, 14, 20, 21, 22, 26, 27, 28, 30, 31, 32, 36, 37, 38, 44, 45, 46, 48, 49, 50, 52]
    # iii
    # del_temp=[1,2,8]
    # iv
    # del_temp=[1,2,11,12,15,16,18,19,20,25,26,39,40,46,47,48,62,63,64,67,68,72]
    # ix
    # del_temp=[1, 2, 30, 31, 32, 34, 35, 36, 38]
    # v
    # del_temp=[1,2,7,8,11,12,20]
    # vi
    # del_temp=[1,2,11,12,31,32,39,40,42,43,44,51,52,58,59,60,65,66,69,70,72,73,74,78,79,80]
    # vii
    # del_temp=[1,2,3,4,27,28]
    # viii
    # del_temp=[1,2,8,9,10,14,15,16]
    # x
    # del_temp=[1, 2, 5, 6, 7, 8, 23, 24]
    # xi
    # del_temp=[1,2,3,4,6,7,8,14,15,16,26,27,28,31,32,41,42,52,53,54,59,60,66,67,68,72,73,74,78,79,80,90,91,92,110,111,112,128,129,130]
    # xii
    # del_temp=[1,2,14,15,16,19,20,22,23,24,26]
    # xiii
    # del_temp=[1,2,9,10,17,18]
    # xiv
    # del_temp=[1,2,8]
    # xix
    # del_temp=[1,2,6]
    # xv
    # del_temp=[1,2,3,4,5,6,40,41,42,65,71,72,74,75,76,83,84,86,87,88,90,91,92,94,95,96,99,100,107,108]
    # xvi
    # del_temp=[1,2,3,4,73,74,75,76]
    # xvii
    # del_temp=[1,2,5,6,41,42,45,46,52]
    # xviii
    # del_temp=[1,2,14,15,16,20,21,22]
    # xx
    # del_temp=[1,2,14,15,16,23,24,36]
    # xxi
    # del_temp=[1,2,4,6]

    if int(page_p) in del_temp:
        return False, None
    else:
        # 열에 있는 데이터를 테이블 안으로 삽입
        df_ta = ta_df[0]
        # 빈 데이터프레임 생성
        df_mk = pd.DataFrame(columns=col_list)
        # 임시 배열 생성
        df_tem = {col1: '', col2: '', col3: '', col4: '', col5: '', col6: ''}
        # 우선적으로 열 내에 있는 데이터를 임시 배열에 삽입
        df_tem[col1] = df_ta.columns[0]
        df_tem[col2] = df_ta.columns[1]
        df_tem[col3] = df_ta.columns[2]
        df_tem[col4] = df_ta.columns[3]
        df_tem[col5] = df_ta.columns[4]
        df_tem[col6] = df_ta.columns[5]
        # 데이터프레임에 삽입
        df_mk = df_mk.append(df_tem, ignore_index=True)

        # 열 데이터는 삽입했으므로, 원본을 지정해야하는 열로 교체
        df_ta.columns = col_list
        for i in range(0, len(df_ta)):
            # 불필요한 부분 제거하고, 필요한 부분만 추출
            df_tem[col1] = df_ta[col1][i]
            df_tem[col2] = df_ta[col2][i]
            df_tem[col3] = df_ta[col3][i]
            df_tem[col4] = df_ta[col4][i]
            df_tem[col5] = df_ta[col5][i]
            df_tem[col6] = df_ta[col6][i]
            df_mk = df_mk.append(df_tem, ignore_index=True)
        return True, df_mk

# 불필요한 데이터를 제거하는 함수
def remo(df_mk):
    df_mk_len=len(df_mk)
    for i in range(0, df_mk_len):
        if (df_mk[col1][i]=='Unnamed: 0' or df_mk[col1][i]=='Unnamed: 1' or df_mk[col1][i]=='Unnamed: 2' or
            df_mk[col1][i]=='Unnamed: 3' or df_mk[col1][i]=='Unnamed: 4' or df_mk[col1][i]=='Unnamed: 5'):
            df_mk[col1][i]=''
        if (df_mk[col2][i]=='Unnamed: 0' or df_mk[col2][i]=='Unnamed: 1' or df_mk[col2][i]=='Unnamed: 2' or
            df_mk[col2][i]=='Unnamed: 3' or df_mk[col2][i]=='Unnamed: 4' or df_mk[col2][i]=='Unnamed: 5'):
            df_mk[col2][i]=''
        if (df_mk[col3][i]=='Unnamed: 0' or df_mk[col3][i]=='Unnamed: 1' or
            df_mk[col3][i]=='Unnamed: 2' or df_mk[col3][i]=='Unnamed: 3' or df_mk[col3][i]=='Unnamed: 4' or
            df_mk[col3][i]=='Unnamed: 5'):
            df_mk[col3][i]=''
        if (df_mk[col4][i]=='Unnamed: 0' or df_mk[col4][i]=='Unnamed: 1' or
            df_mk[col4][i]=='Unnamed: 2' or df_mk[col4][i]=='Unnamed: 3' or df_mk[col4][i]=='Unnamed: 4' or
            df_mk[col4][i]=='Unnamed: 5' or df_mk[col4][i]=='* * *'):
            df_mk[col4][i]=''
        if (df_mk[col5][i]=='Unnamed: 0' or df_mk[col5][i]=='Unnamed: 1' or
            df_mk[col5][i]=='Unnamed: 2' or df_mk[col5][i]=='Unnamed: 3' or df_mk[col5][i]=='Unnamed: 4' or
            df_mk[col5][i]=='Unnamed: 5'):
            df_mk[col5][i]=''
        if (df_mk[col6][i]=='Unnamed: 0' or df_mk[col6][i]=='Unnamed: 1' or
            df_mk[col6][i]=='Unnamed: 2' or df_mk[col6][i]=='Unnamed: 3' or df_mk[col6][i]=='Unnamed: 4' or
            df_mk[col6][i]=='Unnamed: 5'):
            df_mk[col6][i]=''
        df_mk=df_mk.fillna('')
    return df_mk


def mkcell(df_mk):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=[col1, col2, col3, col4, col5, col6])
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: '', col5: '', col6: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    desbuffer = ""
    bgfbuffer = ""
    adobuffer = ""
    pakbuffer = ""
    for i in range(0, df_mk_len):
        # col1
        a_mkcell = df_mk[col1][i]
        # col2
        b_mkcell = df_mk[col2][i]
        # col3
        if pd.isnull(df_mk[col3][i]):
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])
        # col4
        if pd.isnull(df_mk[col4][i]):
            d_mkcell = ""
        else:
            d_mkcell = str(df_mk[col4][i])
        # col5
        if pd.isnull(df_mk[col5][i]):
            e_mkcell = ""
        else:
            e_mkcell = str(df_mk[col5][i])
        # col6
        if pd.isnull(df_mk[col6][i]):
            f_mkcell = ""
        else:
            f_mkcell = str(df_mk[col6][i])

        # col1이 비어있지 않거나 col4에 해당하는 값의 맨 앞이 -이나 .으로 시작하는 경우
        if (not a_mkcell == "") or ((not d_mkcell == "") and (d_mkcell[0] == '–')) or (
                (not d_mkcell == "") and (d_mkcell[0] == '.')):
            # i가 0일 경우
            if not i == 0:
                temp_dic_mk[col3] = pakbuffer
                temp_dic_mk[col4] = desbuffer
                temp_dic_mk[col5] = adobuffer
                temp_dic_mk[col6] = bgfbuffer
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

            # i가 0이 아닐 경우
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]
            temp_dic_mk[col4] = df_mk[col4][i]
            temp_dic_mk[col5] = df_mk[col5][i]
            temp_dic_mk[col6] = df_mk[col6][i]

            # col3, col4, col5, col6 내의 내용은 줄바꿈된 문장이 존재하기 때문에 다음과 같이 예외적으로 처리
            pakbuffer = c_mkcell
            desbuffer = d_mkcell
            adobuffer = e_mkcell
            bgfbuffer = f_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙음
        else:
            pakbuffer = pakbuffer + "\n" + c_mkcell
            desbuffer = desbuffer + "\n" + d_mkcell
            adobuffer = adobuffer + "\n" + e_mkcell
            bgfbuffer = bgfbuffer + "\n" + f_mkcell
    # 최종적으로 쌓인 desbuffer에 대한 값을 넣어줌
    temp_dic_mk[col3] = pakbuffer
    temp_dic_mk[col4] = desbuffer
    temp_dic_mk[col5] = adobuffer
    temp_dic_mk[col6] = bgfbuffer
    # append를 통해 병합, 인덱스의 이름은 무시
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 최종 데이터프레임 생성
df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(1, num_of_pages+1):
    # 필요없는 행이나 열을 제거
    tab_bool, df_tem = tabu(i)
    if tab_bool:
        # Unnamed: 0, NaN, 반복적인 열과 같은 데이터에 대한 처리
        df_temp_dic = remo(df_tem)
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat = mkcell(df_temp_dic)
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_forcat],ignore_index=True)
    print(i)
print("success")

# 엑셀에 작성하기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()