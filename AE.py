# pdf 내의 데이터 추출을 위한 라이브러리
import pandas as pd
# 벡터, 행렬과 같은 수치연산을 수행하는 라이브러리
import tabula
# pdf 변환 시에 사용하는 라이브러리
import PyPDF2

# 열을 string으로 변환
col2str = {'dtype': str}
kwargs = {
          'pandas_options': col2str,
          'stream': True}

# 읽을 pdf 경로
pdf_path="C:\\datapare\\AE\\AE.pdf"
# 생성할 excel 경로
excel_path="C:\\datapare\\AE\\AE.xlsx"

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()
num_of_pages

# 가져올 데이터 내의 열 값을 리스트에 담기
col1="DUTY فـئة الرسم RATE"
col2="وحدة الاستيفاء UINT"
col3="DESCRIPTION"
col4="الصنف"
col5="رمز النظام المنسق H.S CODE"
col6="رقم البند HEADING NO"

col_list=[col1, col2, col3, col4, col5, col6]


# 테이블의 개수가 2개 이상인데, 다른 테이블이 삽입된 경우
def tabu(page_p):
    ta_df = tabula.read_pdf(pdf_path, multiple_tables=True, pages=page_p, **kwargs)
    df_temp = pd.DataFrame(columns=col_list)
    flag = False
    for i in range(0, len(ta_df)):
        if len(ta_df[i].columns) > 6:
            # 테이블 가져오기
            df_ta = ta_df[i]
            # 열의 개수가 9, 8, 7인 경우에 따라 제거할 열을 선택
            if len(df_ta.columns) == 9:
                if df_ta['Unnamed: 1'][1] == 'وحدة الاستيفاء UINT':
                    df_ta = df_ta.drop(columns=['Unnamed: 2', 'Unnamed: 4', 'Unnamed: 6'])
                elif df_ta['Unnamed: 2'][1] == 'وحدة الاستيفاء UINT':
                    df_ta = df_ta.drop(columns=['Unnamed: 1', 'Unnamed: 4', 'Unnamed: 6'])

            elif len(df_ta.columns) == 8:
                if df_ta['Unnamed: 1'][1] == 'وحدة الاستيفاء UINT':
                    df_ta = df_ta.drop(columns=['Unnamed: 2', 'Unnamed: 5'])
                elif df_ta['Unnamed: 2'][1] == 'وحدة الاستيفاء UINT':
                    df_ta = df_ta.drop(columns=['Unnamed: 1', 'Unnamed: 5'])

            elif len(df_ta.columns) == 7:
                if df_ta['Unnamed: 1'][1] == 'وحدة الاستيفاء UINT' or df_ta['Unnamed: 1'][2] == 'UNITS':
                    df_ta = df_ta.drop(columns=['Unnamed: 2'])
                elif df_ta['Unnamed: 2'][1] == 'وحدة الاستيفاء UINT':
                    df_ta = df_ta.drop(columns=['Unnamed: 1'])

            # 필요없는 부분 제거
            df_ta = df_ta.drop(index=[0, 1, 2, 3], axis=0)

            # 열을 동일하게 설정
            df_ta.columns = col_list
            df_temp = pd.concat([df_temp, df_ta], ignore_index=True)
            flag = True
        else:
            print("필요 없는 테이블")
            flag = False

    if flag == True:
        return True, df_temp
    else:
        return False, None


def mkcell(df_mk):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=col_list)
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: '', col5: '', col6: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # 하나의 데이터가 줄바꿈을 기준으로 여러 셀에 삽입되는 것을 처리
    # 빈 문자열 생성
    desbuffer = ""
    bgobuffer = ""
    blabuffer = ""
    for i in range(0, df_mk_len):
        # col2, col4의 값이 여러 셀의 삽입되는 경우가 발생하므로 아래와 같이 처리
        a_mkcell = df_mk[col1][i]
        b_mkcell = df_mk[col2][i]
        if pd.isnull(df_mk[col3][i]):
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])
        if pd.isnull(df_mk[col4][i]):
            d_mkcell = ""
        else:
            d_mkcell = str(df_mk[col4][i])
        e_mkcell = df_mk[col5][i]
        if pd.isnull(df_mk[col6][i]):
            f_mkcell = ""
        else:
            f_mkcell = str(df_mk[col6][i])

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if ((not c_mkcell == "") and (c_mkcell[0] == '-')) or ((not c_mkcell == "") and (c_mkcell[0].isupper())):
            if not i == 0:
                temp_dic_mk[col3] = desbuffer
                temp_dic_mk[col4] = bgobuffer
                temp_dic_mk[col6] = blabuffer
                # append 함수는 표 내의 데이터를 합치는 함수
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)
                # temp_dic_mk 에 전달 받은 df_mk 값 삽입
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]
            temp_dic_mk[col4] = df_mk[col4][i]
            temp_dic_mk[col5] = df_mk[col5][i]
            temp_dic_mk[col6] = df_mk[col6][i]

            desbuffer = c_mkcell
            bgobuffer = d_mkcell
            blabuffer = f_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙기
        else:
            desbuffer = desbuffer + "\n" + c_mkcell
            bgobuffer = bgobuffer + "\n" + d_mkcell
            blabuffer = blabuffer + "\n" + f_mkcell
    # 최종적으로 쌓인 desbuffer, bgobuffer를 col2, col4에 삽입
    temp_dic_mk[col3] = desbuffer
    temp_dic_mk[col4] = bgobuffer
    temp_dic_mk[col6] = blabuffer
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 빈 데이터프레임 생성f
df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(1, num_of_pages+1):
    # 필요없는 행이나 열을 제거
    tab_bool, df_tem = tabu(i)
    if tab_bool:
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat=mkcell(df_tem)
        # NaN 값 ''으로 처리
        df_forcat=df_forcat.fillna('')
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_forcat],ignore_index=True)
        print(i)
print("success")
