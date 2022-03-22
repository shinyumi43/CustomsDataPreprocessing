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

# pdf 파일 추출 경로
pdf_path="C:\\datapare\\ZA\\ZA.pdf"
# excel 파일 생성 경로
excel_path="C:\\datapare\\ZA\\ZA.xlsx"

top=(21.85*72)/25.4
left=(12.70*72)/25.4
width=(271.64*72)/25.4
height=(172.72*72)/25.4

col1=(24.14*72)/25.4+left
col2=(12.08*72)/25.4+col1
col3=(102.62*72)/25.4+col2
col4=(18.11*72)/25.4+col3
col5=(18.11*72)/25.4+col4
col6=(18.11*72)/25.4+col5
col7=(18.11*72)/25.4+col6
col8=(18.11*72)/25.4+col7
col9=(24.15*72)/25.4+col8
col10=(18.11*72)/25.4+col9

y1=top
x1=left
y2=top+height
x2=left+width

p_cols=(col1, col2, col3, col4, col5, col6, col7, col8, col9, col10)
p_area=(y1, x1, y2, x2)

# 가져올 데이터 내의 열 값을 리스트에 담기
col1="Heading/Subheading"
col2="CD"
col3="Article Description"
col4="Statistical Unit"
col5="General"
col6="EU"
col7="EFTA"
col8="SADC"
col9="MERCOSUR"
col10="AfCFTA"
col_list=[col1, col2, col3, col4, col5, col6, col7, col8, col9, col10]

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()

def tabu(page_p):
    # pdf 파일 읽기
    ta_df = tabula.read_pdf(pdf_path, columns=p_cols, area=p_area, pages=page_p, **kwargs)
    # 테이블이 존재하는 경우
    if len(ta_df)==1:
        # 열의 개수를 파악
        df_ta = ta_df[0]
        # 열의 개수가 10이 아니거나 10으로 인식되지만 불필요한 테이블인 경우
        if (len(df_ta.columns) != 10) or (len(df_ta.columns)==10 and df_ta.columns[0]!='Heading /'):
            return False, None
        # 열의 개수가 10이고, 필요한 테이블인 경우
        elif (len(df_ta.columns)==10 and df_ta.columns[0]=='Heading /'):
            df_ta=df_ta.drop(index=0, axis=0)
            df_ta.columns=col_list
            df_ta=df_ta.fillna('')
            return True, df_ta
    # 테이블이 존재하지 않는 경우(예외처리)
    else:
        print("테이블의 개수 : "+str(len(ta_df)))
        return False, None


def mkcell(df_mk):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=[col1, col2, col3, col4, col5, col6, col7, col8, col9, col10])
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: '', col5: '', col6: '', col7: '', col8: '', col9: '', col10: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    desbuffer = ""
    bgfbuffer = ""
    bgxbuffer = ""
    bgebuffer = ""
    bggbuffer = ""
    bgnbuffer = ""
    bgtbuffer = ""
    # 삭제한 행으로 인해 1부터 시작
    for i in range(1, df_mk_len + 1):
        # 맨 뒤로 가져올 때, NaN으로 나오는 경우도 있기 때문에 이를 고려하여 col3에 대해 유의
        # col1
        a_mkcell = df_mk[col1][i]
        # col2
        b_mkcell = df_mk[col2][i]
        # col3
        if df_mk[col3][i] == '':
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])
        # col4
        d_mkcell = df_mk[col4][i]
        # col5
        if df_mk[col5][i] == '':
            e_mkcell = ""
        else:
            e_mkcell = str(df_mk[col5][i])
        # col6
        if df_mk[col6][i] == '':
            f_mkcell = ""
        else:
            f_mkcell = str(df_mk[col6][i])
        # col7
        if df_mk[col7][i] == '':
            g_mkcell = ""
        else:
            g_mkcell = str(df_mk[col7][i])
        # col8
        if df_mk[col8][i] == '':
            h_mkcell = ""
        else:
            h_mkcell = str(df_mk[col8][i])
        # col9
        if df_mk[col9][i] == '':
            i_mkcell = ""
        else:
            i_mkcell = str(df_mk[col9][i])
        # col10
        if df_mk[col10][i] == '':
            j_mkcell = ""
        else:
            j_mkcell = str(df_mk[col10][i])

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if (not a_mkcell == "") or ((not c_mkcell == "") and (c_mkcell[0] == '-')):
            # i가 1일 경우
            if not i == 1:
                temp_dic_mk[col3] = desbuffer
                temp_dic_mk[col5] = bgfbuffer
                temp_dic_mk[col6] = bgxbuffer
                temp_dic_mk[col7] = bgebuffer
                temp_dic_mk[col8] = bggbuffer
                temp_dic_mk[col9] = bgnbuffer
                temp_dic_mk[col10] = bgtbuffer
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

            # i가 1이 아닐 경우
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]
            temp_dic_mk[col4] = df_mk[col4][i]
            temp_dic_mk[col5] = df_mk[col5][i]
            temp_dic_mk[col6] = df_mk[col6][i]
            temp_dic_mk[col7] = df_mk[col7][i]
            temp_dic_mk[col8] = df_mk[col8][i]
            temp_dic_mk[col9] = df_mk[col9][i]
            temp_dic_mk[col10] = df_mk[col10][i]

            # col3, col5, col6, col7, col8, col9, col10 내의 내용은 줄바꿈된 문장이 존재하기 때문에 다음과 같이 예외적으로 처리
            desbuffer = c_mkcell
            bgfbuffer = e_mkcell
            bgxbuffer = f_mkcell
            bgebuffer = g_mkcell
            bggbuffer = h_mkcell
            bgnbuffer = i_mkcell
            bgtbuffer = j_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙음
        else:
            desbuffer = desbuffer + "\n" + c_mkcell
            bgfbuffer = bgfbuffer + "\n" + e_mkcell
            bgxbuffer = bgxbuffer + "\n" + f_mkcell
            bgebuffer = bgebuffer + "\n" + g_mkcell
            bggbuffer = bggbuffer + "\n" + h_mkcell
            bgnbuffer = bgnbuffer + "\n" + i_mkcell
            bgtbuffer = bgtbuffer + "\n" + j_mkcell
    # 최종적으로 쌓인 desbuffer에 대한 값을 넣어줌
    temp_dic_mk[col3] = desbuffer
    temp_dic_mk[col5] = bgfbuffer
    temp_dic_mk[col6] = bgxbuffer
    temp_dic_mk[col7] = bgebuffer
    temp_dic_mk[col8] = bggbuffer
    temp_dic_mk[col9] = bgnbuffer
    temp_dic_mk[col10] = bgtbuffer
    # append를 통해 병합, 인덱스의 이름은 무시
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# adobe 측정으로 인해 분리된 데이터를 본래 자리에 결합
def cobi(df_mk):
    # 테이블 길이 측정
    df_mk_len=len(df_mk)
    for i in range(1, df_mk_len+1):
        # col2의 데이터가 ''인 경우에 분리된 데이터 삽입이 필요
        if df_mk[col2][i]=='' and df_mk[col4][i]!='':
            # col3에 모든 데이터를 결합
            df_mk[col3][i]=df_mk[col3][i]+df_mk[col4][i]+df_mk[col5][i]+df_mk[col6][i]+df_mk[col7][i]+df_mk[col8][i]+df_mk[col9][i]+df_mk[col10][i]
            # col3에 결합된 이후의 데이터들은 ''로 초기화
            df_mk[col4][i]=''
            df_mk[col5][i]=''
            df_mk[col6][i]=''
            df_mk[col7][i]=''
            df_mk[col8][i]=''
            df_mk[col9][i]=''
            df_mk[col10][i]=''
    return df_mk

# 최종 데이터프레임 생성
df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(1, num_of_pages+1):
    # 필요없는 행이나 열을 제거
    tab_bool, df_tem = tabu(i)
    if tab_bool:
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat=mkcell(df_tem)
        # 여러 열로 분리된 데이터에 대한 처리
        df_temp_mk=cobi(df_forcat)
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_temp_mk],ignore_index=True)
    print(i)

# 엑셀에 작성하기
excel_path="C:\\datapare\\ZA\\ZA.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()