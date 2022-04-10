import pandas as pd
import tabula
import PyPDF2

# 열을 string으로 변환
col2str = {'dtype': str}
kwargs = {
          'pandas_options': col2str,
          'stream': True}

# pdf 경로
pdf_path="C:\\datapare\\CL\\2838.pdf"
# excel 경로
excel_path="C:\\datapare\\CL\\2838.xlsx"

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()

num_of_pages

# pdf 파일 내 표의 길이가 다르므로 재측정
top=(22.40*72)/25.4
left=(27.71*72)/25.4
width=(138.55*72)/25.4
height=(288.64*72)/25.4

col1=(11.08*72)/25.4+left
col2=(16.86*72)/25.4+col1
col3=(80.57*72)/25.4+col2
col4=(9.40*72)/25.4+col3
col5=(7.74*72)/25.4+col4
col6=(12.72*72)/25.4+col5

y1=top
x1=left
y2=top+height
x2=left+width

p_area=(y1, x1, y2, x2)
p_cols=(col1, col2, col3, col4, col5, col6)

# 지정할 열
col1="Partida"
col2="Código del S.A."
col3="Glosa"
col4="U.A."
col5="Adv."
col6="Estad. Unidad Código"
col_list=[col1, col2, col3, col4, col5, col6]


# 아래 네 가지 유형 중 페이지가 어떤 것에 해당하는지 파악하고, 페이지만 바꿔서 적용
# 불필요 + 필요 : 12, 10, 28, 35, 39, 42, 45, 47, 48, 51
# 필요 + 불필요 + 필요 : X
# 필요 + 불필요 : 27
# 필요 : 나머지
# 불필요 : 1, 9, 33, 38
# 열 삭제 되지 않은 부분이 존재
def tabu(page_p):
    # 불필요 + 필요
    temp = [2, 10, 28, 35, 39, 42, 45, 47, 48, 51]
    del_temp = [1, 9, 33, 38]
    # pdf 파일 읽기
    ta_df = tabula.read_pdf(pdf_path, columns=p_cols, area=p_area, pages=page_p, **kwargs)
    # 테이블은 무조건 존재함을 확인, 페이지 유형에 따라 다르게 정제
    # 불필요 + 필요
    if int(page_p) in temp:
        # 열 내에 데이터가 삽입되지 않고, 앞부분에만 불필요한 부분이 존재하는 경우
        df_ta = ta_df[0]
        df_ta.columns = col_list

        # 불필요한 행을 탐색하면서 해당 행들을 추출
        del_index = []
        for i in range(0, len(df_ta)):
            if df_ta[col1][i] == col1:
                break
            del_index.append(i)

        # 불필요한 부분 제거 및 아래 코드 실행 후 추가적인 3행 삭제 진행
        df_ta = df_ta.drop(index=del_index, axis=0)

        # 3행 추가 삭제
        del_index_len = len(del_index)
        df_ta = df_ta.drop(index=[del_index_len, del_index_len + 1, del_index_len + 2], axis=0)

        return True, df_ta, (del_index_len + 2) + 1
    # 필요 + 불필요
    elif int(page_p) == 27:
        df_ta = ta_df[0]
        df_len = len(df_ta)

        # 삭제할 지점을 찾는 과정
        idx = []
        for i in range(0, df_len):
            if df_ta[df_ta.columns[2]][i] == '_____________________':
                idx.append(i)

        # 삭제할 곳의 범위 설정
        start = idx[0]  # 시작 지점
        # 끝 지점, 뒤가 전부 불필요하므로 행의 길이까지 제거
        end = df_len
        # 삭제할 행 전부 배열에 삽입
        del_idx = []
        for i in range(start, end):
            del_idx.append(i)

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

        # 열 데이터는 삽입했으므로, 지정해야하는 열로 교체
        df_ta.columns = col_list
        for i in range(0, len(df_ta)):
            # 불필요한 부분 제거하고, 필요한 부분만 추출
            if not i in del_idx:
                df_tem[col1] = df_ta[col1][i]
                df_tem[col2] = df_ta[col2][i]
                df_tem[col3] = df_ta[col3][i]
                df_tem[col4] = df_ta[col4][i]
                df_tem[col5] = df_ta[col5][i]
                df_tem[col6] = df_ta[col6][i]
                df_mk = df_mk.append(df_tem, ignore_index=True)
        return True, df_mk, 0
    # 불필요
    elif int(page_p) in del_temp:
        return False, None, 0
    # 필요
    else:
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

        # 불필요한 부분이 존재하지 않으므로 del_idx = []
        del_idx = []

        # 열 데이터는 삽입했으므로, 지정해야하는 열로 교체
        df_ta.columns = col_list
        for i in range(0, len(df_ta)):
            # 불필요한 부분 제거하고, 필요한 부분만 추출
            if not i in del_idx:
                df_tem[col1] = df_ta[col1][i]
                df_tem[col2] = df_ta[col2][i]
                df_tem[col3] = df_ta[col3][i]
                df_tem[col4] = df_ta[col4][i]
                df_tem[col5] = df_ta[col5][i]
                df_tem[col6] = df_ta[col6][i]
                df_mk = df_mk.append(df_tem, ignore_index=True)
        return True, df_mk, 0

# 불필요한 데이터를 제거하는 함수
def remo(df_mk, df_del_len):
    df_mk_len=len(df_mk)
    for i in range(df_del_len, df_mk_len+df_del_len):
        if (df_mk[col1][i]=='Unnamed: 0' or df_mk[col1][i]=='Unnamed: 1' or df_mk[col1][i]=='Unnamed: 2' or
            df_mk[col1][i]=='Unnamed: 3' or df_mk[col1][i]=='Unnamed: 4' or df_mk[col1][i]=='Unnamed: 5' or
            df_mk[col1][i]=='Partida'):
            df_mk[col1][i]=''
        if (df_mk[col2][i]=='Código' or df_mk[col2][i]=='del S.A.' or df_mk[col2][i]=='Unnamed: 0' or
            df_mk[col2][i]=='Unnamed: 1' or df_mk[col2][i]=='Unnamed: 2' or df_mk[col2][i]=='Unnamed: 3' or
            df_mk[col2][i]=='Unnamed: 4' or df_mk[col2][i]=='Unnamed: 5'):
            df_mk[col2][i]=''
        if (df_mk[col3][i]=='Glosa' or df_mk[col3][i]=='Unnamed: 0' or df_mk[col3][i]=='Unnamed: 1' or
            df_mk[col3][i]=='Unnamed: 2' or df_mk[col3][i]=='Unnamed: 3' or df_mk[col3][i]=='Unnamed: 4' or
            df_mk[col3][i]=='Unnamed: 5' or df_mk[col3][i]=='_____________________'):
            df_mk[col3][i]=''
        if (df_mk[col4][i]=='U.A.' or df_mk[col4][i]=='Unnamed: 0' or df_mk[col4][i]=='Unnamed: 1' or
            df_mk[col4][i]=='Unnamed: 2' or df_mk[col4][i]=='Unnamed: 3' or df_mk[col4][i]=='Unnamed: 4' or
            df_mk[col4][i]=='Unnamed: 5'):
            df_mk[col4][i]=''
        if (df_mk[col5][i]=='Adv.' or df_mk[col5][i]=='Unnamed: 0' or df_mk[col5][i]=='Unnamed: 1' or
            df_mk[col5][i]=='Unnamed: 2' or df_mk[col5][i]=='Unnamed: 3' or df_mk[col5][i]=='Unnamed: 4' or
            df_mk[col5][i]=='Unnamed: 5'):
            df_mk[col5][i]=''
        if (df_mk[col6][i]=='Estad.' or df_mk[col6][i]=='Unidad' or
            df_mk[col6][i]=='Código' or df_mk[col6][i]=='Unnamed: 0' or df_mk[col6][i]=='Unnamed: 1' or
            df_mk[col6][i]=='Unnamed: 2' or df_mk[col6][i]=='Unnamed: 3' or df_mk[col6][i]=='Unnamed: 4' or
            df_mk[col6][i]=='Unnamed: 5'):
            df_mk[col6][i]=''
        df_mk=df_mk.fillna('')
    return df_mk, df_del_len


def mkcell(df_mk, del_len):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=[col1, col2, col3, col4, col5, col6])
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: '', col5: '', col6: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # bgrbuffer=""
    # bgebuffer=""
    desbuffer = ""
    # bgfbuffer=""
    # bgibuffer=""
    # bgxbuffer=""
    # 삭제한 행으로 인해 1부터 시작
    for i in range(del_len, df_mk_len + del_len):
        # 맨 뒤로 가져올 때, NaN으로 나오는 경우도 있기 때문에 이를 고려하여 col3에 대해 유의
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
        d_mkcell = df_mk[col4][i]
        # col5
        e_mkcell = df_mk[col5][i]
        # col6
        f_mkcell = df_mk[col6][i]

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if (not a_mkcell == "") or ((not c_mkcell == "") and (c_mkcell[0] == '-')):
            # i가 del_len일 경우
            if not i == del_len:
                # temp_dic_mk[col1]=bgrbuffer
                # temp_dic_mk[col2]=bgebuffer
                temp_dic_mk[col3] = desbuffer
                # temp_dic_mk[col4]=bgfbuffer
                # temp_dic_mk[col5]=bgibuffer
                # temp_dic_mk[col6]=bgxbuffer
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

            # i가 del_len이 아닐 경우
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]
            temp_dic_mk[col4] = df_mk[col4][i]
            temp_dic_mk[col5] = df_mk[col5][i]
            temp_dic_mk[col6] = df_mk[col6][i]

            # col1, col2, col3, col4, col5, col6 내의 내용은 줄바꿈된 문장이 존재하기 때문에 다음과 같이 예외적으로 처리
            # desbuffer=a_mkcell
            # bgebuffer=b_mkcell
            desbuffer = c_mkcell
            # bgfbuffer=d_mkcell
            # bgibuffer=f_mkcell
            # bgxbuffer=e_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙음
        else:
            # desbuffer=bgrbuffer+"\n"+a_mkcell
            # bgebuffer=bgebuffer+"\n"+b_mkcell
            desbuffer = desbuffer + "\n" + c_mkcell
            # bgfbuffer=bgfbuffer+"\n"+d_mkcell
            # bgibuffer=bgibuffer+"\n"+f_mkcell
            # bgxbuffer=bgxbuffer+"\n"+e_mkcell
    # 최종적으로 쌓인 desbuffer에 대한 값을 넣어줌
    # temp_dic_mk[col1] = bgrbuffer
    # temp_dic_mk[col2] = bgebuffer
    temp_dic_mk[col3] = desbuffer
    # temp_dic_mk[col4] = bgfbuffer
    # temp_dic_mk[col5] = bgibuffer
    # temp_dic_mk[col6] = bgxbuffer
    # append를 통해 병합, 인덱스의 이름은 무시
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(1, num_of_pages+1):
    # 필요없는 행이나 열을 제거
    tab_bool, df_tem, del_idx_len = tabu(i)
    if tab_bool:
        # Unnamed: 0, NaN, 반복적인 열과 같은 데이터에 대한 처리
        df_temp_dic, del_len = remo(df_tem, del_idx_len)
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat = mkcell(df_temp_dic, del_len)
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_forcat],ignore_index=True)
    print(i)
print("success")

# 엑셀에 작성하기
excel_path="C:\\datapare\\CL\\2838.xlsx"
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()