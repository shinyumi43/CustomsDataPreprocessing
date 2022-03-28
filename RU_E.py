# 영문 버전에서 사용
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

# 경로만 변경해서 사용
# 읽을 pdf 경로
pdf_path="C:\\datapre\\RU\\RU.pdf"
# 생성할 excel 경로
excel_path="C:\\datapre\\RU\\RU.xlsx"

# 가져올 데이터 내의 열 값을 리스트에 담기
col1="HS Code"
col2="Description"
col3="Add. units"
col4="The rate of the import customs duty (in percentage of customs cost either in euro or in US dollars)"
col_list=[col1,col2,col3,col4]

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()
num_of_pages


# 테이블 정리에 필요한 함수(영문)
def tabu(page_p):
    # pdf 파일 읽기
    ta_df = tabula.read_pdf(pdf_path, multiple_tables=True, pages=page_p, **kwargs)
    # 테이블의 개수가 1개나 2개인 경우
    if len(ta_df) == 1 or len(ta_df) == 2:
        # 테이블의 개수가 1개인 경우
        if len(ta_df) == 1:
            df_ta = ta_df[0]
        # 테이블의 개수가 2개인 경우
        elif len(ta_df) == 2:
            df_ta = ta_df[1]
        # 열의 개수가 4개인 경우
        if len(df_ta.columns) == 4:
            # 열 이름 변경
            df_ta.columns = col_list
            # 추출이 필요한 테이블일 경우
            if df_ta[col1][2] == col1:
                df_ta = df_ta.drop(index=[0, 1, 2, 3, 4, 5], axis=0)
                ta_len = 6
                df_ta = df_ta.fillna('')
                return df_ta, True, ta_len
            elif df_ta[col1][3] == col1:
                df_ta = df_ta.drop(index=[0, 1, 2, 3, 4, 5, 6, 7], axis=0)
                ta_len = 8
                df_ta = df_ta.fillna('')
                return df_ta, True, ta_len
            # 추출이 필요하지 않은 테이블일 경우
            else:
                print("not need to extract")
                return None, False, 0
        # 열의 개수가 5개나 6개인 경우
        elif len(df_ta.columns) == 5 or len(df_ta.columns) == 6:
            # 추출이 필요한 테이블일 경우(테이블의 형식에 따라 분류)
            if df_ta["Unnamed: 0"][2] == col1:
                df_ta = df_ta.drop(index=[0, 1, 2, 3, 4, 5], axis=0)
                ta_len = 6

                # 제거가 필요한 열인지 확인
                del_col_name = []
                for i in range(0, len(df_ta.columns)):
                    # 전부 다 NaN이 담긴 경우, 해당 열은 삭제
                    if df_ta[df_ta.columns[i]].isnull().sum() == len(df_ta):
                        # 삭제가 필요한 열들을 저장
                        del_col_name.append(df_ta.columns[i])

                df_ta = df_ta.fillna('')

                # 테이블에서 삭제할 열을 제거해도 columns의 개수가 4가 되지 않는 경우
                if (len(df_ta.columns) - len(del_col_name)) != 4:
                    df_col_len = len(df_ta.columns)
                    if page_p == 898:
                        # 898 페이지에서 유일하게 나타나는 오류
                        mov_col = df_ta.columns[df_col_len - 3]
                        rec_col = df_ta.columns[df_col_len - 4]
                        df_ta[rec_col] = df_ta[rec_col] + " " + df_ta[mov_col]
                        del_col_name.append(mov_col)
                    else:
                        # 열을 옮길 부분을 지정
                        mov_col = df_ta.columns[df_col_len - 1]
                        rec_col = df_ta.columns[df_col_len - 2]
                        df_ta[rec_col] = df_ta[rec_col] + df_ta[mov_col]
                        del_col_name.append(mov_col)

                # 해당 열 삭제 진행
                for i in range(0, len(del_col_name)):
                    df_ta = df_ta.drop([del_col_name[i]], axis=1)

                # 열의 이름을 변경
                df_ta.columns = col_list
                return df_ta, True, ta_len
            # 추출이 필요한 테이블일 경우(테이블의 형식에 따라 분류)
            elif df_ta["Unnamed: 0"][3] == col1:
                df_ta = df_ta.drop(index=[0, 1, 2, 3, 4, 5, 6, 7], axis=0)
                ta_len = 8
                # 제거가 필요한 열인지 확인
                del_col_name = []
                for i in range(0, len(df_ta.columns)):
                    # 전부 다 NaN이 담긴 경우, 해당 열은 삭제
                    if df_ta[df_ta.columns[i]].isnull().sum() == len(df_ta):
                        # 삭제가 필요한 열들을 저장
                        del_col_name.append(df_ta.columns[i])

                df_ta = df_ta.fillna('')
                # 테이블에서 삭제할 열을 제거해도 columns의 개수가 4가 되지 않는 경우
                if (len(df_ta.columns) - len(del_col_name)) != 4:
                    # 열을 옮길 부분을 지정
                    df_col_len = len(df_ta.columns)
                    mov_col = df_ta.columns[df_col_len - 1]
                    rec_col = df_ta.columns[df_col_len - 2]
                    df_ta[rec_col] = df_ta[rec_col] + df_ta[mov_col]
                    del_col_name.append(mov_col)

                # 해당 열 삭제 진행
                for i in range(0, len(del_col_name)):
                    df_ta = df_ta.drop([del_col_name[i]], axis=1)

                # 열의 이름을 변경
                df_ta.columns = col_list
                return df_ta, True, ta_len
            # 추출이 필요하지 않은 테이블일 경우
            else:
                print("not need to extract")
                return None, False, 0
        # 열의 개수가 3 이하 이거나 7 이상인 경우
        else:
            print("열의 개수 : " + str(len(df_ta.columns)))
            return None, False, 0
    # 테이블의 개수가 0개인 경우
    elif len(ta_df) == 0:
        print("테이블이 존재하지 않음")
        return None, False, 0
    # 테이블의 개수가 3개 이상인 경우
    else:
        print("테이블의 개수 : " + str(len(ta_df)))
        return None, False, 0


def mkcell(df_mk, df_len):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=col_list)
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # 하나의 데이터가 줄바꿈을 기준으로 여러 셀에 삽입되는 것을 처리
    # 빈 문자열 생성
    desbuffer = ""
    dbobuffer = ""
    bgobuffer = ""
    # 삭제한 행을 고려하여 시작점은 6이나 8
    for i in range(df_len, df_mk_len + df_len):
        # col2, col4의 값이 여러 셀의 삽입되는 경우가 발생하므로 아래와 같이 처리
        a_mkcell = df_mk[col1][i]
        if df_mk[col2][i] == '':
            b_mkcell = ""
        else:
            b_mkcell = str(df_mk[col2][i])
        if df_mk[col3][i] == '':
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])
        if df_mk[col4][i] == '':
            d_mkcell = ""
        else:
            d_mkcell = str(df_mk[col4][i])

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if (not a_mkcell == '') or ((not b_mkcell == "") and (b_mkcell[0] == '-')):
            if not i == df_len:
                temp_dic_mk[col2] = desbuffer
                temp_dic_mk[col3] = dbobuffer
                temp_dic_mk[col4] = bgobuffer
                # append 함수는 표 내의 데이터를 합치는 함수
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)
                # temp_dic_mk 에 전달 받은 df_mk 값 삽입
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]
            temp_dic_mk[col4] = df_mk[col4][i]

            desbuffer = b_mkcell
            dbobuffer = c_mkcell
            bgobuffer = d_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙기
        else:
            desbuffer = desbuffer + "\n" + b_mkcell
            dbobuffer = dbobuffer + "\n" + c_mkcell
            bgobuffer = bgobuffer + "\n" + d_mkcell
    # 최종적으로 쌓인 desbuffer, bgobuffer를 col2, col4에 삽입
    temp_dic_mk[col2] = desbuffer
    temp_dic_mk[col3] = dbobuffer
    temp_dic_mk[col4] = bgobuffer
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 빈 데이터프레임 생성
df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(1, num_of_pages + 1):
    # 필요없는 행이나 열을 제거
    df_tem, tab_bool, dk_len = tabu(i)
    if tab_bool:
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat = mkcell(df_tem, dk_len)
        # NaN 값 ''으로 처리
        df_forcat = df_forcat.fillna('')
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_forcat], ignore_index=True)
        print(i)

# xlsxwriter 엔진으로 pandas writer 객체 만들기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
# dataframe을 xlsx에 쓰기
df.to_excel(writer, index=False)
# pandas excel writer을 닫고, 엑셀 파일을 출력
writer.save()
