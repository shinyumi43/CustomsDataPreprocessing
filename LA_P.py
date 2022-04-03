# pdf 내의 데이터 추출을 위한 라이브러리
import pandas as pd
import numpy as np
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
pdf_path="C:\\datapare\\LA\\LA_P.pdf"
# 생성할 excel 경로
excel_path="C:\\datapare\\LA\\LA_P_xlsx"

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()


def tabu(page_p):
    # 열에 들어갈 값 설정
    col1 = "AHTN 2017 CODES"
    col2 = "NaN"
    col3 = "ລາຍການສນຄາ"
    col4 = "DESCRIPTIONS"
    col5 = "NaN2"

    ta_df = tabula.read_pdf(pdf_path, multiple_tables=True, pages=page_p, **kwargs)
    for i in range(0, len(ta_df)):
        df_ta = ta_df[i]

        # 열의 개수가 11개인 경우
        if ((len(df_ta.columns) == 11) and (page_p != 9)):
            # 필요 없는 열 삭제
            del_col = []
            for i in range(4, len(df_ta.columns)):
                del_col.append(df_ta.columns[i])

            df_ta = df_ta.drop(columns=del_col)

            # 해당 경우에 대한 열의 값
            col_list = [col1, col2, col3, col4]
            # 열 내에 삽입되어야 할 데이터가 포함된 경우
            if df_ta.columns[0] != 'AHTN 2017 CODES':
                # 열 내의 담긴 데이터를 행에 담기도록 처리
                df = pd.DataFrame(columns=col_list)

                # 임시 데이터프레임 내에 삽입
                df_tem = {col1: '', col2: '', col3: '', col4: ''}
                df_tem[col1] = df_ta.columns[0]
                df_tem[col2] = df_ta.columns[1]
                df_tem[col3] = df_ta.columns[2]
                df_tem[col4] = df_ta.columns[3]

                # 임시 데이터프레임과 결과 데이터프레임에 통합
                df = df.append(df_tem, ignore_index=True)

            # 본래 바로 추출된 데이터프레임의 열을 교체
            df_ta.columns = col_list

            # 최종적으로 세 개의 열로 구성하기 위해 두 열의 데이터를 결합
            for i in range(0, len(df_ta)):
                if not pd.isnull(df_ta[col2][i]):
                    df_ta[col1][i] = str(df_ta[col1][i]) + str(df_ta[col2][i])

            # 결과 데이터프레임에 통합
            df = pd.concat([df, df_ta], ignore_index=True)

            # df의 NaN은 삭제 진행
            df = df.drop(columns=col2, axis=0)

            return True, df

        # 열의 개수가 11개이고, page가 9인 경우(column 형태가 유일하게 다름)
        elif ((len(df_ta.columns) == 11) and (page_p == 9)):
            # 필요 없는 열 삭제
            del_col = []
            for i in range(4, len(df_ta.columns)):
                del_col.append(df_ta.columns[i])

            df_ta = df_ta.drop(columns=del_col)

            # 해당 경우에 대한 열의 값
            col_list = [col1, col3, col2, col4]

            # 열 내에 삽입되어야 할 데이터가 포함된 경우
            if df_ta.columns[0] != 'AHTN 2017 CODES':
                # 열 내의 담긴 데이터를 행에 담기도록 처리
                df = pd.DataFrame(columns=col_list)

                # 임시 데이터프레임 내에 삽입
                df_tem = {col1: '', col3: '', col2: '', col4: ''}
                df_tem[col1] = df_ta.columns[0]
                df_tem[col3] = df_ta.columns[1]
                df_tem[col2] = df_ta.columns[2]
                df_tem[col4] = df_ta.columns[3]

                # 임시 데이터프레임과 결과 데이터프레임에 통합
                df = df.append(df_tem, ignore_index=True)

            # 본래 바로 추출된 데이터프레임의 열을 교체
            df_ta.columns = col_list

            # 결과 데이터프레임에 통합
            df = pd.concat([df, df_ta], ignore_index=True)

            # df의 NaN은 삭제 진행
            df = df.drop(columns=col2, axis=0)

            return True, df

        # 열의 개수가 9개인 경우
        elif len(df_ta.columns) == 9:
            # 필요 없는 열 삭제
            del_col = []
            for i in range(3, len(df_ta.columns)):
                del_col.append(df_ta.columns[i])

            df_ta = df_ta.drop(columns=del_col)
            # 해당 경우에 대한 열의 값
            col_list = [col1, col3, col4]

            df = pd.DataFrame(columns=col_list)

            # 본래 바로 추출된 데이터프레임의 열을 교체
            df_ta.columns = col_list

            # 결과 데이터프레임에 통합
            df = pd.concat([df, df_ta], ignore_index=True)

            return True, df

        # 열의 개수가 13개일 경우
        elif len(df_ta.columns) == 13:
            # 필요 없는 열 삭제
            del_col = []
            for i in range(5, len(df_ta.columns)):
                del_col.append(df_ta.columns[i])

            df_ta = df_ta.drop(columns=del_col)
            # 해당 경우에 대한 열의 값
            col_list = [col1, col2, col3, col5, col4]

            if df_ta.columns[0] != 'AHTN 2017 CODES':
                # 열 내의 담긴 데이터를 행에 담기도록 처리
                df = pd.DataFrame(columns=col_list)

                # 임시 데이터프레임 내에 삽입
                df_tem = {col1: '', col2: '', col3: '', col5: '', col4: ''}
                df_tem[col1] = df_ta.columns[0]
                df_tem[col2] = df_ta.columns[1]
                df_tem[col3] = df_ta.columns[2]
                df_tem[col5] = df_ta.columns[3]
                df_tem[col4] = df_ta.columns[4]

                # 임시 데이터프레임과 결과 데이터프레임에 통합
                df = df.append(df_tem, ignore_index=True)
            # 본래 바로 추출된 데이터프레임의 열을 교체
            df_ta.columns = col_list
            # 결과 데이터프레임에 통합
            df = pd.concat([df, df_ta], ignore_index=True)
            # df의 NaN과 NaN2는 삭제 진행
            df = df.drop(columns=[col2, col5], axis=0)

            return True, df

    # 위와 같은 경우에 해당하지 않는 경우
    return False, None


def mkcell(df_mk):
    # 한 행을 구성하는데에 사용
    col1 = "AHTN 2017 CODES"
    col2 = "ລາຍການສນຄາ"
    col3 = "DESCRIPTIONS"
    col_list = [col1, col2, col3]

    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=col_list)
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # 하나의 데이터가 줄바꿈을 기준으로 여러 셀에 삽입되는 것을 처리
    # 빈 문자열 생성
    # up=0
    blabuffer = ""
    desbuffer = ""
    bgobuffer = ""
    for i in range(0, df_mk_len):
        if ((pd.isnull(df_mk[col1][i])) or (df_mk[col1][i] == 'Unnamed: 0') or (df_mk[col1][i] == 'Unnamed: 1') or (
                df_mk[col1][i] == 'Unnamed: 2') or (df_mk[col1][i] == 'Unnamed: 3')):
            a_mkcell = ""
        else:
            a_mkcell = str(df_mk[col1][i])
        if ((pd.isnull(df_mk[col2][i])) or (df_mk[col2][i] == 'Unnamed: 0') or (df_mk[col2][i] == 'Unnamed: 1') or (
                df_mk[col2][i] == 'Unnamed: 2') or (df_mk[col2][i] == 'Unnamed: 3')):
            b_mkcell = ""
        else:
            b_mkcell = str(df_mk[col2][i])
        if ((pd.isnull(df_mk[col3][i])) or (df_mk[col3][i] == 'Unnamed: 0') or (df_mk[col3][i] == 'Unnamed: 1') or (
                df_mk[col3][i] == 'Unnamed: 2') or (df_mk[col3][i] == 'Unnamed: 3')):
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])

        if a_mkcell == 'AHTN 2017 CODES':
            a_mkcell = ""
            b_mkcell = ""
            c_mkcell = ""

        # if ((not c_mkcell=="") and (c_mkcell[0].isupper()) and (not c_mkcell[0]=='-')):
        # up=up+1

        if ((not a_mkcell == "") or ((not c_mkcell == "") and (c_mkcell[0] == '-')) or (
                (not c_mkcell == "") and (c_mkcell[0].isupper()))):
            if not i == 0:
                temp_dic_mk[col1] = blabuffer
                temp_dic_mk[col2] = desbuffer
                temp_dic_mk[col3] = bgobuffer
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)
                # 넣어주고, 조건식은 초기화
                # up=0
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]

            blabuffer = a_mkcell
            desbuffer = b_mkcell
            bgobuffer = c_mkcell
        else:
            blabuffer = blabuffer + "\n" + a_mkcell
            desbuffer = desbuffer + "\n" + b_mkcell
            bgobuffer = bgobuffer + "\n" + c_mkcell

    temp_dic_mk[col1] = blabuffer
    temp_dic_mk[col2] = desbuffer
    temp_dic_mk[col3] = bgobuffer
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 최종 데이터프레임 생성
col1="AHTN 2017 CODES"
col2="ລາຍການສນຄາ"
col3="DESCRIPTIONS"
col_list=[col1, col2, col3]
df=pd.DataFrame(columns=col_list)

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

# xlsxwriter 엔진으로 pandas writer 객체 만들기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
# dataframe을 xlsx에 쓰기
df.to_excel(writer, index=False)
# pandas excel writer을 닫고, 엑셀 파일을 출력
writer.save()