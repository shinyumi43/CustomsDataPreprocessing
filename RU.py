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

def tabu(page_p):
    # pdf 파일 읽기
    ta_df = tabula.read_pdf(pdf_path, multiple_tables=True, pages=page_p, **kwargs)
    # 테이블이 존재하는 경우
    if len(ta_df)==1 or len(ta_df)==2:
        # 테이블 하나 갖고 오기
        df_ta = ta_df[0]
        # 테이블이 2개인데, ta_df[1]에 추출해야할 테이블이 존재하는 경우
        if len(ta_df)==2:
            if ((not df_ta.columns[0]=='Unnamed: 0') or (len(df_ta) < 9) or (not len(df_ta.columns)==4)):
                df_ta = ta_df[1]
        # 열의 개수가 4개인 경우
        if len(df_ta.columns)==4:
            # 추출하고자 하는 열로 구성하기|
            df_ta.columns=col_list
            # 필요없는 행 제거하기
            df_ta = df_ta.drop(index=[0,1,2,3,4,5,6,7,8,9], axis=0)
            return True, df_ta
        elif len(df_ta.columns)==5 and df_ta.columns[0]=='Unnamed: 0':
            # 불필요한 열을 제거하기
            df_ta = df_ta.drop(columns='Unnamed: 2')
            # 추출하고자 하는 열로 구성하기
            df_ta.columns=col_list
            # 필요없는 행 제거하기
            df_ta = df_ta.drop(index=[0,1,2,3,4,5,6,7,8,9], axis=0)
            return True, df_ta
        else:
            # 열의 개수가 4개나 5개가 아닌 경우(예외처리)
            print("열의 개수 : "+str(len(df_ta.columns)))
            return False, None
    # 테이블이 존재하지 않는 경우(예외처리)
    else:
        print("테이블의 개수 : "+str(len(ta_df)))
        return False, None


def mkcell(df_mk):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=col_list)
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # 하나의 데이터가 줄바꿈을 기준으로 여러 셀에 삽입되는 것을 처리
    # 빈 문자열 생성
    desbuffer = ""
    bgobuffer = ""
    dbobuffer = ""
    # 삭제한 행을 고려하여 시작점은 10
    for i in range(10, df_mk_len + 10):
        # col2, col4의 값이 여러 셀의 삽입되는 경우가 발생하므로 아래와 같이 처리
        a_mkcell = df_mk[col1][i]
        if pd.isnull(df_mk[col2][i]):
            b_mkcell = ""
        else:
            b_mkcell = str(df_mk[col2][i])
        if pd.isnull(df_mk[col3][i]):
            c_mkcell = ""
        else:
            c_mkcell = str(df_mk[col3][i])
        if pd.isnull(df_mk[col4][i]):
            d_mkcell = ""
        else:
            d_mkcell = str(df_mk[col4][i])

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if (not pd.isnull(a_mkcell)) or ((not b_mkcell == "") and (b_mkcell[0] == '–')) or (
                ((len(d_mkcell) == 3) or (len(d_mkcell) == 4)) and ((d_mkcell[2] == 'С') or (d_mkcell[2] == ')'))):
            if not i == 10:
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

start=1
end=97

for i in range(start, end+1):
    # 읽을 pdf 경로
    pdf_path = "C:\\datapare\\RU\\pdf\\" + str(i) + ".pdf"
    # 생성할 excel 경로
    excel_path = "C:\\datapare\\RU\\excel\\" + str(i) + ".xlsx"

    # 가져올 데이터 내의 열 값을 리스트에 담기
    col1 = "Код ТН ВЭД"
    col2 = "Наименование позиции"
    col3 = "Доп. ед. изм."
    col4 = "Ставка ввозной таможенной пошлины (в процентах от таможенной стоимости либо в евро, либо в долларах США)"
    col_list = [col1, col2, col3, col4]

    # 전체 페이지 수를 확인
    num_of_pages = 0
    with open(pdf_path, 'rb') as f:
        read_pdf = PyPDF2.PdfFileReader(f)
        num_of_pages = read_pdf.getNumPages()

    # 빈 데이터프레임 생성
    df = pd.DataFrame(columns=col_list)

    # 전체 페이지에 대하여 데이터 추출을 반복
    for j in range(1, num_of_pages + 1):
        # 필요없는 행이나 열을 제거
        tab_bool, df_tem = tabu(j)
        if tab_bool:
            # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
            df_forcat = mkcell(df_tem)
            # NaN 값 ''으로 처리
            df_forcat = df_forcat.fillna('')
            # 여러 페이지에서 추출한 표를 하나의 표에 합침
            df = pd.concat([df, df_forcat], ignore_index=True)
        print(j)

    print(str(i) + "p success")

    # xlsxwriter 엔진으로 pandas writer 객체 만들기
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    # dataframe을 xlsx에 쓰기
    df.to_excel(writer, index=False)
    # pandas excel writer을 닫고, 엑셀 파일을 출력
    writer.save()