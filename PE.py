import pandas as pd
import tabula
import numpy as np
import PyPDF2

# pdf 및 excel 경로 설정
pdf_path = "C:\\datapare\\PE\\PE.pdf"
excel_path = "C:\\datapare\\PE\\PE.xlsx"

# 열을 string 형태로 변환
col2str = {'dtype': str}
kwargs = {'output_format': 'dataframe',
          'pandas_options': col2str,
          'stream': True}

# 전체 페이지 수 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()

top=(16.44*72)/25.4
left=(22.25*72)/25.4
width=(172.58*72)/25.4
height=(271.17*72)/25.4

# *72/25.4 적용한 결과
y1=top
x1=left
y2=top+height
x2=left+width

# area 좌표 데이터 (top, left, top+height, left+width)
p_area=(y1, x1, y2, x2)

col1=(27.52*72)/25.4+left
col2=(132.53*72)/25.4+col1
col3=(12.53*72)/25.4+col2
# cols 좌표 데이터 (col1, col2, col3)
p_cols=(col1, col2, col3)

# 가져올 열 값을 리스트에 추가
col1 = "Código"
col2 = "Designación de la Mercancía"
col3 = "A/V"
col_list=[col1, col2, col3]


def tabu(page_p):
    try:
        ta_df = tabula.read_pdf(pdf_path, columns=p_cols, area=p_area, pages=page_p, **kwargs)
    except:
        print("read_pdf 오류 발생 예외 처리")
        return False, None

    if len(ta_df) == 1:
        # 테이블 하나 갖고 오기
        df_ta = ta_df[0]
        if len(df_ta.columns) == 3:
            # 열 지정
            df_ta.columns = col_list
        else:
            # 열의 개수가 3개가 아닌 경우는 제외
            print("열의 개수 : " + str(len(df_ta.columns)))
            return False, None, 0

        # 불필요한 부분을 제거하기 위해 행의 길이만큼 탐색 진행
        del_index = []
        for i in range(0, len(df_ta)):
            # 해당 지점부터 표 시작
            if df_ta[col1][i] == col1:
                break
            # 지점을 만나기 전까지는 전부 삭제할 행으로 간주
            del_index.append(i)

        # 불필요한 부분이 존재하는지 길이로 확인
        del_index_len = len(del_index)

        # 바로 표 시작 부분일 경우
        if del_index_len == 0:
            df_ta = df_ta.drop(index=[0, 1, 2], axis=0)
            return True, df_ta, del_index_len

        # 불필요한 부분이 존재할 경우
        elif ((del_index_len > 0) and (del_index_len < len(df_ta))):
            df_ta = df_ta.drop(index=del_index, axis=0)
            # 표 시작 부분을 발견했고, 그 아래 열 데이터도 불필요하므로 추가적인 삭제
            df_ta = df_ta.drop(index=[del_index_len, del_index_len + 1, del_index_len + 2], axis=0)
            return True, df_ta, del_index_len

        # 전부 불필요한 경우
        elif del_index_len == len(df_ta):
            return False, None, 0

    # 테이블이 존재하지 않는 경우(예외처리)
    else:
        print("테이블의 개수 : " + str(len(ta_df)))
        return False, None, 0


def mkcell(df_mk, df_del_len):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=[col1, col2, col3])
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    desbuffer = ""
    for i in range(df_del_len + 3, df_mk_len + df_del_len + 3):
        # 맨 뒤로 가져올 때, NaN으로 나오는 경우도 있기 때문에 이를 고려하여 col3에 대해 유의
        a_mkcell = df_mk[col1][i]
        if pd.isnull(df_mk[col2][i]):
            b_mkcell = ""
        else:
            b_mkcell = str(df_mk[col2][i])

        # col1이 비어있지 않거나 col2에 해당하는 값의 맨 앞이 -로 시작하는 경우
        if (not pd.isnull(a_mkcell)) or ((not b_mkcell == "") and (b_mkcell[0] == '-')):
            # i가 df_del_len+3일 경우
            if not i == df_del_len + 3:
                temp_dic_mk[col2] = desbuffer
                df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

            # i가 df_del_len+3이 아닐 경우
            temp_dic_mk[col1] = df_mk[col1][i]
            temp_dic_mk[col2] = df_mk[col2][i]
            temp_dic_mk[col3] = df_mk[col3][i]

            # col3 내의 내용은 줄바꿈된 문장이 존재하기 때문에 다음과 같이 예외적으로 처리
            desbuffer = b_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙음
        else:
            desbuffer = desbuffer + "\n" + b_mkcell
    # 최종적으로 쌓인 desbuffer에 대한 값을 넣어줌
    temp_dic_mk[col2] = desbuffer
    # append를 통해 병합, 인덱스의 이름은 무시
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 비어있는 데이터프레임 준비
col_list=[col1,col2,col3]
df=pd.DataFrame(columns=col_list)

for i in range(13, num_of_pages+1):
    bool_temp, df_tem, del_len=tabu(i)
    if bool_temp:   # 필요한 테이블만
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat = mkcell(df_tem, del_len)
        # NaN 값 ''으로 처리
        df_forcat=df_forcat.fillna('')
        # 열의 이름에 맞추어 테이블 병합, 인덱스 이름은 무시
        df = pd.concat([df,df_forcat], ignore_index=True)
    print(i)

# 엑셀에 작성하기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df.to_excel(writer, index=False)
writer.save()