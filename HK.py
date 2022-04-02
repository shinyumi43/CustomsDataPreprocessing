import pandas as pd
import numpy as np
import tabula
# pdf 변환 시에 사용하는 라이브러리
import PyPDF2

# pdf 내용 그대로 추출하는 코드
# 엑셀은 그대로 추출한 것과 hs code만 정제한 것 두 가지로 업로드

pdf_path = 'C:\\datapre\\HK\\HK.pdf'
excel_path = 'C:\\datapre\\HK\\HK.xlsx'
col2str = {'dtype': str}
kwargs = {'output_format': 'dataframe',
          'pandas_options': col2str,
          'stream': True}

# 전체 페이지 수를 확인
num_of_pages=0
with open(pdf_path, 'rb') as f:
    read_pdf = PyPDF2.PdfFileReader(f)
    num_of_pages = read_pdf.getNumPages()
num_of_pages

# 첫 페이지와 마지막 페이지를 제외하고의 테이블의 형식
top=(20.13*72)/25.4
left=(148.61*72)/25.4
width=(147.95*72)/25.4
height=(174.01*72)/25.4

col1=(24.09*72)/25.4+left
col2=(46.03*72)/25.4+col1
col3=(59.01*72)/25.4+col2
col4=(18.82*72)/25.4+col3

y1=top
x1=left
y2=top+height
x2=left+width

p_cols=(col1, col2, col3, col4)
p_area=(y1, x1, y2, x2)

# 첫 페이지 테이블의 형식
ftop=(39.02*72)/25.4
fheight=(152.35*72)/25.4

y2=ftop+fheight
f_area=(y1, x1, y2, x2)

# 47 페이지의 예외적인 형식
ntop=(0*72)/25.4
nheight=(193.41*72)/25.4

y2=ntop+nheight
n_area=(y1, x1, y2, x2)

col1="類/章/註釋/港貨協制 編號 Sect/Ch/Note/HKHS Code"
col2="描述/貨物名稱"
col3="Description"
col4="數量 單位 Unit of Quantity"
col_list=[col1, col2, col3, col4]

def tabu(page_p):
    # pdf 파일 읽기
    if page_p == 3:
        ta_df = tabula.read_pdf(pdf_path, columns=p_cols, area=f_area, pages=page_p, **kwargs)
        df_ta = ta_df[0]
        # 불필요한 행 제거
        df_ta = df_ta.drop(index=[0,1,2,3,4,5,6,7,8], axis=0)
        # 열 이름 변경
        df_ta.columns = col_list
        # 난수 공백으로 변경
        df_ta = df_ta.fillna('')
        # 삭제 열 기준 설정
        df_len = 9
        return df_ta, True, df_len
    elif page_p == 47:
        # 추출할 내용이 담긴 테이블 형식과 다름
        return None, False, 0
    else:
        ta_df = tabula.read_pdf(pdf_path, columns=p_cols, area=p_area, pages=page_p, **kwargs)
        df_ta = ta_df[0]
        # 불필요한 행 제거
        df_ta = df_ta.drop(index=[0,1,2,3,4,5], axis=0)
        # 열 이름 변경
        df_ta.columns = col_list
        # 난수 공백으로 변경
        df_ta = df_ta.fillna('')
         # 삭제 열 기준 설정
        df_len = 6
        return df_ta, True, df_len


def mkcell(df_mk, mk_len):
    # pandas dataframe 형태로 변환
    df_temp_mk = pd.DataFrame(columns=col_list)
    # column 값을 지정
    temp_dic_mk = {col1: '', col2: '', col3: '', col4: ''}
    # 행의 길이를 저장
    df_mk_len = len(df_mk)

    # 하나의 데이터가 줄바꿈을 기준으로 여러 셀에 삽입되는 것을 처리
    # 빈 문자열 생성
    bufbuffer = ""
    desbuffer = ""
    dbobuffer = ""
    bgobuffer = ""
    # 삭제한 행을 고려하여 시작점은 6이나 9
    for i in range(mk_len, df_mk_len + mk_len):
        # col2, col4의 값이 여러 셀의 삽입되는 경우가 발생하므로 아래와 같이 처리
        if df_mk[col1][i] == '':
            a_mkcell = ""
        else:
            a_mkcell = str(df_mk[col1][i])
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
        if (((not a_mkcell == '') and ((a_mkcell[0] == '註') or (a_mkcell[0] == '第') or (a_mkcell[0] == '分')))
                or ((not a_mkcell == '') and (len(a_mkcell) == 4))
                or ((not b_mkcell == "") and (b_mkcell[0] == '-'))):
            if not i == mk_len:
                temp_dic_mk[col1] = bufbuffer
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

            bufbuffer = a_mkcell
            desbuffer = b_mkcell
            dbobuffer = c_mkcell
            bgobuffer = d_mkcell
        # 앞이 -로 시작하지 않으면, 줄바꿈한 뒤에 붙기
        else:
            bufbuffer = bufbuffer + "\n" + a_mkcell
            desbuffer = desbuffer + "\n" + b_mkcell
            dbobuffer = dbobuffer + "\n" + c_mkcell
            bgobuffer = bgobuffer + "\n" + d_mkcell
    # 최종적으로 쌓인 desbuffer, bgobuffer를 col2, col4에 삽입
    temp_dic_mk[col1] = bufbuffer
    temp_dic_mk[col2] = desbuffer
    temp_dic_mk[col3] = dbobuffer
    temp_dic_mk[col4] = bgobuffer
    df_temp_mk = df_temp_mk.append(temp_dic_mk, ignore_index=True)

    return df_temp_mk

# 빈 데이터프레임 생성
df=pd.DataFrame(columns=col_list)

# 전체 페이지에 대하여 데이터 추출을 반복
for i in range(3, num_of_pages):
    # 필요없는 행이나 열을 제거
    df_tem, tab_bool, dk_len = tabu(i)
    if tab_bool:
        # 하나의 데이터가 여러 셀에 삽입되는 것에 대한 처리
        df_forcat = mkcell(df_tem, dk_len)
        # NaN 값 ''으로 처리
        df_forcat = df_forcat.fillna('')
        # 여러 페이지에서 추출한 표를 하나의 표에 합침
        df = pd.concat([df, df_forcat],ignore_index=True)
        print(i)

# 마지막 테이블 형식
top=(37.54*72)/25.4
left=(148.61*72)/25.4
width=(147.95*72)/25.4
height=(28.04*72)/25.4

c1=(20.12*72)/25.4+left
c2=(51.06*72)/25.4+c1
c3=(69.86*72)/25.4+c2

yl1=top
xl1=left
yl2=top+height
xl2=left+width

l_cols=(c1, c2, c3)
l_area=(yl1, xl1, yl2, xl2)

# 다른 형식의 테이블 열 설정
lcol1="編號 Code"
lcol2="國家/地區"
lcol3="Country/Territory"
lcol_list=[lcol1, lcol2, lcol3]

# 빈 데이터프레임 생성
ta_df = tabula.read_pdf(pdf_path, columns=l_cols, area=l_area, pages=num_of_pages, **kwargs)
ta = ta_df[0]
ta = ta.drop(index=[0,1], axis=0)
ta.columns=lcol_list

# xlsxwriter 엔진으로 pandas writer 객체 만들기
writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
# dataframe을 xlsx에 쓰기
df.to_excel(writer, sheet_name='Table 1', index=False)
ta.to_excel(writer, sheet_name='Table 2', index=False)
# pandas excel writer을 닫고, 엑셀 파일을 출력
writer.save()