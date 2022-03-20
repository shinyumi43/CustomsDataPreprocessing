import pandas as pd
import tabula
import PyPDF2

col2str = {'dtype': str}
kwargs = {
          'pandas_options': col2str,
          'stream': True}

pdf_path="C:\\datapre\\MM\\1.pdf"
excel_path="C:\\datapre\\MM\\1.xlsx"

col1="WCO H.S Code"
col2="AHTN Code"
col3="Stat.Code"
col4="MACS Code"
col5="Description"
col6="Unit"
col7="MCT Rate(%)"
col_list=[col1, col2, col3, col4, col5, col6, col7]

# 예외적으로 추가된 열에 대한 처리
col8="del_col1"
col9="del_col2"
col_del1=[col1, col2, col3, col4, col5, col6, col7, col8, col9] # col의 개수 9
col_del2=[col1, col2, col3, col4, col5, col6, col7, col8] # col의 개수 8


def tabu(page_p):
    ta_df = tabula.read_pdf(pdf_path, lattice=True, pages=page_p, **kwargs)

    # table이 존재하지 않는 경우, false를 반환
    if not len(ta_df) == 1:
        print(len(ta_df))
        return False, None
    else:
        if ta_df[0].columns[0] == 'WCO\rH.S\rCode':
            df_ta = ta_df[0]
            if len(df_ta.columns) == 7:
                df_ta.columns = col_list
                df_ta = df_ta.drop(columns=[col4, col7])
                df_ta = df_ta.fillna('')
                return True, df_ta
            else:
                if len(df_ta.columns) == 8:
                    df_ta.columns = col_del2
                    df_ta = df_ta.drop(columns=[col4, col7, col8])
                    df_ta = df_ta.fillna('')
                    return True, df_ta
                elif len(df_ta.columns) == 9:
                    df_ta.columns = col_del1
                    df_ta = df_ta.drop(columns=[col4, col7, col8, col9])
                    df_ta = df_ta.fillna('')
                    return True, df_ta
                else:
                    print("col의 개수 : " + str(len(df_ta.columns)))
                    return False, df_ta

        else:
            return False, None

tab_bool,df_origin = tabu(27)
col_ori=df_origin.columns

for i in range(28, num_of_pages + 1):
    tab_bool, df_tem = tabu(i)
    if tab_bool:
        df_tem.columns = col_ori
        df_origin = pd.concat([df_origin, df_tem], ignore_index=True)
    print(i)
print("success")

writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
df_origin.to_excel(writer, index = False)
writer.save()