import pandas as pd
import pickle

# 엑셀 파일 경로와 저장할 피클 파일 경로를 지정합니다.
excel_path = r'\\NAS451\team451\DB\통합매출데이터.xlsx'
pickle_path = r'\\NAS451\team451\DB\통합매출데이터.pickle'

try:
    # 엑셀 파일을 DataFrame으로 읽어옵니다.
    df = pd.read_excel(excel_path, engine='openpyxl')
    # DataFrame을 피클 파일로 저장합니다.
    with open(pickle_path, 'wb') as fw:
        pickle.dump(df, fw, protocol=pickle.HIGHEST_PROTOCOL)
    print("피클 파일이 성공적으로 생성되었습니다.")
except Exception as e:
    print(f"피클 파일 생성 중 오류 발생: {e}")
print(df)