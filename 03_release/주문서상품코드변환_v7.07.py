import os
import sys
import datetime
import pandas as pd
import tkinter
from tksheet import Sheet
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import *
from tkinter import filedialog
import warnings
import matplotlib.pyplot as plt
import traceback

if getattr(sys, 'frozen', False):
    #test.exe로 실행한 경우,test.exe를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    #python test.py로 실행한 경우,test.py를 보관한 디렉토리의 full path를 취득
    program_directory = os.path.dirname(os.path.abspath(__file__))
#현재 작업 디렉토리를 변경
os.chdir(program_directory)

#글꼴 설정
plt.rc('font', family='NanumGothic')

with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")

deldf = pd.DataFrame([])

now = datetime.datetime.now()
thisyear = now.strftime('%Y')
# print(thisyear)

# 경로 목록을 리스트로 정의
dir_paths = [
    r"\\NAS451\team451",
    r"\\192.168.0.101\team451",
    r"n:\개인\nSync\Coding"
    r"d:\hSync\Coding"
]

# 유효한 경로를 찾을 때까지 반복
path = next((p for p in dir_paths if os.path.isdir(p)), None)

if path is None:
    print("유효한 경로를 찾을 수 없습니다.")
else:
    path

sfile_dir = path + r"\04-주문택배업로드\택배업로드 및 발주서\\" + thisyear
stock_file = path + r"\DB\반품티셔츠재고장.xlsx"
opt_file = path + r"\DB\변환코드.xlsx"

code_sr = pd.read_excel(opt_file, sheet_name="상품코드").iloc[:, 0]
color_sr = pd.read_excel(opt_file, sheet_name="컬러").iloc[:, 0]
size_sr = pd.read_excel(opt_file, sheet_name="사이즈").iloc[:, 0]
df_w = pd.read_excel(opt_file, sheet_name="여자상품코드")
df_m = pd.read_excel(opt_file, sheet_name="남자상품코드")
df_t = pd.read_excel(opt_file, sheet_name="상의상품코드")
df_w2 = pd.read_excel(opt_file, sheet_name="바진여자상품코드")
df_m2 = pd.read_excel(opt_file, sheet_name="바진남자상품코드")
df_t2 = pd.read_excel(opt_file, sheet_name="바진상의상품코드")

def import_order(): # 주문서 가져오기(택배업로드 화일 불러오기 함수)
    global gDate
    if combo.get() == " +수집 프로그램 선택":
        msgbox.showwarning("경고", "수집 프로그램을 선택해 주세요")
        return
    
    load_file = filedialog.askopenfilename(title="다운로드 주문서를 선택해 주세요.", \
        filetypes=(("모든 파일", "*.*"),("xlsx 파일", "*.xlsx"), ("xls 파일", "*.xls")), \
        initialdir=sfile_dir)

    txt_load_file.delete(0, END)
    txt_load_file.insert(END, load_file)

    ds = load_file.find("_")
    gDate = load_file[ds-8:ds]
    gDate = gDate[0:4] + "-" + gDate[4:6] + "-" + gDate[6:8]

def reset(): #초기화 함수
    txt_load_file.delete(0, END)
    txt_save_path.delete(0, END)
    txt_vol.delete(0, END)

    output_sheet(frame_L1, deldf)
    output_sheet(frame_L2, deldf)
    output_sheet(frame_R1, deldf)
    output_sheet(frame_R2, deldf)

    progress = 0
    p_var.set(progress)
    progress_bar.update()

def order_product_code_extraction(): #주문서에서 상품코드 추출 함수
    global sdf, code_sr, color_sr, size_sr, df_w, df_m, df_t, df_w2, df_m2, df_t2
    global odNo, odName, odPdname, odQty, odSite, gCode, optCode, gColor, gSize, gCode2, optCode2, sum1

    sFile = txt_load_file.get() # 주문서(택배업로드파일)명 가져오기
    sdf = pd.read_excel(sFile) # 주문서엑셀파일 데이터프레임

    sdf["상품코드"] = ""
    sdf["컬러"] = ""
    sdf["사이즈"] = ""
    sdf["옵션코드"] = ""
    sdf["상품코드2"] = ""
    sdf["옵션코드2"] = ""

    if combo.get() == " 2.이지어드민":
        sdf = sdf.reindex(columns=["수령자", "전화번호", "핸드폰", "우편번호", "주소", "옵션",
                                   "주문수량", "특이사항", "판매처", "배송비", "상품번호", "주문번호", "주문자명", 
                                   "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2"])
        sdf = sdf.rename(columns={"수령자": "성명","옵션":"상품명","주문수량":"수량","특이사항":"배송메세지","판매처":"주문처"})
        sdf['수량'] = sdf['수량'].astype(int)
    else:
        sdf = sdf.reindex(columns=["성명", "전화번호", "핸드폰", "우편번호", "주소", "상품명", 
                                   "수량", "배송메세지", "주문처", "요금구분", "운송장번호", "사방넷주문번호", "쇼핑몰아이디", 
                                   "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2"])

    sum1 = str(sdf["수량"].sum())
    idx = 1
    
    for r in range(0, sdf.shape[0]):
        odNo = idx # 넘버링
        odName = sdf.iloc[r, 0] # 성명(주문자명)
        odPdname = sdf.iloc[r, 5] # 상품명
        odQty = sdf.iloc[r, 6] # 수량
        odSite = sdf.iloc[r, 8] # 주문처
        gCode = sdf.iloc[r, 13] # 추출 상품코드
        gColor = sdf.iloc[r, 14] # 추출 컬러
        gSize = sdf.iloc[r, 15] # 추출 사이즈
        optCode = sdf.iloc[r, 16] # 옵션코드
        gCode2 = sdf.iloc[r, 17] # 상품코드2
        optCode2 = sdf.iloc[r, 18] # 옵션코드2

        x = odPdname

        if odSite == "하프클럽(신)":
            x = x.replace("_", "-").replace(" ", "").replace("/", " : ")

        if odSite == "(주)진마니아":
            del_texts = ["모델명/색상:", "모델명:사이즈:", ",사이즈", "사이즈:", "MODEL:SIZE:", " "]
            for del_text in del_texts:
                x = x.replace(del_text, "")
            x = x.replace(",", ":").replace(":", " : ")

        if odSite == "롯데닷컷" or odSite == "롯데홈쇼핑(신)":
            del_texts = ["모델명/색상:", "모델명:", ",사이즈", "MODEL:", ",SIZE"]
            for del_text in del_texts:
                x = x.replace(del_text, "")

        if odSite == "현대홈쇼핑(신)":
            x = x.replace(":", " : ").replace("/", " : ")

        if odSite == "패션플러스":
            x = x.replace(" (", "(").replace(" ", " : ")

        if odSite == "GS shop":
            x = x.replace(",", ":")
            
        if odSite == "ESM지마켓":
            del_texts = ["1000원", "2000원", "3000원", "4000원","5000원", "6000원", "7000원", "8000원", "9000원"]
            for del_text in del_texts:
                x = x.replace(del_text, "")
            x = x.replace(" ", "")
            s = x.find("_")
            e = x.find("/")
            x = x[s+1:e]

        if odSite == "ESM옥션":
            del_texts = ["1000원", "2000원", "3000원","4000원", "5000원", "6000원", "7000원", "8000원", "9000원"]
            for del_text in del_texts:
                x = x.replace(del_text, "")
            x = x.replace(" ", "")
            s = x.find("_")
            e = x.find("[")
            x = x[s+1:e]

        if odSite == "11번가":
            s = x.find("_")
            e = x.find("+")
            x = x[s+1:e-1]

        if odSite == "쿠팡":
            if x.startswith(("MBP")):
                x = x
            else:
                x = "+" + x

        if odSite == "스마트스토어":
            del_texts = ["모델명/색상:", " / 사이즈", " "]
            for del_text in del_texts:
                x = x.replace(del_text, "")
            x = x.replace(":", " : ")
            
        if odSite == "티몬":
            x = x.replace("|", "")

        if odSite == "T deal":
            x = x.replace("모델명(색상)|사이즈:", "")
            x = x.replace("|", ":")

        gCode = ""
        

        for code in code_sr:
            if x.find(code) < 0:
                continue
            s = x.find(code)

            if len(code) > 7:
                if code[0:1] in ["R"]:
                    code = x[s:s+13]
                elif code[0:1] in ["W", "B"]:
                    code = x[s:s+15]
                else:
                    code = x[s:s+22]           
            elif code[:4] in ["TPNS", "TPWS", "TPXS"]:
                code = x[s:s+7]
            elif code[0:3] in ["TAS", "TAL", "TAG", "TAO", "TAN", "TLJ", "TGJ", "UAA", "UAB", "UAC", "UAD", "UAF", "SPH"]:
                code = x[s:s+7]
            elif code[0:2] in ["TP", "DP", "EP", "CP", "EA"]:
                code = x[s:s+7]
            elif code[0:1] in ["T", "D", "E", "R"]:
                code = x[s:s+6]
            else:
                code = x[s:s+7]
            gCode = code.upper()
        
        gColor = ""
        for color in color_sr:
            if x.find(color) > 0:
                gColor = color
            else:
                continue

        gSize = ""
        for size in size_sr:
            if x.find(str(size)) < 0:
                continue
            gSize = str(size).strip(" "":""/""+""(")

        sdf.iloc[r, 13] = gCode
        sdf.iloc[r, 14] = gColor
        sdf.iloc[r, 15] = gSize

def set_product_separation(): #세트상품 상품분리 행추가 함수
    global sdf2
    sdf2 = pd.DataFrame(columns=sdf.columns)
    i = 0

    cnt = int(sdf.shape[0])
    for r in range(0, cnt):
        codeVar = sdf.iloc[r, 13]
        colorVar = sdf.iloc[r, 14]
        cutA = int(codeVar.find('+'))
        cutB = int(colorVar.find("+"))
        set1 = codeVar[:cutA]
        set2 = codeVar[cutA+1:]
        # if codeVar[0:4] in ["TPNS", "TPWS", "TPXS"]:
        if colorVar.find('+') > 0:
            S = colorVar.split("+")
            for c in range(0, len(S)):
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0, 19)]
                sdf2.iloc[i, 13] = codeVar[:7]
                sdf2.iloc[i, 14] = S[c]
                i += 1
        elif codeVar.find('+') > 0:
            if codeVar[0:4] == "BHP":
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0, 19)]
                sdf2.iloc[i, 13] = codeVar[:cutA]
                i += 1
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0, 19)]
                sdf2.iloc[i, 13] = codeVar[cutA+1:cutA+8]
                i += 1
            if codeVar[0:6] == "TB-001":
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0, 19)]
                sdf2.iloc[i, 13] = set1[:6]
                sdf2.iloc[i, 14] = set1[7:int(set1.find(")"))]
                i += 1
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0, 19)]
                sdf2.iloc[i, 13] = set2[:6]
                sdf2.iloc[i, 14] = set2[7:int(set2.find(")"))]
                i += 1
            else:
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0,  19)]
                sdf2.iloc[i, 13] = str(set1)
                i += 1
                sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0,  19)]
                sdf2.iloc[i, 13] = str(set2)
                i += 1
        else:
            sdf2.loc[i] = [sdf.iloc[r, a] for a in range(0,  19)]
            i += 1
def product_code_conversion(): # 상품코드 변환함수
    cnt = sdf2.shape[0]
    for e in range(0, cnt):
        sdf2.iloc[e, 17] = extract_itemcode(sdf2.iloc[e, 13])

def ouput_upload_excel_list(): # 엑셀업로드 리스트 출력 함수
    idx = 1
    cnt = int(sdf2.shape[0])
    for r in range(0, cnt):
        gCode = sdf2.iloc[r, 13]
        gColor = sdf2.iloc[r, 14]
        gSize = sdf2.iloc[r, 15]
        # print(r)

        if gColor and gCode.startswith(("T", "U", "BHP-7", "BHP-5", "WJL-0", "EA", "DP", "EP", "CP", "DL0", "EL0", "RL0", "WBL", "SPH")):
            optCode = f"{gCode}({gColor}) : {gSize}"
        elif gCode[3] == "T" or gCode[4] == "T":
            optCode = f"{gCode}({gColor}) : {gSize}"
        else:
            optCode = f"{gCode} : {gSize}"
        sdf2.iloc[r, 16] = optCode
        
        gCode2 = extract_itemcode(sdf2.iloc[r, 13])
        sdf2.iloc[r, 17] = gCode2
        if gColor and gCode2.startswith(("T", "U", "BHP-7", "BHP-5", "WJL-0", "EA", "DP", "EP", "CP", "DL0", "EL0", "RL0", "WBL", "SPH")):
            optCode2 = f"{gCode2}({gColor}) : {gSize}"
        else:
            optCode2 =f"{gCode2} : {gSize}"
        
        sdf2.iloc[r, 18] = optCode2
        idx += 1
        progress = idx / int(sdf.shape[0]) * 100
        progress_label.config(text=f"ROW : {r+1}")
        p_var.set(progress)
        progress_bar.update()

    global sdf2_separation
    sdf2_separation = sdf2.loc[:,["성명", "상품명", "옵션코드", "수량", "주문처", "상품코드2", "컬러", "사이즈", "옵션코드2"]]
    output_sheet(frame_L1, sdf2_separation) # 상품코드 추출리스트

def extract_itemcode(inCode):
    inCode = inCode.upper()
    prefix_map = {
        "S": "JS", "U": "JU", "A": "BA", "C": "BC", 
        "L": "JL", "M": "JM", "N": "NA", "J": "EJ",
        "F": "JF"
    }

    if inCode.startswith(("W", "M", "T", "U", "BHP", "GS", "DS", "CS", "PAC", "SPH")):
        if inCode[3] == "A":
            goodNum = inCode[4:]
            return f"{inCode[:3]}-{goodNum}"
        if inCode[3] =="F":
            goodNum = inCode[4:]
            return f"{inCode[0]}JF-{goodNum}"
        elif inCode.startswith(("TS")): ##
            goodNum = inCode[3:]
            return f"TPS-{goodNum}"
        elif inCode.startswith(("TPNS")):
            goodNum = inCode[4:]
            return f"{inCode[:3]}-{goodNum}"
        elif inCode.startswith(("TGJS", "TLJS")):
            goodNum = inCode[4:]
            return f"{inCode[:2]}-{goodNum}"
        elif inCode.startswith(("UAA-", "UAB-", "UAC-", "UAD-", "UAE-", "UAF-", "UAR-")):
            goodNum = inCode[4:]
            return f"{inCode[:3]}O{goodNum}"
        return inCode
    elif inCode.startswith(("C", "D", "E", "H", "R", "Z")):
        if inCode[1:3] == "P7":
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"BHP-{goodNum}"
        elif inCode[3]== "T" and inCode[:2] == "RS": ##
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"TPS-{goodNum}"
        elif inCode[1:] in ("PN7T13"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"BHP-{goodNum}"
        elif inCode[1:3] in ("PP", "PK"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"BHP-{goodNum}"
        elif inCode[1:3] in ("PL", "PG", "PM", "PH"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"
        elif inCode[3] in ("T", "U"):
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"{inCode[3]}{inCode[1]}-{goodNum}"
        elif inCode[3] in ("W", "M"):
            goodNum = inCode[2] + inCode[4] + inCode[5]
            return f"{inCode[3]}{prefix_map.get(inCode[1], 'XX')}-{goodNum}"
        elif inCode[4] in ("T"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"
        elif inCode[4] in ("U"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}{inCode[1]}{inCode[2]}O{goodNum}"
        elif inCode[4] in ("W", "M"):
            goodNum = inCode[3] + inCode[5] + inCode[6]
            return f"{inCode[4]}BP-{goodNum}"
    return ""

def output_sale_data(): # 판매데이터 출력 함수
    # product_code_conversion() # 옵션코드 변환함수

    global df_sale_data
    idx = 1
    df_sale_data = sdf2.loc[:,["주문처", "상품코드2", "컬러", "사이즈", "수량"]]
    df_sale_data = df_sale_data.groupby(["주문처", "상품코드2", "컬러", "사이즈"], as_index=False).agg({'수량':'sum'})
    
    df_sale_data["날짜"] = gDate
    cnt = df_sale_data.shape[0]
    df_sale_data = df_sale_data[["날짜", "상품코드2", "컬러", "사이즈", "수량", "주문처"]]
    for r in range(0, cnt):
        df_sale_data.iloc[r, 1] = df_sale_data.iloc[r, 1]
        idx += 1
       
    progress = 0    
    global sum2
    sum2 = str(df_sale_data["수량"].sum())
    sum_txt = "주문서 주문수량 : " + sum1 + "    /    " + "총 출고수량 : " + sum2
    txt_vol.delete(0, END)
    txt_vol.insert(END, sum_txt)

    output_sheet(frame_L2, df_sale_data) # 판매데이터

def create_forwarding_list():  
    # 전체 출고리스트 생성---------------
    progress = 0
    p_var.set(progress)
    progress_bar.update()

    global pdf
    pdf = sdf2.pivot_table("수량", index="옵션코드2", aggfunc="sum")
    pdf = pdf.reset_index()

    global odf
    odf = pd.DataFrame(columns=["옵션코드", "수량"])
    odf.reset_index()
    idx = 1
    for p in range(0, int(pdf.shape[0])):
        eNo = idx
        eCode = pdf.iloc[p, 0]
        Ea = pdf.iloc[p, 1]
        box_list2 = [eCode, Ea]
        odf.loc[eNo-1] = box_list2
        idx += 1

        progress = idx / int(pdf.shape[0]) * 100
        p_var.set(progress)
        progress_bar.update()

    st_df = pd.read_excel(stock_file, skiprows=1)
    st_df = st_df.loc[:, ["상품코드", "재고수량"]]
    st_df.reset_index()

    output_sheet(frame_R1, pdf) # 전체상품 출고리스트
    
    progress = 0
    p_var.set(progress)
    progress_bar.update()

    # 티셔츠 출고리스트 생성----------
    global ndf
    ndf = pd.DataFrame(columns=["주문자명", "옵션코드", "수량", "재고"])
    ndf.reset_index()
    idx = 1
    nu = 1
    for n in range(0, int(sdf2.shape[0])):
        if sdf2.iloc[n, 18][0] == "T" or sdf2.iloc[n, 18][0] == "S":
            eNo = nu
            eName = sdf2.iloc[n, 0]
            eCode = sdf2.iloc[n, 18]
            Ea = sdf2.iloc[n, 6]
            sEa = ""
            for k in range(0, st_df.shape[0]):
                if sdf2.iloc[n, 18] == st_df.iloc[k, 0]:
                    sEa = st_df.iloc[k, 1]
            box_list3 = [eName, eCode, Ea, sEa]
            ndf.loc[eNo-1] = box_list3
            nu += 1
        idx += 1

        progress = idx / int(sdf2.shape[0]) * 100
        p_var.set(progress)
        progress_bar.update()

    output_sheet(frame_R2, ndf) # 티셔츠 출고리스트

def output_sheet(frame, dfName): # tksheet 출력함수
    sheet = dfName.values.tolist()
    col_List = list(dfName.columns)
    # if dfName == df10:
    #     col_width = None
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    frame = tk.Frame(frame)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)
    sheet = Sheet(frame, data=sheet, height=520,
                            headers=col_List, header_height=28, header_bg="#444444", header_fg="#FFFFFF")
    sheet.header_font(('NanumGothic', 10, 'bold'))
    sheet.font(('NanumGothic', 9, 'normal'))
    sheet.set_all_column_widths(
        width=None, only_set_if_too_small=False, redraw=True, recreate_selection_boxes=True)
    sheet.set_all_row_heights(
        height = 25, only_set_if_too_small = False, redraw = True, recreate_selection_boxes = True)
    sheet.enable_bindings()
    frame.grid(row=0, column=0, sticky="nswe")
    sheet.grid(row=0, column=0, sticky="nswe")

# 실행함수 #########################################################################################
def output_product_code_extraction(): # 
    if combo.get() == " +수집 프로그램 선택":
        msgbox.showwarning("경고", "수집 프로그램을 선택해 주세요")
        return
    elif combo.get() == " 1.사방넷" and "우체국택배업로드" not in txt_load_file.get():
        msgbox.showwarning("경고", "우체국택배업로드 주문서를 선택해 주세요")
        return
    elif combo.get() == " 2.이지어드민" and "이지어드민" not in txt_load_file.get():
        msgbox.showwarning("경고", "이지어드민 주문서를 선택해 주세요")
        return
    
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일을 선택해 주세요.")
        return
    
    order_product_code_extraction() # 주문서 상품코드 추출 함수
    set_product_separation() # 세트상품 분리 행추가 함수
    ouput_upload_excel_list() # 엑셀업로드 리스트 출력 함수

    global df_excel_upload # 엑셀 업로드 파일 데이터프레임 변환
    df_excel_upload = sdf2.loc[:, ["성명", "전화번호", "핸드폰", "우편번호", "주소", "옵션코드",
        "수량", "배송메세지", "주문처", "요금구분", "운송장번호", "사방넷주문번호", "쇼핑몰아이디"]]
    df_excel_upload = df_excel_upload.rename(columns={"옵션코드": "상품명"})
    msgbox.showinfo("알림", "주문서에서 상품코드 추출 변환이 완료 되었습니다.")

def create_factory_list():  # 출고리스트 생성 출력 함수cls
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
        return
    
    output_sale_data()
    create_forwarding_list()
    
    msgbox.showinfo("알림", "출고리스트 생성이 완료 되었습니다")


####################################################################
def code_ext_copy(): # 업로드 상품코드 복사
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
        return
    df_excel_upload.to_clipboard(index=False)
    msgbox.showinfo("알림", "상품코드 및 수량 복사완료. 엑셀시트에 붙여넣기 해 주세요")

def output_list1_copy():  # 전체 출고리스트 복사
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환 및 출고리스트 생성을 실행해 주세요.")
        return
    odf.to_clipboard(index=False)
    msgbox.showinfo("알림", "출고리스트 복사완료")

def output_list2_copy():  # 티셔츠 출고리스트 복사
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환 및 출고리스트 생성을 실행해 주세요.")
        return
    ndf.to_clipboard(index=False)
    msgbox.showinfo("알림", "티셔츠 출고리스트 복사완료")


def sale_data_copy():  # 매출데이터 복사
    if len(txt_load_file.get()) == 0:
        msgbox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
        return
    df_sale_data.to_clipboard(index=False)
    msgbox.showinfo("알림", "매출데이터 복사완료")

def menual():
    desc = """
    주문서 파일명 체크항목
    1. 사방넷 다운로드 주문서
       - 파일명 형식 예)20250301_우체국택배업로드.xlsx
    2. 이지어드민 다운로드 주문서
       - 파일명 형식 예)20250301_이지어드민.xls
    
    오류 체크사항
    1. 코드변환이 실행되지 않을 경우 
       주문서 파일을 엑셀형식 문서로 다시 저장해 주세요.
    2. 상품코드 변환 오류가 있을 경우
       NAS451 DB폴더에 변환코드 시트를 수정해 주세요

    프로그램 실행순서
    1. 주문서 선택 실행
    2. 상품코드 변환 실행
    3. 출고리스트 생성 실행
    4. 상품코드 복사 > 업로드시트에 붙여넣기
    5. 출고리스트 복사 > 출고리스트시트에 붙여널기
    """
    msgbox.showinfo("알림", desc)

def browse_save_path():
    folder_selected = filedialog.askdirectory()
    if folder_selected == "":  # 사용자가 취소를 누를 때
        print("폴더 선택 취소")
        return
    txt_save_path.delete(0, END)
    txt_save_path.insert(0, folder_selected)

def save_file():
    if len(txt_save_path.get()) == 0:
        msgbox.showwarning("경고", "저장할 위치를 선택해 주세요.")
        return
    save_dir = txt_save_path.get()
    df_excel_upload.to_excel(save_dir + "/택배업로드.xlsx", sheet_name="택배업로드", index=False)
    # odf.to_excel(save_dir + "/출고리스트.xlsx", sheet_name="출고리스트", index=False)
    # ndf.to_excel(save_dir + "/출고리스트.xlsx", sheet_name="티셔츠출고리스트", index=False)
    msgbox.showinfo("알림", "출고리스트가 저장되었습니다.")


# TK UI ----------------------------------------------------------------
window = Tk()
window.title("주문서 상품코드 변환 & 출고리스트 생성 프로그램 v7.07")
window.geometry("1600x850+20+20")
window.minsize(1600, 850)
window.state('zoomed')
window.resizable(True, True)

## frame topmenu
btn_relief = "flat" #flat, groove, raised, ridge, solid, sunken
btn_bg = "#333333"
btn_fg = "#FFFFFF"

frm_topmenu = Frame(window)
frm_topmenu.pack(fill="x", padx=5, pady=5)

combo_list = [" +수집 프로그램 선택", " 1.사방넷" , " 2.이지어드민"]
combo = ttk.Combobox(frm_topmenu, state="readonly", values=combo_list, width=19)
combo.current(0)
combo.pack(side="left", padx=5, pady=5, ipady=3)

txt_load_file = Entry(frm_topmenu, relief="flat", background="#FFFFFF")
txt_load_file.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4)

btn_addfile = Button(frm_topmenu, width=15, text="주문서 선택", relief=btn_relief, bg=btn_bg, fg=btn_fg, command=import_order)
btn_addfile.pack(side="left", padx=5, pady=5)

btn_close = Button(frm_topmenu, text="닫기", width=12,
                   relief=btn_relief, bg=btn_bg, fg=btn_fg, command=window.quit)
btn_close.pack(side="right", padx=5, pady=5)

btn_reset = Button(frm_topmenu, width=12, text="초기화",
                   relief=btn_relief, bg=btn_bg, fg=btn_fg, command=reset)
btn_reset.pack(side="right", padx=5, pady=5)

btn_output = Button(frm_topmenu, text="출고리스트 생성",
                    relief=btn_relief, bg=btn_bg, fg=btn_fg, width=14, command=create_factory_list)
btn_output.pack(side="right", padx=5, pady=5)

btn_code_ext = Button(frm_topmenu, text="상품코드 추출",
                      relief=btn_relief, bg=btn_bg, fg=btn_fg, width=14, command=output_product_code_extraction)
btn_code_ext.pack(side="right", padx=5, pady=5)

## frame container
frm_container = Frame(window)
frm_container.pack(fill="both", padx=1, pady=1, ipady=5, expand=True)

## frame container > ecchanghe Listbox
frm_exchange = Frame(frm_container, width=1100)
frm_exchange.pack(side="left", fill="both", padx=5, pady=5, ipady=5, expand=True)

notebook2 = ttk.Notebook(frm_exchange, width=1100)
notebook2.pack(fill="both", padx=3, pady=3, ipady=2, expand=True)

frame_L1 = tkinter.Frame(frm_exchange)
notebook2.add(frame_L1, text="  상품코드 추출리스트  ", padding=2)

frame_L2 = tkinter.Frame(frm_exchange)
notebook2.add(frame_L2, text="  판매데이터  ", padding=2)

## frame container > output Listbox
frm_output = Frame(frm_container)
frm_output.pack(side="right", fill="both", padx=5, pady=5, ipady=5, expand=True)

## frame container > output Listbox
notebook2 = ttk.Notebook(frm_output, width=500)
notebook2.pack(fill="both", padx=3, pady=3, ipady=2, expand=True)

frame_R1 = tkinter.Frame(frm_output)
notebook2.add(frame_R1, text="  전체상품 출고리스트  ", padding=2)

frame_R2 = tkinter.Frame(frm_output)
notebook2.add(frame_R2, text="  티셔츠 츨고리스트  ", padding=2)

## frame Progress Bar
frm_progress = LabelFrame(window, text="진행상황")
frm_progress.pack(fill="x", padx=5, pady=5, ipady=5)

p_var = DoubleVar()
progress_bar = ttk.Progressbar(frm_progress, maximum=100, variable=p_var)
progress_bar.pack(side="left", fill="x", expand=True, padx=5, pady=5)

progress_label = tk.Label(frm_progress, width="12", text="ROW", relief="flat", background="#FFFFFF")
progress_label.pack(side="right", padx=5, pady=5)


## frame path
frm_path = LabelFrame(window, text="업로드 엑셀시트 파일로 저장하기 (저장경로를 입력해 주세요!)")
frm_path.pack(fill="x", padx=5, pady=5, ipady=5)

txt_save_path = Entry(frm_path, relief="flat", background="#FFFFFF")
txt_save_path.pack(side="left", fill="x", expand=True, padx=5, pady=5, ipady=4)

btn_save_file = Button(frm_path, text="저장하기", width=12,
                       relief=btn_relief, bg=btn_bg, fg=btn_fg, command=save_file)
btn_save_file.pack(side="right", padx=5, pady=5)

btn_destpath = Button(frm_path, text="찾아보기", width=12, relief=btn_relief, bg=btn_bg, fg=btn_fg,
                      command=browse_save_path)
btn_destpath.pack(side="right", padx=5, pady=5)

## frame run
frm_run = Frame(window)
frm_run.pack(fill="x", padx=5, pady=5, ipady=5)

txt_vol = Entry(frm_run, width=50, borderwidth=0, background="#EFEFEF", insertofftime=600)
txt_vol.pack(side="left", padx=5, pady=5, ipady=3)

btn_reset = Button(frm_run, width=12, text="사용설명",
                   relief=btn_relief, bg=btn_bg, fg=btn_fg, command=menual)
btn_reset.pack(side="right", padx=5, pady=5)

btn_saledata_copy = Button(frm_run, text="판매데이터 복사",
                           width=20, relief=btn_relief, bg=btn_bg, fg=btn_fg, command=sale_data_copy)
btn_saledata_copy.pack(side="right", padx=5, pady=5)

btn_output_list2_copy = Button(frm_run, text="티셔츠 출고리스트 복사",
                               width=20, relief=btn_relief, bg=btn_bg, fg=btn_fg, command=output_list2_copy)
btn_output_list2_copy.pack(side="right", padx=5, pady=5)

btn_output_list1_copy = Button(frm_run, text="전체 출고리스트 복사",
                               width=20, relief=btn_relief, bg=btn_bg, fg=btn_fg, command=output_list1_copy)
btn_output_list1_copy.pack(side="right", padx=5, pady=5)

btn_code_ext = Button(frm_run, text="업로드 상품코드 복사", width=20,
                      relief=btn_relief, bg=btn_bg, fg=btn_fg, command=code_ext_copy)
btn_code_ext.pack(side="right", padx=5, pady=5)

window.mainloop()