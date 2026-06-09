import datetime
import pandas as pd
from tksheet import Sheet
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import *
from tkinter import filedialog
import tkinter.font as tkFont
from pathlib import Path
from openpyxl import Workbook
from openpyxl import load_workbook
import warnings
import matplotlib.pyplot as plt

plt.rc('font', family='NanumGothic')  # For Windows

with warnings.catch_warnings(record=True):
    warnings.simplefilter("always")

now = datetime.datetime.now()
today = str(now.strftime("%Y")) + \
    str(now.strftime("%m")) + str(now.strftime("%d"))


def load_sorce_data():
    global deldf, stock_file, stodf, ydf
    global code_sr, color_sr, size_sr, df_w, df_m, df_t, df_w2, df_m2, df_t2

    deldf = pd.DataFrame([])

    stock_file = r"\\NAS451\team451\DB\반품티셔츠재고장.xlsx"
    opt_file = r"\\NAS451\team451\DB\변환코드.xlsx"

    code_sr = pd.read_excel(opt_file, sheet_name="상품코드").iloc[:, 0]
    color_sr = pd.read_excel(opt_file, sheet_name="컬러").iloc[:, 0]
    size_sr = pd.read_excel(opt_file, sheet_name="사이즈").iloc[:, 0]
    df_w = pd.read_excel(opt_file, sheet_name="여자상품코드")
    df_m = pd.read_excel(opt_file, sheet_name="남자상품코드")
    df_t = pd.read_excel(opt_file, sheet_name="상의상품코드")
    df_w2 = pd.read_excel(opt_file, sheet_name="바진여자상품코드")
    df_m2 = pd.read_excel(opt_file, sheet_name="바진남자상품코드")
    df_t2 = pd.read_excel(opt_file, sheet_name="바진상의상품코드")

    stodf = pd.read_excel(stock_file, skiprows=1)
    ydf = pd.DataFrame(stodf, columns=['상품코드', '재고수량'])

def add_file():  # 주문서(택배업로드 화일 불러오기 함수)
    global idx, gDate
    load_file = filedialog.askopenfilename(title="다운로드 주문서를 선택해 주세요.",
                                           filetypes=(
                                               ("모든 파일", "*.*"), ("xlsx 파일", "*.xlsx"), ("xls 파일", "*.xls")),
                                           initialdir='')

    txt_load_file.delete(0, END)
    txt_load_file.insert(END, load_file)
    idx = 0
    ds = load_file.find("_")
    gDate = load_file[ds-8:ds]
    gDate = gDate[0:4] + "-" + gDate[4:6] + "-" + gDate[6:8]
    
    if "주문서" not in txt_load_file.get():
        msgbox.showwarning("경고", "파일의 형식이 올바르지 않습니다.\n사방넷에서 다운로드\n주문서확정관리_다운로드.xlsx 파일을 선택해 주세요")
        return

def reset(): #초기화 함수
    txt_load_file.delete(0, END)
    txt_save_path.delete(0, END)
    txt_vol.delete(0, END)

    tkSheet(frm_exchange, deldf)
    tkSheet(frm_output, deldf)

    progress = 0
    p_var.set(progress)
    progress_bar.update()

def tkSheet(frameName, dfName):
    global tkdata
    dataSheet = dfName.values.tolist()
    col_List = list(dfName.columns)
    frameName.grid_columnconfigure(0, weight=1)
    frameName.grid_rowconfigure(0, weight=1)
    frameName.frame = tk.Frame(frameName)
    frameName.frame.grid_columnconfigure(0, weight=1)
    frameName.frame.grid_rowconfigure(0, weight=1)
    frameName.sheet = Sheet(frameName.frame, data=dataSheet, height=560,
                            headers=col_List, header_height=30, header_fg="#FFFFFF", header_bg="#333333")
    frameName.sheet.header_font(('NanumGothic', 10, 'normal'))
    frameName.sheet.font(('NanumGothic', 10, 'normal'))
    frameName.sheet.table_align(align="left")
    frameName.sheet["A"].align("c")
    frameName.sheet["E:G"].align("c")
    frameName.sheet.set_all_column_widths(
        width=None, only_set_if_too_small=False, redraw=True, recreate_selection_boxes=True)
    frameName.sheet.enable_bindings()
    frameName.frame.grid(row=0, column=0, sticky="nswe")
    frameName.sheet.grid(row=0, column=0, sticky="nswe")


# 품번코드 변환 함수 --------------------------------------
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

def order_list_ext():
    global sdf, xdf, df1, df2
    sFile = txt_load_file.get()  # 주문서(택배업로드파일)명 가져오기
    sdf = pd.read_excel(sFile, skiprows=0)
    xdf = pd.DataFrame(columns=['주문처', '주문자명','상품옵션코드', '수량', "주문상품코드", '품번', "컬러", "사이즈", "재고체크"])
    
    i = 1
    idx = 2
    for r in range(2, sdf.shape[0]-1):
        varR = r % 2
        if varR == 0:
            odSite = sdf.iloc[r, 1] #쇼핑몰
            shopID = sdf.iloc[r+1, 1] #쇼핑몰ID
            odName = sdf.iloc[r+1, 4] #주문자
            odPdname = sdf.iloc[r+1, 6] #옵셥명
            odQty = sdf.iloc[r, 7] #수량
            stock = ""
            x = odPdname

            if odSite == "하프클럽(신)":
                x = x.replace("_", "-").replace(" ", "").replace("/", " : ")

            if odSite == "(주)진마니아":
                del_texts = ["모델명/색상:", "모델명:사이즈:",
                            ",사이즈", "사이즈:", "MODEL:SIZE:", " "]
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

            if odSite == "shop by":
                x = x.replace("|", " ")

            if odSite == "ESM지마켓":
                del_texts = ["1000원", "2000원", "3000원", "4000원",
                            "5000원", "6000원", "7000원", "8000원", "9000원"]
                for del_text in del_texts:
                    x = x.replace(del_text, "")
                x = x.replace(" ", "")
                s = x.find("_")
                e = x.find("/")
                x = x[s+1:e]

            if odSite == "ESM옥션":
                del_texts = ["1000원", "2000원", "3000원", "4000원",
                            "5000원", "6000원", "7000원", "8000원", "9000원"]
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
                x = x.replace("|", ":")

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
                elif code[0:3] in ["TAS", "TAL", "TAG", "TAO", "TAN", "TLJ", "TGJ", "UAA", "UAB", "UAC", "UAD"]:
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
                if x.find(color) < 0 or gCode.find("+") > 0:
                    continue
                gColor = color

            gSize = ""
            for size in size_sr:
                if x.find(str(size)) < 0:
                    continue
                gSize = str(size).strip(" "":""/""+""(")
            
            inCode = gCode
            xgCode = extract_itemcode(inCode)
            odPdname = xgCode + '(' + gColor + ') : ' + gSize
            
            for r in range(0, ydf.shape[0]):
                if odPdname == ydf.iloc[r, 0]:
                    stock = '재고 : ' + str(ydf.iloc[r, 1])
                else:
                    pass
            
            items= [odSite, odName, odPdname, odQty, gCode, xgCode, gColor, gSize, stock]
            xdf.loc[i] = items
            # 추출검증 출력코드 #############################################################################################
            print(odSite + " " + odName + " " + odPdname + " || " + gCode + " > " + xgCode + " " + gColor + " " + gSize)
            ################################################################################################################
            i += 1
        else:
            pass
        idx += 1
        progress = idx / int(sdf.shape[0]-1) * 100
        progress_label.config(text=f"ROW : {r+1}")
        p_var.set(progress)
        progress_bar.update()
        r += 2
                 
    df1 = pd.DataFrame(
        xdf, columns=['주문처', '주문자명', '주문상품코드', '컬러', '사이즈', '수량']).fillna('')
            
    df2 = pd.DataFrame(
        xdf, columns=['주문처', '주문자명', '상품옵션코드', '수량', '재고체크']).fillna('')
    df2 = df2[df2['상품옵션코드'].str.contains('T')]
    
    global order_info
    sum1 = str(df1["수량"].sum())
    sum2 = str(df2["수량"].sum())
    order_info = "총 주문수량 : " + sum1 + "    /    " + "티셔츠 주문수량 : " + sum2
    
    tkSheet(frm_exchange, df1)
    tkSheet(frm_output, df2)
    
    txt_vol.delete(0, END)
    txt_vol.insert(END,order_info)


## 실행함수 #########
def creat_forwarding_list(): # 출고리스트 생성
    load_sorce_data()
    order_list_ext()
    msgbox.showinfo("알림", "티셔츠 출고리스트가 생성되었습니다.")

def copy_forwarding_list():  # 출고리스트 생성
    df2.to_clipboard(index=False, header=False)
    msgbox.showinfo("알림", "티셔츠 출고리스트가 복사되었습니다.")
    
def menual():
    desc = """
    프로그램 실행순서
    1. 사방넷에서 주문관리>주문서확정관리 메뉴에서
    2. 항목 전체 선택 다운로드
    3. 다운로드 파일명 '주문서확정관리_다운로드.xlsx' 불러오기
    4. 출고 작업리스트 생성 실행
    
    방법1. 출고 작업리스트를 복사 엑셀에 붙여넣고 출력하기
    방법2. 아래 저장할 폴더를 선택하고 저장하여 출력하기
    
    오류 체크사항
    1. 코드변환이 실행되지 않을 경우
       주문서 파일을 엑셀형식 문서로 다시 저장해 주세요.
    2. 상품코드 변환 오류가 있을 경우
       NAS451 DB폴더에 변환코드 시트를 수정해 주세요
    """
    msgbox.showinfo("프로그램 사용 설명", desc)

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
    df2.to_excel(excel_writer=save_dir + "/" + today +"_티셔츠출고리스트.xlsx", sheet_name="sheet1", freeze_panes=(1,0), index=False)
    msgbox.showinfo("알림", "티셔츠 출고리스트가 저장되었습니다.")


# TK UI ----------------------------------------------------------------
window = Tk()
window.title("티셔츠 출고 작업리스트 출력 프로그램 v2.04")
window.geometry("1500x900+160+20")
window.resizable(True, True)


relief = "flat"  # flat / groove / raised / ridge /
btn_bg = "#FFFFFF"

### Frame Layout ----------------------------------------------------------------
frm_top = Frame(window, padx=5, pady=2, bg='#FFFFFF',)
frm_top.pack(fill="x", padx=10, pady=2)
frm_run_menu = LabelFrame(window, padx=5, pady=2, relief='groove')
frm_run_menu.pack(fill="x", padx=10, pady=5)
frm_container = Frame(window)
frm_container.pack(fill="both", padx=1, pady=1, ipady=5)
frm_progress = LabelFrame(window, text="진행상황")
frm_progress.pack(fill="x", padx=10, pady=5, ipady=5)
frm_path = LabelFrame(window, text="엑셀 파일로 저장하기 (저장 경로를 선택해 주세요.)")
frm_path.pack(fill="x", padx=10, pady=5, ipady=5)

## Frame top
label_text = today + '_티셔츠 출고리스트를 생성합니다. ※주의 : 파일명에 반드시 주문서 단어가 포함 되어 있어야만 합니다'
label = Label(frm_top, bg='#FFFFFF', text=label_text)
label.pack(side="left", padx=10, pady=5)
txt_vol = Entry(frm_top, width=50, borderwidth=0, fg='red', bg="#FFFFFF", justify='center', insertofftime=600)
txt_vol.pack(side="left", padx=20, pady=5, ipady=3)
btn_close = Button(frm_top, text="프로그램 종료", width=18,
                   relief=relief, fg='#FFFFFF', bg='#333333', command=window.quit)
btn_close.pack(side="right", padx=5, pady=5)
btn_reset = Button(frm_top, width=19, text="프로그램 도움말",
                   relief=relief, fg='#FFFFFF', bg='#333333', command=menual)
btn_reset.pack(side="right", padx=10, pady=5)

## Frame Run_menu
txt_load_file = Entry(frm_run_menu, width=100,
                      relief="flat", background=btn_bg)
txt_load_file.pack(side="left", padx=5, pady=5, ipadx=5, ipady=4)
btn_addfile = Button(frm_run_menu, width=22, text="다운로드 주문서 선택",
                     relief=relief, fg='#FFFFFF', bg='#333333', command=add_file)
btn_addfile.pack(side="left", padx=10, pady=5)
btn_output = Button(frm_run_menu, text="출고리스트 생성하기",
                    relief=relief, width=22, fg='#FFFFFF', bg='#3162C7', command=creat_forwarding_list)
btn_output.pack(side="left", padx=10, pady=5)
btn_output_copy = Button(frm_run_menu, text="출고리스트 복사하기",
                         relief=relief, width=22, fg='#FFFFFF', bg='#2F9D27', command=copy_forwarding_list)
btn_output_copy.pack(side="left", padx=10, pady=5)
btn_reset = Button(frm_run_menu, width=12, text="초기화",
                   relief=relief, fg='#FFFFFF', bg='#333333', command=reset)
btn_reset.pack(side="right", padx=5, pady=5)

## Frame Container > ecchanghe Listbox
frm_exchange = LabelFrame(frm_container, padx=10,
                          pady=10, height=600, text='   1. 전체 주문서 상품 리스트   ')
frm_exchange.pack(side="left", fill="both", padx=10, pady=5, ipady=5, expand=True)

## Frame container > Output Listbox
frm_output = LabelFrame(frm_container, padx=10, pady=10,
                        height=600, text='   2. 티셔츠 출고 리스트   ')
frm_output.pack(side="right", fill="both", padx=10, pady=5, ipady=5, expand=True)

## Frame Progress Bar
p_var = DoubleVar()
progress_bar = ttk.Progressbar(frm_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x", padx=10, pady=5)

progress_label = tk.Label(frm_progress, width="12", text="ROW", relief="flat", background="#FFFFFF")
progress_label.pack(side="right", padx=5, pady=5)

## Frame Path
txt_save_path = Entry(frm_path, relief="flat", background=btn_bg)
txt_save_path.pack(side="left", fill="x", expand=True, padx=10, pady=5, ipady=3)
btn_save_file = Button(frm_path, text="저장하기", width=12,
                       relief=relief, bg=btn_bg, command=save_file)
btn_save_file.pack(side="right", padx=10, pady=5)
btn_destpath = Button(frm_path, text="저장폴더 선택", width=12, relief=relief, bg=btn_bg,
                      command=browse_save_path)
btn_destpath.pack(side="right", padx=10, pady=5)


window.mainloop()
# TK UI ----------------------------------------------------------------------



