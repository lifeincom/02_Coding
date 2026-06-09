import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import customtkinter as ctk
import pandas as pd
import itertools
from typing import List, Dict, Optional, Any, Union
import webbrowser
import os
import tempfile
from datetime import datetime
from PIL import Image, ImageTk

# tksheet 라이브러리 체크
try:
    from tksheet import Sheet
    HAS_TKSHEET = True
except ImportError:
    HAS_TKSHEET = False
    print("tksheet 라이브러리가 필요합니다. 설치 명령: pip install tksheet")

# --- Configuration ---
ctk.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

COLUMN_KEYWORDS = {
    'item_code': ['상품코드', '모델명', '관리코드'],
    'color': ['옵션컬러', '색상', '컬러', 'Color'],
    'size': ['옵션사이즈', '사이즈', 'Size'],
    'price': ['행사가', '할인가', '가격', '판매가', 'Price'],
    'option_no': ['옵션번호', '상품번호', '순번', 'No']
}

DEFAULT_BASE_PRICE = ""
DEFAULT_STOCK_QTY = 500
TEMPLATE_COLUMNS = ["상품번호", "상품코드", "상품명", "컬러", "사이즈", "정상가", "할인가", "쿠폰적용가", "이미지"]
TEMPLATE_WIDTHS = [8, 10, 30, 30, 10, 10, 10]


# --- Logic Class ---
class DealOptionLogic:
    """
    Handles data processing logic for GMarket Deal Options.
    """
    @staticmethod
    def normalize_cell(val: Any) -> str:
        """데이터 정제 함수 (공백, 소수점 제거)"""
        if pd.isna(val) or val is None:
            return ""
        s = str(val).strip()
        if s.endswith(".0"): 
            s = s[:-2]
        return s

    @staticmethod
    def find_column_name(columns: list, candidates: list) -> Optional[str]:
        """컬럼명 자동 찾기 (유사한 이름 검색)"""
        for col in columns:
            col_str = str(col).strip()
            for cand in candidates:
                if cand in col_str: 
                    return col
        return None

    def process_data(self, df: pd.DataFrame, base_price: int, option_type: str = "3S") -> pd.DataFrame:
        """
        Processes the input DataFrame and generates the option combination DataFrame.
        """
        headers = df.columns.tolist()
        
        # 필수 컬럼 찾기
        col_option_no = self.find_column_name(headers, COLUMN_KEYWORDS['option_no'])
        col_item_code = self.find_column_name(headers, COLUMN_KEYWORDS['item_code'])
        col_color = self.find_column_name(headers, COLUMN_KEYWORDS['color'])
        col_size = self.find_column_name(headers, COLUMN_KEYWORDS['size'])
        col_price = self.find_column_name(headers, COLUMN_KEYWORDS['price'])

        # 검증
        if not col_color or not col_size:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다.\n필요 컬럼: {COLUMN_KEYWORDS['color']} 중 하나, {COLUMN_KEYWORDS['size']} 중 하나")
        
        results = []
        
        for index, row in df.iterrows():
            try:
                # 데이터 정제 및 추출
                raw_colors = self.normalize_cell(row.get(col_color, ''))
                raw_sizes = self.normalize_cell(row.get(col_size, ''))
                
                if not raw_colors or not raw_sizes:
                    continue
                
                # 콤마로 구분된 옵션 분리
                colors = [self.normalize_cell(c) for c in raw_colors.split(',') if c.strip()]
                sizes = [self.normalize_cell(s) for s in raw_sizes.split(',') if s.strip()]

                # 옵션 번호 및 코드
                val_no = row.get(col_option_no, '') if col_option_no else str(index + 1)
                val_code = row.get(col_item_code, '') if col_item_code else ""
                
                str_no = self.normalize_cell(val_no)
                str_code = self.normalize_cell(val_code)
                
                # 옵션번호 포맷팅 (2자리)
                if not str_no: str_no = "00"
                prefix = str_no.zfill(2)
                
                # 옵션명 포맷팅
                final_option_name = f"{prefix}_{str_code}" if str_code else f"{prefix}_CodeMissing"

                # 가격 계산
                try:
                    p_val_str = str(row.get(col_price, '0')).replace(",", "")
                    p_val = int(float(p_val_str))
                except (ValueError, TypeError):
                    p_val = 0
                    
                add_price = p_val - base_price

                # 조합 생성 Loop
                if option_type == "2S":
                    # 2S: Color 무시, Size를 옵션명2로 설정
                    for size in sizes:
                        item = {
                            '옵션타입': option_type,
                            '옵션명1': final_option_name, 
                            '옵션명2': size,
                            # 옵션명3 제거
                            '영문옵션명1': final_option_name,
                            '영문옵션명2': size,
                            # 영문옵션명3 제거
                            '중문옵션명1': final_option_name,
                            '중문옵션명2': size,
                            # 중문옵션명3 제거
                            '상태': '정상',
                            '노출여부': 'Y',
                            '추가금액': add_price,
                            '관리코드': str_code,
                            '재고수량(A)': DEFAULT_STOCK_QTY,
                            '재고수량(G)': DEFAULT_STOCK_QTY
                        }
                        results.append(item)
                else:
                    # 3S (기본): Color x Size combination
                    for color, size in itertools.product(colors, sizes):
                        item = {
                            '옵션타입': option_type,
                            '옵션명1': final_option_name, 
                            '옵션명2': color,
                            '옵션명3': size,
                            '영문옵션명1': final_option_name,
                            '영문옵션명2': color,
                            '영문옵션명3': size,
                            '중문옵션명1': final_option_name,
                            '중문옵션명2': color,
                            '중문옵션명3': size,
                            '상태': '정상',
                            '노출여부': 'Y',
                            '추가금액': add_price,
                            '관리코드': str_code,
                            '재고수량(A)': DEFAULT_STOCK_QTY,
                            '재고수량(G)': DEFAULT_STOCK_QTY
                        }
                        results.append(item)

            except Exception as e:
                print(f"Error processing row {index}: {e}")
                continue
                
        if not results:
            raise ValueError("생성된 옵션이 없습니다. 데이터 형식을 확인해주세요.")
            
        return pd.DataFrame(results)


# --- UI Components (Frames) ---

class OptionMakerFrame(ctk.CTkFrame):
    """
    Tab 1: Main Option Maker
    """
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.logic = DealOptionLogic()
        
        # Grid Layout Config
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Main content area expands
        
        # 1. Top Control Panel
        self.ctrl_frame = ctk.CTkFrame(self, corner_radius=10)
        self.ctrl_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
        
        self._init_controls()
        
        # 2. Main Sheet Area (Notebook styled via CTkTabview or manual switching)
        # Using inner Tabview for Input/Output visualization
        self.sheet_tabview = ctk.CTkTabview(self)
        self.sheet_tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        self.sheet_tabview.add("  Input Data  ")
        self.sheet_tabview.add("  Output Result  ")
        
        self._init_sheets()
        
    def _init_controls(self):
        # Card 1: Load Data
        f1 = ctk.CTkFrame(self.ctrl_frame, fg_color="transparent")
        f1.pack(side="left", padx=10, pady=10, fill="y")
        
        ctk.CTkLabel(f1, text="1. Data Source", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        
        self.btn_load = ctk.CTkButton(f1, text="📂 Open Excel/CSV", command=self.load_file, fg_color="#E0a800", hover_color="#c69500", width=140)
        self.btn_load.pack(side="left", padx=(0, 5), pady=5)
        
        self.btn_template = ctk.CTkButton(f1, text="⬇️ Template", command=self.download_template, fg_color="gray", width=100)
        self.btn_template.pack(side="left", pady=5)
        
        self.lbl_file_info = ctk.CTkLabel(f1, text="No File Selected", text_color="gray")
        self.lbl_file_info.pack(side="left", anchor="w", padx=5)

        # Separator
        ttk.Separator(self.ctrl_frame, orient="vertical").pack(side="left", fill="y", padx=5, pady=10)

        # Card 2: Settings
        f2 = ctk.CTkFrame(self.ctrl_frame, fg_color="transparent")
        f2.pack(side="left", padx=10, pady=10, fill="y")
        
        ctk.CTkLabel(f2, text="2. Settings", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        
        row_s = ctk.CTkFrame(f2, fg_color="transparent")
        row_s.pack(fill="x", padx=(0, 5), pady=5)
        ctk.CTkLabel(row_s, text="Base Price:").pack(side="left")
        
        self.var_base_price = ctk.StringVar(value=str(DEFAULT_BASE_PRICE))
        self.entry_price = ctk.CTkEntry(row_s, textvariable=self.var_base_price, width=80)
        self.entry_price.pack(side="left", pady=5, padx=5)
        ctk.CTkLabel(row_s, text="KRW").pack(side="left")
        
        ctk.CTkLabel(row_s, text="    Option Type:").pack(side="left")
        self.cbo_option_type = ctk.CTkOptionMenu(row_s, values=["2S", "3S"], width=80)
        self.cbo_option_type.set("3S")
        self.cbo_option_type.pack(side="left", pady=5, padx=5)
        
        # self.lbl_price_range = ctk.CTkLabel(f2, text="", text_color="#1f77b4", font=("Segoe UI", 11))
        # self.lbl_price_range.pack(anchor="w", pady=2)

        # Separator
        ttk.Separator(self.ctrl_frame, orient="vertical").pack(side="left", fill="y", padx=5, pady=10)

        # Card 3: Actions
        f3 = ctk.CTkFrame(self.ctrl_frame, fg_color="transparent")
        f3.pack(side="left", padx=10, pady=10, fill="both", expand=True)
        
        ctk.CTkLabel(f3, text="3. Actions", font=("Segoe UI", 13, "bold")).pack(anchor="w")
        
        self.btn_run = ctk.CTkButton(f3, text="⚙️ Run Conversion", command=self.run_conversion, state="disabled", width=160)
        self.btn_run.pack(side="left", padx=(0, 10), pady=10)
        
        self.btn_save = ctk.CTkButton(f3, text="💾 Save Result", command=self.save_file, state="disabled", fg_color="#28a745", hover_color="#218838", width=160)
        self.btn_save.pack(side="left", pady=10)
        
        self.btn_reset = ctk.CTkButton(f3, text="🔄 Reset", command=self.reset_all, fg_color="transparent", border_width=1, text_color=("gray10", "gray90"), width=80)
        self.btn_reset.pack(side="right", anchor="n")

    def _init_sheets(self):
        if not HAS_TKSHEET:
            ctk.CTkLabel(self.sheet_tabview.tab("  Input Data  "), text="Error: tksheet not installed").pack()
            return

        # Use frame to hold tksheet to prevent resizing issues
        frame_input = ctk.CTkFrame(self.sheet_tabview.tab("  Input Data  "))
        frame_input.pack(fill="both", expand=True)
        
        self.sheet_input = Sheet(frame_input, headers=["Load File First"])
        self.sheet_input.set_all_column_widths(width=None)
        self.sheet_input.enable_bindings()

        self.sheet_input.pack(fill="both", expand=True)

        frame_output = ctk.CTkFrame(self.sheet_tabview.tab("  Output Result  "))
        frame_output.pack(fill="both", expand=True)
        
        self.sheet_output = Sheet(frame_output, headers=["Result will appear here"])
        self.sheet_output.set_all_column_widths(width=None)
        self.sheet_output.enable_bindings()
        self.sheet_output.pack(fill="both", expand=True)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")])
        if not file_path: return

        try:
            if file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            
            df.columns = df.columns.str.strip()
            df = df.map(lambda x: DealOptionLogic.normalize_cell(x))
            
            if HAS_TKSHEET:
                self.sheet_input.headers(df.columns.tolist())
                self.sheet_input.set_sheet_data(df.values.tolist())
            
            self.lbl_file_info.configure(text=f"...{file_path[-30:]} ({len(df)} rows)")
            
            # Auto Calc Price Range
            col_price = DealOptionLogic.find_column_name(df.columns.tolist(), COLUMN_KEYWORDS['price'])
            if col_price:
                try:
                    price_series = df[col_price].astype(str).str.replace(',', '', regex=True)
                    price_numeric = pd.to_numeric(price_series, errors='coerce')
                    min_p = price_numeric.min()
                    max_p = price_numeric.max()
                    if pd.notna(min_p) and pd.notna(max_p):
                        self.lbl_price_range.configure(text=f"(Price Range: {int(min_p):,} ~ {int(max_p):,})")
                except:
                    pass
            
            self.btn_run.configure(state="normal")
            self.sheet_tabview.set("  Input Data  ")
            
        except Exception as e:
            messagebox.showerror("Load Failed", str(e))

    def run_conversion(self):
        if not HAS_TKSHEET: return
        data = self.sheet_input.get_sheet_data()
        if not data: return

        try:
            headers = self.sheet_input.headers()
            df_input = pd.DataFrame(data, columns=headers)
            
            try:
                base_price_str = self.var_base_price.get().replace(",", "")
                base_price = int(base_price_str) if base_price_str else 0
            except ValueError:
                messagebox.showerror("Error", "Check Base Price")
                return
            
            # Base Price 값이 없거나 0일 경우 경고 메시지 출력
            if base_price == 0:
                if not messagebox.askokcancel("Warning", "Base Price 값이 설정되지 않았습니다.\n추가금액이 정확하지 않을 수 있습니다.\n계속 진행하시겠습니까?"):
                    return

            option_type = self.cbo_option_type.get()
            df_result = self.logic.process_data(df_input, base_price, option_type)
            
            self.sheet_output.headers(df_result.columns.tolist())
            self.sheet_output.set_sheet_data(df_result.values.tolist())
            
            # Highlighting Logic
            try:
                col_idx = df_result.columns.get_loc("추가금액")
                threshold = base_price / 2 if base_price else 99999999
                self.sheet_output.dehighlight_all()
                if base_price > 0:
                    for row_idx, val in enumerate(df_result["추가금액"]):
                        try:
                            num_val = int(val)
                            if num_val > threshold or num_val < -threshold:
                                self.sheet_output.highlight_cells(row=row_idx, column=col_idx, bg="red", fg="white")
                        except: continue
            except: pass
            
            self.sheet_tabview.set("  Output Result  ")
            self.btn_save.configure(state="normal")
            messagebox.showinfo("Done", f"Generated {len(df_result)} options.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {e}")

    def save_file(self):
        if not HAS_TKSHEET: return
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not file_path: return
        try:
            data = self.sheet_output.get_sheet_data()
            headers = self.sheet_output.headers()
            df = pd.DataFrame(data, columns=headers)
            if file_path.lower().endswith('.xlsx'):
                df.to_excel(file_path, index=False)
            else:
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            messagebox.showinfo("Saved", "File saved successfully.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def download_template(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="deal_template.xlsx")
        if not file_path: return
        try:
            df = pd.DataFrame(columns=TEMPLATE_COLUMNS)
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
                ws = writer.sheets['Sheet1']
                for idx, w in enumerate(TEMPLATE_WIDTHS):
                    ws.column_dimensions[chr(65+idx)].width = w
            messagebox.showinfo("Success", "Template downloaded.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def reset_all(self):
        if messagebox.askyesno("Reset", "Clear all data?"):
            if HAS_TKSHEET:
                self.sheet_input.set_sheet_data([])
                self.sheet_output.set_sheet_data([])
            self.var_base_price.set(str(DEFAULT_BASE_PRICE))
            self.lbl_file_info.configure(text="No File Selected")
            self.lbl_price_range.configure(text="")
            self.btn_run.configure(state="disabled")
            self.btn_save.configure(state="disabled")


class HtmlGeneratorFrame(ctk.CTkFrame):
    """
    Tab 2: HTML Generator
    """
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.image_items = []
        self.df = None
        
        self._init_ui()
        
    def _init_ui(self):
        # Top Settings
        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(top_frame, text="Table Width(px):").pack(side="left")
        self.width_var = ctk.StringVar(value="860")
        ctk.CTkEntry(top_frame, textvariable=self.width_var, width=60).pack(side="left", padx=5)
        
        ctk.CTkLabel(top_frame, text="| Layout:").pack(side="left", padx=10)
        self.column_mode = ctk.IntVar(value=2)
        ctk.CTkRadioButton(top_frame, text="1 Col", variable=self.column_mode, value=1).pack(side="left", padx=5)
        ctk.CTkRadioButton(top_frame, text="2 Cols", variable=self.column_mode, value=2).pack(side="left", padx=5)
        
        ctk.CTkButton(top_frame, text="🔄 Reset", command=self.reset_all, fg_color="gray", width=60).pack(side="right")
        
        # Excel Integration
        ex_frame = ctk.CTkFrame(self)
        ex_frame.pack(fill="x", padx=10, pady=5)
        
        row1 = ctk.CTkFrame(ex_frame, fg_color="transparent")
        row1.pack(fill="x", padx=10, pady=5)
        ctk.CTkButton(row1, text="📂 Load Excel", command=self.load_excel, fg_color="#E0a800", height=28).pack(side="left")
        self.lbl_excel = ctk.CTkLabel(row1, text="No File", text_color="gray")
        self.lbl_excel.pack(side="left", padx=10)
        
        ctk.CTkLabel(row1, text="Code Col:").pack(side="left")
        self.combo_columns = ctk.CTkComboBox(row1, values=[], state="readonly", width=120)
        self.combo_columns.pack(side="left", padx=5)
        
        ctk.CTkButton(row1, text="⚡ Create List", command=self.populate_list, fg_color="#28a745", height=28).pack(side="left", padx=10)

        row2 = ctk.CTkFrame(ex_frame, fg_color="transparent")
        row2.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(row2, text="Common Img Path:", width=120, anchor="e").pack(side="left")
        self.entry_img_path = ctk.CTkEntry(row2)
        self.entry_img_path.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_img_path.insert(0, "https://gi.esmplus.com/pch7412/")
        
        row3 = ctk.CTkFrame(ex_frame, fg_color="transparent")
        row3.pack(fill="x", padx=10, pady=(0, 5))
        ctk.CTkLabel(row3, text="Common Link Path:", width=120, anchor="e").pack(side="left")
        self.entry_link_path = ctk.CTkEntry(row3)
        self.entry_link_path.pack(side="left", fill="x", expand=True, padx=5)
        self.entry_link_path.insert(0, "https://gi.esmplus.com/pch7412/ht/")

        # List Area
        self.scroll_frame = ctk.CTkScrollableFrame(self, label_text="Image List")
        self.scroll_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Bottom Actions
        bot_frame = ctk.CTkFrame(self, fg_color="transparent")
        bot_frame.pack(fill="x", padx=10, pady=10)
        
        self.txt_result = ctk.CTkTextbox(bot_frame, height=200)
        self.txt_result.pack(side="top", fill="both", expand=True, padx=10, pady=(0, 10))

        btn_box = ctk.CTkFrame(bot_frame, fg_color="transparent")
        btn_box.pack(side="bottom", fill="x", pady=5)
        
        # Center the buttons by packing them into a frame that centers itself or using pack side with expand
        # Using a container frame for buttons to center them group-wise
        center_container = ctk.CTkFrame(btn_box, fg_color="transparent")
        center_container.pack(anchor="center")

        ctk.CTkButton(center_container, text="Generate HTML", command=self.generate_html, width=150).pack(side="left", padx=5)
        ctk.CTkButton(center_container, text="Copy", command=self.copy_code, fg_color="gray", width=100).pack(side="left", padx=5)
        ctk.CTkButton(center_container, text="Preview", command=self.preview_html, fg_color="#dc3545", width=100).pack(side="left", padx=5)

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if not path: return
        try:
            self.df = pd.read_excel(path)
            # Preprocessing: convert specific columns to string if needed
            if '상품번호' in self.df.columns:
                self.df['상품번호'] = self.df['상품번호'].astype(str).str.replace(r'\.0$', '', regex=True)
            self.lbl_excel.configure(text=os.path.basename(path))
            self.combo_columns.configure(values=list(self.df.columns))
            if len(self.df.columns) > 0: self.combo_columns.set(self.df.columns[0])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def populate_list(self):
        if self.df is None: return
        col = self.combo_columns.get()
        if not col: return
        
        c_img = self.entry_img_path.get().strip()
        c_link = self.entry_link_path.get().strip()
        
        count = 0
        for _, row in self.df.iterrows():
            code = str(row[col]).strip().lower()
            if not code or code == 'nan': continue
            
            img_url = f"{c_img}{code}.jpg" if c_img else f"{code}.jpg"
            link_url = f"{c_link}{code}.html" if c_link else f"{code}.html"
            self._add_row(code, img_url, link_url)
            count += 1
        
        messagebox.showinfo("Done", f"Added {count} items.")

    def _add_row(self, label, u_img, u_link):
        row = ctk.CTkFrame(self.scroll_frame, fg_color=("gray90", "gray20"))
        row.pack(fill="x", pady=2)
        
        ctk.CTkLabel(row, text=label, width=100, anchor="w", font=("Segoe UI", 11, "bold")).pack(side="left", padx=5)
        
        ctk.CTkLabel(row, text="IMG:").pack(side="left")
        e_img = ctk.CTkEntry(row)
        e_img.pack(side="left", fill="x", expand=True, padx=6)
        e_img.insert(0, u_img)
        
        ctk.CTkLabel(row, text="LNK:").pack(side="left")
        e_link = ctk.CTkEntry(row)
        e_link.pack(side="left", fill="x", expand=True, padx=6)
        e_link.insert(0, u_link)
        
        btn_del = ctk.CTkButton(row, text="X", width=30, fg_color="transparent", text_color="red", command=lambda f=row: self._remove_row(f))
        btn_del.pack(side="right", padx=5)
        
        self.image_items.append({'frame': row, 'e_img': e_img, 'e_link': e_link})

    def _remove_row(self, frame):
        frame.destroy()
        self.image_items = [i for i in self.image_items if i['frame'] != frame]

    def reset_all(self):
        for i in self.image_items: i['frame'].destroy()
        self.image_items.clear()
        self.df = None
        self.lbl_excel.configure(text="No File")
        self.txt_result.delete("0.0", "end")

    def generate_html(self):
        if not self.image_items: return
        width = self.width_var.get()
        mode = self.column_mode.get()
        
        html = f'<table border="0" cellpadding="0" cellspacing="0" width="{width}" style="width:{width}px; margin:0 auto;">\n'
        
        if mode == 1:
            for item in self.image_items:
                src = item['e_img'].get()
                href = item['e_link'].get()
                html += '  <tr><td style="font-size:0; line-height:0;">\n'
                html += f'      <a href="{href}" target="_blank"><img src="{src}" width="100%" border="0" style="display:block;"></a>\n'
                html += '  </td></tr>\n'
        else:
            w_int = int(width) // 2
            for i in range(0, len(self.image_items), 2):
                html += '  <tr>\n'
                html += self._td(self.image_items[i], w_int)
                if i+1 < len(self.image_items):
                    html += self._td(self.image_items[i+1], w_int)
                else:
                    html += f'    <td width="{w_int}">&nbsp;</td>\n'
                html += '  </tr>\n'
        html += '</table>'
        self.txt_result.delete("0.0", "end")
        self.txt_result.insert("0.0", html)

    def _td(self, item, w):
        src = item['e_img'].get()
        href = item['e_link'].get()
        return (f'    <td width="{w}" style="width:{w}px; font-size:0; line-height:0; vertical-align:top;">\n'
                f'      <a href="{href}" target="_blank"><img src="{src}" width="100%" border="0" style="display:block;"></a>\n'
                f'    </td>\n')

    def copy_code(self):
        code = self.txt_result.get("0.0", "end").strip()
        if code:
            self.clipboard_clear()
            self.clipboard_append(code)
            messagebox.showinfo("Copied", "HTML copied to clipboard!")

    def preview_html(self):
        code = self.txt_result.get("0.0", "end").strip()
        if not code: return
        try:
            with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
                f.write(f"<html><body style='text-align:center; background:#eee;'>{code}</body></html>")
                webbrowser.open('file://' + f.name)
        except Exception as e:
            messagebox.showerror("Error", str(e))


class FileRenameFrame(ctk.CTkFrame):
    """
    Tab 3: File Renamer - 엑셀 리스트에서 파일명 편집 및 저장 기능 지원
    """
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.files = []
        self.codes = []
        self.df_src = None
        self.source_file_path = None  # 불러온 엑셀 파일 경로 저장
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=3)
        self.grid_rowconfigure(0, weight=1)
        
        # Left Panel (File Selection)
        panel_l = ctk.CTkFrame(self)
        panel_l.grid(row=0, column=0, sticky="nsew", padx=5, pady=5)
        
        ctk.CTkLabel(panel_l, text="1. Select Files", font=("Segoe UI", 13, "bold")).pack(padx=5, anchor="w")
        btn_box = ctk.CTkFrame(panel_l, fg_color="transparent")
        btn_box.pack(fill="x", padx=5, pady=5)
        ctk.CTkButton(btn_box, text="1. Add Image Files", command=self.add_files, width=150).pack(side="left", padx=5)
        ctk.CTkButton(btn_box, text="Clear", command=self.clear_files, fg_color="gray", width=100).pack(side="left", padx=5)
        
        self.list_files = tk.Listbox(panel_l, selectmode='extended', relief="flat", bg="#f0f0f0")
        self.list_files.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Right Panel (Source & Preview)
        panel_r = ctk.CTkFrame(self)
        panel_r.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)
        
        # Source Config - 상단 컨트롤 영역
        src_box = ctk.CTkFrame(panel_r, fg_color="transparent")
        src_box.pack(fill="x", padx=5, pady=5)
        ctk.CTkLabel(src_box, text="2. Source Codes (Editable)", font=("Segoe UI", 13, "bold")).pack(padx=5, anchor="w")
        
        # 버튼 행 1: 파일 로드 및 컬럼 선택
        row = ctk.CTkFrame(src_box, fg_color="transparent")
        row.pack(fill="x", pady=2)
        ctk.CTkButton(row, text="2. Load Excel/CSV", command=self.load_source, fg_color="#E0a800", height=28).pack(side="left")
        self.lbl_file = ctk.CTkLabel(row, text="No file", text_color="gray")
        self.lbl_file.pack(side="left", padx=10)
        
        ctk.CTkLabel(row, text="3. 파일명 변경 컬럼:").pack(side="left", padx=(20, 5))
        self.combo_src = ctk.CTkComboBox(row, values=[], state="readonly", command=self.on_col_select, width=120)
        self.combo_src.pack(side="left", padx=5)
        self.lbl_src = ctk.CTkLabel(row, text="(0 codes)")
        self.lbl_src.pack(side="left")

        self.check_ext = ctk.CTkCheckBox(row, text="Preserve Ext", onvalue=True, offvalue=False)
        self.check_ext.select()
        self.check_ext.pack(side="left", padx=20)       
        
        ctk.CTkButton(row, text="4. Preview", command=self.preview, width=100).pack(side="left", padx=10)  

        self.btn_exec = ctk.CTkButton(row, text="5. EXECUTE RENAME", command=self.execute, fg_color="#dc3545", state="disabled")
        self.btn_exec.pack(side="left", padx=5)

        self.btn_save_excel = ctk.CTkButton(row, text="💾 Save Excel", command=self.save_excel, fg_color="#28a745", height=28, state="disabled")
        self.btn_save_excel.pack(side="right", padx=5)
      
        
        # tksheet로 엑셀 데이터 편집 가능하게 표시
        if HAS_TKSHEET:
            sheet_frame = ctk.CTkFrame(panel_r)
            sheet_frame.pack(fill="both", expand=True, padx=5, pady=5)
            
            self.sheet_src = Sheet(sheet_frame, headers=["Load Excel First"])
            self.sheet_src.enable_bindings()  # 셀 편집, 선택 등 활성화
            self.sheet_src.pack(fill="both", expand=True)
        else:
            ctk.CTkLabel(panel_r, text="tksheet 라이브러리 필요 (pip install tksheet)").pack()
            self.sheet_src = None
        
        # 하단: 미리보기 Treeview (파일 변경 결과 표시)
        preview_label = ctk.CTkLabel(panel_r, text="3. Rename Preview", font=("Segoe UI", 13, "bold"))
        preview_label.pack(padx=5, pady=(10, 2), anchor="w")
        
        tree_frame = ctk.CTkFrame(panel_r)
        tree_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(tree_frame, columns=("old", "arr", "new", "stat"), show="headings", height=6)
        self.tree.heading("old", text="Current Name")
        self.tree.heading("arr", text="->")
        self.tree.heading("new", text="New Name")
        self.tree.heading("stat", text="Status")
        self.tree.column("old", width=100)
        self.tree.column("arr", width=20, anchor="center")
        self.tree.column("new", width=100)
        self.tree.column("stat", width=100)
        self.tree.pack(fill="both", expand=True)

    def add_files(self):
        fs = filedialog.askopenfilenames()
        if not fs: return
        for f in fs:
            if f not in [x['path'] for x in self.files]:
                self.files.append({'path': f, 'name': os.path.basename(f)})
        self._refresh_file_list()

    def clear_files(self):
        self.files.clear()
        self._refresh_file_list()
        self.tree.delete(*self.tree.get_children())

    def _refresh_file_list(self):
        self.list_files.delete(0, tk.END)
        for f in self.files:
            self.list_files.insert(tk.END, f['name'])

    def load_source(self):
        p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not p: return
        try:
            if p.lower().endswith('.csv'): 
                self.df_src = pd.read_csv(p)
            else: 
                self.df_src = pd.read_excel(p)
            
            self.source_file_path = p
            self.lbl_file.configure(text=os.path.basename(p))
            self.combo_src.configure(values=list(self.df_src.columns))
            
            # tksheet에 데이터 표시
            if self.sheet_src:
                self.sheet_src.headers(self.df_src.columns.tolist())
                # 데이터 정제 후 표시
                data = self.df_src.map(lambda x: DealOptionLogic.normalize_cell(x)).values.tolist()
                self.sheet_src.set_sheet_data(data)
                self.btn_save_excel.configure(state="normal")
            
            if len(self.df_src.columns) > 0:
                self.combo_src.set(self.df_src.columns[0])
                self.on_col_select(None)
                
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def on_col_select(self, _):
        """선택된 컬럼의 코드 데이터를 갱신 (시트에서 직접 읽음)"""
        if not self.sheet_src: return
        col = self.combo_src.get()
        if not col: return
        
        try:
            headers = self.sheet_src.headers()
            col_idx = headers.index(col)
            data = self.sheet_src.get_sheet_data()
            
            codes = []
            for row in data:
                if col_idx < len(row):
                    val = str(row[col_idx]).strip().lower()
                    if val and val != 'nan':
                        codes.append(val)
            
            self.codes = codes
            self.lbl_src.configure(text=f"({len(self.codes)} codes)")
        except Exception as e:
            print(f"Error in on_col_select: {e}")

    def save_excel(self):
        """편집된 시트 데이터를 엑셀 파일로 저장"""
        if not self.sheet_src: return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
            initialfile=os.path.basename(self.source_file_path) if self.source_file_path else "renamed_list.xlsx"
        )
        if not file_path: return
        
        try:
            headers = self.sheet_src.headers()
            data = self.sheet_src.get_sheet_data()
            df = pd.DataFrame(data, columns=headers)
            
            if file_path.lower().endswith('.xlsx'):
                df.to_excel(file_path, index=False)
            else:
                df.to_csv(file_path, index=False, encoding='utf-8-sig')
            
            messagebox.showinfo("Saved", f"파일이 저장되었습니다:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def preview(self):
        """시트에서 선택된 컬럼의 현재 데이터로 미리보기 갱신"""
        # 먼저 시트 데이터에서 codes 다시 읽기
        self.on_col_select(None)
        
        self.tree.delete(*self.tree.get_children())
        self.btn_exec.configure(state="disabled")
        if not self.files: 
            messagebox.showwarning("Warning", "먼저 이름을 변경할 파일을 선택하세요.")
            return
        if not self.codes:
            messagebox.showwarning("Warning", "엑셀에서 코드 컬럼을 선택하세요.")
            return
        
        ready = False
        for i, f in enumerate(self.files):
            old = f['name']
            ext = os.path.splitext(old)[1]
            new_n = ""
            stat = ""
            
            if i < len(self.codes):
                code = self.codes[i]
                if any(c in '<>:"/\\|?*' for c in code):
                    stat = "Invalid Char"
                    new_n = code
                else:
                    new_n = f"{code}{ext}" if self.check_ext.get() else code
                    stat = "Ready"
                    ready = True
            else:
                stat = "Skip (No Code)"
            
            self.tree.insert("", "end", values=(old, "->", new_n, stat))
            
        if ready: self.btn_exec.configure(state="normal")

    def execute(self):
        if not messagebox.askyesno("Confirm", "선택한 파일의 이름을 변경하시겠습니까?"): return
        count = 0
        kids = self.tree.get_children()
        for i, kid in enumerate(kids):
            vals = self.tree.item(kid)['values']
            if vals[3] != "Ready": continue
            
            try:
                f_info = self.files[i]
                old_p = f_info['path']
                d = os.path.dirname(old_p)
                new_p = os.path.join(d, vals[2])
                
                if old_p == new_p: continue
                if os.path.exists(new_p): raise FileExistsError("Target exists")
                
                os.rename(old_p, new_p)
                self.files[i]['path'] = new_p
                self.files[i]['name'] = vals[2]
                self.tree.set(kid, "stat", "Done")
                count += 1
            except Exception as e:
                self.tree.set(kid, "stat", str(e))
                
        messagebox.showinfo("Result", f"{count}개 파일의 이름이 변경되었습니다.")
        self._refresh_file_list()
        self.btn_exec.configure(state="disabled")


# --- Main Application ---
class ESMToolApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("ESM Deal Option & Tools v3.02")
        self.geometry("1600x900+10+10")
        self.minsize(1600, 900)
        
        # Configure Grid
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        # TabView
        self.tabview = ctk.CTkTabview(self)
        self.tabview.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        
        self.tabview.add("  ESM Option Maker  ")
        self.tabview.add("  HTML Generator  ")
        self.tabview.add("  File Renamer  ")

        
        # Initialize Frames
        OptionMakerFrame(self.tabview.tab("  ESM Option Maker  ")).pack(fill="both", expand=True)
        HtmlGeneratorFrame(self.tabview.tab("  HTML Generator  ")).pack(fill="both", expand=True)
        FileRenameFrame(self.tabview.tab("  File Renamer  ")).pack(fill="both", expand=True)
        
        # Footer
        self.lbl_status = ctk.CTkLabel(self, text="Ready", anchor="w")
        self.lbl_status.grid(row=1, column=0, sticky="ew", padx=10, pady=2)

if __name__ == "__main__":
    app = ESMToolApp()
    app.mainloop()
