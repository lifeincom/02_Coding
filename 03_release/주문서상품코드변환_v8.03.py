"""
주문서 상품코드 변환 및 출고리스트 생성 프로그램 v8.02

이 프로그램은 다양한 쇼핑몰에서 다운로드한 주문서를 읽어들여
상품코드를 변환하고 출고리스트를 자동으로 생성하는 GUI 애플리케이션입니다.

주요 기능:
- 주문서 Excel 파일 불러오기 (사방넷, 이지어드민 지원)
- 상품코드 자동 추출 및 변환
- 세트상품 분리 처리
- 출고리스트 자동 생성
- 티셔츠 별도 출고리스트 생성
- 판매 데이터 집계
"""

# ===== 라이브러리 임포트 =====
import os                # 파일 및 디렉토리 경로 처리
import sys               # 시스템 관련 기능 (실행 경로 확인 등)
import datetime          # 날짜/시간 처리
import pandas as pd      # Excel 파일 읽기/쓰기 및 데이터 처리
import tkinter as tk     # 기본 GUI 라이브러리
from tkinter import filedialog, messagebox  # 파일 선택 및 메시지 박스 다이얼로그
import customtkinter as ctk  # 모던한 UI를 위한 CustomTkinter
from tksheet import Sheet    # Excel 형태의 테이블 표시 위젯
import warnings          # 경고 메시지 관리
import matplotlib.pyplot as plt  # 그래프 표시 (한글 폰트 설정용)

# ===== CustomTkinter 전역 설정 =====
# CustomTkinter의 외관 모드 설정 (라이트 모드 고정)
ctk.set_appearance_mode("light")  # 선택 가능: "System", "Dark", "Light"
# CustomTkinter의 기본 컬러 테마 설정
ctk.set_default_color_theme("blue")  # 선택 가능: "blue", "green", "dark-blue"

# ===== UI 컬러 테마 정의 =====
# 애플리케이션 전체에서 사용할 컬러 팔레트
# 모던하고 일관성 있는 디자인을 위해 미리 정의된 색상 코드
COLORS = {
    # 주요 색상 - 생동감 있는 블루 그라데이션
    "primary": "#2563EB",         # 모던 브라이트 블루
    "primary_hover": "#1D4ED8",   # 진한 블루
    "primary_light": "#DBEAFE",   # 연한 블루 배경
    
    # 보조 색상 - 모던 그린
    "secondary": "#10B981",       # 에메랄드 그린
    "secondary_hover": "#059669", # 진한 에메랄드
    "secondary_light": "#D1FAE5", # 연한 그린 배경
    
    # 강조 색상
    "accent": "#2563EB",         # 모던 브라이트 블루
    "accent_hover": "#1D4ED8",   # 진한 블루
    
    # 경고/위험 색상
    "danger": "#EF4444",          # 모던 레드
    "danger_hover": "#DC2626",    # 진한 레드
    "warning": "#F59E0B",         # 앰버 경고색
    "warning_hover": "#D97706",   # 진한 앰버
    
    # 정보 색상
    "info": "#2563EB",         # 모던 브라이트 블루
    "info_hover": "#1D4ED8",   # 진한 블루
    
    # 배경 색상
    "frame_bg": "#F9FAFB",        # 밝은 회색 배경
    "card_bg": "#FFFFFF",         # 깔끔한 화이트 카드
    "entry_bg": "#FFFFFF",        # 화이트 입력 배경
    "border": "#E5E7EB",          # 부드러운 테두리
    
    # 텍스트 색상
    "text_primary": "#111111",    # 진한 텍스트
    "text_secondary": "#6B7280",  # 중간 회색 텍스트
    "text_muted": "#9CA3AF",      # 흐린 회색
    
    # 시트 색상
    "sheet_header_bg": "#333333", # 다크 블랙 헤더
    "sheet_header_fg": "#FFFFFF", # 화이트 헤더 텍스트
    "sheet_row_alt": "#F9FAFB",   # 교차 행 색상
}

# ===== 폰트 설정 =====
# 애플리케이션 전체에서 사용할 폰트 패밀리 및 크기
FONT_FAMILY = "nanumgothic"  # Windows에서 기본 제공되는 모던한 폰트

# 용도별 폰트 크기 정의
FONT_SIZES = {
    "title": 13,      # 제목용
    "heading": 12,    # 소제목용
    "normal": 12,     # 일반 텍스트
    "small": 12,      # 작은 텍스트
}

# ===== 메인 애플리케이션 클래스 =====
class OrderProcessingApp(ctk.CTk):
    """
    주문서 상품코드 변환 및 출고리스트 생성 GUI 애플리케이션
    
    이 클래스는 CustomTkinter를 기반으로 한 메인 윈도우를 생성하고,
    주문서 처리에 필요한 모든 기능을 제공합니다.
    
    주요 기능:
    - Excel 형식의 주문서 파일 불러오기
    - 상품코드 추출 및 변환 (다양한 쇼핑몰 형식 지원)
    - 세트상품 자동 분리
    - 전체 출고리스트 및 티셔츠 전용 출고리스트 생성
    - 판매 데이터 집계 및 표시
    - 클립보드 복사 기능
    - Excel 파일로 저장
    """
    def __init__(self):
        """
        OrderProcessingApp 클래스의 생성자
        
        애플리케이션 윈도우를 초기화하고 기본 설정을 수행합니다.
        - 윈도우 크기 및 제목 설정
        - 변수 초기화
        - UI 구성요소 생성
        """
        super().__init__()

        # 윈도우 기본 설정
        self.title("주문서 상품코드 변환 & 출고리스트 생성 프로그램 v8.03")
        self.geometry("1600x850+10+10")  # 창 크기 및 위치 설정
        self.minsize(1600, 850)  # 최소 창 크기 제한
        
        # 윈도우 최대화 시도 (Windows에서만 동작)
        try:
            self.state('zoomed')
        except:
            pass  # Linux/Mac에서는 'zoomed' 미지원, 무시

        # 변수 초기화 및 UI 구성
        self.init_variables()
        self.setup_ui()
        
    def init_variables(self):
        # Determine program directory
        if getattr(sys, 'frozen', False):
            self.program_directory = os.path.dirname(os.path.abspath(sys.executable))
        else:
            self.program_directory = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.program_directory)

        # Font setting for matplotlib (kept from original)
        plt.rc('font', family='NanumGothic')

        # Data placeholders
        self.deldf = pd.DataFrame([])
        self.sdf = None
        self.sdf2 = None
        self.sdf2_separation = None
        self.odf = None
        self.ndf = None
        self.df_sale_data = None
        self.df_excel_upload = None
        
        self.gDate = ""
        self.sum1 = "0"
        self.sum2 = "0"

        # Path setup
        now = datetime.datetime.now()
        thisyear = now.strftime('%Y')
        dir_paths = [
            r"\\NAS451\team451",
            r"\\192.168.0.101\team451",
            r"n:\개인\nSync\Coding",
            r"d:\hSync\Coding"
        ]
        self.path = next((p for p in dir_paths if os.path.isdir(p)), None)
        
        if self.path:
            self.sfile_dir = self.path + r"\04-주문택배업로드\택배업로드 및 발주서\\" + thisyear
            self.stock_file = self.path + r"\DB\반품티셔츠재고장.xlsx"
            self.opt_file = self.path + r"\DB\변환코드.xlsx"
            self.load_reference_data()
        else:
            print("유효한 경로를 찾을 수 없습니다.")
            self.sfile_dir = "/"
            self.stock_file = ""
            self.opt_file = ""

    def load_reference_data(self):
        """
        변환코드 Excel 파일에서 참조 데이터를 로드합니다.
        
        변환코드.xlsx 파일에는 다음 시트들이 포함되어 있습니다:
        - 상품코드 시트: 인식 가능한 상품코드 목록
        - 컬러 시트: 인식 가능한 컬러명 목록
        - 사이즈 시트: 인식 가능한 사이즈 목록
        
        이 데이터들은 주문서에서 상품 정보를 추출할 때 사용됩니다.
        """
        try:
            # 상품코드 시트의 첫 번째 열 읽기
            self.code_sr = pd.read_excel(self.opt_file, sheet_name="상품코드").iloc[:, 0]
            # 컬러 시트의 첫 번째 열 읽기
            self.color_sr = pd.read_excel(self.opt_file, sheet_name="컬러").iloc[:, 0]
            # 사이즈 시트의 첫 번째 열 읽기
            self.size_sr = pd.read_excel(self.opt_file, sheet_name="사이즈").iloc[:, 0]
        except Exception as e:
            # 파일을 찾을 수 없거나 형식이 잘못된 경우 에러 표시
            messagebox.showerror("에러", f"설정 파일을 불러올 수 없습니다.\n{e}")

    def setup_ui(self):
        """
        애플리케이션의 전체 UI 레이아웃을 구성합니다.
        
        UI 구성:
        1. 상단 메뉴 프레임 (row=0): 
           - 수집 프로그램 선택 콤보박스
           - 주문서 파일 경로 표시 및 선택 버튼
           - 주요 실행 버튼들 (상품코드 추출, 출고리스트 생성)
           - 초기화, 닫기 버튼
        
        2. 메인 컨텐츠 프레임 (row=1):
           - 왼쪽 탭뷰 (60%): 상품코드 추출리스트, 판매데이터
           - 오른쪽 탭뷰 (40%): 전체상품 출고리스트, 티셔츠 출고리스트
        
        3. 진행률 프레임 (row=2): 작업 진행 상태 표시
        
        4. 저장 경로 프레임 (row=3): 파일 저장 경로 선택 및 저장 버튼
        
        5. 실행/액션 프레임 (row=4): 수량 요약, 복사 버튼들, 도움말 버튼
        """
        # 그리드 레이아웃 설정
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # Main content area expands

        # --- 상단 메뉴 프레임 ---
        self.frm_topmenu = ctk.CTkFrame(self, fg_color=COLORS["card_bg"], corner_radius=12, border_width=1, border_color=COLORS["border"])
        self.frm_topmenu.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))

        self.combo_list = [" +수집 프로그램 선택", " 1.사방넷" , " 2.이지어드민"]
        self.combo = ctk.CTkComboBox(
            self.frm_topmenu, values=self.combo_list, width=200, height=32,
            state="readonly", font=(FONT_FAMILY, FONT_SIZES["normal"]),
            dropdown_font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["entry_bg"], border_color=COLORS["border"],
            button_color=COLORS["primary"], button_hover_color=COLORS["primary_hover"],
            dropdown_fg_color=COLORS["card_bg"]
        )
        self.combo.set(self.combo_list[0])
        self.combo.pack(side="left", padx=(14, 10), pady=12)

        self.txt_load_file = ctk.CTkEntry(
            self.frm_topmenu, width=500, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["entry_bg"], text_color=COLORS["text_primary"],
            border_color=COLORS["border"], border_width=1,
            placeholder_text="주문서 파일을 선택해주세요..."
        )
        self.txt_load_file.pack(side="left", fill="x", expand=True, padx=10, pady=12)

        self.btn_addfile = ctk.CTkButton(
            self.frm_topmenu, text="📂 1. 주문서 선택", width=145, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
            corner_radius=8,
            command=self.import_order
        )
        self.btn_addfile.pack(side="left", padx=10, pady=12)

        # 우측 버튼 그룹 - 시각적 그룹화
        # 닫기 버튼 (경고색)
        self.btn_close = ctk.CTkButton(
            self.frm_topmenu, text="✕ 닫기", width=85, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["danger"], hover_color=COLORS["danger_hover"],
            corner_radius=8,
            command=self.quit
        )
        self.btn_close.pack(side="right", padx=(10, 14), pady=12)

        # 초기화 버튼 (경고색)
        self.btn_reset = ctk.CTkButton(
            self.frm_topmenu, text="🔄 초기화", width=95, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["warning"], hover_color=COLORS["warning_hover"],
            corner_radius=8,
            command=self.reset
        )
        self.btn_reset.pack(side="right", padx=6, pady=12)

        # 구분선
        separator1 = ctk.CTkFrame(self.frm_topmenu, width=2, height=28, fg_color=COLORS["border"])
        separator1.pack(side="right", padx=12, pady=12)

        # 액션 버튼들 (보조색)
        self.btn_output = ctk.CTkButton(
            self.frm_topmenu, text="📦 3. 출고리스트 생성", width=165, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["secondary"], hover_color=COLORS["secondary_hover"],
            corner_radius=8,
            command=self.create_factory_list
        )
        self.btn_output.pack(side="right", padx=6, pady=12)

        self.btn_code_ext = ctk.CTkButton(
            self.frm_topmenu, text="🏷️ 2. 상품코드 추출", width=165, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["accent"], hover_color=COLORS["accent_hover"],
            corner_radius=8,
            command=self.output_product_code_extraction
        )
        self.btn_code_ext.pack(side="right", padx=6, pady=12)

        # --- 메인 컨텐츠 프레임 ---
        self.frm_container = ctk.CTkFrame(self, fg_color="transparent")
        self.frm_container.grid(row=1, column=0, sticky="nsew", padx=16, pady=8)
        self.frm_container.grid_columnconfigure(0, weight=8)  # 80% 비율
        self.frm_container.grid_columnconfigure(1, weight=2)  # 20% 비율
        self.frm_container.grid_rowconfigure(0, weight=1)

        # 왼쪽 탭뷰 (변환)
        self.tabview_left = ctk.CTkTabview(
            self.frm_container, corner_radius=12,
            segmented_button_fg_color=COLORS["frame_bg"],
            segmented_button_selected_color=COLORS["primary"],
            segmented_button_selected_hover_color=COLORS["primary_hover"],
            border_width=1, border_color=COLORS["border"]
        )
        self.tabview_left.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=0)
        self.tab_l1 = self.tabview_left.add("📦 상품코드 추출리스트")
        self.tab_l2 = self.tabview_left.add("� 판매데이터")

        # 오른쪽 탭뷰 (출력)
        self.tabview_right = ctk.CTkTabview(
            self.frm_container, corner_radius=12,
            segmented_button_fg_color=COLORS["frame_bg"],
            segmented_button_selected_color=COLORS["secondary"],
            segmented_button_selected_hover_color=COLORS["secondary_hover"],
            border_width=1, border_color=COLORS["border"]
        )
        self.tabview_right.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=0)
        self.tab_r1 = self.tabview_right.add("📋 전체상품 출고리스트")
        self.tab_r2 = self.tabview_right.add("👕 티셔츠 출고리스트")

        # Create Sheets (Placeholders initially)
        self.sheet_l1 = self.create_sheet(self.tab_l1, self.deldf)
        self.sheet_l2 = self.create_sheet(self.tab_l2, self.deldf)
        self.sheet_r1 = self.create_sheet(self.tab_r1, self.deldf)
        self.sheet_r2 = self.create_sheet(self.tab_r2, self.deldf)

        # --- 진행률 프레임 ---
        self.frm_progress = ctk.CTkFrame(self, fg_color=COLORS["card_bg"], corner_radius=12, border_width=1, border_color=COLORS["border"])
        self.frm_progress.grid(row=2, column=0, sticky="ew", padx=16, pady=8)
        
        lbl_progress = ctk.CTkLabel(
            self.frm_progress, text="🕑 진행률:",
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            text_color=COLORS["text_secondary"]
        )
        lbl_progress.pack(side="left", padx=(16, 10), pady=12)
        
        self.p_var = tk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(
            self.frm_progress, variable=self.p_var,
            progress_color=COLORS["primary"], height=20,
            corner_radius=10
        )
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=10, pady=12)
        self.progress_bar.set(0)

        self.progress_label = ctk.CTkLabel(
            self.frm_progress, text="ROW : 0", width=100,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            text_color=COLORS["primary"]
        )
        self.progress_label.pack(side="right", padx=16, pady=12)

        # --- 저장 경로 프레임 ---
        self.frm_path = ctk.CTkFrame(self, fg_color=COLORS["card_bg"], corner_radius=12, border_width=1, border_color=COLORS["border"])
        self.frm_path.grid(row=3, column=0, sticky="ew", padx=16, pady=8)
        
        lbl_path = ctk.CTkLabel(
            self.frm_path, text="💾 저장 경로:",
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            text_color=COLORS["text_primary"]
        )
        lbl_path.pack(side="left", padx=16, pady=12)

        self.txt_save_path = ctk.CTkEntry(
            self.frm_path, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["entry_bg"], text_color=COLORS["text_primary"],
            border_color=COLORS["border"], border_width=1,
            placeholder_text="저장할 폴더를 선택하세요..."
        )
        self.txt_save_path.pack(side="left", fill="x", expand=True, padx=10, pady=12)

        self.btn_save_file = ctk.CTkButton(
            self.frm_path, text="💾 저장하기", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["secondary"], hover_color=COLORS["secondary_hover"],
            corner_radius=8,
            command=self.save_file
        )
        self.btn_save_file.pack(side="right", padx=(10, 16), pady=12)

        self.btn_destpath = ctk.CTkButton(
            self.frm_path, text="📁 찾아보기", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["info"], hover_color=COLORS["info_hover"],
            corner_radius=8,
            command=self.browse_save_path
        )
        self.btn_destpath.pack(side="right", padx=6, pady=12)

        # --- 실행/액션 프레임 ---
        self.frm_run = ctk.CTkFrame(self, fg_color=COLORS["card_bg"], corner_radius=12, border_width=1, border_color=COLORS["border"])
        self.frm_run.grid(row=4, column=0, sticky="ew", padx=16, pady=(8, 16))

        # 요약 텍스트 (읽기 전용 스타일)
        self.txt_vol = ctk.CTkEntry(
            self.frm_run, width=480, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            fg_color=COLORS["primary_light"], text_color=COLORS["primary"],
            border_color=COLORS["primary"], border_width=1
        )
        self.txt_vol.pack(side="left", padx=16, pady=12)

        # 도움말 버튼 (정보 스타일)
        self.btn_help = ctk.CTkButton(
            self.frm_run, text="❓ 사용설명", width=110, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color="#9E9E9E", hover_color="#757575",
            corner_radius=8,
            command=self.manual
        )
        self.btn_help.pack(side="right", padx=(10, 16), pady=12)

        # 구분선
        separator2 = ctk.CTkFrame(self.frm_run, width=2, height=28, fg_color=COLORS["border"])
        separator2.pack(side="right", padx=12, pady=12)

        # 복사 버튼들 (모듀 일관된 정보색)
        self.btn_saledata_copy = ctk.CTkButton(
            self.frm_run, text="판매데이터 복사", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
            corner_radius=8,
            command=self.sale_data_copy
        )
        self.btn_saledata_copy.pack(side="right", padx=6, pady=12)

        self.btn_output_list2_copy = ctk.CTkButton(
            self.frm_run, text="티셔츠출고 복사", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
            corner_radius=8,
            command=self.output_list2_copy
        )
        self.btn_output_list2_copy.pack(side="right", padx=6, pady=12)

        self.btn_output_list1_copy = ctk.CTkButton(
            self.frm_run, text="전체출고 복사", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
            corner_radius=8,
            command=self.output_list1_copy
        )
        self.btn_output_list1_copy.pack(side="right", padx=6, pady=12)

        self.btn_copy_code = ctk.CTkButton(
            self.frm_run, text="상품코드 복사", width=120, height=32,
            font=(FONT_FAMILY, FONT_SIZES["normal"]),
            fg_color=COLORS["primary"], hover_color=COLORS["primary_hover"],
            corner_radius=8,
            command=self.code_ext_copy
        )
        self.btn_copy_code.pack(side="right", padx=6, pady=12)

    def create_sheet(self, parent, data):
        """
        tksheet를 사용하여 DataFrame 데이터를 표시하는 시트를 생성합니다.
        
        Args:
            parent: 시트가 배치될 부모 위젯 (탭 등)
            data: 표시할 pandas DataFrame
        
        Returns:
            Sheet: 생성된 tksheet 객체
        
        기능:
        - 기존 위젯 제거 (갱신용)
        - DataFrame 데이터를 2D 리스트로 변환하여 시트에 표시
        - 모던한 색상 테마 적용 (헤더, 셀, 그리드)
        - 사용자 상호작용 활성화 (복사, 선택 등)
        """
        # 자식 위젯이 있으면 제거 (새로고침 로직)
        for widget in parent.winfo_children():
            widget.destroy()

        parent_frame = tk.Frame(parent, bg=COLORS["card_bg"])
        parent_frame.pack(fill="both", expand=True, padx=2, pady=2)

        sheet = Sheet(
            parent_frame, 
            data=data.values.tolist(), 
            headers=list(data.columns),
            height=500,
            # 향상된 헤더 스타일 - 가독성 개선
            header_bg=COLORS["sheet_header_bg"], 
            header_fg=COLORS["sheet_header_fg"],
            header_font=(FONT_FAMILY, FONT_SIZES["normal"], "bold"),
            # 개선된 셀 스타일
            index_bg=COLORS["frame_bg"],
            top_left_bg=COLORS["frame_bg"],
            # 행 색상 - 가독성 개선
            table_bg=COLORS["card_bg"],
            table_fg=COLORS["text_primary"],
            # 선택된 셀 색상
            table_selected_cells_bg=COLORS["primary_light"],
            table_selected_cells_fg=COLORS["text_primary"],
            # 테두리
            show_vertical_grid=True,
            show_horizontal_grid=True,
            vertical_grid_to_end_of_window=True,
            horizontal_grid_to_end_of_window=True,
            grid_color=COLORS["border"],
        )
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)
        return sheet

    # --- 로직 메서드 ---

    def import_order(self):
        """
        주문서 Excel 파일을 선택하고 기본 정보를 추출합니다.
        
        동작 과정:
        1. 수집 프로그램 선택 여부 확인 (사방넷/이지어드민)
        2. 파일 선택 다이얼로그 표시
        3. 선택한 파일 경로를 텍스트 필드에 표시
        4. 파일명에서 날짜 추출 (형식: YYYYMMDD_파일명.xlsx)
           - 날짜가 없는 경우 현재 날짜 사용
        
        Note:
            실제 파일 로드는 하지 않으며, 경로 확인과 날짜 추출만 수행합니다.
        """
        if self.combo.get() == " +수집 프로그램 선택":
            messagebox.showwarning("경고", "수집 프로그램을 선택해 주세요")
            return
        
        load_file = filedialog.askopenfilename(title="다운로드 주문서를 선택해 주세요.", 
            filetypes=(("모든 파일", "*.*"),("xlsx 파일", "*.xlsx"), ("xls 파일", "*.xls")), 
            initialdir=self.sfile_dir)

        if not load_file:
            return

        self.txt_load_file.delete(0, tk.END)
        self.txt_load_file.insert(tk.END, load_file)

        ds = load_file.find("_")
        if ds != -1:
            self.gDate = load_file[ds-8:ds]
            self.gDate = self.gDate[0:4] + "-" + self.gDate[4:6] + "-" + self.gDate[6:8]
        else:
            self.gDate = datetime.datetime.now().strftime("%Y-%m-%d")

    def reset(self):
        """
        애플리케이션을 초기 상태로 리셋합니다.
        
        리셋 항목:
        - 주문서 파일 경로 입력 필드 초기화
        - 저장 경로 입력 필드 초기화
        - 수량 요약 필드 초기화
        - 모든 시트 데이터 삭제 (빈 DataFrame으로 갱신)
        - 진행률 바 0%로 리셋
        """
        self.txt_load_file.delete(0, tk.END)
        self.txt_save_path.delete(0, tk.END)
        self.txt_vol.delete(0, tk.END)

        self.create_sheet(self.tab_l1, self.deldf)
        self.create_sheet(self.tab_l2, self.deldf)
        self.create_sheet(self.tab_r1, self.deldf)
        self.create_sheet(self.tab_r2, self.deldf)

        self.p_var.set(0)
        self.progress_label.configure(text="ROW : 0")

    def extract_itemcode(self, inCode):
        """
        원본 상품코드를 표준화된 형식으로 변환합니다.
        
        이 메서드는 다양한 쇼핑몰에서 사용하는 상품코드 형식을
        회사 내부의 표준 형식으로 변환하는 핵심 로직입니다.
        
        Args:
            inCode: 변환할 원본 상품코드 (문자열)
        
        Returns:
            str: 변환된 표준 상품코드
                 변환할 수 없는 경우 빈 문자열 반환
        
        변환 규칙:
        1. 특정 접두사로 시작하는 코드 (W, M, T, U, BHP, GS, DS 등):
           - 위치별 문자열 추출 및 재조합
           - 예: "WBA-123" 형식으로 변환
        
        2. C, D, E, H, R, Z로 시작하는 코드:
           - 복잡한 패턴 매칭 및 변환
           - 위치별 문자 추출 후 prefix_map 사용하여 변환
        
        3. prefix_map: 특정 문자를 표준 접두사로 매핑
           - S -> JS, U -> JU, A -> BA 등
        
        Note:
            대소문자 구분 없음 (자동으로 대문자로 변환)
        """
        inCode = str(inCode).upper()  # 대문자로 변환하여 대소문자 구분 제거
        
        # 문자 매핑 테이블 (특정 문자를 표준 접두사로 변환)
        prefix_map = {
            "S": "JS", "U": "JU", "A": "BA", "C": "BC", 
            "L": "JL", "M": "JM", "N": "NA", "J": "EJ",
            "F": "JF"
        }

        if inCode.startswith(("W", "M", "T", "U", "BHP", "GS", "DS", "CS", "PAC", "SPH")):
            if len(inCode) > 3 and inCode[3] == "A":
                goodNum = inCode[4:]
                return f"{inCode[:3]}-{goodNum}"
            if len(inCode) > 3 and inCode[3] =="F":
                goodNum = inCode[4:]
                return f"{inCode[0]}JF-{goodNum}"
            elif inCode.startswith(("TS")): 
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
            if len(inCode) > 3 and inCode[1:3] == "P7":
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"BHP-{goodNum}"
            elif len(inCode) > 3 and inCode[3]== "T" and inCode[:2] == "RS": 
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
            elif len(inCode) > 3 and inCode[3] in ("T", "U"):
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"{inCode[3]}{inCode[1]}-{goodNum}"
            elif len(inCode) > 3 and inCode[3] in ("W", "M"):
                goodNum = inCode[2] + inCode[4] + inCode[5]
                return f"{inCode[3]}{prefix_map.get(inCode[1], 'XX')}-{goodNum}"
            elif len(inCode) > 4 and inCode[4] in ("T"):
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}{inCode[1]}{inCode[2]}-{goodNum}"
            elif len(inCode) > 4 and inCode[4] in ("U"):
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}{inCode[1]}{inCode[2]}O{goodNum}"
            elif len(inCode) > 4 and inCode[4] in ("W", "M"):
                goodNum = inCode[3] + inCode[5] + inCode[6]
                return f"{inCode[4]}BP-{goodNum}"
        return ""

    def order_product_code_extraction(self):
        """
        주문서 Excel 파일에서 상품코드, 컬러, 사이즈 정보를 추출합니다.
        
        동작 과정:
        1. 주문서 Excel 파일 로드
        2. 필요한 컬럼 추가 (상품코드, 컬러, 사이즈, 옵션코드 등)
        3. 수집 프로그램에 따라 컬럼 재배치 (사방넷 vs 이지어드민)
        4. 각 주문 행을 순회하며:
           - 쇼핑몰별 상품명 형식 정규화
           - 참조 데이터(변환코드.xlsx)를 사용하여 상품코드, 컬러, 사이즈 추출
        
        쇼핑몰별 처리:
        - 하프클럽: "_" → "-", 공백 제거
        - 진마니아: 특정 텍스트 제거 및 구분자 정규화
        - 롯데닷컷/롯데홈쇼핑: 모델명, 색상 표기 제거
        - 현대홈쇼핑: 구분자 정규화
        - ESM지마켓/ESM옥션: 가격 텍스트 제거, 문자열 추출
        - 11번가: 특정 텍스트 제거 및 특정 구간 문자열 추출
        - 쿠팡: MBP가 아닌 경우 "+" 접두사 추가
        - 스마트스토어: 특정 텍스트 제거
        - 티몬: "|" 제거
        
        Note:
            추출된 데이터는 self.sdf DataFrame에 저장됩니다.
        """
        sFile = self.txt_load_file.get()
        if not os.path.isfile(sFile):
            messagebox.showwarning("오류", "파일이 존재하지 않습니다.")
            return

        self.sdf = pd.read_excel(sFile)

        self.sdf["상품코드"] = ""
        self.sdf["컬러"] = ""
        self.sdf["사이즈"] = ""
        self.sdf["옵션코드"] = ""
        self.sdf["상품코드2"] = ""
        self.sdf["옵션코드2"] = ""

        if self.combo.get() == " 2.이지어드민":
            self.sdf = self.sdf.reindex(columns=["수령자", "전화번호", "핸드폰", "우편번호", "주소", "옵션",
                                       "주문수량", "특이사항", "판매처", "배송비", "상품번호", "주문번호", "주문자명", 
                                       "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2"])
            self.sdf = self.sdf.rename(columns={"수령자": "성명","옵션":"상품명","주문수량":"수량","특이사항":"배송메세지","판매처":"주문처"})
            self.sdf['수량'] = self.sdf['수량'].astype(int)
        else:
            self.sdf = self.sdf.reindex(columns=["성명", "전화번호", "핸드폰", "우편번호", "주소", "상품명", 
                                       "수량", "배송메세지", "주문처", "요금구분", "운송장번호", "사방넷주문번호", "쇼핑몰아이디", 
                                       "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2"])

        self.sum1 = str(self.sdf["수량"].sum())
        
        for r in range(0, self.sdf.shape[0]):
            odName = self.sdf.iloc[r, 0] # 성명
            odPdname = self.sdf.iloc[r, 5] # 상품명
            odQty = self.sdf.iloc[r, 6] # 수량
            odSite = self.sdf.iloc[r, 8] # 주문처

            x = str(odPdname)
            
            # Site specific processing
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
                s_idx = x.find("_")
                e_idx = x.find("/")
                if s_idx != -1 and e_idx != -1:
                    x = x[s_idx+1:e_idx]

            if odSite == "ESM옥션":
                del_texts = ["1000원", "2000원", "3000원","4000원", "5000원", "6000원", "7000원", "8000원", "9000원"]
                for del_text in del_texts:
                    x = x.replace(del_text, "")
                x = x.replace(" ", "")
                s_idx = x.find("_")
                e_idx = x.find("[")
                if s_idx != -1 and e_idx != -1:
                    x = x[s_idx+1:e_idx]

            if odSite == "11번가":
                del_texts = ["색상:", "사이즈"]
                for del_text in del_texts:
                    x = x.replace(del_text, "")
                s_idx = x.find("_")
                e_idx = x.find("+")
                if s_idx != -1 and e_idx != -1:
                    x = x[s_idx+1:e_idx-1]

            if odSite == "쿠팡":
                if not x.startswith(("MBP")):
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
            for code in self.code_sr:
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
            for color in self.color_sr:
                if x.find(color) > 0:
                    gColor = color
                else:
                    continue

            gSize = ""
            for size in self.size_sr:
                if x.find(str(size)) < 0:
                    continue
                gSize = str(size).strip(" "":""/""+""(")

            self.sdf.iloc[r, 13] = gCode
            self.sdf.iloc[r, 14] = gColor
            self.sdf.iloc[r, 15] = gSize

    def set_product_separation(self):
        """
        세트 상품을 개별 상품으로 분리합니다.
        
        "+" 기호로 연결된 세트 상품을 개별 행으로 분리하여 처리합니다.
        
        처리 규칙:
        1. 컬러에 "+"가 있는 경우:
           - 컬러를 "+"로 분할하여 각 컬러별로 행 생성
           - 상품코드는 상위 7자리만 유지
        
        2. 상품코드에 "+"가 있는 경우:
           a) BHP로 시작: 2개 상품으로 분리
           b) TB-001로 시작: 특수 처리 (컬러 정보 포함)
           c) 기타: 일반 분리
        
        3. "+"가 없는 경우:
           - 그대로 복사
        
        결과:
            self.sdf2에 분리된 데이터 저장
        
        Note:
            원본 데이터(self.sdf)는 변경하지 않습니다.
        """
        # *** PRESERVED LOGIC AS REQUESTED ***
        self.sdf2 = pd.DataFrame(columns=self.sdf.columns)
        i = 0

        cnt = int(self.sdf.shape[0])
        for r in range(0, cnt):
            codeVar = self.sdf.iloc[r, 13]
            colorVar = self.sdf.iloc[r, 14]
            cutA = int(codeVar.find('+'))
            cutB = int(colorVar.find("+"))
                 
            set1 = codeVar[:cutA] if cutA != -1 else ""
            set2 = codeVar[cutA+1:] if cutA != -1 else ""
            
            if colorVar.find('+') > 0:
                S = colorVar.split("+")
                for c in range(0, len(S)):
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0, 19)]
                    self.sdf2.iloc[i, 13] = codeVar[:7]
                    self.sdf2.iloc[i, 14] = S[c]
                    i += 1
            elif codeVar.find('+') > 0:
                if codeVar[0:4] == "BHP":
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0, 19)]
                    self.sdf2.iloc[i, 13] = codeVar[:cutA]
                    i += 1
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0, 19)]
                    self.sdf2.iloc[i, 13] = codeVar[cutA+1:cutA+8]
                    i += 1
                elif codeVar[0:6] == "TB-001":
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0, 19)]
                    self.sdf2.iloc[i, 13] = set1[:6]
                    self.sdf2.iloc[i, 14] = set1[7:int(set1.find(")"))]
                    i += 1
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0, 19)]
                    self.sdf2.iloc[i, 13] = set2[:6]
                    self.sdf2.iloc[i, 14] = set2[7:int(set2.find(")"))]
                    i += 1
                else:
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0,  19)]
                    self.sdf2.iloc[i, 13] = str(set1)
                    i += 1
                    self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0,  19)]
                    self.sdf2.iloc[i, 13] = str(set2)
                    i += 1
            else:
                self.sdf2.loc[i] = [self.sdf.iloc[r, a] for a in range(0,  19)]
                i += 1

    def output_product_code_extraction(self):
        """
        주문서 상품코드 추출 전체 프로세스를 실행합니다.
        
        실행 단계:
        1. 선택된 파일과 수집 프로그램 유효성 검증
        2. order_product_code_extraction(): 상품정보 추출
        3. set_product_separation(): 세트상품 분리
        4. 상품코드 변환 (extract_itemcode 사용)
        5. 옵션코드 생성 (상품코드, 컬러, 사이즈 조합)
        6. 화면에 결과 표시 (왼쪽 탭에 표시)
        7. Excel 업로드용 데이터 준비
        
        진행 상황:
            진행률 바를 통해 처리 진행 상황 표시
        
        결과:
            - self.sdf2_separation: 화면 표시용
            - self.df_excel_upload: Excel 업로드용
        """
        if self.combo.get() == " +수집 프로그램 선택":
            messagebox.showwarning("경고", "수집 프로그램을 선택해 주세요")
            return
        elif self.combo.get() == " 1.사방넷" and "우체국택배업로드" not in self.txt_load_file.get():
            messagebox.showwarning("경고", "우체국택배업로드 주문서를 선택해 주세요")
            return
        elif self.combo.get() == " 2.이지어드민" and "이지어드민" not in self.txt_load_file.get():
            messagebox.showwarning("경고", "이지어드민 주문서를 선택해 주세요")
            return
        
        if len(self.txt_load_file.get()) == 0:
            messagebox.showwarning("경고", "주문서 파일을 선택해 주세요.")
            return
        
        # 1. code extraction
        self.order_product_code_extraction()
        # 2. separation
        self.set_product_separation()
        
        # 3. product code conversion (inline loop in original code, adapted here)
        cnt = self.sdf2.shape[0]
        idx = 1
        for r in range(0, cnt):
            self.sdf2.iloc[r, 17] = self.extract_itemcode(self.sdf2.iloc[r, 13])
            
            gCode = self.sdf2.iloc[r, 13]
            gColor = self.sdf2.iloc[r, 14]
            gSize = self.sdf2.iloc[r, 15]

            if gColor and str(gCode).startswith(("T", "U", "BHP-7", "BHP-5", "WJL-0", "EA", "DP", "EP", "CP", "DL0", "EL0", "RL0", "WBL", "SPH")):
                optCode = f"{gCode}({gColor}) : {gSize}"
            elif len(str(gCode)) > 4 and (str(gCode)[3] == "T" or str(gCode)[4] == "T"):
                optCode = f"{gCode}({gColor}) : {gSize}"
            else:
                optCode = f"{gCode} : {gSize}"
            self.sdf2.iloc[r, 16] = optCode
            
            gCode2 = self.extract_itemcode(self.sdf2.iloc[r, 13])
            self.sdf2.iloc[r, 17] = gCode2
            if gColor and str(gCode2).startswith(("T", "U", "BHP-7", "BHP-5", "WJL-0", "EA", "DP", "EP", "CP", "DL0", "EL0", "RL0", "WBL", "SPH")):
                optCode2 = f"{gCode2}({gColor}) : {gSize}"
            else:
                optCode2 =f"{gCode2} : {gSize}"
            
            self.sdf2.iloc[r, 18] = optCode2
            idx += 1
            
            progress = idx / int(self.sdf.shape[0]) * 100
            self.progress_label.configure(text=f"ROW : {r+1}")
            self.p_var.set(progress)
            self.update_idletasks()

        self.sdf2_separation = self.sdf2.loc[:,["성명", "상품명", "옵션코드", "수량", "주문처", "상품코드2", "컬러", "사이즈", "옵션코드2"]]
        self.create_sheet(self.tab_l1, self.sdf2_separation)

        self.df_excel_upload = self.sdf2.loc[:, ["성명", "전화번호", "핸드폰", "우편번호", "주소", "옵션코드",
            "수량", "배송메세지", "주문처", "요금구분", "운송장번호", "사방넷주문번호", "쇼핑몰아이디"]]
        self.df_excel_upload = self.df_excel_upload.rename(columns={"옵션코드": "상품명"})
        messagebox.showinfo("알림", "주문서에서 상품코드 추출 변환이 완료 되었습니다.")

    def create_factory_list(self):
        """
        출고리스트를 생성합니다.
        
        생성하는 리스트:
        1. 판매 데이터 (self.df_sale_data):
           - 주문처별, 상품코드별 수량 집계
           - 날짜 정보 추가
        
        2. 전체 출고리스트 (self.odf):
           - 옵션코드별 총 수량 집계
           - 피벗 테이블 사용
        
        3. 티셔츠 출고리스트 (self.ndf):
           - 'T' 또는 'S'로 시작하는 상품만 필터링
           - 재고 정보 조회 (반품티셔츠재고장.xlsx)
           - 주문자명, 옵션코드, 수량, 재고 포함
        
        진행 상황:
            진행률 바를 통해 처리 진행 상황 표시
        
        결과:
            생성된 리스트를 각각 오른쪽 탭에 표시
        """
        if len(self.txt_load_file.get()) == 0:
            messagebox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
            return

        # Output sale data
        self.df_sale_data = self.sdf2.loc[:,["주문처", "상품코드2", "컬러", "사이즈", "수량"]]
        self.df_sale_data = self.df_sale_data.groupby(["주문처", "상품코드2", "컬러", "사이즈"], as_index=False).agg({'수량':'sum'})
        self.df_sale_data["날짜"] = self.gDate
        self.df_sale_data = self.df_sale_data[["날짜", "상품코드2", "컬러", "사이즈", "수량", "주문처"]]
        
        self.sum2 = str(self.df_sale_data["수량"].sum())
        sum_txt = f"주문서 주문수량 : {self.sum1}    /    총 출고수량 : {self.sum2}"
        self.txt_vol.configure(state="normal")
        self.txt_vol.delete(0, tk.END)
        self.txt_vol.insert(0, sum_txt)
        self.txt_vol.configure(state="disabled")

        self.create_sheet(self.tab_l2, self.df_sale_data)

        # Create forwarding list
        self.pdf = self.sdf2.pivot_table("수량", index="옵션코드2", aggfunc="sum")
        self.pdf = self.pdf.reset_index()

        self.odf = pd.DataFrame(columns=["옵션코드", "수량"])
        
        idx = 1
        for p in range(0, int(self.pdf.shape[0])):
            eNo = idx
            eCode = self.pdf.iloc[p, 0]
            Ea = self.pdf.iloc[p, 1]
            box_list2 = [eCode, Ea]
            self.odf.loc[eNo-1] = box_list2
            idx += 1
            progress = idx / int(self.pdf.shape[0]) * 100
            self.p_var.set(progress)
            self.update_idletasks()

        if os.path.isfile(self.stock_file):
            st_df = pd.read_excel(self.stock_file, skiprows=1)
            st_df = st_df.loc[:, ["상품코드", "재고수량"]]
        else:
            st_df = pd.DataFrame(columns=["상품코드", "재고수량"])

        self.create_sheet(self.tab_r1, self.pdf)
        self.p_var.set(0)

        # Tshirt list
        self.ndf = pd.DataFrame(columns=["주문자명", "옵션코드", "수량", "재고"])
        idx = 1
        nu = 1
        for n in range(0, int(self.sdf2.shape[0])):
            if str(self.sdf2.iloc[n, 18])[0] in ["T", "S"]:
                eNo = nu
                eName = self.sdf2.iloc[n, 0]
                eCode = self.sdf2.iloc[n, 18]
                Ea = self.sdf2.iloc[n, 6]
                sEa = ""
                # Simple lookup (can be optimized but keeping original logic structure)
                for k in range(0, st_df.shape[0]):
                    if self.sdf2.iloc[n, 18] == st_df.iloc[k, 0]:
                        sEa = st_df.iloc[k, 1]
                
                box_list3 = [eName, eCode, Ea, sEa]
                self.ndf.loc[eNo-1] = box_list3
                nu += 1
            idx += 1
            progress = idx / int(self.sdf2.shape[0]) * 100
            self.p_var.set(progress)
            self.update_idletasks()

        self.create_sheet(self.tab_r2, self.ndf)
        messagebox.showinfo("알림", "출고리스트 생성이 완료 되었습니다")

    def code_ext_copy(self):
        """
        상품코드 추출 데이터를 클립보드에 복사합니다.
        Excel 업로드용 데이터를 복사하여 다른 프로그램에 붙여넣을 수 있습니다.
        """
        if self.df_excel_upload is None:
            messagebox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
            return
        self.df_excel_upload.to_clipboard(index=False)
        messagebox.showinfo("알림", "상품코드 및 수량 복사완료. 엑셀시트에 붙여넣기 해 주세요")

    def output_list1_copy(self):
        """
        전체 출고리스트를 클립보드에 복사합니다.
        """
        if self.odf is None:
            messagebox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환 및 출고리스트 생성을 실행해 주세요.")
            return
        self.odf.to_clipboard(index=False)
        messagebox.showinfo("알림", "출고리스트 복사완료")

    def output_list2_copy(self):
        """
        티셔츠 출고리스트를 클립보드에 복사합니다.
        """
        if self.ndf is None:
            messagebox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환 및 출고리스트 생성을 실행해 주세요.")
            return
        self.ndf.to_clipboard(index=False)
        messagebox.showinfo("알림", "티셔츠 출고리스트 복사완료")

    def sale_data_copy(self):
        """
        판매 데이터를 클립보드에 복사합니다.
        """
        if self.df_sale_data is None:
            messagebox.showwarning("경고", "주문서 파일 선택 후 상품코드 변환을 실행해 주세요.")
            return
        self.df_sale_data.to_clipboard(index=False)
        messagebox.showinfo("알림", "매출데이터 복사완료")

    def manual(self):
        """
        사용 설명서를 표시합니다.
        
        포함 내용:
        - 주문서 파일명 형식 요구사항
        - 오류 체크사항 및 해결 방법
        - 프로그램 실행 순서
        """
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
        5. 출고리스트 복사 > 출고리스트시트에 붙여넣기
        """
        messagebox.showinfo("알림", desc)

    def browse_save_path(self):
        """
        파일 저장 경로를 선택하는 다이얼로그를 표시합니다.
        선택한 경로를 저장 경로 입력 필드에 표시합니다.
        """
        folder_selected = filedialog.askdirectory()
        if not folder_selected:
            return
        self.txt_save_path.delete(0, tk.END)
        self.txt_save_path.insert(0, folder_selected)

    def save_file(self):
        """
        택배 업로드용 Excel 파일을 저장합니다.
        
        동작:
        1. 저장 경로 유효성 검증
        2. df_excel_upload를 Excel 파일로 저장
        3. 파일명: "택배업로드.xlsx"
        
        Note:
            저장 실패 시 에러 메시지 표시
        """
        if len(self.txt_save_path.get()) == 0:
            messagebox.showwarning("경고", "저장할 위치를 선택해 주세요.")
            return
        save_dir = self.txt_save_path.get()
        try:
            self.df_excel_upload.to_excel(save_dir + "/택배업로드.xlsx", sheet_name="택배업로드", index=False)
            messagebox.showinfo("알림", "출고리스트가 저장되었습니다.")
        except Exception as e:
             messagebox.showerror("에러", f"파일 저장 중 에러가 발생했습니다.\n{e}")

if __name__ == "__main__":
    app = OrderProcessingApp()
    app.mainloop()