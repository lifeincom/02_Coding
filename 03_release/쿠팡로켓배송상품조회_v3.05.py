# 쿠팡 로켓배송 상품조회 v3.05

# [주요 기능]
# - 쿠팡 SKU 데이터베이스 조회 및 검색
# - 상품코드, SKU ID, 또는 쿠팡상품코드로 검색
# - 테이블 헤더 클릭으로 오름차순/내림차순 정렬
# - 다중 스레드를 이용한 비동기 데이터 로딩 및 검색
# - 모던한 UI/UX (오버레이 로딩창, 호버 효과)

# [개선 사항]
# - 코드 구조 개선 (상수 분리, 메서드 그룹화)
# - 속도 개선 (검색 스레드 분리)
# - UI 개선 (모던 테마, 오버레이 로딩창)
# - 헤더 클릭 정렬 기능 추가

# [데이터 소스]
# - NAS 서버: \\NAS451\team451\DB\쿠팡SKU리스트.xlsm
# - 로컬 백업: D:\NAS451\DB\쿠팡SKU리스트.xlsm

# 표준 라이브러리
import os  # 파일 및 디렉토리 경로 처리
import pickle  # 데이터 직렬화 (빠른 로딩을 위해 사용)
import re  # 정규표현식 처리
import threading  # 멀티스레딩 (UI 블로킹 방지)
from dataclasses import dataclass  # 불변 설정 클래스 생성
from typing import Optional, List  # 타입 힌팅

# 서드파티 라이브러리
import pandas as pd  # 데이터 처리 및 분석
import tkinter as tk  # GUI 프레임워크
from tkinter import ttk, messagebox  # ttk 위젯 및 메시지박스
from tksheet import Sheet  # 엑셀과 유사한 시트 위젯
from openpyxl import load_workbook  # 엑셀 파일 읽기


# ============================================================================
# 설정 상수
# ============================================================================
@dataclass(frozen=True)
class AppConfig:
    """
    애플리케이션 설정 상수
    
    frozen=True: 인스턴스 생성 후 값 변경 불가 (불변 객체)
    이를 통해 설정값의 안전성을 보장하고 실수로 인한 값 변경을 방지합니다.
    """
    # 윈도우 설정
    TITLE: str = "쿠팡 로켓배송 상품조회 v3.05"  # 프로그램 제목
    GEOMETRY: str = "1800x900+60+60"  # 윈도우 크기 및 위치 (너비x높이+x좌표+y좌표)
    
    # 데이터베이스 파일 경로
    # NAS 서버 경로가 우선, 연결 실패 시 로컬 경로 사용
    NAS_PATH: str = r"\\NAS451\team451\DB"  # NAS 서버 DB 폴더
    LOCAL_PATH: str = r"D:\NAS451\DB"  # 로컬 백업 DB 폴더
    
    # 데이터 파일명
    PICKLE_FILE: str = r"\쿠팡SKU리스트.pickle"  # 피클 파일 (빠른 로딩용)
    EXCEL_FILE: str = r"\쿠팡SKU리스트.xlsm"  # 엑셀 원본 파일
    SHEET_NAME: str = "SKU리스트"  # 엑셀 시트명


@dataclass(frozen=True)
class Theme:
    """
    UI 테마 색상 팔레트
    
    모던하고 전문적인 느낌의 색상 조합을 사용합니다.
    모든 색상은 16진수(HEX) 형식으로 정의됩니다.
    """
    # 배경색 (Background)
    BG_PRIMARY: str = "#F5F7FA"  # 주 배경색 (연한 회색-파랑)
    BG_SECONDARY: str = "#FFFFFF"  # 보조 배경색 (흰색, 카드/패널용)
    BG_OVERLAY: str = "#00000080"  # 오버레이 배경 (반투명 검정, 로딩창용)
    
    # 텍스트 색상 (Foreground)
    FG_PRIMARY: str = "#2D3748"  # 주 텍스트 색상 (진한 회색)
    FG_SECONDARY: str = "#718096"  # 보조 텍스트 색상 (중간 회색, 부가 정보용)
    FG_LIGHT: str = "#FFFFFF"  # 밝은 텍스트 색상 (흰색, 버튼 텍스트용)
    
    # 강조 색상 (Accent)
    ACCENT_PRIMARY: str = "#4A90D9"  # 주 강조색 (파랑, 버튼/링크용)
    ACCENT_SUCCESS: str = "#48BB78"  # 성공 색상 (녹색, 완료 메시지용)
    ACCENT_DANGER: str = "#E53E3E"  # 경고/오류 색상 (빨강)
    ACCENT_DARK: str = "#2D3748"  # 어두운 강조색 (진한 회색, 보조 버튼용)
    
    # 테두리 (Border)
    BORDER: str = "#E2E8F0"  # 테두리 색상 (연한 회색)
    
    # 테이블 헤더
    HEADER_BG: str = "#4A5568"  # 헤더 배경색 (중간 회색)
    HEADER_FG: str = "#FFFFFF"  # 헤더 텍스트 색상 (흰색)


@dataclass(frozen=True)
class Fonts:
    """
    폰트 설정
    
    나눔고딕 폰트를 사용하여 한글 가독성을 최적화합니다.
    각 용도별로 적절한 크기와 굵기를 설정합니다.
    
    튜플 형식: (폰트명, 크기, 스타일)
    """
    FAMILY: str = "NanumGothic"  # 기본 폰트 패밀리
    TITLE: tuple = ("NanumGothic", 11, "bold")  # 제목용 (굵게, 크게)
    NORMAL: tuple = ("NanumGothic", 10, "normal")  # 일반 텍스트용
    SMALL: tuple = ("NanumGothic", 9, "normal")  # 작은 텍스트용 (상태 메시지 등)
    BUTTON: tuple = ("NanumGothic", 10, "normal")  # 버튼 텍스트용
    LOADING: tuple = ("NanumGothic", 12, "bold")  # 로딩 메시지용 (굵게, 눈에 띄게)

# 전역 설정 인스턴스
CONFIG = AppConfig()
THEME = Theme()
FONTS = Fonts()

# 컬럼별 너비 설정 (픽셀 단위, 딕셔너리에 없는 컬럼은 자동 너비 유지)
# 컬럼명: 너비(px) 형식으로 필요한 항목을 추가/수정하세요
COLUMN_WIDTHS: dict = {
    "SKU ID":       70,
    "상품명":        300,
    "바코드":        140,
    "발주가능상태":   90,
    "공급상태":       60,
    "쿠팡옵션코드":   140,
    "옵션코드":      140,
    "SKU":          70,
    "상품코드":      90,
    "컬러":         100,
    "사이즈":        60,
    "쿠팡상품코드":    90,
}


# ============================================================================
# 메인 애플리케이션 클래스
# ============================================================================
class CoupangSKUApp:
    """쿠팡 SKU 조회 애플리케이션"""
    
    def __init__(self, root: tk.Tk):
        """
        애플리케이션 초기화
        
        Args:
            root: Tkinter 루트 윈도우 객체
        
        초기화 순서:
        1. Tkinter 루트 윈도우 및 스타일 설정
        2. 데이터 상태 변수 초기화 (원본 데이터, 필터링 데이터, 로딩 상태 등)
        3. 정렬 상태 변수 초기화 (현재 정렬된 컬럼, 정렬 방향 등)
        4. UI 컴포넌트 생성 (프레임, 버튼, 시트 등)
        5. 로딩 오버레이 UI 생성
        6. 비동기 데이터 로딩 시작
        """
        self.root = root
        self._configure_root()  # 윈도우 기본 설정 (크기, 제목, 배경색)
        self._configure_styles()  # ttk 위젯 스타일 설정
        
        # 데이터 상태 변수
        self.data: pd.DataFrame = pd.DataFrame()  # 원본 데이터 (전체 SKU 리스트)
        self.filtered_data: pd.DataFrame = pd.DataFrame()  # 필터링된 데이터 (검색 결과)
        self.base_path: str = ""  # 데이터 파일 기본 경로 (NAS 또는 로컬)
        self._is_loading: bool = False  # 로딩 상태 플래그 (중복 요청 방지용)
        
        # 정렬 상태 추적 변수
        self.sort_column: Optional[int] = None  # 현재 정렬된 컬럼의 인덱스 (None이면 정렬 안됨)
        self.sort_ascending: bool = True  # 정렬 방향 (True: 오름차순, False: 내림차순)
        self.current_sheet: Optional[Sheet] = None  # 현재 표시된 시트 객체 (정렬 시 참조용)
        
        
        # UI 구성 단계별 실행
        self._build_ui()  # 메인 UI 컴포넌트 생성
        self._create_loading_overlay()  # 로딩 오버레이 창 생성
        
        # 데이터 로드 (비동기 방식으로 UI 블로킹 방지)
        self.load_data_threaded()
    
    # ========================================================================
    # 초기 설정
    # ========================================================================
    def _configure_root(self) -> None:
        """
        루트 윈도우 기본 설정
        
        설정 항목:
        - 윈도우 제목 (CONFIG.TITLE)
        - 윈도우 초기 크기 및 위치 (CONFIG.GEOMETRY)
        - 윈도우 최대화 상태 ('zoomed')
        - 메인 배경색 (THEME.BG_PRIMARY)
        """
        self.root.title(CONFIG.TITLE)  # 윈도우 타이틀바에 표시될 제목
        self.root.geometry(CONFIG.GEOMETRY)  # 윈도우의 초기 크기와 화면 위치
        self.root.state("zoomed")  # 윈도우를 전체화면 모드로 시작 (Windows 환경)
        self.root.config(background=THEME.BG_PRIMARY)  # 윈도우 배경색 설정
    
    def _configure_styles(self) -> None:
        """ttk 스타일 설정"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Combobox 스타일
        style.configure("TCombobox",
                        fieldbackground=THEME.BG_SECONDARY,
                        background=THEME.BG_SECONDARY,
                        foreground=THEME.FG_PRIMARY)
        
        # Progressbar 스타일
        style.configure("TProgressbar",
                        troughcolor=THEME.BORDER,
                        background=THEME.ACCENT_PRIMARY,
                        thickness=8)
        
        # LabelFrame 스타일
        style.configure("TLabelframe",
                        background=THEME.BG_PRIMARY,
                        foreground=THEME.FG_PRIMARY)
        style.configure("TLabelframe.Label",
                        background=THEME.BG_PRIMARY,
                        foreground=THEME.FG_PRIMARY,
                        font=FONTS.TITLE)
    
    # ========================================================================
    # UI 구성
    # ========================================================================
    def _build_ui(self) -> None:
        """
        메인 UI 구성
        
        UI는 3개의 주요 영역으로 분리됩니다:
        1. 상단 프레임: 프로그램 설명 및 종료 버튼
        2. 기능 프레임: 검색 옵션, 검색/초기화/업데이트 버튼
        3. 메인 컨테이너: 입력창과 데이터 시트
        """
        self._build_top_frame()  # 1. 상단 프레임 생성
        self._build_function_frame()  # 2. 기능 프레임 생성
        self._build_main_container()  # 3. 메인 컨테이너 생성
    
    def _build_top_frame(self) -> None:
        """
        상단 프레임 생성 (설명 텍스트 및 종료 버튼)
        
        상단 프레임에는 다음 요소가 포함됩니다:
        - 좌측: 프로그램 사용법 및 데이터 소스 설명
        - 우측: '프로그램 종료' 버튼
        """
        # 상단 프레임 생성 (흰색 배경, x축 방향으로 가득 채움)
        self.frm_top = tk.Frame(self.root, background=THEME.BG_SECONDARY)
        self.frm_top.pack(padx=15, pady=5, fill="x")
        
        # 설명 텍스트 (사용법 및 데이터 소스 정보)
        desc = "1. 검색유형을 선택하고 검색 정보를 입력해 주세요.   2. 소스데이터 : /nas451/DB/쿠팡SKU리스트.xlsm"
        tk.Label(
            self.frm_top, text=desc, height=2, justify="left",
            bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY, font=FONTS.NORMAL
        ).pack(side="left", padx=10, pady=5)
        
        # 종료 버튼 (우측 정렬, 어두운 회색)
        self._create_button(
            self.frm_top, "프로그램 종료", self.root.quit,
            bg=THEME.ACCENT_DARK, side="right"
        )
    
    def _build_function_frame(self) -> None:
        """
        기능 프레임 생성 (검색 옵션 및 버튼들)
        
        기능 프레임은 2개의 그룹으로 구성됩니다:
        1. 좌측 그룹: 검색 유형 선택 콤보박스 + 검색 버튼 + 초기화 버튼
        2. 우측 그룹: 데이터 업데이트 버튼 + 전체 보기 버튼
        """
        # 기능 프레임 생성 (흰색 배경)
        self.frm_function = tk.Frame(self.root, background=THEME.BG_SECONDARY, pady=5)
        self.frm_function.pack(padx=15, pady=5, fill="x")
        
        # 좌측 그룹: 검색 관련 기능
        left_group = tk.Frame(self.frm_function, background=THEME.BG_SECONDARY)
        left_group.pack(side="left", padx=5)
        
        # 검색 유형 선택 콤보박스 (상품코드 or SKU_ID or 쿠팡상품코드)
        self.search_opt = ttk.Combobox(
            left_group, state="readonly",  # 읽기 전용 (직접 입력 불가)
            values=["상품코드로 검색", "쿠팡상품코드로 검색", "SKU_ID로 검색"], width=18
        )
        self.search_opt.current(0)  # 기본값: '상품코드로 검색'
        self.search_opt.pack(side="left", padx=5, ipady=5)
        
        # 검색 버튼 (파랑색 강조)
        self._create_button(left_group, "🔍 입력 데이터 검색", self.filter_data_threaded,
                           bg=THEME.ACCENT_PRIMARY)
        # 초기화 버튼 (검색 결과 및 입력 내용 초기화)
        self._create_button(left_group, "↺ 초기화", self.reset,
                           bg=THEME.ACCENT_DARK)
        
        
        # 우측 그룹: 데이터 관리 기능
        right_group = tk.Frame(self.frm_function, background=THEME.BG_SECONDARY)
        right_group.pack(side="right", padx=5)
        
        # 업데이트 버튼 (엑셀 파일에서 최신 데이터 로드)
        self._create_button(right_group, "등록상품 업데이트", self.update_data_threaded,
                           bg=THEME.ACCENT_DARK)
        # 전체 보기 버튼 (필터링 제거하고 전체 데이터 표시)
        self._create_button(right_group, "전체 등록상품 보기", self.show_all_data,
                           bg=THEME.ACCENT_DARK)
    
    def _build_main_container(self) -> None:
        """
        메인 컨테이너 생성 (입력창 + 데이터 시트)
        
        메인 컨테이너는 2개의 주요 영역으로 구성됩니다:
        1. 좌측: 검색할 상품코드 또는 SKU 입력 텍스트 창
        2. 우측: 검색 결과를 표시하는 데이터 시트
        """
        # 메인 컨테이너 프레임
        self.frm_container = tk.Frame(self.root, background=THEME.BG_PRIMARY)
        self.frm_container.pack(padx=15, pady=5, fill="both", expand=True)
        
        # 좌측: 검색 리스트 입력 영역
        self.frm_input = ttk.LabelFrame(self.frm_container, text="  검색 리스트 입력  ")
        self.frm_input.pack(padx=5, pady=5, side="left", fill="both")
        
        # 여러 줄 입력을 지원하는 텍스트 위젯
        self.input_text = tk.Text(
            self.frm_input, width=20, font=FONTS.NORMAL,
            bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
            relief="flat", borderwidth=1, highlightthickness=1,
            highlightbackground=THEME.BORDER,  # 비활성 상태 테두리
            highlightcolor=THEME.ACCENT_PRIMARY  # 활성 상태 테두리 (파랑)
        )
        self.input_text.pack(padx=8, pady=8, fill="both", expand=True)
        
        # 우측: 데이터 시트 표시 영역
        self.frm_sheet = ttk.LabelFrame(self.frm_container, text="  쿠팡 로켓배송 상품정보 출력  ")
        self.frm_sheet.pack(padx=5, pady=5, side="right", fill="both", expand=True)
    
    def _create_button(self, parent: tk.Widget, text: str, command, 
                       bg: str = None, side: str = "left") -> tk.Button:
        """
        스타일이 적용된 버튼 생성
        
        Args:
            parent: 버튼을 배치할 부모 위젯
            text: 버튼에 표시될 텍스트
            command: 버튼 클릭 시 실행될 함수
            bg: 배경색 (기본값: THEME.ACCENT_DARK)
            side: 배치 방향 ('left' 또는 'right')
        
        Returns:
            생성된 버튼 객체
        
        Features:
            - 호버 효과: 마우스 올리면 파랑색으로 변경
            - 손모양 커서를 표시하여 클릭 가능하다는 것을 시각적으로 표현
        """
        # 버튼 위젯 생성
        btn = tk.Button(
            parent, text=text, command=command,
            font=FONTS.BUTTON, fg=THEME.FG_LIGHT,  # 흰색 텍스트
            bg=bg or THEME.ACCENT_DARK,  # 배경색 (지정 안하면 어두운 회색)
            activebackground=THEME.ACCENT_PRIMARY,  # 클릭 시 배경색
            activeforeground=THEME.FG_LIGHT,  # 클릭 시 텍스트 색상
            relief="flat",  # 평평한 모양 (입체감 없음)
            padx=15, pady=5,  # 내부 여백
            cursor="hand2"  # 손모양 커서
        )
        btn.pack(side=side, padx=5, pady=5)
        
        # 호버 효과 구현 (마우스 올리면 색상 변경)
        original_bg = bg or THEME.ACCENT_DARK
        def on_enter(e): btn.config(bg=THEME.ACCENT_PRIMARY)  # 마우스 진입: 파랑색
        def on_leave(e): btn.config(bg=original_bg)  # 마우스 이탈: 원래 색상
        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        
        return btn
    
    # ========================================================================
    # 로딩 오버레이
    # ========================================================================
    def _create_loading_overlay(self) -> None:
        """
        오버레이 스타일의 로딩 UI 생성
        
        사용자가 데이터 로딩 중임을 인지할 수 있도록 화면 중앙에 로딩 메시지를 표시합니다.
        오버레이 방식으로 화면 전체를 덕고, 중앙에 로딩 박스를 배치합니다.
        
        구성요소:
        - 전체 화면 오버레이 (사용자 클릭 차단)
        - 중앙 로딩 박스 (메시지 + 프로그레스바 + 상태 텍스트)
        """
        # 전체 화면을 덕는 오버레이 프레임
        self.loading_overlay = tk.Frame(self.root, bg=THEME.BG_PRIMARY)
        
        # 중앙에 배치될 로딩 박스 (흰색 배경, 테두리 있음)
        self.loading_box = tk.Frame(
            self.loading_overlay, bg=THEME.BG_SECONDARY,
            padx=40, pady=25,  # 내부 여백
            highlightbackground=THEME.BORDER,  # 테두리 색상
            highlightthickness=1  # 테두리 두께
        )
        # 화면 중앙에 배치 (relx=0.5, rely=0.5 = 중앙, anchor="center" = 중심점 기준)
        self.loading_box.place(relx=0.5, rely=0.5, anchor="center")
        
        # 주요 로딩 메시지 (모래시계 이모티콘 + 텍스트)
        self.loading_label = tk.Label(
            self.loading_box, text="⌛ 데이터 로딩 중...",
            font=FONTS.LOADING,  # 크고 굵은 폰트
            bg=THEME.BG_SECONDARY, fg=THEME.ACCENT_PRIMARY  # 파랑색 텍스트
        )
        self.loading_label.pack(pady=(0, 15))  # 아래 여백 15px
        
        # 무한 회전 프로그레스바 (진행률을 알 수 없을 때 사용)
        self.loading_bar = ttk.Progressbar(
            self.loading_box, 
            mode='indeterminate',  # 무한 회전 모드
            length=280  # 프로그레스바 길이 (px)
        )
        self.loading_bar.pack()
        
        # 보조 상태 텍스트 (작고 회색)
        self.loading_status = tk.Label(
            self.loading_box, text="잠시만 기다려주세요...",
            font=FONTS.SMALL,  # 작은 폰트
            bg=THEME.BG_SECONDARY, fg=THEME.FG_SECONDARY  # 회색 텍스트
        )
        self.loading_status.pack(pady=(10, 0))  # 위 여백 10px
    
    def _show_loading(self, message: str = "데이터 로딩 중...") -> None:
        """로딩 오버레이 표시"""
        self.loading_label.config(text=f"⏳ {message}")
        self.loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.loading_bar.start(15)
        self.root.update_idletasks()
    
    def _hide_loading(self) -> None:
        """로딩 오버레이 숨김"""
        self.loading_bar.stop()
        self.loading_overlay.place_forget()
        self.root.update_idletasks()
    
    # ========================================================================
    # 데이터 로드
    # ========================================================================
    def load_data_threaded(self) -> None:
        """데이터 로드 (스레드)"""
        if self._is_loading:
            return
        self._is_loading = True
        self._show_loading("데이터 로딩 중...")
        threading.Thread(target=self._load_data_worker, daemon=True).start()
    
    def _load_data_worker(self) -> None:
        """데이터 로드 작업 (백그라운드)"""
        try:
            # 경로 탐색
            if os.path.isdir(CONFIG.NAS_PATH):
                self.base_path = CONFIG.NAS_PATH
            elif os.path.isdir(CONFIG.LOCAL_PATH):
                self.base_path = CONFIG.LOCAL_PATH
            else:
                raise FileNotFoundError('NAS451서버 또는 로컬 DB 폴더를 찾을 수 없습니다.')
            
            pickle_path = self.base_path + CONFIG.PICKLE_FILE
            excel_path = self.base_path + CONFIG.EXCEL_FILE
            
            # 피클 우선, 없으면 엑셀
            if os.path.exists(pickle_path):
                with open(pickle_path, 'rb') as fr:
                    data = pickle.load(fr)
            else:
                data = pd.read_excel(excel_path, sheet_name=CONFIG.SHEET_NAME, header=0)
            
            # 문자열 변환 (fillna를 먼저 적용해야 NaN이 "nan" 문자열로 변환되는 것을 방지)
            data = data.fillna("").astype(str)
            
            self.root.after(0, self._on_data_loaded, data)
        except Exception as e:
            self.root.after(0, self._on_data_load_error, str(e))
    
    def _on_data_loaded(self, data: pd.DataFrame) -> None:
        """데이터 로드 완료 콜백"""
        self.data = data
        self.filtered_data = data.copy()
        self._update_sheet(self.data)
        self._hide_loading()
        self._is_loading = False
    
    def _on_data_load_error(self, error_msg: str) -> None:
        """데이터 로드 에러 콜백"""
        messagebox.showerror('에러', f'상품정보 파일을 불러올 수 없습니다!\n{error_msg}')
        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()
        self._hide_loading()
        self._is_loading = False
    
    # ========================================================================
    # 데이터 업데이트
    # ========================================================================
    def update_data_threaded(self) -> None:
        """등록상품 업데이트 (스레드)"""
        if self._is_loading:
            messagebox.showinfo('알림', '데이터 처리 중입니다. 잠시 후 다시 시도해주세요.')
            return
        self._is_loading = True
        self._show_loading("등록상품 업데이트 중...")
        threading.Thread(target=self._update_data_worker, daemon=True).start()
    
    def _update_data_worker(self) -> None:
        """데이터 업데이트 작업 (백그라운드)"""
        try:
            if not self.base_path:
                raise FileNotFoundError("DB 경로를 찾을 수 없습니다.")
            
            excel_path = self.base_path + CONFIG.EXCEL_FILE
            pickle_path = self.base_path + CONFIG.PICKLE_FILE
            
            df = pd.read_excel(excel_path, sheet_name=CONFIG.SHEET_NAME, header=0)
            
            # 피클 저장
            with open(pickle_path, 'wb') as fw:
                pickle.dump(df, fw)
            
            df = df.fillna("").astype(str)
            
            self.root.after(0, self._on_data_updated, df)
        except Exception as e:
            self.root.after(0, self._on_data_update_error, str(e))
    
    def _on_data_updated(self, df: pd.DataFrame) -> None:
        """데이터 업데이트 완료 콜백"""
        messagebox.showinfo('알림', '업데이트가 완료되었습니다')
        self.data = df
        self.filtered_data = df.copy()
        self._update_sheet(self.data)
        self._hide_loading()
        self._is_loading = False
    
    def _on_data_update_error(self, error_msg: str) -> None:
        """데이터 업데이트 에러 콜백"""
        messagebox.showerror('에러', f'업데이트 실패: {error_msg}')
        self._hide_loading()
        self._is_loading = False
    
    # ========================================================================
    # 검색 필터링
    # ========================================================================
    def filter_data_threaded(self) -> None:
        """검색 필터링 (스레드)"""
        if self.data.empty:
            messagebox.showinfo("알림", "데이터가 로드되지 않았습니다.")
            return
        
        input_str = self.input_text.get("1.0", tk.END).strip()
        if not input_str:
            messagebox.showinfo("알림", "상품코드 또는 SKU ID를 입력해 주세요!")
            return
        
        if self._is_loading:
            return
        
        self._is_loading = True
        self._show_loading("검색 중...")
        threading.Thread(
            target=self._filter_data_worker, 
            args=(input_str,), 
            daemon=True
        ).start()
    
    def _filter_data_worker(self, input_str: str) -> None:
        """검색 작업 (백그라운드)"""
        try:
            # 검색어 파싱
            search_items = [s.strip() for s in input_str.replace('\n', ' ').split() if s.strip()]
            if not search_items:
                self.root.after(0, self._on_filter_error, "검색어가 없습니다.")
                return
            
            # 검색 컬럼 결정
            option = self.search_opt.get()
            if option == "상품코드로 검색":
                key = "상품코드"
            elif option == "SKU_ID로 검색":
                key = "SKU"
            elif option == "쿠팡상품코드로 검색":
                key = "쿠팡상품코드"
            else:
                key = "상품코드"  # 기본값
            
            if key not in self.data.columns:
                self.root.after(0, self._on_filter_error, f"'{key}' 컬럼이 존재하지 않습니다.")
                return
            
            # 검색 실행
            if len(search_items) == 1:
                mask = self.data[key].str.contains(search_items[0], na=False, regex=False)
            else:
                pattern = '|'.join([re.escape(item) for item in search_items])
                mask = self.data[key].str.contains(pattern, na=False, regex=True)
            
            filtered = self.data[mask].copy()
            self.root.after(0, self._on_filter_complete, filtered)
            
        except Exception as e:
            self.root.after(0, self._on_filter_error, str(e))
    
    def _on_filter_complete(self, filtered: pd.DataFrame) -> None:
        """검색 완료 콜백"""
        # 상품코드, 컬러, 사이즈 순으로 오름차순 정렬
        sort_cols = [col for col in ["쿠팡상품코드", "컬러", "사이즈"] if col in filtered.columns]
        if sort_cols:
            filtered = filtered.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)
        self.filtered_data = filtered
        self._update_sheet(filtered)
        self._hide_loading()
        self._is_loading = False
    
    def _on_filter_error(self, error_msg: str) -> None:
        """검색 에러 콜백"""
        messagebox.showerror("에러", error_msg)
        self._hide_loading()
        self._is_loading = False
    
    # ========================================================================
    # 유틸리티 메서드
    # ========================================================================
    def show_all_data(self) -> None:
        """전체 데이터 보기 (스레드)"""
        if self.data.empty:
            messagebox.showinfo("알림", "데이터가 로드되지 않았습니다.")
            return
        
        if self._is_loading:
            return
        
        self.input_text.delete("1.0", tk.END)
        self._is_loading = True
        self._show_loading("전체 데이터 불러오는 중...")
        threading.Thread(target=self._show_all_data_worker, daemon=True).start()
    
    def _show_all_data_worker(self) -> None:
        """전체 데이터 보기 작업 (백그라운드)"""
        try:
            filtered = self.data.copy()
            self.root.after(0, self._on_show_all_complete, filtered)
        except Exception as e:
            self.root.after(0, self._on_show_all_error, str(e))
    
    def _on_show_all_complete(self, data: pd.DataFrame) -> None:
        """전체 데이터 보기 완료 콜백"""
        self.filtered_data = data
        self._update_sheet(data)
        self._hide_loading()
        self._is_loading = False
    
    def _on_show_all_error(self, error_msg: str) -> None:
        """전체 데이터 보기 에러 콜백"""
        messagebox.showerror("에러", error_msg)
        self._hide_loading()
        self._is_loading = False
    
    def reset(self) -> None:
        """입력/검색 초기화"""
        self.input_text.delete("1.0", tk.END)
        self.filtered_data = self.data.copy()
        self.sort_column = None  # 정렬 상태 초기화
        self.sort_ascending = True
        for widget in self.frm_sheet.winfo_children():
            widget.destroy()
    
    def _sort_by_column(self, event) -> None:
        """헤더 클릭 시 정렬 수행 (토글 방식)"""
        if not hasattr(event, 'value') or self.filtered_data.empty:
            return
        
        column_idx = event.value
        
        # 동일 컬럼 클릭 시 정렬 순서 토글 (오름차순 ↔ 내림차순)
        if self.sort_column == column_idx:
            self.sort_ascending = not self.sort_ascending
        else:
            # 새로운 컬럼 클릭 시 오름차순으로 시작
            self.sort_column = column_idx
            self.sort_ascending = True
        
        # 정렬 수행
        column_name = self.filtered_data.columns[column_idx]
        self.filtered_data = self.filtered_data.sort_values(
            by=column_name, 
            ascending=self.sort_ascending
        ).reset_index(drop=True)
        
        # 시트 업데이트 (정렬 표시 포함)
        self._update_sheet(self.filtered_data, preserve_sort_indicator=True)
    

    
    def _update_sheet(self, df: pd.DataFrame, preserve_sort_indicator: bool = False) -> None:
        """시트 업데이트"""
        for widget in self.frm_sheet.winfo_children():
            widget.destroy()
        
        if df.empty:
            tk.Label(
                self.frm_sheet, text="조회할 데이터가 없습니다.",
                fg=THEME.ACCENT_DANGER, bg=THEME.BG_PRIMARY, font=FONTS.NORMAL
            ).pack(padx=10, pady=20)
            return
        
        data_list = df.values.tolist()
        headers_list = list(df.columns)
        
        # 정렬 표시 추가 (▲: 오름차순, ▼: 내림차순)
        if preserve_sort_indicator and self.sort_column is not None:
            sort_indicator = " ▲" if self.sort_ascending else " ▼"
            headers_list[self.sort_column] = headers_list[self.sort_column] + sort_indicator
        
        sheet = Sheet(
            self.frm_sheet, data=data_list, headers=headers_list,
            header_height=28, header_fg=THEME.HEADER_FG, header_bg=THEME.HEADER_BG
        )
        sheet.header_font((FONTS.FAMILY, 10, 'bold'))
        sheet.font((FONTS.FAMILY, 9, 'normal'))
        sheet.table_align(align="left")
        sheet.set_all_column_widths(width=None)  # 기본: 자동 너비
        
        # 컬럼명 기반 너비 적용
        for col_idx, col_name in enumerate(df.columns):
            if col_name in COLUMN_WIDTHS:
                sheet.column_width(column=col_idx, width=COLUMN_WIDTHS[col_name])
        
        # 헤더 클릭 바인딩 활성화
        sheet.enable_bindings()
        sheet.bind("<ButtonPress-1>", self._on_sheet_click, add=True)
        
        sheet.pack(fill="both", expand=True)
        self.current_sheet = sheet
    
    def _on_sheet_click(self, event) -> None:
        """시트 클릭 이벤트 핸들러"""
        if self.current_sheet and hasattr(event, 'name'):
            # 헤더 영역 클릭 확인
            region = self.current_sheet.identify_region(event)
            if region and region.type_ == "header":
                # 클릭한 컬럼 인덱스로 정렬
                class ColumnEvent:
                    def __init__(self, col_idx):
                        self.value = col_idx
                
                if region.column is not None:
                    self._sort_by_column(ColumnEvent(region.column))
    




# ============================================================================
# 메인 실행
# ============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = CoupangSKUApp(root)
    root.mainloop()
