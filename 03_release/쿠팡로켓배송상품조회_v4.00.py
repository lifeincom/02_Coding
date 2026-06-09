# 쿠팡 로켓배송 상품조회 v4.00

# [주요 기능]
# - 쿠팡 SKU 데이터베이스 조회 및 검색
# - 상품코드 / SKU_ID / 쿠팡상품코드 / 상품명 / 바코드 검색 (5가지)
# - 발주가능상태 · 공급상태 추가 필터 동시 적용
# - 검색 히스토리 (최근 10개 저장 · 재사용)
# - 검색 결과 Excel/CSV 내보내기
# - 행 더블클릭 상세 정보 팝업 + 클립보드 복사
# - 하단 상태바 (전체 건수 · 검색 결과 건수 · 파일 경로)
# - 테이블 헤더 클릭 오름차순/내림차순 정렬
# - 다중 스레드를 이용한 비동기 데이터 로딩 및 검색
# - Pickle 신선도 자동 체크 (Excel 변경 시 자동 갱신)
# - Ctrl+Enter 검색 단축키

# [v4.00 개선 사항]
# - 속도: pickle 신선도 검사 추가, 단계별 로딩 메시지
# - 검색: 상품명·바코드 검색 추가, 발주가능상태·공급상태 필터 추가
# - UI: 하단 상태바, 결과 건수 뱃지, Ctrl+Enter 단축키, 검색 타입 레이블
# - 로딩: 단계별 진행 메시지 (경로탐색 → 파일읽기 → 데이터처리)
# - 기능: Excel/CSV 내보내기, 검색 히스토리, 행 더블클릭 상세 팝업

# [데이터 소스]
# - NAS 서버: \\NAS451\team451\DB\쿠팡SKU리스트.xlsm
# - 로컬 백업: D:\NAS451\DB\쿠팡SKU리스트.xlsm

import os
import pickle
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from typing import Optional, List

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tksheet import Sheet


# ============================================================================
# 설정 상수
# ============================================================================
@dataclass(frozen=True)
class AppConfig:
    TITLE: str = "쿠팡 로켓배송 상품조회 v4.00"
    GEOMETRY: str = "1800x900+60+60"
    NAS_PATH: str = r"\\NAS451\team451\DB"
    LOCAL_PATH: str = r"D:\NAS451\DB"
    PICKLE_FILE: str = r"\쿠팡SKU리스트.pickle"
    EXCEL_FILE: str = r"\쿠팡SKU리스트.xlsm"
    SHEET_NAME: str = "SKU리스트"
    MAX_HISTORY: int = 10


@dataclass(frozen=True)
class Theme:
    BG_PRIMARY: str = "#F5F7FA"
    BG_SECONDARY: str = "#FFFFFF"
    BG_STATUS: str = "#EDF2F7"
    FG_PRIMARY: str = "#2D3748"
    FG_SECONDARY: str = "#718096"
    FG_LIGHT: str = "#FFFFFF"
    ACCENT_PRIMARY: str = "#4A90D9"
    ACCENT_SUCCESS: str = "#48BB78"
    ACCENT_DANGER: str = "#E53E3E"
    ACCENT_DARK: str = "#2D3748"
    BORDER: str = "#E2E8F0"
    HEADER_BG: str = "#4A5568"
    HEADER_FG: str = "#FFFFFF"


@dataclass(frozen=True)
class Fonts:
    FAMILY: str = "NanumGothic"
    TITLE: tuple = ("NanumGothic", 11, "bold")
    NORMAL: tuple = ("NanumGothic", 10, "normal")
    SMALL: tuple = ("NanumGothic", 9, "normal")
    BUTTON: tuple = ("NanumGothic", 10, "normal")
    LOADING: tuple = ("NanumGothic", 12, "bold")
    STATUS: tuple = ("NanumGothic", 9, "normal")


CONFIG = AppConfig()
THEME = Theme()
FONTS = Fonts()

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

# 검색 옵션: 표시명 → 컬럼명
SEARCH_OPTIONS: dict = {
    "상품코드로 검색":     "상품코드",
    "쿠팡상품코드로 검색":  "쿠팡상품코드",
    "SKU_ID로 검색":       "SKU",
    "상품명으로 검색":     "상품명",
    "바코드로 검색":       "바코드",
}


# ============================================================================
# 메인 애플리케이션
# ============================================================================
class CoupangSKUApp:
    """쿠팡 SKU 조회 애플리케이션 v4.00"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self._configure_root()
        self._configure_styles()

        # 데이터 상태
        self.data: pd.DataFrame = pd.DataFrame()
        self.filtered_data: pd.DataFrame = pd.DataFrame()
        self.base_path: str = ""
        self._is_loading: bool = False

        # 정렬 상태
        self.sort_column: Optional[int] = None
        self.sort_ascending: bool = True
        self.current_sheet: Optional[Sheet] = None

        # 검색 히스토리
        self.search_history: List[str] = []

        self._build_ui()
        self._create_loading_overlay()
        self.load_data_threaded()

    # ========================================================================
    # 초기 설정
    # ========================================================================
    def _configure_root(self) -> None:
        self.root.title(CONFIG.TITLE)
        self.root.geometry(CONFIG.GEOMETRY)
        self.root.state("zoomed")
        self.root.config(background=THEME.BG_PRIMARY)
        self.root.protocol("WM_DELETE_WINDOW", self.root.quit)

    def _configure_styles(self) -> None:
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TCombobox",
                        fieldbackground=THEME.BG_SECONDARY,
                        background=THEME.BG_SECONDARY,
                        foreground=THEME.FG_PRIMARY)
        style.configure("TProgressbar",
                        troughcolor=THEME.BORDER,
                        background=THEME.ACCENT_PRIMARY,
                        thickness=8)
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
        self._build_top_frame()
        self._build_function_frame()
        self._build_filter_frame()
        self._build_main_container()
        self._build_status_bar()

    def _build_top_frame(self) -> None:
        self.frm_top = tk.Frame(self.root, background=THEME.BG_SECONDARY)
        self.frm_top.pack(padx=15, pady=(5, 0), fill="x")

        desc = "1. 검색유형을 선택하고 검색 정보를 입력해 주세요.   2. 소스데이터 : /nas451/DB/쿠팡SKU리스트.xlsm"
        tk.Label(
            self.frm_top, text=desc, height=2, justify="left",
            bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY, font=FONTS.NORMAL
        ).pack(side="left", padx=10, pady=5)

        self._create_button(self.frm_top, "프로그램 종료", self.root.quit,
                            bg=THEME.ACCENT_DARK, side="right")

    def _build_function_frame(self) -> None:
        self.frm_function = tk.Frame(self.root, background=THEME.BG_SECONDARY, pady=5)
        self.frm_function.pack(padx=15, pady=(0, 2), fill="x")

        # 좌측: 검색
        left_group = tk.Frame(self.frm_function, background=THEME.BG_SECONDARY)
        left_group.pack(side="left", padx=5)

        tk.Label(left_group, text="검색유형:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_PRIMARY, font=FONTS.NORMAL).pack(side="left", padx=(5, 2))

        self.search_opt = ttk.Combobox(
            left_group, state="readonly",
            values=list(SEARCH_OPTIONS.keys()), width=18
        )
        self.search_opt.current(0)
        self.search_opt.pack(side="left", padx=5, ipady=5)

        self._create_button(left_group, "🔍 검색", self.filter_data_threaded,
                            bg=THEME.ACCENT_PRIMARY)
        self._create_button(left_group, "↺ 초기화", self.reset,
                            bg=THEME.ACCENT_DARK)
        self._create_button(left_group, "📋 검색 히스토리", self._show_history,
                            bg=THEME.ACCENT_DARK)

        # 우측: 데이터 관리
        right_group = tk.Frame(self.frm_function, background=THEME.BG_SECONDARY)
        right_group.pack(side="right", padx=5)

        self._create_button(right_group, "📤 결과 내보내기", self._export_data,
                            bg=THEME.ACCENT_SUCCESS)
        self._create_button(right_group, "등록상품 업데이트", self.update_data_threaded,
                            bg=THEME.ACCENT_DARK)
        self._create_button(right_group, "전체 등록상품 보기", self.show_all_data,
                            bg=THEME.ACCENT_DARK)

    def _build_filter_frame(self) -> None:
        """발주가능상태 · 공급상태 · 등록계정 · 상품분류 추가 필터"""
        self.frm_filter = tk.Frame(self.root, background=THEME.BG_SECONDARY, pady=3)
        self.frm_filter.pack(padx=15, pady=(0, 5), fill="x")

        tk.Label(self.frm_filter, text="추가 필터:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_SECONDARY, font=FONTS.SMALL).pack(side="left", padx=(10, 5))

        tk.Label(self.frm_filter, text="발주가능상태:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_PRIMARY, font=FONTS.SMALL).pack(side="left", padx=(5, 2))
        self.filter_order = ttk.Combobox(
            self.frm_filter, state="readonly",
            values=["전체", "가능", "불가능"], width=8
        )
        self.filter_order.current(0)
        self.filter_order.pack(side="left", padx=(0, 10), ipady=3)

        tk.Label(self.frm_filter, text="공급상태:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_PRIMARY, font=FONTS.SMALL).pack(side="left", padx=(5, 2))
        self.filter_supply = ttk.Combobox(
            self.frm_filter, state="readonly",
            values=["전체", "정상", "단종", "중단"], width=8
        )
        self.filter_supply.current(0)
        self.filter_supply.pack(side="left", padx=(0, 10), ipady=3)

        tk.Label(self.frm_filter, text="등록계정:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_PRIMARY, font=FONTS.SMALL).pack(side="left", padx=(5, 2))
        self.filter_account = ttk.Combobox(
            self.frm_filter, state="readonly",
            values=["전체"], width=12
        )
        self.filter_account.current(0)
        self.filter_account.pack(side="left", padx=(0, 10), ipady=3)

        tk.Label(self.frm_filter, text="상품분류:", bg=THEME.BG_SECONDARY,
                 fg=THEME.FG_PRIMARY, font=FONTS.SMALL).pack(side="left", padx=(5, 2))
        self.filter_category = ttk.Combobox(
            self.frm_filter, state="readonly",
            values=["전체"], width=12
        )
        self.filter_category.current(0)
        self.filter_category.pack(side="left", padx=(0, 10), ipady=3)

        # 검색 결과 건수 뱃지
        self.lbl_count = tk.Label(
            self.frm_filter, text="",
            bg=THEME.BG_SECONDARY, fg=THEME.ACCENT_PRIMARY, font=FONTS.NORMAL
        )
        self.lbl_count.pack(side="right", padx=15)

    def _build_main_container(self) -> None:
        self.frm_container = tk.Frame(self.root, background=THEME.BG_PRIMARY)
        self.frm_container.pack(padx=15, pady=5, fill="both", expand=True)

        # 좌측: 검색 입력
        self.frm_input = ttk.LabelFrame(self.frm_container, text="  검색 리스트 입력  ")
        self.frm_input.pack(padx=5, pady=5, side="left", fill="both")

        self.input_text = tk.Text(
            self.frm_input, width=20, font=FONTS.NORMAL,
            bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
            relief="flat", borderwidth=1, highlightthickness=1,
            highlightbackground=THEME.BORDER,
            highlightcolor=THEME.ACCENT_PRIMARY
        )
        self.input_text.pack(padx=8, pady=8, fill="both", expand=True)

        # Ctrl+Enter 검색 / Ctrl+A 전체선택
        self.input_text.bind("<Control-Return>", lambda e: self.filter_data_threaded())
        self.input_text.bind("<Control-a>", lambda e: (
            self.input_text.tag_add("sel", "1.0", "end"), "break"
        ))

        tk.Label(
            self.frm_input, text="Ctrl+Enter: 검색",
            bg=THEME.BG_SECONDARY, fg=THEME.FG_SECONDARY, font=FONTS.SMALL
        ).pack(pady=(0, 5))

        # 우측: 결과 시트
        self.frm_sheet = ttk.LabelFrame(self.frm_container, text="  쿠팡 로켓배송 상품정보 출력  ")
        self.frm_sheet.pack(padx=5, pady=5, side="right", fill="both", expand=True)

    def _build_status_bar(self) -> None:
        """하단 상태바: 전체 건수 · 검색 결과 · 파일 경로 · 갱신 시각"""
        self.frm_status = tk.Frame(self.root, bg=THEME.BG_STATUS, height=25)
        self.frm_status.pack(side="bottom", fill="x")
        self.frm_status.pack_propagate(False)

        def sep():
            tk.Label(self.frm_status, text="|", bg=THEME.BG_STATUS,
                     fg=THEME.BORDER, font=FONTS.STATUS).pack(side="left")

        self.lbl_status_total = tk.Label(
            self.frm_status, text="전체: 0건",
            bg=THEME.BG_STATUS, fg=THEME.FG_SECONDARY, font=FONTS.STATUS
        )
        self.lbl_status_total.pack(side="left", padx=15)
        sep()

        self.lbl_status_filtered = tk.Label(
            self.frm_status, text="검색결과: -",
            bg=THEME.BG_STATUS, fg=THEME.FG_SECONDARY, font=FONTS.STATUS
        )
        self.lbl_status_filtered.pack(side="left", padx=15)
        sep()

        self.lbl_status_path = tk.Label(
            self.frm_status, text="데이터 경로: -",
            bg=THEME.BG_STATUS, fg=THEME.FG_SECONDARY, font=FONTS.STATUS
        )
        self.lbl_status_path.pack(side="left", padx=15)

        self.lbl_status_time = tk.Label(
            self.frm_status, text="",
            bg=THEME.BG_STATUS, fg=THEME.FG_SECONDARY, font=FONTS.STATUS
        )
        self.lbl_status_time.pack(side="right", padx=15)

    def _create_button(self, parent: tk.Widget, text: str, command,
                       bg: str = None, side: str = "left") -> tk.Button:
        btn = tk.Button(
            parent, text=text, command=command,
            font=FONTS.BUTTON, fg=THEME.FG_LIGHT,
            bg=bg or THEME.ACCENT_DARK,
            activebackground=THEME.ACCENT_PRIMARY,
            activeforeground=THEME.FG_LIGHT,
            relief="flat", padx=15, pady=5, cursor="hand2"
        )
        btn.pack(side=side, padx=5, pady=5)

        original_bg = bg or THEME.ACCENT_DARK
        btn.bind("<Enter>", lambda e: btn.config(bg=THEME.ACCENT_PRIMARY))
        btn.bind("<Leave>", lambda e: btn.config(bg=original_bg))
        return btn

    # ========================================================================
    # 로딩 오버레이
    # ========================================================================
    def _create_loading_overlay(self) -> None:
        self.loading_overlay = tk.Frame(self.root, bg=THEME.BG_PRIMARY)

        self.loading_box = tk.Frame(
            self.loading_overlay, bg=THEME.BG_SECONDARY,
            padx=40, pady=25,
            highlightbackground=THEME.BORDER,
            highlightthickness=1
        )
        self.loading_box.place(relx=0.5, rely=0.5, anchor="center")

        self.loading_label = tk.Label(
            self.loading_box, text="⌛ 데이터 로딩 중...",
            font=FONTS.LOADING, bg=THEME.BG_SECONDARY, fg=THEME.ACCENT_PRIMARY
        )
        self.loading_label.pack(pady=(0, 15))

        self.loading_bar = ttk.Progressbar(
            self.loading_box, mode='indeterminate', length=280
        )
        self.loading_bar.pack()

        self.loading_status = tk.Label(
            self.loading_box, text="잠시만 기다려주세요...",
            font=FONTS.SMALL, bg=THEME.BG_SECONDARY, fg=THEME.FG_SECONDARY
        )
        self.loading_status.pack(pady=(10, 0))

        # 단계별 진행 메시지
        self.loading_step = tk.Label(
            self.loading_box, text="",
            font=FONTS.SMALL, bg=THEME.BG_SECONDARY, fg=THEME.FG_SECONDARY
        )
        self.loading_step.pack(pady=(4, 0))

    def _show_loading(self, message: str = "데이터 로딩 중...",
                      status: str = "잠시만 기다려주세요...") -> None:
        self.loading_label.config(text=f"⏳ {message}")
        self.loading_status.config(text=status)
        self.loading_step.config(text="")
        self.loading_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.loading_bar.start(15)
        self.root.update_idletasks()

    def _set_loading_step(self, step: str) -> None:
        """백그라운드 스레드에서 단계 메시지 업데이트"""
        self.root.after(0, lambda: self.loading_step.config(text=step))

    def _hide_loading(self) -> None:
        self.loading_bar.stop()
        self.loading_overlay.place_forget()
        self.root.update_idletasks()

    # ========================================================================
    # 데이터 로드
    # ========================================================================
    def load_data_threaded(self) -> None:
        if self._is_loading:
            return
        self._is_loading = True
        self._show_loading("데이터 로딩 중...", "파일 경로를 확인하는 중...")
        threading.Thread(target=self._load_data_worker, daemon=True).start()

    def _load_data_worker(self) -> None:
        try:
            self._set_loading_step("1/3  경로 탐색 중...")
            if os.path.isdir(CONFIG.NAS_PATH):
                self.base_path = CONFIG.NAS_PATH
            elif os.path.isdir(CONFIG.LOCAL_PATH):
                self.base_path = CONFIG.LOCAL_PATH
            else:
                raise FileNotFoundError("NAS451서버 또는 로컬 DB 폴더를 찾을 수 없습니다.")

            pickle_path = self.base_path + CONFIG.PICKLE_FILE
            excel_path  = self.base_path + CONFIG.EXCEL_FILE

            # pickle 신선도 체크: Excel이 pickle보다 최신이면 재생성
            use_pickle = False
            if os.path.exists(pickle_path):
                p_mtime = os.path.getmtime(pickle_path)
                e_mtime = os.path.getmtime(excel_path) if os.path.exists(excel_path) else 0
                use_pickle = p_mtime >= e_mtime

            if use_pickle:
                self._set_loading_step("2/3  캐시 파일 읽는 중...")
                with open(pickle_path, 'rb') as f:
                    data = pickle.load(f)
            else:
                self._set_loading_step("2/3  엑셀 파일 읽는 중...")
                data = pd.read_excel(excel_path, sheet_name=CONFIG.SHEET_NAME, header=0)
                # 다음 실행을 위해 pickle 자동 저장
                try:
                    with open(pickle_path, 'wb') as f:
                        pickle.dump(data, f)
                except Exception:
                    pass

            self._set_loading_step("3/3  데이터 처리 중...")
            for col in ["견적공급가", "판매가"]:
                if col in data.columns:
                    data[col] = pd.to_numeric(data[col], errors='coerce').fillna(0).astype('int64')
            data = data.fillna("").astype(str)
            for col in ["견적공급가", "판매가"]:
                if col in data.columns:
                    data[col] = data[col].replace('0', '')

            self.root.after(0, self._on_data_loaded, data)
        except Exception as e:
            self.root.after(0, self._on_data_load_error, str(e))

    def _populate_dynamic_filters(self) -> None:
        """등록계정 · 상품분류 콤보박스를 데이터의 고유값으로 채움"""
        for col, combobox in [("등록계정", self.filter_account), ("상품분류", self.filter_category)]:
            if col in self.data.columns:
                vals = sorted({v for v in self.data[col].tolist() if v and v != "nan"})
                combobox["values"] = ["전체"] + vals
                combobox.current(0)

    def _on_data_loaded(self, data: pd.DataFrame) -> None:
        self.data = data
        self.filtered_data = data.copy()
        self._populate_dynamic_filters()
        self._update_sheet(self.data)
        self._refresh_status_bar()
        self._hide_loading()
        self._is_loading = False

    def _on_data_load_error(self, error_msg: str) -> None:
        messagebox.showerror("에러", f"상품정보 파일을 불러올 수 없습니다!\n{error_msg}")
        self.data = pd.DataFrame()
        self.filtered_data = pd.DataFrame()
        self._hide_loading()
        self._is_loading = False

    # ========================================================================
    # 데이터 업데이트
    # ========================================================================
    def update_data_threaded(self) -> None:
        if self._is_loading:
            messagebox.showinfo("알림", "데이터 처리 중입니다. 잠시 후 다시 시도해주세요.")
            return
        self._is_loading = True
        self._show_loading("등록상품 업데이트 중...", "엑셀 파일을 읽는 중...")
        threading.Thread(target=self._update_data_worker, daemon=True).start()

    def _update_data_worker(self) -> None:
        try:
            if not self.base_path:
                raise FileNotFoundError("DB 경로를 찾을 수 없습니다.")

            excel_path  = self.base_path + CONFIG.EXCEL_FILE
            pickle_path = self.base_path + CONFIG.PICKLE_FILE

            self._set_loading_step("1/2  엑셀 파일 읽는 중...")
            df = pd.read_excel(excel_path, sheet_name=CONFIG.SHEET_NAME, header=0)

            self._set_loading_step("2/2  캐시 저장 중...")
            with open(pickle_path, 'wb') as f:
                pickle.dump(df, f)

            for col in ["견적공급가", "판매가"]:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype('int64')
            df = df.fillna("").astype(str)
            for col in ["견적공급가", "판매가"]:
                if col in df.columns:
                    df[col] = df[col].replace('0', '')
            self.root.after(0, self._on_data_updated, df)
        except Exception as e:
            self.root.after(0, self._on_data_update_error, str(e))

    def _on_data_updated(self, df: pd.DataFrame) -> None:
        messagebox.showinfo("알림", f"업데이트가 완료되었습니다.\n총 {len(df):,}개 상품")
        self.data = df
        self.filtered_data = df.copy()
        self._populate_dynamic_filters()
        self._update_sheet(self.data)
        self._refresh_status_bar()
        self._hide_loading()
        self._is_loading = False

    def _on_data_update_error(self, error_msg: str) -> None:
        messagebox.showerror("에러", f"업데이트 실패: {error_msg}")
        self._hide_loading()
        self._is_loading = False

    # ========================================================================
    # 검색 필터링
    # ========================================================================
    def filter_data_threaded(self) -> None:
        if self.data.empty:
            messagebox.showinfo("알림", "데이터가 로드되지 않았습니다.")
            return

        input_str = self.input_text.get("1.0", tk.END).strip()
        if not input_str:
            messagebox.showinfo("알림", "검색할 코드 또는 텍스트를 입력해 주세요!")
            return

        if self._is_loading:
            return

        # UI에서 필터 값을 미리 읽어 스레드에 전달 (스레드 안전)
        order_filter    = self.filter_order.get()
        supply_filter   = self.filter_supply.get()
        account_filter  = self.filter_account.get()
        category_filter = self.filter_category.get()

        self._is_loading = True
        self._show_loading("검색 중...", "데이터를 필터링하는 중...")
        threading.Thread(
            target=self._filter_data_worker,
            args=(input_str, order_filter, supply_filter, account_filter, category_filter),
            daemon=True
        ).start()

    def _filter_data_worker(self, input_str: str,
                            order_filter: str, supply_filter: str,
                            account_filter: str, category_filter: str) -> None:
        try:
            search_items = [s.strip() for s in input_str.replace('\n', ' ').split() if s.strip()]
            if not search_items:
                self.root.after(0, self._on_filter_error, "검색어가 없습니다.")
                return

            key = SEARCH_OPTIONS.get(self.search_opt.get(), "상품코드")
            if key not in self.data.columns:
                self.root.after(0, self._on_filter_error, f"'{key}' 컬럼이 존재하지 않습니다.")
                return

            # 키워드 검색
            if len(search_items) == 1:
                mask = self.data[key].str.contains(search_items[0], na=False, regex=False)
            else:
                pattern = '|'.join(re.escape(item) for item in search_items)
                mask = self.data[key].str.contains(pattern, na=False, regex=True)

            filtered = self.data[mask].copy()

            # 발주가능상태 필터
            if order_filter != "전체" and "발주가능상태" in filtered.columns:
                filtered = filtered[filtered["발주가능상태"].str.contains(
                    order_filter, na=False, regex=False)]

            # 공급상태 필터
            if supply_filter != "전체" and "공급상태" in filtered.columns:
                filtered = filtered[filtered["공급상태"].str.contains(
                    supply_filter, na=False, regex=False)]

            # 등록계정 필터
            if account_filter != "전체" and "등록계정" in filtered.columns:
                filtered = filtered[filtered["등록계정"] == account_filter]

            # 상품분류 필터
            if category_filter != "전체" and "상품분류" in filtered.columns:
                filtered = filtered[filtered["상품분류"] == category_filter]

            self.root.after(0, self._add_to_history, input_str.strip())
            self.root.after(0, self._on_filter_complete, filtered)

        except Exception as e:
            self.root.after(0, self._on_filter_error, str(e))

    def _on_filter_complete(self, filtered: pd.DataFrame) -> None:
        sort_cols = [c for c in ["쿠팡상품코드", "컬러", "사이즈"] if c in filtered.columns]
        if sort_cols:
            filtered = filtered.sort_values(by=sort_cols).reset_index(drop=True)
        self.filtered_data = filtered
        self._update_sheet(filtered)
        self._refresh_status_bar(filtered_count=len(filtered))
        self._hide_loading()
        self._is_loading = False

    def _on_filter_error(self, error_msg: str) -> None:
        messagebox.showerror("에러", error_msg)
        self._hide_loading()
        self._is_loading = False

    # ========================================================================
    # 검색 히스토리
    # ========================================================================
    def _add_to_history(self, text: str) -> None:
        if text and text not in self.search_history:
            self.search_history.insert(0, text)
            if len(self.search_history) > CONFIG.MAX_HISTORY:
                self.search_history.pop()

    def _show_history(self) -> None:
        if not self.search_history:
            messagebox.showinfo("검색 히스토리", "검색 기록이 없습니다.")
            return

        popup = tk.Toplevel(self.root)
        popup.title("검색 히스토리")
        popup.geometry("420x320")
        popup.grab_set()
        popup.configure(bg=THEME.BG_SECONDARY)

        tk.Label(popup, text="최근 검색 목록  (더블클릭하여 적용)",
                 bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
                 font=FONTS.TITLE).pack(pady=10)

        frm = tk.Frame(popup, bg=THEME.BG_SECONDARY)
        frm.pack(fill="both", expand=True, padx=10, pady=5)

        sb = tk.Scrollbar(frm)
        sb.pack(side="right", fill="y")

        lb = tk.Listbox(frm, font=FONTS.NORMAL, yscrollcommand=sb.set,
                        bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
                        selectbackground=THEME.ACCENT_PRIMARY, activestyle="none")
        for item in self.search_history:
            lb.insert(tk.END, item.replace('\n', ' '))
        lb.pack(fill="both", expand=True)
        sb.config(command=lb.yview)

        def apply_selected():
            sel = lb.curselection()
            if sel:
                self.input_text.delete("1.0", tk.END)
                self.input_text.insert("1.0", self.search_history[sel[0]])
                popup.destroy()

        def delete_selected():
            sel = lb.curselection()
            if sel:
                self.search_history.pop(sel[0])
                lb.delete(sel[0])

        lb.bind("<Double-Button-1>", lambda e: apply_selected())

        btn_frm = tk.Frame(popup, bg=THEME.BG_SECONDARY)
        btn_frm.pack(pady=5)
        self._create_button(btn_frm, "적용",  apply_selected,  bg=THEME.ACCENT_PRIMARY)
        self._create_button(btn_frm, "삭제",  delete_selected, bg=THEME.ACCENT_DANGER)
        self._create_button(btn_frm, "닫기",  popup.destroy,   bg=THEME.ACCENT_DARK)

    # ========================================================================
    # 내보내기
    # ========================================================================
    def _export_data(self) -> None:
        if self.filtered_data.empty:
            messagebox.showinfo("알림", "내보낼 데이터가 없습니다.")
            return

        default_name = f"쿠팡상품조회_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx"), ("CSV 파일", "*.csv"), ("모든 파일", "*.*")],
            initialfile=default_name
        )
        if not filepath:
            return

        try:
            if filepath.lower().endswith(".csv"):
                self.filtered_data.to_csv(filepath, index=False, encoding="utf-8-sig")
            else:
                self.filtered_data.to_excel(filepath, index=False)
            messagebox.showinfo("완료", f"파일이 저장되었습니다.\n{filepath}")
        except Exception as e:
            messagebox.showerror("오류", f"저장 실패: {e}")

    # ========================================================================
    # 전체 보기 · 초기화
    # ========================================================================
    def show_all_data(self) -> None:
        if self.data.empty:
            messagebox.showinfo("알림", "데이터가 로드되지 않았습니다.")
            return
        if self._is_loading:
            return

        self.input_text.delete("1.0", tk.END)
        self.filter_order.current(0)
        self.filter_supply.current(0)
        self.filter_account.current(0)
        self.filter_category.current(0)
        self._is_loading = True
        self._show_loading("전체 데이터 불러오는 중...")
        threading.Thread(target=self._show_all_worker, daemon=True).start()

    def _show_all_worker(self) -> None:
        try:
            data = self.data.copy()
            self.root.after(0, self._on_show_all_complete, data)
        except Exception as e:
            self.root.after(0, lambda: (
                messagebox.showerror("에러", str(e)),
                self._hide_loading(),
                setattr(self, "_is_loading", False)
            ))

    def _on_show_all_complete(self, data: pd.DataFrame) -> None:
        self.filtered_data = data
        self._update_sheet(data)
        self._refresh_status_bar(filtered_count=len(data))
        self._hide_loading()
        self._is_loading = False

    def reset(self) -> None:
        self.input_text.delete("1.0", tk.END)
        self.filter_order.current(0)
        self.filter_supply.current(0)
        self.filter_account.current(0)
        self.filter_category.current(0)
        self.filtered_data = self.data.copy()
        self.sort_column = None
        self.sort_ascending = True
        for w in self.frm_sheet.winfo_children():
            w.destroy()
        self.lbl_count.config(text="")
        self.lbl_status_filtered.config(text="검색결과: -")

    # ========================================================================
    # 상태바
    # ========================================================================
    def _refresh_status_bar(self, filtered_count: Optional[int] = None) -> None:
        total = len(self.data) if not self.data.empty else 0
        self.lbl_status_total.config(text=f"전체: {total:,}건")

        if filtered_count is not None:
            self.lbl_status_filtered.config(text=f"검색결과: {filtered_count:,}건")
            self.lbl_count.config(text=f"검색결과  {filtered_count:,}건")
        else:
            self.lbl_status_filtered.config(text="검색결과: -")
            self.lbl_count.config(text="")

        if self.base_path:
            self.lbl_status_path.config(text=f"데이터 경로: {self.base_path}")

        self.lbl_status_time.config(
            text=f"갱신: {datetime.now().strftime('%H:%M:%S')}"
        )

    # ========================================================================
    # 정렬 및 시트 업데이트
    # ========================================================================
    def _sort_by_column(self, col_idx: int) -> None:
        if self.filtered_data.empty:
            return

        if self.sort_column == col_idx:
            self.sort_ascending = not self.sort_ascending
        else:
            self.sort_column = col_idx
            self.sort_ascending = True

        col_name = self.filtered_data.columns[col_idx]
        self.filtered_data = self.filtered_data.sort_values(
            by=col_name, ascending=self.sort_ascending
        ).reset_index(drop=True)

        self._update_sheet(self.filtered_data, preserve_sort_indicator=True)

    def _update_sheet(self, df: pd.DataFrame,
                      preserve_sort_indicator: bool = False) -> None:
        for w in self.frm_sheet.winfo_children():
            w.destroy()

        if df.empty:
            tk.Label(
                self.frm_sheet, text="조회할 데이터가 없습니다.",
                fg=THEME.ACCENT_DANGER, bg=THEME.BG_PRIMARY, font=FONTS.NORMAL
            ).pack(padx=10, pady=20)
            return

        headers = list(df.columns)
        if preserve_sort_indicator and self.sort_column is not None:
            indicator = " ▲" if self.sort_ascending else " ▼"
            headers[self.sort_column] += indicator

        sheet = Sheet(
            self.frm_sheet,
            data=df.values.tolist(),
            headers=headers,
            header_height=28,
            header_fg=THEME.HEADER_FG,
            header_bg=THEME.HEADER_BG
        )
        sheet.header_font((FONTS.FAMILY, 10, 'bold'))
        sheet.font((FONTS.FAMILY, 9, 'normal'))
        sheet.table_align(align="left")
        sheet.set_all_column_widths(width=None)

        for idx, col in enumerate(df.columns):
            if col in COLUMN_WIDTHS:
                sheet.column_width(column=idx, width=COLUMN_WIDTHS[col])

        sheet.enable_bindings()
        sheet.bind("<ButtonPress-1>",   self._on_sheet_click,        add=True)
        sheet.bind("<Double-Button-1>", self._on_sheet_double_click, add=True)

        sheet.pack(fill="both", expand=True)
        self.current_sheet = sheet

    def _on_sheet_click(self, event) -> None:
        if not self.current_sheet:
            return
        region = self.current_sheet.identify_region(event)
        if region and region.type_ == "header" and region.column is not None:
            self._sort_by_column(region.column)

    def _on_sheet_double_click(self, event) -> None:
        """행 더블클릭 → 상세 정보 팝업"""
        if not self.current_sheet:
            return
        region = self.current_sheet.identify_region(event)
        if region and region.type_ == "cell" and region.row is not None:
            if region.row < len(self.filtered_data):
                self._show_row_detail(region.row)

    def _show_row_detail(self, row_idx: int) -> None:
        row = self.filtered_data.iloc[row_idx]
        title = str(row.get("상품명", ""))[:30]

        popup = tk.Toplevel(self.root)
        popup.title(f"상세 정보 — {title}")
        popup.geometry("500x620")
        popup.grab_set()
        popup.configure(bg=THEME.BG_SECONDARY)

        tk.Label(popup, text="상품 상세 정보",
                 bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
                 font=FONTS.TITLE).pack(pady=10)

        frm = tk.Frame(popup, bg=THEME.BG_SECONDARY)
        frm.pack(fill="both", expand=True, padx=15, pady=5)

        sb = tk.Scrollbar(frm)
        sb.pack(side="right", fill="y")

        txt = tk.Text(frm, font=FONTS.NORMAL, yscrollcommand=sb.set,
                      bg=THEME.BG_SECONDARY, fg=THEME.FG_PRIMARY,
                      relief="flat", padx=10, pady=5)
        for col, val in row.items():
            txt.insert(tk.END, f"{col}\n", "key")
            txt.insert(tk.END, f"  {val}\n\n", "val")
        txt.tag_configure("key", font=FONTS.TITLE, foreground=THEME.ACCENT_PRIMARY)
        txt.tag_configure("val", font=FONTS.NORMAL)
        txt.config(state="disabled")
        txt.pack(fill="both", expand=True)
        sb.config(command=txt.yview)

        def copy_all():
            content = "\n".join(f"{c}: {v}" for c, v in row.items())
            popup.clipboard_clear()
            popup.clipboard_append(content)
            messagebox.showinfo("완료", "클립보드에 복사되었습니다.", parent=popup)

        btn_frm = tk.Frame(popup, bg=THEME.BG_SECONDARY)
        btn_frm.pack(pady=6)
        self._create_button(btn_frm, "📋 전체 복사", copy_all,       bg=THEME.ACCENT_PRIMARY)
        self._create_button(btn_frm, "닫기",         popup.destroy,  bg=THEME.ACCENT_DARK)


# ============================================================================
# 메인 실행
# ============================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = CoupangSKUApp(root)
    root.mainloop()
