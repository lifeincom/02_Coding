"""
주문서 상품코드 변환 및 출고리스트 생성 프로그램 v8.11

변경사항 (v8.03 → v8.11):
- 코드 구조 최적화 및 리팩토링 (중복 제거, 상수 분리)
- set_product_separation() 로직 검증 및 안정화
- UI/UX 리디자인: 워크플로우 중심 단순화, 섹션 명확화
- 각 출력 시트별 컬럼 너비 조정 기능 추가 (column_width_dialog)
- 진행률 표시 통합 및 상태바 개선
- 에러 처리 강화
"""

# ===== 라이브러리 임포트 =====
import os
import sys
import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import customtkinter as ctk
from tksheet import Sheet
import warnings

warnings.filterwarnings("ignore")

# ===== CustomTkinter 전역 설정 =====
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# ===== UI 컬러 테마 =====
COLORS = {
    "primary":          "#1E40AF",   # 딥 블루
    "primary_hover":    "#1E3A8A",
    "primary_light":    "#DBEAFE",
    "secondary":        "#059669",   # 에메랄드
    "secondary_hover":  "#047857",
    "secondary_light":  "#D1FAE5",
    "accent":           "#7C3AED",   # 보라 (세트 분리 강조용)
    "accent_hover":     "#6D28D9",
    "danger":           "#DC2626",
    "danger_hover":     "#B91C1C",
    "warning":          "#D97706",
    "warning_hover":    "#B45309",
    "neutral":          "#6B7280",
    "neutral_hover":    "#4B5563",
    "frame_bg":         "#F3F4F6",
    "card_bg":          "#FFFFFF",
    "entry_bg":         "#FFFFFF",
    "border":           "#D1D5DB",
    "text_primary":     "#111827",
    "text_secondary":   "#6B7280",
    "sheet_header_bg":  "#1F2937",
    "sheet_header_fg":  "#FFFFFF",
    "step_active":      "#1E40AF",
    "step_done":        "#059669",
    "step_inactive":    "#9CA3AF",
}

# ===== 폰트 설정 =====
FONT = "nanumgothic"
FS = {"title": 14, "heading": 12, "normal": 12, "small": 11}

# ===== 시트별 기본 컬럼 너비 설정 =====
# 각 시트의 컬럼 인덱스별 기본 픽셀 너비
DEFAULT_COL_WIDTHS = {
    "l1": {0: 80, 1: 350, 2: 200, 3: 50, 4: 100, 5: 120, 6: 70, 7: 60, 8: 200},
    "l2": {0: 100, 1: 120, 2: 70, 3: 60, 4: 60, 5: 100},
    "r1": {0: 220, 1: 60},
    "r2": {0: 80, 1: 220, 2: 50, 3: 60},
}

# ===== 수집 프로그램 옵션 =====
PROGRAM_OPTIONS = [" + 수집 프로그램 선택", " 1. 사방넷", " 2. 이지어드민"]

# ===== 사방넷 컬럼 순서 =====
SABANG_COLS = [
    "성명", "전화번호", "핸드폰", "우편번호", "주소", "상품명",
    "수량", "배송메세지", "주문처", "요금구분", "운송장번호",
    "사방넷주문번호", "쇼핑몰아이디",
    "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2",
]

# ===== 이지어드민 원본→변환 컬럼 =====
EASY_SRC_COLS = [
    "수령자", "전화번호", "핸드폰", "우편번호", "주소", "옵션",
    "주문수량", "특이사항", "판매처", "배송비", "상품번호",
    "주문번호", "주문자명",
    "상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2",
]
EASY_RENAME = {
    "수령자": "성명", "옵션": "상품명", "주문수량": "수량",
    "특이사항": "배송메세지", "판매처": "주문처",
}

# ===== 티셔츠 접두사 =====
TSHIRT_PREFIXES = ("T", "S")

# ===== 옵션코드 컬러 포함 접두사 =====
COLOR_CODE_PREFIXES = (
    "T", "U", "BHP-7", "BHP-5", "WJL-0",
    "EA", "DP", "EP", "CP", "DL0", "EL0", "RL0", "WBL", "SPH",
)


# ============================================================
#  헬퍼 함수
# ============================================================

def make_btn(parent, text, cmd, color_key="primary", width=140, height=34):
    """공통 CTkButton 생성 헬퍼"""
    return ctk.CTkButton(
        parent, text=text, command=cmd,
        width=width, height=height,
        font=(FONT, FS["normal"], "bold"),
        fg_color=COLORS[color_key],
        hover_color=COLORS[f"{color_key}_hover"],
        corner_radius=8,
    )


def apply_col_widths(sheet: Sheet, widths: dict):
    """tksheet 컬럼 너비 일괄 적용"""
    for col_idx, px in widths.items():
        try:
            sheet.column_width(column=col_idx, width=px)
        except Exception:
            pass


# ============================================================
#  메인 애플리케이션
# ============================================================

class OrderProcessingApp(ctk.CTk):
    """
    주문서 상품코드 변환 & 출고리스트 생성 v8.11
    워크플로우:  ①주문서 선택  →  ②상품코드 추출  →  ③출고리스트 생성  →  ④저장/복사
    """

    def __init__(self):
        super().__init__()
        self.title("주문서 상품코드 변환 & 출고리스트 생성  v8.11")
        self.geometry("1600x900+10+10")
        self.minsize(1400, 800)
        try:
            self.state("zoomed")
        except Exception:
            pass

        # ── 컬럼 너비 설정 (사용자 조정 가능, 기본값으로 초기화) ──
        self.col_widths = {k: dict(v) for k, v in DEFAULT_COL_WIDTHS.items()}

        self._init_variables()
        self._build_ui()

    # ----------------------------------------------------------
    #  초기화
    # ----------------------------------------------------------

    def _init_variables(self):
        """데이터 및 경로 변수 초기화"""
        if getattr(sys, "frozen", False):
            self.program_dir = os.path.dirname(os.path.abspath(sys.executable))
        else:
            self.program_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.program_dir)

        # DataFrame 플레이스홀더
        self.empty_df   = pd.DataFrame([])
        self.sdf        = None   # 원본 주문서
        self.sdf2       = None   # 세트분리 후
        self.sdf2_view  = None   # 화면 표시용
        self.odf        = None   # 전체 출고리스트
        self.ndf        = None   # 티셔츠 출고리스트
        self.df_sale    = None   # 판매 데이터
        self.df_upload  = None   # 업로드용

        self.gDate = ""
        self.sum_order = "0"
        self.sum_ship  = "0"

        # 참조 데이터
        self.code_sr  = pd.Series(dtype=str)
        self.color_sr = pd.Series(dtype=str)
        self.size_sr  = pd.Series(dtype=str)

        # 경로 설정
        now = datetime.datetime.now()
        year = now.strftime("%Y")
        candidates = [
            r"\\NAS451\team451",
            r"\\192.168.0.101\team451",
            r"n:\개인\nSync\Coding",
            r"d:\hSync\Coding",
        ]
        self.base_path = next((p for p in candidates if os.path.isdir(p)), None)

        if self.base_path:
            self.sfile_dir  = os.path.join(self.base_path, "04-주문택배업로드", "택배업로드 및 발주서", year)
            self.stock_file = os.path.join(self.base_path, "DB", "반품티셔츠재고장.xlsx")
            self.opt_file   = os.path.join(self.base_path, "DB", "변환코드.xlsx")
            self._load_reference_data()
        else:
            self.sfile_dir  = "/"
            self.stock_file = ""
            self.opt_file   = ""

    def _load_reference_data(self):
        """변환코드.xlsx 참조 데이터 로드"""
        try:
            self.code_sr  = pd.read_excel(self.opt_file, sheet_name="상품코드").iloc[:, 0]
            self.color_sr = pd.read_excel(self.opt_file, sheet_name="컬러").iloc[:, 0]
            self.size_sr  = pd.read_excel(self.opt_file, sheet_name="사이즈").iloc[:, 0]
        except Exception as e:
            messagebox.showerror("에러", f"변환코드 파일 로드 실패:\n{e}")

    # ----------------------------------------------------------
    #  UI 빌드
    # ----------------------------------------------------------

    def _build_ui(self):
        """전체 UI 레이아웃 구성"""
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)  # 메인 컨텐츠 영역 확장

        self._build_step_bar()      # row 0: 워크플로우 단계 표시
        self._build_top_toolbar()   # row 1: 파일 선택 + 실행 버튼
        self._build_main_content()  # row 2: 4-패널 데이터 뷰
        self._build_status_bar()    # row 3: 진행률 + 수량 요약
        self._build_action_bar()    # row 4: 저장/복사 액션

    # ── Step Bar ──────────────────────────────────────────────
    def _build_step_bar(self):
        frm = ctk.CTkFrame(self, fg_color=COLORS["primary"], corner_radius=0, height=40)
        frm.grid(row=0, column=0, sticky="ew")
        frm.grid_propagate(False)
        frm.grid_columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)

        steps = [
            ("①", "수집 프로그램 선택"),
            ("→", ""),
            ("②", "주문서 선택"),
            ("→", ""),
            ("③", "상품코드 추출"),
            ("→", ""),
            ("④", "출고리스트 생성"),
        ]
        for i, (num, label) in enumerate(steps):
            color = COLORS["card_bg"] if num not in ("→",) else "#93C5FD"
            ctk.CTkLabel(
                frm,
                text=f" {num} {label} " if label else f"  {num}  ",
                font=(FONT, FS["small"], "bold"),
                text_color=color,
            ).grid(row=0, column=i, padx=4, pady=8)

    # ── 상단 툴바 ─────────────────────────────────────────────
    def _build_top_toolbar(self):
        frm = ctk.CTkFrame(
            self, fg_color=COLORS["card_bg"],
            corner_radius=10, border_width=1, border_color=COLORS["border"]
        )
        frm.grid(row=1, column=0, sticky="ew", padx=12, pady=(10, 4))

        # 수집 프로그램 콤보
        self.combo = ctk.CTkComboBox(
            frm, values=PROGRAM_OPTIONS, width=190, height=34,
            state="readonly",
            font=(FONT, FS["normal"]),
            dropdown_font=(FONT, FS["normal"]),
            fg_color=COLORS["entry_bg"], border_color=COLORS["border"],
            button_color=COLORS["primary"], button_hover_color=COLORS["primary_hover"],
        )
        self.combo.set(PROGRAM_OPTIONS[0])
        self.combo.pack(side="left", padx=(12, 8), pady=10)

        # 파일 경로 입력
        self.ent_file = ctk.CTkEntry(
            frm, height=34,
            font=(FONT, FS["normal"]),
            fg_color=COLORS["entry_bg"], text_color=COLORS["text_primary"],
            border_color=COLORS["border"], border_width=1,
            placeholder_text="주문서 파일을 선택해주세요...",
        )
        self.ent_file.pack(side="left", fill="x", expand=True, padx=6, pady=10)

        # 우측: 닫기, 초기화
        make_btn(frm, "✕  닫기",   self.quit,  "danger",  85, 34).pack(side="right", padx=(8, 12), pady=10)
        make_btn(frm, "↺  초기화", self.reset, "warning", 95, 34).pack(side="right", padx=4, pady=10)

        ctk.CTkFrame(frm, width=1, height=30, fg_color=COLORS["border"]).pack(side="right", padx=10, pady=10)

        # 주요 워크플로우 버튼 (우→좌 순서)
        make_btn(frm, "③  출고리스트 생성",  self.create_factory_list,            "secondary", 165, 34).pack(side="right", padx=4, pady=10)
        make_btn(frm, "②  상품코드 추출",    self.output_product_code_extraction, "primary",   155, 34).pack(side="right", padx=4, pady=10)
        make_btn(frm, "①  주문서 선택",      self.import_order,                   "accent",    145, 34).pack(side="right", padx=4, pady=10)

    # ── 메인 컨텐츠 (4패널) ────────────────────────────────────
    def _build_main_content(self):
        frm = ctk.CTkFrame(self, fg_color="transparent")
        frm.grid(row=2, column=0, sticky="nsew", padx=12, pady=4)
        frm.grid_columnconfigure(0, weight=9)
        frm.grid_columnconfigure(1, weight=1)
        frm.grid_rowconfigure(0, weight=1)

        # ─ 왼쪽 탭뷰 ─
        self.tab_left = ctk.CTkTabview(
            frm, corner_radius=10,
            segmented_button_fg_color=COLORS["frame_bg"],
            segmented_button_selected_color=COLORS["primary"],
            segmented_button_selected_hover_color=COLORS["primary_hover"],
            border_width=1, border_color=COLORS["border"],
        )
        self.tab_left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.tab_l1 = self.tab_left.add("📦  상품코드 추출리스트")
        self.tab_l2 = self.tab_left.add("📊  판매데이터")

        # ─ 오른쪽 탭뷰 ─
        self.tab_right = ctk.CTkTabview(
            frm, corner_radius=10,
            segmented_button_fg_color=COLORS["frame_bg"],
            segmented_button_selected_color=COLORS["secondary"],
            segmented_button_selected_hover_color=COLORS["secondary_hover"],
            border_width=1, border_color=COLORS["border"],
        )
        self.tab_right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        self.tab_r1 = self.tab_right.add("📋  전체상품 출고리스트")
        self.tab_r2 = self.tab_right.add("👕  티셔츠 출고리스트")

        # 초기 빈 시트
        self.sheet_l1 = self._make_sheet(self.tab_l1, self.empty_df, "l1")
        self.sheet_l2 = self._make_sheet(self.tab_l2, self.empty_df, "l2")
        self.sheet_r1 = self._make_sheet(self.tab_r1, self.empty_df, "r1")
        self.sheet_r2 = self._make_sheet(self.tab_r2, self.empty_df, "r2")

    # ── 상태바 ────────────────────────────────────────────────
    def _build_status_bar(self):
        frm = ctk.CTkFrame(
            self, fg_color=COLORS["card_bg"],
            corner_radius=10, border_width=1, border_color=COLORS["border"]
        )
        frm.grid(row=3, column=0, sticky="ew", padx=12, pady=4)

        ctk.CTkLabel(
            frm, text="진행률:", font=(FONT, FS["normal"], "bold"),
            text_color=COLORS["text_secondary"]
        ).pack(side="left", padx=(14, 6), pady=10)

        self.p_var = tk.DoubleVar()
        self.prog_bar = ctk.CTkProgressBar(
            frm, variable=self.p_var,
            progress_color=COLORS["primary"], height=18, corner_radius=9
        )
        self.prog_bar.pack(side="left", fill="x", expand=True, padx=8, pady=10)
        self.prog_bar.set(0)

        self.lbl_row = ctk.CTkLabel(
            frm, text="ROW : 0", width=90,
            font=(FONT, FS["normal"], "bold"), text_color=COLORS["primary"]
        )
        self.lbl_row.pack(side="left", padx=6, pady=10)

        # 수량 요약
        self.lbl_summary = ctk.CTkLabel(
            frm, text="주문수량 : -    /    출고수량 : -",
            font=(FONT, FS["normal"], "bold"),
            text_color=COLORS["secondary"],
        )
        self.lbl_summary.pack(side="right", padx=16, pady=10)

    # ── 액션바 ────────────────────────────────────────────────
    def _build_action_bar(self):
        frm = ctk.CTkFrame(
            self, fg_color=COLORS["card_bg"],
            corner_radius=10, border_width=1, border_color=COLORS["border"]
        )
        frm.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 12))

        # 왼쪽: 저장 섹션
        ctk.CTkLabel(frm, text="💾 저장 경로:", font=(FONT, FS["normal"], "bold"),
                     text_color=COLORS["text_primary"]).pack(side="left", padx=(14, 6), pady=10)

        self.ent_savepath = ctk.CTkEntry(
            frm, height=32, font=(FONT, FS["normal"]),
            fg_color=COLORS["entry_bg"], text_color=COLORS["text_primary"],
            border_color=COLORS["border"], border_width=1,
            placeholder_text="저장할 폴더를 선택하세요...",
        )
        self.ent_savepath.pack(side="left", fill="x", expand=True, padx=6, pady=10)

        make_btn(frm, "📁 폴더 선택", self.browse_save_path, "neutral",  110, 32).pack(side="left", padx=4, pady=10)
        make_btn(frm, "💾 저장하기",  self.save_file,        "secondary", 110, 32).pack(side="left", padx=4, pady=10)

        ctk.CTkFrame(frm, width=1, height=28, fg_color=COLORS["border"]).pack(side="left", padx=14, pady=10)

        # 오른쪽: 복사 + 컬럼너비 + 도움말
        make_btn(frm, "❓ 도움말",    self.manual,               "neutral",  100, 32).pack(side="right", padx=(8, 14), pady=10)
        make_btn(frm, "↔ 컬럼 너비", self.open_col_width_dialog, "neutral",  110, 32).pack(side="right", padx=4, pady=10)

        ctk.CTkFrame(frm, width=1, height=28, fg_color=COLORS["border"]).pack(side="right", padx=10, pady=10)

        for label, cmd in [
            ("판매데이터 복사",    self.copy_sale_data),
            ("티셔츠출고 복사",    self.copy_output_list2),
            ("전체출고 복사",      self.copy_output_list1),
            ("상품코드 복사",      self.copy_code_ext),
        ]:
            make_btn(frm, label, cmd, "primary", 115, 32).pack(side="right", padx=4, pady=10)

    # ----------------------------------------------------------
    #  시트 생성 / 갱신
    # ----------------------------------------------------------

    def _make_sheet(self, parent, data: pd.DataFrame, sheet_key: str) -> Sheet:
        """tksheet 생성 및 컬럼 너비 적용"""
        for w in parent.winfo_children():
            w.destroy()

        wrap = tk.Frame(parent, bg=COLORS["card_bg"])
        wrap.pack(fill="both", expand=True, padx=2, pady=2)

        sheet = Sheet(
            wrap,
            data=data.values.tolist(),
            headers=list(data.columns),
            header_bg=COLORS["sheet_header_bg"],
            header_fg=COLORS["sheet_header_fg"],
            header_font=(FONT, FS["normal"], "bold"),
            table_bg=COLORS["card_bg"],
            table_fg=COLORS["text_primary"],
            index_bg=COLORS["frame_bg"],
            top_left_bg=COLORS["frame_bg"],
            table_selected_cells_bg=COLORS["primary_light"],
            table_selected_cells_fg=COLORS["text_primary"],
            show_vertical_grid=True,
            show_horizontal_grid=True,
            vertical_grid_to_end_of_window=True,
            horizontal_grid_to_end_of_window=True,
            grid_color=COLORS["border"],
        )
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        # 컬럼 너비 적용
        apply_col_widths(sheet, self.col_widths.get(sheet_key, {}))

        return sheet

    def _refresh_sheet(self, sheet_attr: str, tab, data: pd.DataFrame, sheet_key: str):
        """시트 재생성 후 인스턴스 변수 갱신"""
        new_sheet = self._make_sheet(tab, data, sheet_key)
        setattr(self, sheet_attr, new_sheet)
        return new_sheet

    # ----------------------------------------------------------
    #  컬럼 너비 조정 다이얼로그
    # ----------------------------------------------------------

    def open_col_width_dialog(self):
        """
        컬럼 너비 조정 다이얼로그
        각 시트별 컬럼 너비를 사용자가 직접 입력하여 변경할 수 있습니다.
        """
        dialog = ctk.CTkToplevel(self)
        dialog.title("컬럼 너비 조정")
        dialog.geometry("560x520")
        dialog.grab_set()
        dialog.resizable(False, False)

        sheet_info = [
            ("l1", "📦 상품코드 추출리스트",
             ["성명", "상품명", "옵션코드", "수량", "주문처", "상품코드2", "컬러", "사이즈", "옵션코드2"]),
            ("l2", "📊 판매데이터",
             ["날짜", "상품코드2", "컬러", "사이즈", "수량", "주문처"]),
            ("r1", "📋 전체상품 출고리스트",
             ["옵션코드", "수량"]),
            ("r2", "👕 티셔츠 출고리스트",
             ["주문자명", "옵션코드", "수량", "재고"]),
        ]

        # 탭뷰로 시트 구분
        tab_view = ctk.CTkTabview(dialog, corner_radius=8)
        tab_view.pack(fill="both", expand=True, padx=14, pady=(14, 6))

        entry_map = {}  # sheet_key -> {col_idx -> entry widget}

        for s_key, s_label, col_names in sheet_info:
            tab = tab_view.add(s_label)
            entry_map[s_key] = {}

            scroll = ctk.CTkScrollableFrame(tab, fg_color="transparent")
            scroll.pack(fill="both", expand=True, padx=4, pady=4)

            ctk.CTkLabel(scroll, text="컬럼명", font=(FONT, FS["small"], "bold"),
                         text_color=COLORS["text_secondary"]).grid(row=0, column=0, sticky="w", padx=8, pady=4)
            ctk.CTkLabel(scroll, text="너비(px)", font=(FONT, FS["small"], "bold"),
                         text_color=COLORS["text_secondary"]).grid(row=0, column=1, padx=8, pady=4)

            for i, col_name in enumerate(col_names):
                ctk.CTkLabel(scroll, text=col_name, font=(FONT, FS["normal"]),
                             anchor="w").grid(row=i+1, column=0, sticky="w", padx=12, pady=3)
                ent = ctk.CTkEntry(scroll, width=90, height=30, font=(FONT, FS["normal"]))
                current_w = self.col_widths.get(s_key, {}).get(i, 100)
                ent.insert(0, str(current_w))
                ent.grid(row=i+1, column=1, padx=12, pady=3)
                entry_map[s_key][i] = ent

        def apply_widths():
            for s_key, cols in entry_map.items():
                for col_idx, ent in cols.items():
                    try:
                        val = int(ent.get())
                        if 20 <= val <= 600:
                            self.col_widths[s_key][col_idx] = val
                    except ValueError:
                        pass

            # 현재 표시된 시트에 즉시 반영
            for attr, tab, s_key in [
                ("sheet_l1", self.tab_l1, "l1"),
                ("sheet_l2", self.tab_l2, "l2"),
                ("sheet_r1", self.tab_r1, "r1"),
                ("sheet_r2", self.tab_r2, "r2"),
            ]:
                sheet = getattr(self, attr, None)
                if sheet is not None:
                    apply_col_widths(sheet, self.col_widths.get(s_key, {}))

            messagebox.showinfo("완료", "컬럼 너비가 적용되었습니다.", parent=dialog)
            dialog.destroy()

        def reset_widths():
            self.col_widths = {k: dict(v) for k, v in DEFAULT_COL_WIDTHS.items()}
            dialog.destroy()
            self.open_col_width_dialog()

        btn_frm = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frm.pack(fill="x", padx=14, pady=(4, 14))

        make_btn(btn_frm, "기본값 복원", reset_widths, "neutral",  120, 34).pack(side="left", padx=4)
        make_btn(btn_frm, "취소",        dialog.destroy, "danger",  80, 34).pack(side="right", padx=4)
        make_btn(btn_frm, "✔ 적용",      apply_widths,  "secondary", 110, 34).pack(side="right", padx=4)

    # ----------------------------------------------------------
    #  핵심 로직
    # ----------------------------------------------------------

    def extract_itemcode(self, inCode: str) -> str:
        """
        원본 상품코드를 내부 표준 코드로 변환합니다.
        대소문자 무관하게 처리합니다.
        """
        inCode = str(inCode).upper()

        prefix_map = {
            "S": "JS", "U": "JU", "A": "BA", "C": "BC",
            "L": "JL", "M": "JM", "N": "NA", "J": "EJ", "F": "JF",
        }

        if inCode.startswith(("W", "M", "T", "U", "BHP", "GS", "DS", "CS", "PAC", "SPH")):
            if len(inCode) > 3 and inCode[3] == "A":
                return f"{inCode[:3]}-{inCode[4:]}"
            if len(inCode) > 3 and inCode[3] == "F":
                return f"{inCode[0]}JF-{inCode[4:]}"
            if inCode.startswith("TS"):
                return f"TPS-{inCode[3:]}"
            if inCode.startswith("TPNS"):
                return f"{inCode[:3]}-{inCode[4:]}"
            if inCode.startswith(("TGJS", "TLJS")):
                return f"{inCode[:2]}-{inCode[4:]}"
            if inCode.startswith(("UAA-", "UAB-", "UAC-", "UAD-", "UAE-", "UAF-", "UAR-")):
                return f"{inCode[:3]}O{inCode[4:]}"
            return inCode

        if inCode.startswith(("C", "D", "E", "H", "R", "Z")):
            if len(inCode) > 3 and inCode[1:3] == "P7":
                return f"BHP-{inCode[2]}{inCode[4]}{inCode[5]}"
            if len(inCode) > 3 and inCode[3] == "T" and inCode[:2] == "RS":
                return f"TPS-{inCode[2]}{inCode[4]}{inCode[5]}"
            if inCode[1:] in ("PN7T13",):
                return f"BHP-{inCode[3]}{inCode[5]}{inCode[6]}"
            if inCode[1:3] in ("PP", "PK"):
                return f"BHP-{inCode[3]}{inCode[5]}{inCode[6]}"
            if inCode[1:3] in ("PL", "PG", "PM", "PH"):
                return f"{inCode[4]}{inCode[1]}{inCode[2]}-{inCode[3]}{inCode[5]}{inCode[6]}"
            if len(inCode) > 3 and inCode[3] in ("T", "U"):
                return f"{inCode[3]}{inCode[1]}-{inCode[2]}{inCode[4]}{inCode[5]}"
            if len(inCode) > 3 and inCode[3] in ("W", "M"):
                return f"{inCode[3]}{prefix_map.get(inCode[1], 'XX')}-{inCode[2]}{inCode[4]}{inCode[5]}"
            if len(inCode) > 4 and inCode[4] == "T":
                return f"{inCode[4]}{inCode[1]}{inCode[2]}-{inCode[3]}{inCode[5]}{inCode[6]}"
            if len(inCode) > 4 and inCode[4] == "U":
                return f"{inCode[4]}{inCode[1]}{inCode[2]}O{inCode[3]}{inCode[5]}{inCode[6]}"
            if len(inCode) > 4 and inCode[4] in ("W", "M"):
                return f"{inCode[4]}BP-{inCode[3]}{inCode[5]}{inCode[6]}"

        return ""

    def _build_option_code(self, code: str, color: str, size: str) -> str:
        """옵션코드 조합 (컬러 포함 여부 자동 판단)"""
        include_color = (
            bool(color)
            and (
                str(code).startswith(COLOR_CODE_PREFIXES)
                or (len(str(code)) > 3 and str(code)[3] == "T")
                or (len(str(code)) > 4 and str(code)[4] == "T")
            )
        )
        if include_color:
            return f"{code}({color}) : {size}"
        return f"{code} : {size}"

    def _normalize_product_name(self, name: str, site: str) -> str:
        """
        쇼핑몰별 상품명을 정규화합니다.
        원본 로직을 보존하여 각 쇼핑몰 특수 처리를 유지합니다.
        """
        x = str(name)

        if site == "하프클럽(신)":
            x = x.replace("_", "-").replace(" ", "").replace("/", " : ")

        elif site == "(주)진마니아":
            for t in ["모델명/색상:", "모델명:사이즈:", ",사이즈", "사이즈:", "MODEL:SIZE:", " "]:
                x = x.replace(t, "")
            x = x.replace(",", ":").replace(":", " : ")

        elif site in ("롯데닷컷", "롯데홈쇼핑(신)"):
            for t in ["모델명/색상:", "모델명:", ",사이즈", "MODEL:", ",SIZE"]:
                x = x.replace(t, "")

        elif site == "현대홈쇼핑(신)":
            x = x.replace(":", " : ").replace("/", " : ")

        elif site == "패션플러스":
            x = x.replace(" (", "(").replace(" ", " : ")

        elif site == "GS shop":
            x = x.replace(",", ":")

        elif site in ("ESM지마켓", "ESM옥션"):
            for t in [f"{n}000원" for n in range(1, 10)]:
                x = x.replace(t, "")
            x = x.replace(" ", "")
            s_idx = x.find("_")
            end_ch = "/" if site == "ESM지마켓" else "["
            e_idx = x.find(end_ch)
            if s_idx != -1 and e_idx != -1:
                x = x[s_idx + 1:e_idx]

        elif site == "11번가":
            for t in ["색상:", "사이즈"]:
                x = x.replace(t, "")
            s_idx = x.find("_")
            e_idx = x.find("+")
            if s_idx != -1 and e_idx != -1:
                x = x[s_idx + 1:e_idx - 1]

        elif site == "쿠팡":
            if not x.startswith("MBP"):
                x = "+" + x

        elif site == "스마트스토어":
            for t in ["모델명/색상:", " / 사이즈", " "]:
                x = x.replace(t, "")
            x = x.replace(":", " : ")

        elif site == "티몬":
            x = x.replace("|", "")

        elif site == "T deal":
            x = x.replace("모델명(색상)|사이즈:", "").replace("|", ":")

        return x

    # ----------------------------------------------------------
    #  STEP ①: 주문서 선택
    # ----------------------------------------------------------

    def import_order(self):
        """주문서 파일 선택 및 날짜 추출"""
        if self.combo.get() == PROGRAM_OPTIONS[0]:
            messagebox.showwarning("경고", "수집 프로그램을 먼저 선택해 주세요.")
            return

        path = filedialog.askopenfilename(
            title="주문서 파일 선택",
            filetypes=[("Excel 파일", "*.xlsx *.xls"), ("모든 파일", "*.*")],
            initialdir=self.sfile_dir,
        )
        if not path:
            return

        self.ent_file.delete(0, tk.END)
        self.ent_file.insert(tk.END, path)

        # 파일명에서 날짜 추출 (YYYYMMDD_xxx.xlsx 형식)
        us = path.find("_")
        if us != -1:
            raw = path[us - 8:us]
            self.gDate = f"{raw[:4]}-{raw[4:6]}-{raw[6:8]}"
        else:
            self.gDate = datetime.datetime.now().strftime("%Y-%m-%d")

    # ----------------------------------------------------------
    #  STEP ②: 상품코드 추출
    # ----------------------------------------------------------

    def order_product_code_extraction(self):
        """주문서 Excel에서 상품코드/컬러/사이즈 추출"""
        sfile = self.ent_file.get()
        if not os.path.isfile(sfile):
            messagebox.showwarning("오류", "파일이 존재하지 않습니다.")
            return False

        self.sdf = pd.read_excel(sfile)

        # 추가 컬럼 초기화
        for col in ["상품코드", "컬러", "사이즈", "옵션코드", "상품코드2", "옵션코드2"]:
            self.sdf[col] = ""

        # 프로그램별 컬럼 재배치
        if self.combo.get() == " 2. 이지어드민":
            self.sdf = self.sdf.reindex(columns=EASY_SRC_COLS)
            self.sdf = self.sdf.rename(columns=EASY_RENAME)
            self.sdf["수량"] = self.sdf["수량"].astype(int)
        else:
            self.sdf = self.sdf.reindex(columns=SABANG_COLS)

        self.sum_order = str(self.sdf["수량"].sum())

        rows = self.sdf.shape[0]
        for r in range(rows):
            site  = self.sdf.iloc[r, 8]
            raw_x = self.sdf.iloc[r, 5]
            x = self._normalize_product_name(raw_x, site)

            # 상품코드 추출
            gCode = ""
            for code in self.code_sr:
                pos = x.find(code)
                if pos < 0:
                    continue
                c = str(code)
                if len(c) > 7:
                    length = 13 if c[0] == "R" else (15 if c[0] in ("W", "B") else 22)
                elif c[:4] in ("TPNS", "TPWS", "TPXS"):
                    length = 7
                elif c[:3] in ("TAS", "TAL", "TAG", "TAO", "TAN", "TLJ", "TGJ",
                               "UAA", "UAB", "UAC", "UAD", "UAF", "SPH"):
                    length = 7
                elif c[:2] in ("TP", "DP", "EP", "CP", "EA"):
                    length = 7
                elif c[0] in ("T", "D", "E", "R"):
                    length = 6
                else:
                    length = 7
                gCode = x[pos:pos + length].upper()

            # 컬러 추출
            gColor = ""
            for color in self.color_sr:
                if x.find(color) > 0:
                    gColor = color

            # 사이즈 추출
            gSize = ""
            for size in self.size_sr:
                if x.find(str(size)) >= 0:
                    gSize = str(size).strip(' :"/(+')

            self.sdf.iloc[r, 13] = gCode
            self.sdf.iloc[r, 14] = gColor
            self.sdf.iloc[r, 15] = gSize

        return True

    def set_product_separation(self):
        """
        세트 상품('+' 기호 포함) 분리 처리.

        처리 규칙:
        1. 컬러에 '+' 있음 → 컬러별로 행 복제, 상품코드는 앞 7자리
        2. 상품코드에 '+' 있음:
           a) BHP로 시작 → 앞/뒤 코드로 2행 분리
           b) TB-001로 시작 → 코드+컬러 동시 분리
           c) 기타 → set1/set2 로 2행 분리
        3. '+' 없음 → 그대로 복사
        """
        if self.sdf is None:
            return False

        result_rows = []
        col_count = self.sdf.shape[1]   # 19

        for r in range(self.sdf.shape[0]):
            row_data = [self.sdf.iloc[r, a] for a in range(col_count)]
            code_val  = str(self.sdf.iloc[r, 13])
            color_val = str(self.sdf.iloc[r, 14])

            cut_code  = code_val.find("+")
            has_color_plus = color_val.find("+") > 0
            has_code_plus  = cut_code > 0

            if has_color_plus:
                # ── 컬러 세트 분리 ──
                colors = color_val.split("+")
                for c in colors:
                    new_row = list(row_data)
                    new_row[13] = code_val[:7]
                    new_row[14] = c
                    result_rows.append(new_row)

            elif has_code_plus:
                set1 = code_val[:cut_code]
                set2 = code_val[cut_code + 1:]

                if code_val[:4] == "BHP-" or code_val[:3] == "BHP":
                    # ── BHP 세트 분리 ──
                    row_a = list(row_data); row_a[13] = set1
                    row_b = list(row_data); row_b[13] = set2[:7] if len(set2) >= 7 else set2
                    result_rows.extend([row_a, row_b])

                elif code_val[:6] == "TB-001":
                    # ── TB-001 세트 분리 (코드+컬러 동시) ──
                    def parse_tb(seg):
                        c_end = seg.find(")")
                        return seg[:6], seg[7:c_end] if c_end > 7 else ""

                    c1, cl1 = parse_tb(set1)
                    c2, cl2 = parse_tb(set2)
                    row_a = list(row_data); row_a[13] = c1; row_a[14] = cl1
                    row_b = list(row_data); row_b[13] = c2; row_b[14] = cl2
                    result_rows.extend([row_a, row_b])

                else:
                    # ── 일반 세트 분리 ──
                    row_a = list(row_data); row_a[13] = set1
                    row_b = list(row_data); row_b[13] = set2
                    result_rows.extend([row_a, row_b])

            else:
                result_rows.append(list(row_data))

        self.sdf2 = pd.DataFrame(result_rows, columns=self.sdf.columns)
        return True

    def output_product_code_extraction(self):
        """
        ② 상품코드 추출 전체 파이프라인:
           검증 → 추출 → 세트 분리 → 코드 변환 → 옵션코드 생성 → 화면 표시
        """
        # ─ 유효성 검사 ─
        prog = self.combo.get()
        if prog == PROGRAM_OPTIONS[0]:
            messagebox.showwarning("경고", "수집 프로그램을 선택해 주세요.")
            return
        if prog == " 1. 사방넷" and "우체국택배업로드" not in self.ent_file.get():
            messagebox.showwarning("경고", "사방넷: 우체국택배업로드 주문서를 선택해 주세요.")
            return
        if prog == " 2. 이지어드민" and "이지어드민" not in self.ent_file.get():
            messagebox.showwarning("경고", "이지어드민 주문서를 선택해 주세요.")
            return
        if not self.ent_file.get():
            messagebox.showwarning("경고", "주문서 파일을 먼저 선택해 주세요.")
            return

        # ─ 추출 & 분리 ─
        if not self.order_product_code_extraction():
            return
        if not self.set_product_separation():
            return

        # ─ 코드 변환 & 옵션코드 생성 ─
        cnt = self.sdf2.shape[0]
        for r in range(cnt):
            raw_code = self.sdf2.iloc[r, 13]
            color    = self.sdf2.iloc[r, 14]
            size     = self.sdf2.iloc[r, 15]

            # 옵션코드 (원본 코드 기반)
            self.sdf2.iloc[r, 16] = self._build_option_code(raw_code, color, size)

            # 변환 코드 & 옵션코드2
            code2 = self.extract_itemcode(raw_code)
            self.sdf2.iloc[r, 17] = code2
            self.sdf2.iloc[r, 18] = self._build_option_code(code2, color, size)

            # 진행률
            self.lbl_row.configure(text=f"ROW : {r + 1}")
            self.p_var.set((r + 1) / cnt * 100)
            self.update_idletasks()

        # ─ 화면 표시 ─
        VIEW_COLS_L1 = ["성명", "상품명", "옵션코드", "수량", "주문처",
                        "상품코드2", "컬러", "사이즈", "옵션코드2"]
        self.sdf2_view = self.sdf2.loc[:, VIEW_COLS_L1]
        self.sheet_l1 = self._refresh_sheet("sheet_l1", self.tab_l1, self.sdf2_view, "l1")

        UPLOAD_COLS = ["성명", "전화번호", "핸드폰", "우편번호", "주소",
                       "옵션코드", "수량", "배송메세지", "주문처",
                       "요금구분", "운송장번호", "사방넷주문번호", "쇼핑몰아이디"]
        self.df_upload = self.sdf2.loc[:, UPLOAD_COLS].rename(columns={"옵션코드": "상품명"})

        messagebox.showinfo("완료", "상품코드 추출 및 변환이 완료되었습니다.")

    # ----------------------------------------------------------
    #  STEP ③: 출고리스트 생성
    # ----------------------------------------------------------

    def create_factory_list(self):
        """
        ③ 출고리스트 생성:
           판매데이터 집계 → 전체 출고리스트(피벗) → 티셔츠 전용 출고리스트
        """
        if self.sdf2 is None:
            messagebox.showwarning("경고", "먼저 ② 상품코드 추출을 실행해 주세요.")
            return

        # ─ 판매데이터 ─
        self.df_sale = (
            self.sdf2.loc[:, ["주문처", "상품코드2", "컬러", "사이즈", "수량"]]
            .groupby(["주문처", "상품코드2", "컬러", "사이즈"], as_index=False)
            .agg({"수량": "sum"})
        )
        self.df_sale["날짜"] = self.gDate
        self.df_sale = self.df_sale[["날짜", "상품코드2", "컬러", "사이즈", "수량", "주문처"]]
        self.sum_ship = str(self.df_sale["수량"].sum())

        summary = f"주문수량 : {self.sum_order}    /    출고수량 : {self.sum_ship}"
        self.lbl_summary.configure(text=summary)
        self.sheet_l2 = self._refresh_sheet("sheet_l2", self.tab_l2, self.df_sale, "l2")

        # ─ 전체 출고리스트 (피벗) ─
        pivot = (
            self.sdf2.pivot_table("수량", index="옵션코드2", aggfunc="sum")
            .reset_index()
            .rename(columns={"옵션코드2": "옵션코드"})
        )
        self.odf = pivot

        cnt = self.odf.shape[0]
        for i in range(cnt):
            self.p_var.set((i + 1) / cnt * 100)
            self.update_idletasks()

        self.sheet_r1 = self._refresh_sheet("sheet_r1", self.tab_r1, self.odf, "r1")
        self.p_var.set(0)

        # ─ 티셔츠 출고리스트 ─
        if os.path.isfile(self.stock_file):
            try:
                st_df = pd.read_excel(self.stock_file, skiprows=1)[["상품코드", "재고수량"]]
                stock_map = dict(zip(st_df["상품코드"], st_df["재고수량"]))
            except Exception:
                stock_map = {}
        else:
            stock_map = {}

        tshirt_rows = []
        cnt2 = self.sdf2.shape[0]
        for n in range(cnt2):
            opt2 = str(self.sdf2.iloc[n, 18])
            if opt2 and opt2[0] in TSHIRT_PREFIXES:
                tshirt_rows.append({
                    "주문자명": self.sdf2.iloc[n, 0],
                    "옵션코드": opt2,
                    "수량":    self.sdf2.iloc[n, 6],
                    "재고":    stock_map.get(opt2, ""),
                })
            self.p_var.set((n + 1) / cnt2 * 100)
            self.update_idletasks()

        self.ndf = pd.DataFrame(tshirt_rows, columns=["주문자명", "옵션코드", "수량", "재고"])
        self.sheet_r2 = self._refresh_sheet("sheet_r2", self.tab_r2, self.ndf, "r2")

        self.p_var.set(0)
        messagebox.showinfo("완료", "출고리스트 생성이 완료되었습니다.")

    # ----------------------------------------------------------
    #  STEP ④: 저장 / 복사
    # ----------------------------------------------------------

    def browse_save_path(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ent_savepath.delete(0, tk.END)
            self.ent_savepath.insert(0, folder)

    def save_file(self):
        """택배 업로드용 Excel 저장"""
        if not self.ent_savepath.get():
            messagebox.showwarning("경고", "저장 경로를 선택해 주세요.")
            return
        if self.df_upload is None:
            messagebox.showwarning("경고", "먼저 ② 상품코드 추출을 실행해 주세요.")
            return
        try:
            save_path = os.path.join(self.ent_savepath.get(), "택배업로드.xlsx")
            self.df_upload.to_excel(save_path, sheet_name="택배업로드", index=False)
            messagebox.showinfo("완료", f"파일이 저장되었습니다:\n{save_path}")
        except Exception as e:
            messagebox.showerror("에러", f"저장 실패:\n{e}")

    def _clipboard_copy(self, df, label: str):
        if df is None:
            messagebox.showwarning("경고", f"{label} 데이터가 없습니다.\n먼저 해당 단계를 실행해 주세요.")
            return
        df.to_clipboard(index=False)
        messagebox.showinfo("완료", f"{label} 복사 완료. 엑셀에 붙여넣기 하세요.")

    def copy_code_ext(self):
        self._clipboard_copy(self.df_upload, "상품코드")

    def copy_output_list1(self):
        self._clipboard_copy(self.odf, "전체 출고리스트")

    def copy_output_list2(self):
        self._clipboard_copy(self.ndf, "티셔츠 출고리스트")

    def copy_sale_data(self):
        self._clipboard_copy(self.df_sale, "판매데이터")

    # ----------------------------------------------------------
    #  초기화
    # ----------------------------------------------------------

    def reset(self):
        """애플리케이션 초기화"""
        self.ent_file.delete(0, tk.END)
        self.ent_savepath.delete(0, tk.END)
        self.lbl_summary.configure(text="주문수량 : -    /    출고수량 : -")
        self.lbl_row.configure(text="ROW : 0")
        self.p_var.set(0)

        self.sdf = self.sdf2 = self.sdf2_view = None
        self.odf = self.ndf = self.df_sale = self.df_upload = None

        self.sheet_l1 = self._refresh_sheet("sheet_l1", self.tab_l1, self.empty_df, "l1")
        self.sheet_l2 = self._refresh_sheet("sheet_l2", self.tab_l2, self.empty_df, "l2")
        self.sheet_r1 = self._refresh_sheet("sheet_r1", self.tab_r1, self.empty_df, "r1")
        self.sheet_r2 = self._refresh_sheet("sheet_r2", self.tab_r2, self.empty_df, "r2")

    # ----------------------------------------------------------
    #  도움말
    # ----------------------------------------------------------

    def manual(self):
        msg = (
            "═══ 사용 방법 ═══\n\n"
            "① 수집 프로그램 선택 (사방넷 / 이지어드민)\n"
            "② 주문서 선택 버튼으로 파일 불러오기\n"
            "③ 상품코드 추출 버튼 클릭\n"
            "④ 출고리스트 생성 버튼 클릭\n"
            "⑤ 각 복사 버튼으로 클립보드 복사 후 Excel 붙여넣기\n\n"
            "═══ 파일명 규칙 ═══\n\n"
            "사방넷:     20250301_우체국택배업로드.xlsx\n"
            "이지어드민: 20250301_이지어드민.xls\n\n"
            "═══ 오류 해결 ═══\n\n"
            "• 코드 변환 오류 → NAS DB폴더 '변환코드.xlsx' 수정\n"
            "• 파일 읽기 오류 → Excel로 다시 저장 후 재시도\n\n"
            "═══ 컬럼 너비 ═══\n\n"
            "하단 '↔ 컬럼 너비' 버튼으로 각 시트 컬럼 너비 조정 가능"
        )
        messagebox.showinfo("사용 설명", msg)


# ============================================================
if __name__ == "__main__":
    app = OrderProcessingApp()
    app.mainloop()
