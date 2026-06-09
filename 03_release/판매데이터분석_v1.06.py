"""
========================
엑셀 판매데이터 분석 GUI (CustomTkinter + tksheet) v1.05
========================

[업데이트 내역 v1.05]
1. UI 프레임워크 변경: Tkinter -> CustomTkinter (현대적, 프리미엄 디자인)
2. 성능 최적화: 계층형 데이터(상품-컬러-사이즈) 사전 인덱싱으로 확장 속도 획기적 개선 (O(N) -> O(1))
3. 구조 개선: 데이터 로직(SalesDataManager)과 UI 로직(SalesAnalyzerUI) 분리
4. 다크모드/라이트모드 대응
5. 로딩 오버레이 추가: 데이터 처리 시 프로그레스 표시
6. 속도 개선: 비동기 분석 처리 및 시트 렌더링 최적화

[요구사항]
- 입력: 검색기간(시작/종료), 최근 기간 단축버튼, 쇼핑몰 필터
- 출력: 
  1) [합계기준 정렬] 상품코드 × 일자 (확장시 컬러/사이즈 상세)
  2) [합계기준 정렬] 상품코드 × 쇼핑몰
  3) [일자순 정렬] 일자 × 쇼핑몰
  4) 년도 × 쇼핑몰
- 기능: 엑셀 로딩(Pickle 캐싱), 엑셀 내보내기, 초기화

[필수 라이브러리]
pip install pandas openpyxl tkcalendar tksheet customtkinter
"""

import os
import pickle
import datetime as dt
import threading
from dataclasses import dataclass
from typing import List, Optional, Tuple, Callable, Dict, Any

import tkinter as tk
from tkinter import messagebox, filedialog
import pandas as pd

try:
    import customtkinter as ctk
except ImportError:
    messagebox.showerror("라이브러리 오류", "customtkinter가 설치되지 않았습니다.\n`pip install customtkinter`를 실행해주세요.")
    raise SystemExit("customtkinter 미설치")

try:
    from tkcalendar import DateEntry
except ImportError:
    raise SystemExit("tkcalendar 미설치: `pip install tkcalendar` 후 다시 실행하세요.")
try:
    from tksheet import Sheet
except ImportError:
    raise SystemExit("tksheet 미설치: `pip install tksheet` 후 다시 실행하세요.")

try:
    import openpyxl  # noqa: F401
except ImportError:
    raise SystemExit("openpyxl 미설치: `pip install openpyxl` 후 다시 실행하세요.")


# =========================
# 상수 및 설정
# =========================
DEFAULT_EXCEL_PATH = r"\\NAS451\team451\DB\통합매출데이터.xlsx"
DATE_COL_CANDIDATES = ["주문일자", "일자", "날짜", "order_date", "date"]
SKU_COL_CANDIDATES = ["상품코드", "상품", "product_code", "sku"]
MALL_COL_CANDIDATES = ["쇼핑몰", "몰명", "채널", "mall", "channel"]
QTY_COL_CANDIDATES  = ["수량", "판매수량", "qty", "quantity"]
COLOR_COL_CANDIDATES = ["컬러", "색상", "색깔", "color", "colour"]
SIZE_COL_CANDIDATES = ["사이즈", "크기", "size"]

# 시트 설정
HEADER_HEIGHT = 45         # 헤더 높이
ROW_HEIGHT = 28            # 행 높이
DEFAULT_COLUMN_WIDTH = 80  # 픽셀 단위 (가독성 향상)
MIN_COLUMN_WIDTH = 20       # 픽셀 단위 (최소 넓이)
FONT_NAME = "나눔고딕"    # 가독성 좋은 기본 폰트
FONT_SIZE = 10             # 시트 폰트 크기

# Light 테마 색상 팔레트
THEME_COLORS = {
    "header_row_bg": "#dbeafe",       # 연한 파랑 (합계 행)
    "header_row_fg": "#1e3a5f",       # 진한 파랑 (합계 행 텍스트)
    "total_col_bg": "#e0f2fe",        # 연한 하늘색 (합계 컬럼)
    "total_col_fg": "#0c4a6e",        # 진한 파랑 (합계 컬럼 텍스트)
    "overlay_bg": "#000000",          # 로딩 오버레이 배경
    "overlay_alpha": 0.5,             # 오버레이 투명도
}

ctk.set_appearance_mode("Light")  # Light 테마 적용
ctk.set_default_color_theme("blue")


@dataclass
class ColumnMap:
    date: str
    sku: str
    mall: str
    qty: str
    color: Optional[str] = None
    size: Optional[str] = None


# =========================
# 데이터 관리 클래스 (로직)
# =========================
class SalesDataManager:
    def __init__(self):
        self.df_raw: Optional[pd.DataFrame] = None
        self.colmap: Optional[ColumnMap] = None
        self.df_filtered: Optional[pd.DataFrame] = None
        self.is_dummy: bool = False 
        
    def _auto_map_columns(self, df: pd.DataFrame) -> ColumnMap:
        """컬럼 자동 매핑"""
        columns_lower = {str(c).lower(): c for c in df.columns}
        
        def find(cands: List[str]) -> Optional[str]:
            for c in df.columns:
                if c in cands: return c
            for cand in cands:
                if cand.lower() in columns_lower: return columns_lower[cand.lower()]
            return None

        date_col = find(DATE_COL_CANDIDATES)
        sku_col  = find(SKU_COL_CANDIDATES)
        mall_col = find(MALL_COL_CANDIDATES)
        qty_col  = find(QTY_COL_CANDIDATES)
        color_col = find(COLOR_COL_CANDIDATES)
        size_col = find(SIZE_COL_CANDIDATES)

        missing = [n for n, v in [("날짜", date_col), ("상품코드", sku_col), 
                                  ("쇼핑몰", mall_col), ("수량", qty_col)] if v is None]
        if missing:
            raise ValueError(f"필수 컬럼 인식 실패: {', '.join(missing)}")
            
        return ColumnMap(date=date_col, sku=sku_col, mall=mall_col, qty=qty_col, 
                         color=color_col, size=size_col)

    def _make_dummy_data(self) -> pd.DataFrame:
        rng = pd.date_range(end=dt.date.today(), periods=90, freq="D")
        skus = ["TPX-001", "TPX-002", "TPW-010"]
        malls = ["스마트스토어", "쿠팡", "11번가"]
        colors = ["빨강", "파랑", "검정", "흰색"]
        sizes = ["S", "M", "L", "XL"]
        rows = []
        import random
        for d in rng:
            for sku in skus:
                for mall in malls:
                    if random.random() > 0.3: continue
                    color = random.choice(colors)
                    size = random.choice(sizes)
                    qty = random.randint(1, 5)
                    rows.append([d.date(), sku, mall, qty, color, size])
        return pd.DataFrame(rows, columns=["주문일자", "상품코드", "쇼핑몰", "수량", "컬러", "사이즈"])

    def load_data(self, path: str, callback: Optional[Callable[[str, int], None]] = None) -> Tuple[bool, str]:
        """엑셀/피클 로드 (callback: 진행 메시지, 퍼센트)"""
        try:
            pickle_path = os.path.splitext(path)[0] + ".pkl"
            
            # Pickle 로드 시도
            if os.path.exists(pickle_path) and os.path.exists(path):
                if os.path.getmtime(pickle_path) >= os.path.getmtime(path):
                    if callback: callback("캐시된 데이터 로딩 중...", 20)
                    try:
                        with open(pickle_path, 'rb') as f:
                            data = pickle.load(f)
                            self.df_raw = data['df']
                            self.colmap = data['colmap']
                            self.is_dummy = False
                            if callback: callback("완료!", 100)
                            return True, f"캐시 로드 성공 ({len(self.df_raw):,}행)"
                    except Exception:
                        pass # 피클 로드 실패 시 엑셀 로드 진행

            # 엑셀 로드
            if callback: callback("엑셀 파일 읽는 중...", 10)
            if not os.path.exists(path):
                raise FileNotFoundError(f"파일 없음: {path}")
                
            df = pd.read_excel(path, engine="openpyxl")
            if callback: callback("컬럼 매핑 중...", 40)
            colmap = self._auto_map_columns(df)
            
            if callback: callback("데이터 전처리 중...", 60)
            # 전처리
            if df[colmap.date].dtype == 'object':
                df[colmap.date] = pd.to_datetime(df[colmap.date], errors='coerce').dt.date
            else:
                df[colmap.date] = pd.to_datetime(df[colmap.date]).dt.date
                
            df[colmap.qty] = pd.to_numeric(df[colmap.qty], errors="coerce").fillna(0).astype('int32')
            
            # 필요한 컬럼만 유지
            cols = [colmap.date, colmap.sku, colmap.mall, colmap.qty]
            if colmap.color: cols.append(colmap.color)
            if colmap.size: cols.append(colmap.size)
            self.df_raw = df[cols].copy()
            self.colmap = colmap
            self.is_dummy = False
            
            # Pickle 저장
            if callback: callback("캐시 생성 중...", 85)
            with open(pickle_path, 'wb') as f:
                pickle.dump({'df': self.df_raw, 'colmap': self.colmap}, f, protocol=pickle.HIGHEST_PROTOCOL)
                
            if callback: callback("완료!", 100)
            return True, f"로드 성공 ({len(self.df_raw):,}행)"
            
        except Exception as e:
            if callback: callback("더미 데이터 생성 중...", 50)
            self.df_raw = self._make_dummy_data()
            self.colmap = self._auto_map_columns(self.df_raw)
            self.is_dummy = True
            if callback: callback("더미 데이터 로드 완료", 100)
            return False, f"[더미 데이터] 오류: {e}"

    def filter_data(self, start_date: dt.date, end_date: dt.date, malls: List[str]):
        """데이터 필터링"""
        if self.df_raw is None: return
        
        mask = (self.df_raw[self.colmap.date] >= start_date) & (self.df_raw[self.colmap.date] <= end_date)
        if malls:
            mask &= self.df_raw[self.colmap.mall].isin(set(malls))
            
        self.df_filtered = self.df_raw.loc[mask].copy()

    def pivot_generic(self, index_col: str, columns_col: str, agg_func: str = "sum", df_source: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """일반 피벗 생성 (df_source가 None이면 df_filtered 사용)"""
        df = df_source if df_source is not None else self.df_filtered
        if df is None or df.empty:
            return pd.DataFrame()
            
        pv = pd.pivot_table(
            df, 
            index=index_col, 
            columns=columns_col, 
            values=self.colmap.qty, 
            aggfunc=agg_func, 
            fill_value=0
        )
        
        pv["합계"] = pv.sum(axis=1)
        if index_col == self.colmap.date: 
            pv = pv.sort_index(ascending=True)
        elif index_col == "년도":
             pv = pv.sort_index(ascending=False)
        else: 
            pv = pv.sort_values("합계", ascending=False)
            
        cols = ["합계"] + [c for c in pv.columns if c != "합계"]
        pv = pv[cols]
        pv = pv.reset_index()
        
        new_cols = {}
        for c in pv.columns:
            if isinstance(c, (dt.date, dt.datetime)):
                new_cols[c] = f"{c.year}\n{c.month}/{c.day}"
        pv = pv.rename(columns=new_cols)
        
        sum_row = {}
        for c in pv.columns:
            if c == index_col: sum_row[c] = "합계"
            elif pv[c].dtype in ['int64', 'float64', 'int32']:
                sum_row[c] = pv[c].sum()
            else:
                sum_row[c] = ""
        
        return pd.concat([pd.DataFrame([sum_row]), pv], ignore_index=True)

    def save_excel(self, dfs: Dict[str, pd.DataFrame], file_name: str) -> bool:
        try:
            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                for sheet_name, df in dfs.items():
                    if df.empty: continue
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    ws = writer.sheets[sheet_name]
                    ws.freeze_panes = "C3"
            return True
        except Exception as e:
            print(e)
            return False


# =========================
# 로딩 오버레이 위젯
# =========================
class LoadingOverlay(ctk.CTkFrame):
    """데이터 처리 시 표시되는 로딩 오버레이"""
    
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=("gray85", "gray20"), corner_radius=0, **kwargs)
        
        # 중앙 카드
        self.card = ctk.CTkFrame(self, corner_radius=15, fg_color=("white", "gray25"))
        self.card.place(relx=0.5, rely=0.5, anchor="center")
        
        # 로딩 아이콘 (애니메이션 대신 텍스트)
        self.spinner_label = ctk.CTkLabel(
            self.card, 
            text="⏳", 
            font=(FONT_NAME, 44)
        )
        self.spinner_label.pack(pady=(30, 10))
        
        # 상태 메시지
        self.message_label = ctk.CTkLabel(
            self.card, 
            text="처리 중...", 
            font=(FONT_NAME, 13)
        )
        self.message_label.pack(pady=(5, 10))
        
        # 프로그레스바
        self.progress = ctk.CTkProgressBar(self.card, width=250, height=12, corner_radius=6)
        self.progress.pack(pady=(10, 30), padx=40)
        self.progress.set(0)
        
        # 애니메이션 상태
        self._animation_running = False
        self._spinner_chars = ["⏳", "⌛"]
        self._spinner_index = 0
        
    def show(self, message: str = "처리 중..."):
        """오버레이 표시"""
        self.message_label.configure(text=message)
        self.progress.set(0)
        self.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.lift()
        self._start_animation()
        self.update()
        
    def hide(self):
        """오버레이 숨김"""
        self._stop_animation()
        self.place_forget()
        
    def update_progress(self, message: str, percent: int):
        """진행률 업데이트"""
        self.message_label.configure(text=message)
        self.progress.set(percent / 100)
        self.update()
        
    def _start_animation(self):
        """스피너 애니메이션 시작"""
        self._animation_running = True
        self._animate_spinner()
        
    def _stop_animation(self):
        """스피너 애니메이션 중지"""
        self._animation_running = False
        
    def _animate_spinner(self):
        """스피너 애니메이션 프레임"""
        if not self._animation_running:
            return
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_chars)
        self.spinner_label.configure(text=self._spinner_chars[self._spinner_index])
        self.after(500, self._animate_spinner)


# =========================
# UI 클래스 (CustomTkinter)
# =========================
class SalesAnalyzerUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.data_mgr = SalesDataManager()
        
        self.title("판매데이터 분석 솔루션 v1.06")
        self.geometry("1400x900")
        self.minsize(1024, 768)
        
        # 시트 캐시 (성능 최적화)
        self._sheet_cache: Dict[str, Sheet] = {}
        
        self._setup_layout()
        self._init_data()

    def _setup_layout(self):
        # 1. 상단 컨트롤 패널
        self.top_frame = ctk.CTkFrame(self, corner_radius=10)
        self.top_frame.pack(fill="x", padx=10, pady=10)
        self._setup_controls(self.top_frame)
        
        # 2. 메인 컨테이너 (탭뷰 + 오버레이용)
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        
        # 3. 탭뷰
        self.tabview = ctk.CTkTabview(self.main_container)
        self.tabview.pack(fill="both", expand=True)
        
        self.tabs = {}
        tab_names = ["상품별 분석", "쇼핑몰별 분석", "일자별 추이", "연도별 추이"]
        for name in tab_names:
            self.tabview.add(name)
            self.tabs[name] = self.tabview.tab(name)
        
        # 4. 로딩 오버레이
        self.loading_overlay = LoadingOverlay(self.main_container)
            
        # 5. 하단 상태
        self.bottom_frame = ctk.CTkFrame(self, height=40, fg_color="transparent")
        self.bottom_frame.pack(fill="x", padx=10, pady=5)
        self.lbl_status = ctk.CTkLabel(self.bottom_frame, text="준비됨", text_color="gray")
        self.lbl_status.pack(side="left")
        
        ctk.CTkButton(self.bottom_frame, text="엑셀 저장", command=self._cmd_export, width=100).pack(side="right")
        ctk.CTkButton(self.bottom_frame, text="초기화", command=self._cmd_reset, width=100, fg_color="gray").pack(side="right", padx=5)

    def _setup_controls(self, parent):
        parent.grid_columnconfigure(9, weight=1)
        
        self.btn_load = ctk.CTkButton(parent, text="📂 데이터 파일 열기", command=self._cmd_load_file, width=140)
        self.btn_load.grid(row=0, column=0, padx=10, pady=15)
        
        self.lbl_file = ctk.CTkLabel(parent, text="파일 미선택 (더미 모드)")
        self.lbl_file.grid(row=0, column=1, sticky="w", padx=5)
        
        date_frame = ctk.CTkFrame(parent, fg_color="transparent")
        date_frame.grid(row=0, column=2, padx=20)
        ctk.CTkLabel(date_frame, text="기간:").pack(side="left", padx=5)
        self.de_start = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.de_start.pack(side="left")
        ctk.CTkLabel(date_frame, text="~").pack(side="left", padx=5)
        self.de_end = DateEntry(date_frame, width=12, date_pattern='yyyy-mm-dd')
        self.de_end.pack(side="left")
        
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=0, column=3)
        periods = [("1주", 7), ("2주", 14), ("1개월", 30), ("3개월", 90), ("6개월", 180)]
        for text, days in periods:
            ctk.CTkButton(btn_frame, text=text, width=50, height=24, 
                          command=lambda d=days: self._set_date_range(d)).pack(side="left", padx=2)

        self.listbox_mall = tk.Listbox(parent, selectmode="extended", height=5, width=20, font=(FONT_NAME, 10))
        self.listbox_mall.grid(row=0, column=4, padx=15, pady=5)
        
        ctk.CTkButton(parent, text="▶ 분석 실행", command=self._cmd_run_analysis, 
                      font=(FONT_NAME, 13, "bold"), width=120, height=50).grid(row=0, column=5, padx=15)

    def _init_data(self):
        self._set_date_range(7) 
        self.loading_overlay.show("데이터 로딩 중...")
        threading.Thread(target=self._async_load, args=(DEFAULT_EXCEL_PATH,), daemon=True).start()

    def _async_load(self, path):
        """비동기 데이터 로드"""
        def cb(msg, percent):
            self.after(0, lambda: self.loading_overlay.update_progress(msg, percent))
        success, msg = self.data_mgr.load_data(path, cb)
        self.after(0, lambda: self._on_load_complete(success, msg, path))

    def _on_load_complete(self, success, msg, path):
        """데이터 로드 완료 처리"""
        self.loading_overlay.hide()
        self.lbl_status.configure(text=msg)
        
        if success and not self.data_mgr.is_dummy:
            self.lbl_file.configure(text=os.path.basename(path), text_color="#4CC2FF")
        else:
            self.lbl_file.configure(text="더미 데이터 (실제 파일 없음)", text_color="orange")
            
        self.listbox_mall.delete(0, "end")
        if self.data_mgr.df_raw is not None and self.data_mgr.colmap is not None:
            malls = sorted(self.data_mgr.df_raw[self.data_mgr.colmap.mall].unique().astype(str))
            for m in malls:
                self.listbox_mall.insert("end", m)
            self.listbox_mall.select_set(0, "end")
        
        self._run_analysis_sync()

    def _set_date_range(self, days):
        ed = dt.date.today()
        sd = ed - dt.timedelta(days=days-1)
        self.de_start.set_date(sd)
        self.de_end.set_date(ed)

    def _cmd_load_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if f:
            self.loading_overlay.show("파일 로딩 중...")
            threading.Thread(target=self._async_load, args=(f,), daemon=True).start()

    def _cmd_reset(self):
        self._set_date_range(7)
        self.listbox_mall.select_set(0, "end")
        self._cmd_run_analysis()

    def _cmd_export(self):
        if not hasattr(self, 'current_pivots') or not self.current_pivots:
            messagebox.showwarning("경고", "분석 결과가 없습니다.")
            return
            
        f = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                         filetypes=[("Excel Files", "*.xlsx")],
                                         initialfile=f"판매분석_{dt.datetime.now().strftime('%Y%m%d')}.xlsx")
        if f:
            self.loading_overlay.show("엑셀 저장 중...")
            
            def save_task():
                dfs = {
                    "상품별": self.current_pivots[0],
                    "쇼핑몰별": self.current_pivots[1],
                    "일자별": self.current_pivots[2],
                    "연도별": self.current_pivots[3],
                }
                result = self.data_mgr.save_excel(dfs, f)
                self.after(0, lambda: self._on_save_complete(result))
            
            threading.Thread(target=save_task, daemon=True).start()
    
    def _on_save_complete(self, success):
        """저장 완료 처리"""
        self.loading_overlay.hide()
        if success:
            messagebox.showinfo("완료", "저장되었습니다.")
        else:
            messagebox.showerror("오류", "저장 실패")

    def _get_selected_malls(self):
        idxs = self.listbox_mall.curselection()
        if not idxs: return []
        return [self.listbox_mall.get(i) for i in idxs]

    def _cmd_run_analysis(self):
        """분석 실행 (UI 버튼용 - 로딩 오버레이 표시)"""
        if self.data_mgr.colmap is None:
            self.lbl_status.configure(text="데이터를 먼저 로드해주세요.")
            return
            
        start = self.de_start.get_date()
        end = self.de_end.get_date()
        
        if start > end:
            messagebox.showwarning("오류", "시작일이 종료일보다 큽니다.")
            return
        
        self.loading_overlay.show("분석 처리 중...")
        threading.Thread(target=self._async_analysis, daemon=True).start()

    def _async_analysis(self):
        """비동기 분석 처리"""
        try:
            start = self.de_start.get_date()
            end = self.de_end.get_date()
            malls = self._get_selected_malls()
            
            self.after(0, lambda: self.loading_overlay.update_progress("데이터 필터링 중...", 20))
            self.data_mgr.filter_data(start, end, malls)
            
            self.after(0, lambda: self.loading_overlay.update_progress("상품별 피벗 생성 중...", 40))
            pv_prod = self.data_mgr.pivot_generic(self.data_mgr.colmap.sku, self.data_mgr.colmap.date)
            
            self.after(0, lambda: self.loading_overlay.update_progress("쇼핑몰별 피벗 생성 중...", 55))
            pv_mall_prod = self.data_mgr.pivot_generic(self.data_mgr.colmap.sku, self.data_mgr.colmap.mall)
            
            self.after(0, lambda: self.loading_overlay.update_progress("일자별 피벗 생성 중...", 70))
            pv_date = self.data_mgr.pivot_generic(self.data_mgr.colmap.date, self.data_mgr.colmap.mall)
            
            self.after(0, lambda: self.loading_overlay.update_progress("연도별 피벗 생성 중...", 85))
            # 연도별 추이는 검색기간에 상관없이 전체 데이터 사용
            df_all = self.data_mgr.df_raw.copy()
            if malls:  # 쇼핑몰 필터만 적용
                df_all = df_all[df_all[self.data_mgr.colmap.mall].isin(set(malls))]
            # NA 값 제거 후 연도 추출
            df_all = df_all.dropna(subset=[self.data_mgr.colmap.date])
            df_all["년도"] = pd.to_datetime(df_all[self.data_mgr.colmap.date]).dt.year.astype(int)
            pv_year = self.data_mgr.pivot_generic("년도", self.data_mgr.colmap.mall, df_source=df_all)
            
            pivots = [pv_prod, pv_mall_prod, pv_date, pv_year]
            
            self.after(0, lambda: self._on_analysis_complete(pivots))
            
        except Exception as e:
            self.after(0, lambda: self._on_analysis_error(str(e)))

    def _on_analysis_complete(self, pivots):
        """분석 완료 UI 업데이트"""
        self.current_pivots = pivots
        
        self.loading_overlay.update_progress("시트 렌더링 중...", 95)
        
        self._render_sheet(self.tabs["상품별 분석"], pivots[0])
        self._render_sheet(self.tabs["쇼핑몰별 분석"], pivots[1])
        self._render_sheet(self.tabs["일자별 추이"], pivots[2])
        self._render_sheet(self.tabs["연도별 추이"], pivots[3])
        
        self.loading_overlay.hide()
        
        count = len(self.data_mgr.df_filtered) if self.data_mgr.df_filtered is not None else 0
        self.lbl_status.configure(text=f"분석 완료 ({count:,}건 처리)")

    def _on_analysis_error(self, error_msg):
        """분석 오류 처리"""
        self.loading_overlay.hide()
        self.lbl_status.configure(text=f"분석 오류: {error_msg}")
        messagebox.showerror("분석 오류", f"분석 중 오류가 발생했습니다:\n{error_msg}")
    
    def _run_analysis_sync(self):
        """동기 분석 실행 (초기 로드 후 사용)"""
        if self.data_mgr.colmap is None:
            return
            
        start = self.de_start.get_date()
        end = self.de_end.get_date()
        malls = self._get_selected_malls()
        
        self.data_mgr.filter_data(start, end, malls)

        pv_prod = self.data_mgr.pivot_generic(self.data_mgr.colmap.sku, self.data_mgr.colmap.date)
        pv_mall_prod = self.data_mgr.pivot_generic(self.data_mgr.colmap.sku, self.data_mgr.colmap.mall)
        pv_date = self.data_mgr.pivot_generic(self.data_mgr.colmap.date, self.data_mgr.colmap.mall)
        
        # 연도별 추이
        df_all = self.data_mgr.df_raw.copy()
        if malls:
            df_all = df_all[df_all[self.data_mgr.colmap.mall].isin(set(malls))]
        df_all = df_all.dropna(subset=[self.data_mgr.colmap.date])
        df_all["년도"] = pd.to_datetime(df_all[self.data_mgr.colmap.date]).dt.year.astype(int)
        pv_year = self.data_mgr.pivot_generic("년도", self.data_mgr.colmap.mall, df_source=df_all)

        self.current_pivots = [pv_prod, pv_mall_prod, pv_date, pv_year]

        self._render_sheet(self.tabs["상품별 분석"], pv_prod)
        self._render_sheet(self.tabs["쇼핑몰별 분석"], pv_mall_prod)
        self._render_sheet(self.tabs["일자별 추이"], pv_date)
        self._render_sheet(self.tabs["연도별 추이"], pv_year)
        
        count = len(self.data_mgr.df_filtered) if self.data_mgr.df_filtered is not None else 0
        self.lbl_status.configure(text=f"분석 완료 ({count:,}건 처리)")

    def _render_sheet(self, parent_tab, df: pd.DataFrame):
        """시트 렌더링 (최적화됨)"""
        for widget in parent_tab.winfo_children():
            widget.destroy()
            
        if df.empty:
            ctk.CTkLabel(parent_tab, text="데이터가 없습니다.").pack(pady=20)
            return

        frame = ctk.CTkFrame(parent_tab)
        frame.pack(fill="both", expand=True)

        # 데이터 전처리 (포맷팅)
        data = []
        for r in range(len(df)):
            row = []
            for c in range(len(df.columns)):
                val = df.iloc[r, c]
                if isinstance(val, (int, float)) and not pd.isna(val):
                    row.append(f"{int(val):,}" if isinstance(val, int) or val == int(val) else f"{val:,.2f}")
                else:
                    row.append(str(val) if pd.notna(val) else "")
            data.append(row)

        sheet = Sheet(frame, 
                      header_height=HEADER_HEIGHT,
                      headers=list(df.columns),
                      data=data,
                      theme="light",
                      font=(FONT_NAME, FONT_SIZE, "normal"),
                      header_font=(FONT_NAME, FONT_SIZE, "bold"))
        
        # 시트 옵션 설정
        try:
            sheet.set_options(
                header_height=HEADER_HEIGHT,
                row_height=ROW_HEIGHT,
                default_column_width=DEFAULT_COLUMN_WIDTH,
                min_column_width=MIN_COLUMN_WIDTH,
                header_wrap=True,
            )
        except Exception:
            pass
        
        sheet.table_align(align="left")
        sheet.set_all_column_widths(width=None)
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)
        
        # 첫 번째 컬럼은 왼쪽 정렬, 두 번째 컬럼부터 오른쪽 정렬
        try:
            # 첫 번째 컬럼(0번 인덱스)은 왼쪽 정렬
            sheet.align_columns(columns=[0], align="w", redraw=False)
            
            # 두 번째 컬럼부터(1번 인덱스~) 오른쪽 정렬
            if len(df.columns) > 1:
                for c in range(1, len(df.columns)):
                    sheet.align_columns(columns=[c], align="e", redraw=False)
        except Exception:
            pass

        
        # 합계 행 강조
        try:
            sheet.highlight_rows(rows=[0], bg=THEME_COLORS["header_row_bg"], fg=THEME_COLORS["header_row_fg"])
        except Exception:
            pass
        
        # 합계 컬럼 강조 (첫 번째 컬럼)
        try:
            # 합계 컬럼이 "합계"로 시작하는 경우 찾기
            if len(df.columns) > 0 and "합계" in str(df.columns[0]):
                sheet.highlight_columns(columns=[0], bg=THEME_COLORS["total_col_bg"], fg=THEME_COLORS["total_col_fg"])
        except Exception:
            pass


if __name__ == "__main__":
    app = SalesAnalyzerUI()
    app.mainloop()
