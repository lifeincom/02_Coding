"""
========================
엑셀 판매데이터 분석 GUI (Tkinter + tksheet) v1.03
========================

[기본 기능]
- 입력: 검색기간(시작/종료), 최근 기간 단축버튼, 쇼핑몰 필터
- 출력: 
  1) [합계기준 정렬] 상품코드 × 일자
  2) [합계기준 정렬] 상품코드 × 쇼핑몰
  3) [일자순 정렬] 일자 × 쇼핑몰
  4) 년도 × 쇼핑몰
- 기능: 엑셀 로딩(Pickle 캐싱), 엑셀 내보내기, 초기화

[필수 라이브러리]
pip install pandas openpyxl tkcalendar tksheet
"""

import os
import pickle
import datetime as dt
from dataclasses import dataclass
from typing import List, Optional, Tuple, Dict

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

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
HEADER_HEIGHT = 35
ROW_HEIGHT = 25
DEFAULT_COLUMN_WIDTH = 80
MIN_COLUMN_WIDTH = 50
FONT_NAME = "나눔고딕"
FONT_SIZE = 10


@dataclass
class ColumnMap:
    date: str
    sku: str
    mall: str
    qty: str
    color: Optional[str] = None
    size: Optional[str] = None


# =========================
# 통합 분석기 클래스
# =========================
class SalesAnalyzer(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # 데이터 속성
        self.df_raw: Optional[pd.DataFrame] = None
        self.colmap: Optional[ColumnMap] = None
        self.df_filtered: Optional[pd.DataFrame] = None
        self.is_dummy: bool = False
        self.current_pivots: List[pd.DataFrame] = []
        
        # UI 설정
        self.title("판매데이터 분석 솔루션 v1.03")
        self.geometry("1400x900")
        self.minsize(1024, 768)
        
        self._setup_ui()
        self._init_data()

    # =========================
    # 데이터 처리 메서드
    # =========================
    
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
        """더미 데이터 생성"""
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

    def load_data(self, path: str) -> Tuple[bool, str]:
        """엑셀/피클 로드"""
        try:
            pickle_path = os.path.splitext(path)[0] + ".pkl"
            
            # Pickle 로드 시도
            if os.path.exists(pickle_path) and os.path.exists(path):
                if os.path.getmtime(pickle_path) >= os.path.getmtime(path):
                    try:
                        with open(pickle_path, 'rb') as f:
                            data = pickle.load(f)
                            self.df_raw = data['df']
                            self.colmap = data['colmap']
                            self.is_dummy = False
                            return True, f"캐시 로드 성공 ({len(self.df_raw):,}행)"
                    except Exception:
                        pass  # 피클 로드 실패 시 엑셀 로드 진행

            # 엑셀 로드
            if not os.path.exists(path):
                raise FileNotFoundError(f"파일 없음: {path}")
                
            df = pd.read_excel(path, engine="openpyxl")
            colmap = self._auto_map_columns(df)
            
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
            with open(pickle_path, 'wb') as f:
                pickle.dump({'df': self.df_raw, 'colmap': self.colmap}, f, protocol=pickle.HIGHEST_PROTOCOL)
                
            return True, f"로드 성공 ({len(self.df_raw):,}행)"
            
        except Exception as e:
            self.df_raw = self._make_dummy_data()
            self.colmap = self._auto_map_columns(self.df_raw)
            self.is_dummy = True
            return False, f"[더미 데이터] 오류: {e}"

    def filter_data(self, start_date: dt.date, end_date: dt.date, malls: List[str]):
        """데이터 필터링"""
        if self.df_raw is None: return
        
        mask = (self.df_raw[self.colmap.date] >= start_date) & (self.df_raw[self.colmap.date] <= end_date)
        if malls:
            mask &= self.df_raw[self.colmap.mall].isin(set(malls))
            
        self.df_filtered = self.df_raw.loc[mask].copy()

    def pivot_generic(self, index_col: str, columns_col: str, agg_func: str = "sum", 
                     df_source: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        """일반 피벗 생성"""
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
        
        # 날짜 컬럼 포맷팅
        new_cols = {}
        for c in pv.columns:
            if isinstance(c, (dt.date, dt.datetime)):
                new_cols[c] = f"{c.year}\n{c.month}/{c.day}"
        pv = pv.rename(columns=new_cols)
        
        # 합계 행 추가
        sum_row = {}
        for c in pv.columns:
            if c == index_col: sum_row[c] = "합계"
            elif pv[c].dtype in ['int64', 'float64', 'int32']:
                sum_row[c] = pv[c].sum()
            else:
                sum_row[c] = ""
        
        return pd.concat([pd.DataFrame([sum_row]), pv], ignore_index=True)

    def save_excel(self, dfs: Dict[str, pd.DataFrame], file_name: str) -> bool:
        """엑셀 저장"""
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
    # UI 구성 메서드
    # =========================
    
    def _setup_ui(self):
        """UI 초기화"""
        # 상단 컨트롤 프레임
        top_frame = ttk.Frame(self, padding="10")
        top_frame.pack(fill="x")
        
        # 첫 번째 행: 파일 로드
        file_frame = ttk.Frame(top_frame)
        file_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Button(file_frame, text="📂 데이터 파일 열기", 
                  command=self._cmd_load_file).pack(side="left", padx=(0, 10))
        
        self.lbl_file = ttk.Label(file_frame, text="파일 미선택 (더미 모드)", foreground="gray")
        self.lbl_file.pack(side="left")
        
        # 구분선
        ttk.Separator(top_frame, orient="horizontal").pack(fill="x", pady=5)
        
        # 두 번째 행: 기간 선택 및 쇼핑몰 필터
        control_frame = ttk.Frame(top_frame)
        control_frame.pack(fill="x")
        
        # 기간 선택
        date_frame = ttk.LabelFrame(control_frame, text="검색 기간", padding="5")
        date_frame.pack(side="left", padx=(0, 10))
        
        date_input = ttk.Frame(date_frame)
        date_input.pack()
        self.de_start = DateEntry(date_input, width=12, date_pattern='yyyy-mm-dd')
        self.de_start.pack(side="left", padx=2)
        ttk.Label(date_input, text="~").pack(side="left", padx=5)
        self.de_end = DateEntry(date_input, width=12, date_pattern='yyyy-mm-dd')
        self.de_end.pack(side="left", padx=2)
        
        # 기간 단축 버튼
        btn_frame = ttk.Frame(date_frame)
        btn_frame.pack(pady=(5, 0))
        periods = [("1주", 7), ("2주", 14), ("1개월", 30), ("3개월", 90), ("6개월", 180)]
        for text, days in periods:
            ttk.Button(btn_frame, text=text, width=6, 
                      command=lambda d=days: self._set_date_range(d)).pack(side="left", padx=2)
        
        # 쇼핑몰 선택
        mall_frame = ttk.LabelFrame(control_frame, text="쇼핑몰 필터", padding="5")
        mall_frame.pack(side="left", padx=(0, 10))
        self.listbox_mall = tk.Listbox(mall_frame, selectmode="extended", height=5, width=15)
        self.listbox_mall.pack()
        
        # 분석 실행 버튼
        action_frame = ttk.Frame(control_frame)
        action_frame.pack(side="left")
        ttk.Button(action_frame, text="▶ 분석 실행", 
                  command=self._cmd_run_analysis, width=12).pack(pady=5)
        
        # 중앙 탭뷰
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.tabs = {}
        tab_names = ["상품별 분석", "쇼핑몰별 분석", "일자별 추이", "연도별 추이"]
        for name in tab_names:
            frame = ttk.Frame(self.notebook)
            self.notebook.add(frame, text=name)
            self.tabs[name] = frame
        
        # 하단 상태 및 버튼
        bottom_frame = ttk.Frame(self, padding="8")
        bottom_frame.pack(fill="x")
        
        self.lbl_status = ttk.Label(bottom_frame, text="준비됨", font=(FONT_NAME, 9))
        self.lbl_status.pack(side="left", padx=10)
        
        # 오른쪽 버튼들
        ttk.Button(bottom_frame, text="❌ 종료", 
                  command=self._cmd_exit, width=10).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="💾 엑셀 저장", 
                  command=self._cmd_export, width=12).pack(side="right", padx=5)
        ttk.Button(bottom_frame, text="🔄 초기화", 
                  command=self._cmd_reset, width=10).pack(side="right", padx=5)

    def _init_data(self):
        """초기 데이터 로드"""
        self._set_date_range(7)
        success, msg = self.load_data(DEFAULT_EXCEL_PATH)
        self._on_load_complete(success, msg, DEFAULT_EXCEL_PATH)

    def _on_load_complete(self, success, msg, path):
        """데이터 로드 완료 처리"""
        self.lbl_status.config(text=msg)
        
        if success and not self.is_dummy:
            self.lbl_file.config(text=os.path.basename(path), foreground="blue")
        else:
            self.lbl_file.config(text="더미 데이터 (실제 파일 없음)", foreground="orange")
            
        self.listbox_mall.delete(0, "end")
        if self.df_raw is not None and self.colmap is not None:
            malls = sorted(self.df_raw[self.colmap.mall].unique().astype(str))
            for m in malls:
                self.listbox_mall.insert("end", m)
            self.listbox_mall.select_set(0, "end")
        
        self._run_analysis()

    def _set_date_range(self, days):
        """날짜 범위 설정"""
        ed = dt.date.today()
        sd = ed - dt.timedelta(days=days-1)
        self.de_start.set_date(sd)
        self.de_end.set_date(ed)

    # =========================
    # 명령 핸들러
    # =========================
    
    def _cmd_load_file(self):
        """파일 열기"""
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if f:
            success, msg = self.load_data(f)
            self._on_load_complete(success, msg, f)

    def _cmd_reset(self):
        """초기화"""
        self._set_date_range(7)
        self.listbox_mall.select_set(0, "end")
        self._cmd_run_analysis()
    
    def _cmd_exit(self):
        """프로그램 종료"""
        if messagebox.askokcancel("종료", "프로그램을 종료하시겠습니까?"):
            self.quit()
            self.destroy()

    def _cmd_export(self):
        """엑셀 저장"""
        if not self.current_pivots or len(self.current_pivots) != 4:
            messagebox.showwarning("경고", "분석 결과가 없습니다.")
            return
            
        f = filedialog.asksaveasfilename(
            defaultextension=".xlsx", 
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"판매분석_{dt.datetime.now().strftime('%Y%m%d')}.xlsx"
        )
        if f:
            dfs = {
                "상품별": self.current_pivots[0],
                "쇼핑몰별": self.current_pivots[1],
                "일자별": self.current_pivots[2],
                "연도별": self.current_pivots[3],
            }
            result = self.save_excel(dfs, f)
            if result:
                messagebox.showinfo("완료", "저장되었습니다.")
            else:
                messagebox.showerror("오류", "저장 실패")

    def _get_selected_malls(self):
        """선택된 쇼핑몰 목록"""
        idxs = self.listbox_mall.curselection()
        if not idxs: return []
        return [self.listbox_mall.get(i) for i in idxs]

    def _cmd_run_analysis(self):
        """분석 실행 (UI 버튼용)"""
        if self.colmap is None:
            self.lbl_status.config(text="데이터를 먼저 로드해주세요.")
            return
            
        start = self.de_start.get_date()
        end = self.de_end.get_date()
        
        if start > end:
            messagebox.showwarning("오류", "시작일이 종료일보다 큽니다.")
            return
        
        self._run_analysis()

    def _run_analysis(self):
        """분석 실행"""
        if self.colmap is None:
            return
            
        start = self.de_start.get_date()
        end = self.de_end.get_date()
        malls = self._get_selected_malls()
        
        try:
            # 데이터 필터링
            self.filter_data(start, end, malls)

            # 피벗 생성
            pv_prod = self.pivot_generic(self.colmap.sku, self.colmap.date)
            pv_mall_prod = self.pivot_generic(self.colmap.sku, self.colmap.mall)
            pv_date = self.pivot_generic(self.colmap.date, self.colmap.mall)
            
            # 연도별 추이 (전체 데이터 사용)
            df_all = self.df_raw.copy()
            if malls:
                df_all = df_all[df_all[self.colmap.mall].isin(set(malls))]
            df_all = df_all.dropna(subset=[self.colmap.date])
            df_all["년도"] = pd.to_datetime(df_all[self.colmap.date]).dt.year.astype(int)
            pv_year = self.pivot_generic("년도", self.colmap.mall, df_source=df_all)

            self.current_pivots = [pv_prod, pv_mall_prod, pv_date, pv_year]

            # 시트 렌더링
            self._render_sheet(self.tabs["상품별 분석"], pv_prod)
            self._render_sheet(self.tabs["쇼핑몰별 분석"], pv_mall_prod)
            self._render_sheet(self.tabs["일자별 추이"], pv_date)
            self._render_sheet(self.tabs["연도별 추이"], pv_year)
            
            count = len(self.df_filtered) if self.df_filtered is not None else 0
            self.lbl_status.config(text=f"분석 완료 ({count:,}건 처리)")
            
        except Exception as e:
            self.lbl_status.config(text=f"분석 오류: {e}")
            messagebox.showerror("분석 오류", f"분석 중 오류가 발생했습니다:\n{e}")

    def _render_sheet(self, parent_tab, df: pd.DataFrame):
        """시트 렌더링"""
        for widget in parent_tab.winfo_children():
            widget.destroy()
            
        if df.empty:
            ttk.Label(parent_tab, text="데이터가 없습니다.").pack(pady=20)
            return

        frame = ttk.Frame(parent_tab)
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

        sheet = Sheet(
            frame, 
            header_height=HEADER_HEIGHT,
            headers=list(df.columns),
            data=data,
            theme="light",
            font=(FONT_NAME, FONT_SIZE, "normal"),
            header_font=(FONT_NAME, FONT_SIZE, "bold")
        )
        
        # 시트 옵션 설정
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)
        
        # 정렬: 첫 번째 컬럼은 왼쪽, 나머지는 오른쪽
        try:
            sheet.align_columns(columns=[0], align="w", redraw=False)
            if len(df.columns) > 1:
                for c in range(1, len(df.columns)):
                    sheet.align_columns(columns=[c], align="e", redraw=False)
        except Exception:
            pass

        # 합계 행 강조 (첫 번째 행)
        try:
            sheet.highlight_rows(rows=[0], bg="#E3F2FD", fg="#1565C0")
        except Exception:
            pass


if __name__ == "__main__":
    app = SalesAnalyzer()
    app.mainloop()
