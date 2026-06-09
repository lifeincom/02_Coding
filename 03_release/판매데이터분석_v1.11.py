"""
========================
엑셀 판매데이터 분석 GUI (CustomTkinter + tksheet) v1.11
========================

[업데이트 내역 v1.11]
1. 코드 최적화: compute_pivots()로 피벗 생성 로직 통합 (sync/async 중복 제거)
2. 코드 최적화: _calc_yearly_pivot() 헬퍼 메서드 분리
3. 코드 최적화: 컬럼 정렬 2회 호출로 최적화 (N회 루프 → 2회)
4. UI 개선: 쇼핑몰 필터 CTkScrollableFrame + CTkCheckBox로 교체 (CTK 스타일 통일)
5. UI 개선: 전체선택/전체해제 버튼 추가
6. UI 개선: 요약 카드 패널 추가 (총 주문수량 / 분석 상품수 / 활성 쇼핑몰 / 분석 기간)
7. UI 개선: 기간 단축 버튼 "1년" 추가 (6개 버튼)
8. UI 개선: Enter 키로 분석 실행 단축키 바인딩
9. UI 개선: 상태바 타임스탬프 + 기간 정보 포함, 창 제목에 파일명 표시
10. 신규 기능: 인사이트 탭 (베스트셀러 Top5 / 채널별 점유율 / 기간비교 / 이상치탐지 / 주의상품)
11. 신규 기능: 기간 비교 (이번 기간 vs 이전 동일 기간, +/-% 표시)
12. 신규 기능: z-score 기반 이상치 탐지 (scipy 불필요, numpy/pandas만 사용)
13. 신규 기능: 엑셀 저장 시 인사이트 시트 포함 (최대 7시트)

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
VERSION = "1.11"
DEFAULT_EXCEL_PATH = r"\\NAS451\team451\DB\통합매출데이터.xlsx"
DATE_COL_CANDIDATES  = ["주문일자", "일자", "날짜", "order_date", "date"]
SKU_COL_CANDIDATES   = ["상품코드", "상품", "product_code", "sku"]
MALL_COL_CANDIDATES  = ["쇼핑몰", "몰명", "채널", "mall", "channel"]
QTY_COL_CANDIDATES   = ["수량", "판매수량", "qty", "quantity"]
COLOR_COL_CANDIDATES = ["컬러", "색상", "색깔", "color", "colour"]
SIZE_COL_CANDIDATES  = ["사이즈", "크기", "size"]

HEADER_HEIGHT        = 45
ROW_HEIGHT           = 28
DEFAULT_COLUMN_WIDTH = 80
MIN_COLUMN_WIDTH     = 20
FONT_NAME            = "나눔고딕"
FONT_SIZE            = 10

THEME_COLORS = {
    "header_row_bg": "#dbeafe",
    "header_row_fg": "#1e3a5f",
    "total_col_bg":  "#e0f2fe",
    "total_col_fg":  "#0c4a6e",
    "card_bg":       ("gray92", "gray20"),
    "up_color":      "#16a34a",
    "down_color":    "#dc2626",
}

ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


@dataclass
class ColumnMap:
    date:  str
    sku:   str
    mall:  str
    qty:   str
    color: Optional[str] = None
    size:  Optional[str] = None


# =========================
# 데이터 관리 클래스
# =========================
class SalesDataManager:
    def __init__(self):
        self.df_raw:      Optional[pd.DataFrame] = None
        self.colmap:      Optional[ColumnMap]    = None
        self.df_filtered: Optional[pd.DataFrame] = None
        self.is_dummy:    bool                   = False

    # ── 내부 유틸 ──────────────────────────────────────────────────
    def _auto_map_columns(self, df: pd.DataFrame) -> ColumnMap:
        columns_lower = {str(c).lower(): c for c in df.columns}

        def find(cands: List[str]) -> Optional[str]:
            for c in df.columns:
                if c in cands:
                    return c
            for cand in cands:
                if cand.lower() in columns_lower:
                    return columns_lower[cand.lower()]
            return None

        date_col  = find(DATE_COL_CANDIDATES)
        sku_col   = find(SKU_COL_CANDIDATES)
        mall_col  = find(MALL_COL_CANDIDATES)
        qty_col   = find(QTY_COL_CANDIDATES)
        color_col = find(COLOR_COL_CANDIDATES)
        size_col  = find(SIZE_COL_CANDIDATES)

        missing = [n for n, v in [("날짜", date_col), ("상품코드", sku_col),
                                   ("쇼핑몰", mall_col), ("수량", qty_col)] if v is None]
        if missing:
            raise ValueError(f"필수 컬럼 인식 실패: {', '.join(missing)}")

        return ColumnMap(date=date_col, sku=sku_col, mall=mall_col, qty=qty_col,
                         color=color_col, size=size_col)

    def _make_dummy_data(self) -> pd.DataFrame:
        import random
        rng   = pd.date_range(end=dt.date.today(), periods=90, freq="D")
        skus  = ["TPX-001", "TPX-002", "TPW-010", "TPW-020", "TBX-100"]
        malls = ["스마트스토어", "쿠팡", "11번가", "G마켓", "옥션"]
        colors = ["빨강", "파랑", "검정", "흰색"]
        sizes  = ["S", "M", "L", "XL"]
        rows = []
        random.seed(42)
        for d in rng:
            for sku in skus:
                for mall in malls:
                    if random.random() > 0.35:
                        continue
                    rows.append([d.date(), sku, mall,
                                  random.randint(1, 10),
                                  random.choice(colors),
                                  random.choice(sizes)])
        return pd.DataFrame(rows, columns=["주문일자", "상품코드", "쇼핑몰", "수량", "컬러", "사이즈"])

    # ── 데이터 로드 ────────────────────────────────────────────────
    def load_data(self, path: str,
                  callback: Optional[Callable[[str, int], None]] = None) -> Tuple[bool, str]:
        try:
            pickle_path = os.path.splitext(path)[0] + ".pkl"

            if os.path.exists(pickle_path) and os.path.exists(path):
                if os.path.getmtime(pickle_path) >= os.path.getmtime(path):
                    if callback: callback("캐시된 데이터 로딩 중...", 20)
                    try:
                        with open(pickle_path, "rb") as f:
                            data = pickle.load(f)
                            self.df_raw   = data["df"]
                            self.colmap   = data["colmap"]
                            self.is_dummy = False
                            if callback: callback("완료!", 100)
                            return True, f"캐시 로드 성공 ({len(self.df_raw):,}행)"
                    except Exception:
                        pass

            if callback: callback("엑셀 파일 읽는 중...", 10)
            if not os.path.exists(path):
                raise FileNotFoundError(f"파일 없음: {path}")

            df = pd.read_excel(path, engine="openpyxl")
            if callback: callback("컬럼 매핑 중...", 40)
            colmap = self._auto_map_columns(df)

            if callback: callback("데이터 전처리 중...", 60)
            if df[colmap.date].dtype == "object":
                df[colmap.date] = pd.to_datetime(df[colmap.date], errors="coerce").dt.date
            else:
                df[colmap.date] = pd.to_datetime(df[colmap.date]).dt.date

            df[colmap.qty] = pd.to_numeric(df[colmap.qty], errors="coerce").fillna(0).astype("int32")

            cols = [colmap.date, colmap.sku, colmap.mall, colmap.qty]
            if colmap.color: cols.append(colmap.color)
            if colmap.size:  cols.append(colmap.size)
            self.df_raw   = df[cols]
            self.colmap   = colmap
            self.is_dummy = False

            if callback: callback("캐시 생성 중...", 85)
            with open(pickle_path, "wb") as f:
                pickle.dump({"df": self.df_raw, "colmap": self.colmap}, f,
                             protocol=pickle.HIGHEST_PROTOCOL)

            if callback: callback("완료!", 100)
            return True, f"로드 성공 ({len(self.df_raw):,}행)"

        except Exception as e:
            if callback: callback("더미 데이터 생성 중...", 50)
            self.df_raw   = self._make_dummy_data()
            self.colmap   = self._auto_map_columns(self.df_raw)
            self.is_dummy = True
            if callback: callback("더미 데이터 로드 완료", 100)
            return False, f"[더미 데이터] 오류: {e}"

    # ── 필터 ──────────────────────────────────────────────────────
    def filter_data(self, start_date: dt.date, end_date: dt.date, malls: List[str]):
        if self.df_raw is None:
            return
        mask = ((self.df_raw[self.colmap.date] >= start_date) &
                (self.df_raw[self.colmap.date] <= end_date))
        if malls:
            mask &= self.df_raw[self.colmap.mall].isin(set(malls))
        self.df_filtered = self.df_raw.loc[mask]

    # ── 피벗 ──────────────────────────────────────────────────────
    def pivot_generic(self, index_col: str, columns_col: str,
                      agg_func: str = "sum",
                      df_source: Optional[pd.DataFrame] = None) -> pd.DataFrame:
        df = df_source if df_source is not None else self.df_filtered
        if df is None or df.empty:
            return pd.DataFrame()

        pv = pd.pivot_table(df, index=index_col, columns=columns_col,
                             values=self.colmap.qty, aggfunc=agg_func, fill_value=0)

        pv["합계"] = pv.sum(axis=1)
        if index_col == self.colmap.date:
            pv = pv.sort_index(ascending=True)
        elif index_col == "년도":
            pv = pv.sort_index(ascending=False)
        else:
            pv = pv.sort_values("합계", ascending=False)

        pv = pv[["합계"] + [c for c in pv.columns if c != "합계"]]
        pv = pv.reset_index()

        new_cols = {}
        for c in pv.columns:
            if isinstance(c, (dt.date, dt.datetime)):
                new_cols[c] = f"{c.year}\n{c.month}/{c.day}"
        pv = pv.rename(columns=new_cols)

        sum_row = {index_col: "합계"}
        numeric_cols = pv.columns[pv.dtypes.isin(["int64", "float64", "int32"])]
        for c in numeric_cols:
            sum_row[c] = pv[c].sum()

        return pd.concat([pd.DataFrame([sum_row]), pv], ignore_index=True)

    def _calc_yearly_pivot(self, malls: List[str]) -> pd.DataFrame:
        """전체 기간(df_raw) 기준 연도별 × 쇼핑몰 피벗"""
        if self.df_raw is None:
            return pd.DataFrame()
        df_all = self.df_raw
        if malls:
            df_all = df_all[df_all[self.colmap.mall].isin(set(malls))]
        df_all = df_all.dropna(subset=[self.colmap.date])
        df_all["년도"] = pd.to_datetime(df_all[self.colmap.date]).dt.year.astype(int)
        return self.pivot_generic("년도", self.colmap.mall, df_source=df_all)

    def compute_pivots(self, start: dt.date, end: dt.date, malls: List[str],
                       progress_cb: Optional[Callable[[str, int], None]] = None
                       ) -> Dict[str, pd.DataFrame]:
        """4개 피벗 테이블 일괄 생성 — sync/async 공통 진입점"""
        if progress_cb: progress_cb("데이터 필터링 중...", 20)
        self.filter_data(start, end, malls)

        if progress_cb: progress_cb("상품별 피벗 생성 중...", 40)
        pv_product = self.pivot_generic(self.colmap.sku, self.colmap.date)

        if progress_cb: progress_cb("쇼핑몰별 피벗 생성 중...", 55)
        pv_mall = self.pivot_generic(self.colmap.sku, self.colmap.mall)

        if progress_cb: progress_cb("일자별 피벗 생성 중...", 70)
        pv_daily = self.pivot_generic(self.colmap.date, self.colmap.mall)

        if progress_cb: progress_cb("연도별 피벗 생성 중...", 85)
        pv_yearly = self._calc_yearly_pivot(malls)

        return {"product": pv_product, "mall": pv_mall,
                "daily": pv_daily, "yearly": pv_yearly}

    # ── 인사이트 계산 ──────────────────────────────────────────────
    def compare_periods(self, current_start: dt.date, current_end: dt.date,
                        malls: List[str]) -> Dict[str, Any]:
        """이번 기간 vs 이전 동일 기간 비교"""
        delta      = (current_end - current_start).days + 1
        prev_end   = current_start - dt.timedelta(days=1)
        prev_start = prev_end - dt.timedelta(days=delta - 1)
        cm         = self.colmap

        def _sum(s: dt.date, e: dt.date) -> int:
            if self.df_raw is None:
                return 0
            mask = ((self.df_raw[cm.date] >= s) & (self.df_raw[cm.date] <= e))
            if malls:
                mask &= self.df_raw[cm.mall].isin(set(malls))
            return int(self.df_raw.loc[mask, cm.qty].sum())

        curr = _sum(current_start, current_end)
        prev = _sum(prev_start, prev_end)
        change_pct = round((curr - prev) / prev * 100, 1) if prev > 0 else None

        return {
            "current_period": f"{current_start} ~ {current_end}",
            "prev_period":    f"{prev_start} ~ {prev_end}",
            "current_total":  curr,
            "prev_total":     prev,
            "change_pct":     change_pct,
            "direction":      "up" if (change_pct or 0) > 0 else "down",
        }

    def detect_anomalies(self, df: pd.DataFrame,
                          z_threshold: float = 2.0) -> pd.DataFrame:
        """일별 수량의 z-score 기반 이상치 탐지 (scipy 불필요)"""
        empty = pd.DataFrame(columns=["일자", "수량", "z점수", "방향"])
        if df is None or df.empty:
            return empty

        daily = df.groupby(self.colmap.date)[self.colmap.qty].sum().reset_index()
        daily.columns = ["일자", "수량"]
        if len(daily) < 3:
            return empty

        mean = float(daily["수량"].mean())
        std  = float(daily["수량"].std())
        if std == 0:
            return empty

        daily["z점수"] = ((daily["수량"] - mean) / std).round(2)
        anom = daily[daily["z점수"].abs() > z_threshold].copy()
        anom["방향"] = anom["z점수"].apply(lambda z: "급증 ▲" if z > 0 else "급락 ▼")
        return anom.sort_values("z점수", key=abs, ascending=False).reset_index(drop=True)

    def compute_insights(self, start: dt.date, end: dt.date,
                          malls: List[str]) -> Dict[str, Any]:
        """인사이트 데이터 일괄 계산 (compute_pivots 이후 호출)"""
        df = self.df_filtered
        cm = self.colmap

        _empty: Dict[str, Any] = {
            "bestsellers":  pd.DataFrame(columns=["순위", "상품코드", "수량", "점유율(%)"]),
            "channel_share": pd.DataFrame(columns=["쇼핑몰", "수량", "점유율(%)"]),
            "period_compare": None,
            "anomalies":    pd.DataFrame(columns=["일자", "수량", "z점수", "방향"]),
            "zero_sales":   [],
        }
        if df is None or df.empty:
            _empty["period_compare"] = self.compare_periods(start, end, malls)
            return _empty

        total_qty = int(df[cm.qty].sum())

        sku_totals = df.groupby(cm.sku)[cm.qty].sum().sort_values(ascending=False)
        top5 = sku_totals.head(5).reset_index()
        top5.columns = ["상품코드", "수량"]
        top5.insert(0, "순위", range(1, len(top5) + 1))
        top5["점유율(%)"] = (top5["수량"] / total_qty * 100).round(1) if total_qty > 0 else 0.0

        mall_totals = df.groupby(cm.mall)[cm.qty].sum().sort_values(ascending=False)
        ch = mall_totals.reset_index()
        ch.columns = ["쇼핑몰", "수량"]
        ch["점유율(%)"] = (ch["수량"] / total_qty * 100).round(1) if total_qty > 0 else 0.0

        if self.df_raw is not None:
            zero_list = sorted(set(self.df_raw[cm.sku].unique()) - set(df[cm.sku].unique()))
        else:
            zero_list = []

        return {
            "bestsellers":   top5,
            "channel_share": ch,
            "period_compare": self.compare_periods(start, end, malls),
            "anomalies":     self.detect_anomalies(df),
            "zero_sales":    zero_list,
        }

    # ── 저장 ──────────────────────────────────────────────────────
    def save_excel(self, dfs: Dict[str, pd.DataFrame], file_name: str) -> bool:
        try:
            with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
                for sheet_name, df in dfs.items():
                    if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                        continue
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    writer.sheets[sheet_name].freeze_panes = "C3"
            return True
        except Exception as e:
            print(e)
            return False


# =========================
# 로딩 오버레이 위젯
# =========================
class LoadingOverlay(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color=("gray85", "gray20"), corner_radius=0, **kwargs)

        self.card = ctk.CTkFrame(self, corner_radius=15, fg_color=("white", "gray25"))
        self.card.place(relx=0.5, rely=0.5, anchor="center")

        self.spinner_label = ctk.CTkLabel(self.card, text="⏳", font=(FONT_NAME, 44))
        self.spinner_label.pack(pady=(30, 10))

        self.message_label = ctk.CTkLabel(self.card, text="처리 중...", font=(FONT_NAME, 13))
        self.message_label.pack(pady=(5, 10))

        self.progress = ctk.CTkProgressBar(self.card, width=250, height=12, corner_radius=6)
        self.progress.pack(pady=(10, 30), padx=40)
        self.progress.set(0)

        self._animation_running = False
        self._spinner_chars = ["⏳", "⌛"]
        self._spinner_index = 0

    def show(self, message: str = "처리 중..."):
        self.message_label.configure(text=message)
        self.progress.set(0)
        self.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.lift()
        self._start_animation()
        self.update()

    def hide(self):
        self._stop_animation()
        self.place_forget()

    def update_progress(self, message: str, percent: int):
        self.message_label.configure(text=message)
        self.progress.set(percent / 100)
        self.update()

    def _start_animation(self):
        self._animation_running = True
        self._animate_spinner()

    def _stop_animation(self):
        self._animation_running = False

    def _animate_spinner(self):
        if not self._animation_running:
            return
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_chars)
        self.spinner_label.configure(text=self._spinner_chars[self._spinner_index])
        self.after(500, self._animate_spinner)


# =========================
# 인사이트 패널 위젯
# =========================
class InsightPanel(ctk.CTkFrame):
    """인사이트 탭 전용 패널 (5개 카드 레이아웃)"""

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)

        scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        scroll.pack(fill="both", expand=True)

        # 상단 3 카드
        top_row = ctk.CTkFrame(scroll, fg_color="transparent")
        top_row.pack(fill="x", padx=4, pady=(6, 3))
        top_row.grid_columnconfigure((0, 1, 2), weight=1)

        self._frame_best   = self._make_card(top_row, "베스트셀러 Top 5")
        self._frame_share  = self._make_card(top_row, "채널별 점유율")
        self._frame_period = self._make_card(top_row, "기간 비교")
        self._frame_best.grid(row=0, column=0, padx=4, pady=2, sticky="nsew")
        self._frame_share.grid(row=0, column=1, padx=4, pady=2, sticky="nsew")
        self._frame_period.grid(row=0, column=2, padx=4, pady=2, sticky="nsew")

        # 하단 2 카드
        bot_row = ctk.CTkFrame(scroll, fg_color="transparent")
        bot_row.pack(fill="x", padx=4, pady=(3, 6))
        bot_row.grid_columnconfigure((0, 1), weight=1)

        self._frame_anomaly = self._make_card(bot_row, "이상치 탐지 (z-score ≥ 2.0)")
        self._frame_zero    = self._make_card(bot_row, "주의 상품 — 기간 내 판매 없음")
        self._frame_anomaly.grid(row=0, column=0, padx=4, pady=2, sticky="nsew")
        self._frame_zero.grid(row=0, column=1, padx=4, pady=2, sticky="nsew")

    @staticmethod
    def _make_card(parent: ctk.CTkFrame, title: str) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(parent, corner_radius=8, border_width=1,
                              border_color=("gray75", "gray45"),
                              fg_color=("white", "gray17"))
        ctk.CTkLabel(frame, text=title, font=(FONT_NAME, 12, "bold"),
                     anchor="w").pack(fill="x", padx=10, pady=(8, 3))
        ctk.CTkFrame(frame, height=1, fg_color=("gray82", "gray42")).pack(fill="x", padx=8, pady=(0, 6))
        return frame

    def _clear_card_body(self, frame: ctk.CTkFrame):
        """카드 헤더(제목 + 구분선) 이후 위젯만 제거"""
        for w in list(frame.winfo_children())[2:]:
            w.destroy()

    # ── 렌더 ──────────────────────────────────────────────────────
    def render(self, insights: Dict[str, Any], start: dt.date, end: dt.date):
        self._render_bestsellers(self._frame_best,   insights.get("bestsellers"))
        self._render_channel_share(self._frame_share, insights.get("channel_share"))
        self._render_period_compare(self._frame_period, insights.get("period_compare"))
        self._render_anomalies(self._frame_anomaly,  insights.get("anomalies"))
        self._render_zero_sales(self._frame_zero,    insights.get("zero_sales", []))

    def _render_bestsellers(self, frame: ctk.CTkFrame, df: Optional[pd.DataFrame]):
        self._clear_card_body(frame)
        if df is None or df.empty:
            ctk.CTkLabel(frame, text="데이터 없음", text_color="gray").pack(pady=12)
            return

        data = [
            [str(int(r["순위"])), str(r["상품코드"]),
             f"{int(r['수량']):,}", f"{r['점유율(%)']:.1f}%"]
            for _, r in df.iterrows()
        ]
        sheet = Sheet(frame, headers=["순위", "상품코드", "수량", "점유율(%)"],
                      data=data, theme="light", height=165,
                      font=(FONT_NAME, FONT_SIZE, "normal"),
                      header_font=(FONT_NAME, FONT_SIZE, "bold"))
        try:
            sheet.set_options(row_height=ROW_HEIGHT, header_height=35)
            sheet.align_columns(columns=[0], align="c", redraw=False)
            sheet.align_columns(columns=[1], align="w", redraw=False)
            sheet.align_columns(columns=[2, 3], align="e", redraw=True)
        except Exception:
            pass
        sheet.enable_bindings("copy")
        sheet.pack(fill="x", padx=8, pady=(0, 8))

    def _render_channel_share(self, frame: ctk.CTkFrame, df: Optional[pd.DataFrame]):
        self._clear_card_body(frame)
        if df is None or df.empty:
            ctk.CTkLabel(frame, text="데이터 없음", text_color="gray").pack(pady=12)
            return

        BAR = 18
        lines = []
        for _, r in df.iterrows():
            pct    = float(r["점유율(%)"])
            filled = int(pct / 100 * BAR)
            bar    = "█" * filled + "░" * (BAR - filled)
            mall   = str(r["쇼핑몰"])[:9].ljust(9)
            lines.append(f"{mall} {bar} {pct:5.1f}% ({int(r['수량']):,})")

        ctk.CTkLabel(frame, text="\n".join(lines),
                     font=("Consolas", 10), justify="left", anchor="w"
                     ).pack(fill="x", padx=12, pady=(0, 10))

    def _render_period_compare(self, frame: ctk.CTkFrame, data: Optional[Dict]):
        self._clear_card_body(frame)
        if data is None:
            ctk.CTkLabel(frame, text="데이터 없음", text_color="gray").pack(pady=12)
            return

        ctk.CTkLabel(frame, text=f"이번  {data['current_period']}",
                     font=(FONT_NAME, 10), text_color="gray50",
                     anchor="w").pack(fill="x", padx=14)
        ctk.CTkLabel(frame, text=f"{data['current_total']:,} 개",
                     font=(FONT_NAME, 24, "bold"),
                     anchor="w").pack(fill="x", padx=14, pady=(0, 4))

        if data["change_pct"] is not None:
            pct   = data["change_pct"]
            arrow = "▲" if data["direction"] == "up" else "▼"
            color = THEME_COLORS["up_color"] if data["direction"] == "up" else THEME_COLORS["down_color"]
            sign  = "+" if pct > 0 else ""
            ctk.CTkLabel(frame, text=f"{arrow} {sign}{pct}%  전기간 대비",
                         font=(FONT_NAME, 13, "bold"), text_color=color,
                         anchor="w").pack(fill="x", padx=14)
        else:
            ctk.CTkLabel(frame, text="이전 기간 데이터 없음",
                         font=(FONT_NAME, 11), text_color="gray",
                         anchor="w").pack(fill="x", padx=14)

        ctk.CTkLabel(frame, text=f"\n이전  {data['prev_period']}",
                     font=(FONT_NAME, 10), text_color="gray50",
                     anchor="w").pack(fill="x", padx=14)
        ctk.CTkLabel(frame, text=f"{data['prev_total']:,} 개",
                     font=(FONT_NAME, 16), text_color="gray60",
                     anchor="w").pack(fill="x", padx=14, pady=(0, 10))

    def _render_anomalies(self, frame: ctk.CTkFrame, df: Optional[pd.DataFrame]):
        self._clear_card_body(frame)
        if df is None or df.empty:
            ctk.CTkLabel(frame, text="이상치 없음 ✓",
                         font=(FONT_NAME, 12), text_color="#16a34a").pack(pady=14)
            return

        data = [
            [str(r["일자"]), f"{int(r['수량']):,}", f"{r['z점수']:.2f}", str(r["방향"])]
            for _, r in df.iterrows()
        ]
        sheet = Sheet(frame, headers=["일자", "수량", "z점수", "방향"],
                      data=data, theme="light", height=165,
                      font=(FONT_NAME, FONT_SIZE, "normal"),
                      header_font=(FONT_NAME, FONT_SIZE, "bold"))
        try:
            sheet.set_options(row_height=ROW_HEIGHT, header_height=35)
            sheet.align_columns(columns=[0, 3], align="c", redraw=False)
            sheet.align_columns(columns=[1, 2], align="e", redraw=True)
        except Exception:
            pass
        sheet.enable_bindings("copy")
        sheet.pack(fill="x", padx=8, pady=(0, 8))

    def _render_zero_sales(self, frame: ctk.CTkFrame, items: List[str]):
        self._clear_card_body(frame)
        if not items:
            ctk.CTkLabel(frame, text="모든 상품 판매 중 ✓",
                         font=(FONT_NAME, 12), text_color="#16a34a").pack(pady=14)
            return

        ctk.CTkLabel(frame, text=f"총 {len(items)}개 상품 미판매",
                     font=(FONT_NAME, 10), text_color="gray",
                     anchor="w").pack(fill="x", padx=10)
        tb = ctk.CTkTextbox(frame, height=145, font=("Consolas", 10))
        tb.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        tb.insert("0.0", "\n".join(items))
        tb.configure(state="disabled")


# =========================
# 메인 UI 클래스
# =========================
class SalesAnalyzerUI(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.data_mgr          = SalesDataManager()
        self._mall_vars:       Dict[str, ctk.BooleanVar] = {}
        self._current_insights: Optional[Dict]           = None

        self.title(f"판매데이터 분석 v{VERSION}")
        self.geometry("1460x940")
        self.minsize(1100, 780)

        self._setup_layout()
        self._init_data()
        self.bind("<Return>", lambda e: self._cmd_run_analysis())

    # ── 레이아웃 구성 ──────────────────────────────────────────────
    def _setup_layout(self):
        # 1. 상단 컨트롤 패널
        self.top_frame = ctk.CTkFrame(self, corner_radius=10)
        self.top_frame.pack(fill="x", padx=10, pady=10)
        self._setup_controls(self.top_frame)

        # 2. 요약 카드 패널
        self._setup_summary_bar()

        # 3. 메인 컨테이너
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.pack(fill="both", expand=True, padx=10)

        # 4. 탭뷰
        self.tabview = ctk.CTkTabview(self.main_container)
        self.tabview.pack(fill="both", expand=True)

        self.tabs = {}
        for name in ["상품별 분석", "쇼핑몰별 분석", "일자별 추이", "연도별 추이", "인사이트"]:
            self.tabview.add(name)
            self.tabs[name] = self.tabview.tab(name)

        self.insight_panel = InsightPanel(self.tabs["인사이트"])
        self.insight_panel.pack(fill="both", expand=True)

        # 5. 로딩 오버레이
        self.loading_overlay = LoadingOverlay(self.main_container)

        # 6. 하단 상태바
        self.bottom_frame = ctk.CTkFrame(self, height=40, fg_color="transparent")
        self.bottom_frame.pack(fill="x", padx=10, pady=5)
        self.lbl_status = ctk.CTkLabel(self.bottom_frame, text="준비됨", text_color="gray")
        self.lbl_status.pack(side="left")
        ctk.CTkButton(self.bottom_frame, text="엑셀 저장", command=self._cmd_export,
                      width=100).pack(side="right")
        ctk.CTkButton(self.bottom_frame, text="초기화", command=self._cmd_reset,
                      width=100, fg_color="gray").pack(side="right", padx=5)

    def _setup_controls(self, parent: ctk.CTkFrame):
        parent.grid_columnconfigure(9, weight=1)

        # 파일 열기
        self.btn_load = ctk.CTkButton(parent, text="📂 데이터 열기",
                                       command=self._cmd_load_file, width=130)
        self.btn_load.grid(row=0, column=0, padx=10, pady=12)

        self.lbl_file = ctk.CTkLabel(parent, text="파일 미선택 (더미 모드)")
        self.lbl_file.grid(row=0, column=1, sticky="w", padx=5)

        # 기간 날짜 입력
        date_frame = ctk.CTkFrame(parent, fg_color="transparent")
        date_frame.grid(row=0, column=2, padx=10)
        ctk.CTkLabel(date_frame, text="기간:").pack(side="left", padx=5)
        self.de_start = DateEntry(date_frame, width=12, date_pattern="yyyy-mm-dd")
        self.de_start.pack(side="left")
        ctk.CTkLabel(date_frame, text="~").pack(side="left", padx=5)
        self.de_end = DateEntry(date_frame, width=12, date_pattern="yyyy-mm-dd")
        self.de_end.pack(side="left")

        # 기간 단축 버튼 (6개)
        btn_frame = ctk.CTkFrame(parent, fg_color="transparent")
        btn_frame.grid(row=0, column=3)
        for text, days in [("1주", 7), ("2주", 14), ("1개월", 30),
                            ("3개월", 90), ("6개월", 180), ("1년", 365)]:
            ctk.CTkButton(btn_frame, text=text, width=44, height=24,
                           command=lambda d=days: self._set_date_range(d)
                           ).pack(side="left", padx=2)

        # 쇼핑몰 필터 (CTkScrollableFrame + CTkCheckBox)
        mall_outer = ctk.CTkFrame(parent, corner_radius=6, border_width=1,
                                   border_color=("gray72", "gray50"))
        mall_outer.grid(row=0, column=4, padx=12, pady=5)

        mall_hdr = ctk.CTkFrame(mall_outer, fg_color="transparent")
        mall_hdr.pack(fill="x", padx=6, pady=(5, 2))
        ctk.CTkLabel(mall_hdr, text="쇼핑몰", font=(FONT_NAME, 10, "bold")).pack(side="left")
        ctk.CTkButton(mall_hdr, text="해제", width=40, height=20, font=(FONT_NAME, 9),
                       fg_color="gray", command=self._deselect_all_malls).pack(side="right")
        ctk.CTkButton(mall_hdr, text="전체", width=40, height=20, font=(FONT_NAME, 9),
                       command=self._select_all_malls).pack(side="right", padx=2)

        self._mall_scroll = ctk.CTkScrollableFrame(mall_outer, width=165, height=72,
                                                     fg_color="transparent")
        self._mall_scroll.pack(fill="x", padx=4, pady=(0, 5))

        # 분석 실행
        ctk.CTkButton(parent, text="▶ 분석 실행", command=self._cmd_run_analysis,
                       font=(FONT_NAME, 13, "bold"), width=120, height=50
                       ).grid(row=0, column=5, padx=15)

    def _setup_summary_bar(self):
        self.summary_frame = ctk.CTkFrame(self, height=68, corner_radius=8)
        self.summary_frame.pack(fill="x", padx=10, pady=(0, 6))
        self.summary_frame.pack_propagate(False)

        for title, attr in [("총 주문수량", "lbl_sum_qty"),
                             ("분석 상품수",  "lbl_sum_sku"),
                             ("활성 쇼핑몰",  "lbl_sum_mall"),
                             ("분석 기간",    "lbl_sum_period")]:
            card = ctk.CTkFrame(self.summary_frame, corner_radius=6,
                                 fg_color=THEME_COLORS["card_bg"])
            card.pack(side="left", expand=True, fill="both", padx=4, pady=5)
            ctk.CTkLabel(card, text=title, font=(FONT_NAME, 9),
                          text_color="gray50").pack(pady=(5, 0))
            lbl = ctk.CTkLabel(card, text="—", font=(FONT_NAME, 16, "bold"))
            lbl.pack(pady=(0, 4))
            setattr(self, attr, lbl)

    def _update_summary_bar(self, df: Optional[pd.DataFrame],
                             start: dt.date, end: dt.date):
        if df is None or df.empty:
            for a in ("lbl_sum_qty", "lbl_sum_sku", "lbl_sum_mall", "lbl_sum_period"):
                getattr(self, a).configure(text="—")
            return
        cm = self.data_mgr.colmap
        self.lbl_sum_qty.configure(text=f"{int(df[cm.qty].sum()):,}")
        self.lbl_sum_sku.configure(text=f"{df[cm.sku].nunique()}종")
        self.lbl_sum_mall.configure(text=f"{df[cm.mall].nunique()}개")
        self.lbl_sum_period.configure(text=f"{(end - start).days + 1}일")

    # ── 쇼핑몰 필터 헬퍼 ──────────────────────────────────────────
    def _rebuild_mall_filter(self, malls: List[str]):
        prev = {n: v.get() for n, v in self._mall_vars.items()}
        for w in self._mall_scroll.winfo_children():
            w.destroy()
        self._mall_vars.clear()
        for mall in malls:
            var = ctk.BooleanVar(value=prev.get(mall, True))
            self._mall_vars[mall] = var
            ctk.CTkCheckBox(self._mall_scroll, text=mall, variable=var,
                             font=(FONT_NAME, 10)).pack(anchor="w", pady=1, padx=4)

    def _select_all_malls(self):
        for v in self._mall_vars.values(): v.set(True)

    def _deselect_all_malls(self):
        for v in self._mall_vars.values(): v.set(False)

    def _get_selected_malls(self) -> List[str]:
        return [n for n, v in self._mall_vars.items() if v.get()]

    # ── 초기화 & 로드 ──────────────────────────────────────────────
    def _init_data(self):
        self._set_date_range(7)
        self.loading_overlay.show("데이터 로딩 중...")
        threading.Thread(target=self._async_load,
                          args=(DEFAULT_EXCEL_PATH,), daemon=True).start()

    def _async_load(self, path: str):
        def cb(msg, pct):
            self.after(0, lambda: self.loading_overlay.update_progress(msg, pct))
        success, msg = self.data_mgr.load_data(path, cb)
        self.after(0, lambda: self._on_load_complete(success, msg, path))

    def _on_load_complete(self, success: bool, msg: str, path: str):
        self.loading_overlay.hide()
        self.lbl_status.configure(text=msg)

        if success and not self.data_mgr.is_dummy:
            fname = os.path.basename(path)
            self.lbl_file.configure(text=fname, text_color="#4CC2FF")
            self.title(f"판매데이터 분석 v{VERSION}  |  {fname}")
        else:
            self.lbl_file.configure(text="더미 데이터 (실제 파일 없음)", text_color="orange")
            self.title(f"판매데이터 분석 v{VERSION}  |  더미 데이터")

        if self.data_mgr.df_raw is not None and self.data_mgr.colmap is not None:
            malls = sorted(self.data_mgr.df_raw[self.data_mgr.colmap.mall].unique().astype(str))
            self._rebuild_mall_filter(malls)

        self._run_analysis_sync()

    # ── 날짜 ──────────────────────────────────────────────────────
    def _set_date_range(self, days: int):
        ed = dt.date.today()
        self.de_start.set_date(ed - dt.timedelta(days=days - 1))
        self.de_end.set_date(ed)

    # ── 커맨드 ────────────────────────────────────────────────────
    def _cmd_load_file(self):
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if f:
            self.loading_overlay.show("파일 로딩 중...")
            threading.Thread(target=self._async_load, args=(f,), daemon=True).start()

    def _cmd_reset(self):
        self._set_date_range(7)
        self._select_all_malls()
        self._cmd_run_analysis()

    def _cmd_export(self):
        if not hasattr(self, "current_pivots") or not self.current_pivots:
            messagebox.showwarning("경고", "분석 결과가 없습니다.")
            return
        f = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"판매분석_{dt.datetime.now().strftime('%Y%m%d')}.xlsx",
        )
        if not f:
            return
        self.loading_overlay.show("엑셀 저장 중...")

        def _task():
            dfs: Dict[str, pd.DataFrame] = {
                "상품별":   self.current_pivots.get("product",  pd.DataFrame()),
                "쇼핑몰별": self.current_pivots.get("mall",     pd.DataFrame()),
                "일자별":   self.current_pivots.get("daily",    pd.DataFrame()),
                "연도별":   self.current_pivots.get("yearly",   pd.DataFrame()),
            }
            if self._current_insights:
                ins = self._current_insights
                for key, sname in [("bestsellers", "베스트셀러"),
                                    ("channel_share", "채널점유율"),
                                    ("anomalies", "이상치")]:
                    df_ins = ins.get(key)
                    if df_ins is not None and not df_ins.empty:
                        dfs[sname] = df_ins
            ok = self.data_mgr.save_excel(dfs, f)
            self.after(0, lambda: self._on_save_complete(ok))

        threading.Thread(target=_task, daemon=True).start()

    def _on_save_complete(self, success: bool):
        self.loading_overlay.hide()
        if success:
            messagebox.showinfo("완료", "저장되었습니다.")
        else:
            messagebox.showerror("오류", "저장 실패")

    # ── 분석 실행 ──────────────────────────────────────────────────
    def _cmd_run_analysis(self):
        if self.data_mgr.colmap is None:
            self.lbl_status.configure(text="데이터를 먼저 로드해주세요.")
            return
        start = self.de_start.get_date()
        end   = self.de_end.get_date()
        if start > end:
            messagebox.showwarning("오류", "시작일이 종료일보다 큽니다.")
            return
        self.loading_overlay.show("분석 처리 중...")
        threading.Thread(target=self._async_analysis, daemon=True).start()

    def _async_analysis(self):
        try:
            start  = self.de_start.get_date()
            end    = self.de_end.get_date()
            malls  = self._get_selected_malls()

            def cb(msg, pct):
                self.after(0, lambda: self.loading_overlay.update_progress(msg, pct))

            pivots   = self.data_mgr.compute_pivots(start, end, malls, progress_cb=cb)
            cb("인사이트 분석 중...", 93)
            insights = self.data_mgr.compute_insights(start, end, malls)

            self.after(0, lambda: self._on_analysis_complete(pivots, insights, start, end))
        except Exception as e:
            self.after(0, lambda: self._on_analysis_error(str(e)))

    def _on_analysis_complete(self, pivots: Dict[str, pd.DataFrame],
                               insights: Dict[str, Any],
                               start: dt.date, end: dt.date):
        self.current_pivots     = pivots
        self._current_insights  = insights

        self.loading_overlay.update_progress("시트 렌더링 중...", 97)

        self._render_sheet(self.tabs["상품별 분석"],  pivots["product"])
        self._render_sheet(self.tabs["쇼핑몰별 분석"], pivots["mall"])
        self._render_sheet(self.tabs["일자별 추이"],   pivots["daily"])
        self._render_sheet(self.tabs["연도별 추이"],   pivots["yearly"])
        self.insight_panel.render(insights, start, end)

        self.loading_overlay.hide()

        count = len(self.data_mgr.df_filtered) if self.data_mgr.df_filtered is not None else 0
        ts = dt.datetime.now().strftime("%H:%M:%S")
        self.lbl_status.configure(
            text=f"[{ts}] 분석 완료 — {count:,}건 처리  |  {start} ~ {end}")
        self._update_summary_bar(self.data_mgr.df_filtered, start, end)

    def _on_analysis_error(self, error_msg: str):
        self.loading_overlay.hide()
        self.lbl_status.configure(text=f"분석 오류: {error_msg}")
        messagebox.showerror("분석 오류", f"분석 중 오류가 발생했습니다:\n{error_msg}")

    def _run_analysis_sync(self):
        """초기 로드 후 동기 분석 (로딩 오버레이 없음)"""
        if self.data_mgr.colmap is None:
            return
        start    = self.de_start.get_date()
        end      = self.de_end.get_date()
        malls    = self._get_selected_malls()
        pivots   = self.data_mgr.compute_pivots(start, end, malls)
        insights = self.data_mgr.compute_insights(start, end, malls)
        self._on_analysis_complete(pivots, insights, start, end)

    # ── 시트 렌더링 ────────────────────────────────────────────────
    def _render_sheet(self, parent_tab: ctk.CTkFrame, df: pd.DataFrame):
        for w in parent_tab.winfo_children():
            w.destroy()
        if df.empty:
            ctk.CTkLabel(parent_tab, text="데이터가 없습니다.").pack(pady=20)
            return

        frame = ctk.CTkFrame(parent_tab)
        frame.pack(fill="both", expand=True)

        def format_value(val):
            if pd.isna(val):
                return ""
            if isinstance(val, (int, float)):
                return f"{int(val):,}" if val == int(val) else f"{val:,.2f}"
            return str(val)

        data = df.apply(lambda row: row.apply(format_value)).values.tolist()

        sheet = Sheet(frame,
                       header_height=HEADER_HEIGHT,
                       headers=list(df.columns),
                       data=data,
                       theme="light",
                       font=(FONT_NAME, FONT_SIZE, "normal"),
                       header_font=(FONT_NAME, FONT_SIZE, "bold"))

        try:
            sheet.set_options(header_height=HEADER_HEIGHT, row_height=ROW_HEIGHT,
                               default_column_width=DEFAULT_COLUMN_WIDTH,
                               min_column_width=MIN_COLUMN_WIDTH, header_wrap=True)
        except Exception:
            pass

        sheet.table_align(align="left")
        sheet.set_all_column_widths(width=None)
        sheet.enable_bindings()
        sheet.pack(fill="both", expand=True)

        try:
            n = len(df.columns)
            sheet.align_columns(columns=[0], align="w", redraw=False)
            if n > 1:
                sheet.align_columns(columns=list(range(1, n)), align="e", redraw=True)
        except Exception:
            pass

        try:
            sheet.highlight_rows(rows=[0],
                                   bg=THEME_COLORS["header_row_bg"],
                                   fg=THEME_COLORS["header_row_fg"])
        except Exception:
            pass

        try:
            if len(df.columns) > 0 and "합계" in str(df.columns[0]):
                sheet.highlight_columns(columns=[0],
                                         bg=THEME_COLORS["total_col_bg"],
                                         fg=THEME_COLORS["total_col_fg"])
        except Exception:
            pass


if __name__ == "__main__":
    app = SalesAnalyzerUI()
    app.mainloop()
