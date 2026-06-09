"""
===========================================
분석 모듈 (analyzer.py)
===========================================

판매 데이터를 피벗 테이블로 분석하는 모듈입니다.
기존 Tkinter 애플리케이션의 피벗 로직을 웹 서비스용으로 이식했습니다.

주요 기능:
1. 4가지 피벗 테이블 생성
   - 상품별 분석 (상품코드 × 일자)
   - 쇼핑몰별 분석 (상품코드 × 쇼핑몰)
   - 일자별 추이 (일자 × 쇼핑몰)
   - 연도별 추이 (년도 × 쇼핑몰)
2. 페이지네이션 지원
3. JSON 직렬화 가능한 형태로 반환

작성자: AI Assistant
작성일: 2026-02-09
"""

import datetime as dt
from typing import Dict, List, Optional, Any
import pandas as pd
import numpy as np

from data_loader import DataLoader, ColumnMap


class SalesAnalyzer:
    """
    판매 데이터 분석 클래스
    
    피벗 테이블 생성, 페이지네이션, JSON 변환 등의 기능을 제공합니다.
    """
    
    def __init__(self, data_loader: DataLoader):
        """
        분석기 초기화
        
        Args:
            data_loader: 데이터 로더 인스턴스
        """
        self.loader = data_loader
    
    def pivot_generic(
        self,
        df: pd.DataFrame,
        index_col: str,
        columns_col: str,
        qty_col: str,
        agg_func: str = "sum"
    ) -> pd.DataFrame:
        """
        일반 피벗 테이블을 생성합니다.
        
        프로세스:
        1. 피벗 테이블 생성 (index × columns)
        2. 합계 컬럼 추가
        3. 정렬 (index가 날짜면 오름차순, 년도면 내림차순, 나머지는 합계 기준)
        4. 합계 컬럼을 첫 번째로 이동
        5. 날짜 컬럼 포맷팅
        6. 합계 행 추가
        
        Args:
            df: 원본 DataFrame
            index_col: 행 인덱스로 사용할 컬럼
            columns_col: 열로 사용할 컬럼
            qty_col: 집계할 수량 컬럼
            agg_func: 집계 함수 (기본: sum)
            
        Returns:
            피벗 테이블 DataFrame
        """
        if df is None or df.empty:
            return pd.DataFrame()
        
        # ======================================
        # 1단계: 피벗 테이블 생성
        # ======================================
        pv = pd.pivot_table(
            df,
            index=index_col,
            columns=columns_col,
            values=qty_col,
            aggfunc=agg_func,
            fill_value=0
        )
        
        # ======================================
        # 2단계: 합계 컬럼 추가
        # ======================================
        pv["합계"] = pv.sum(axis=1)
        
        # ======================================
        # 3단계: 정렬
        # ======================================
        if index_col in ["주문일자", "일자", "날짜", "order_date", "date"]:
            # 날짜는 오름차순
            pv = pv.sort_index(ascending=True)
        elif index_col == "년도":
            # 년도는 내림차순 (최신 먼저)
            pv = pv.sort_index(ascending=False)
        else:
            # 나머지는 합계 기준 내림차순
            pv = pv.sort_values("합계", ascending=False)
        
        # ======================================
        # 4단계: 합계 컬럼을 첫 번째로 이동
        # ======================================
        cols = ["합계"] + [c for c in pv.columns if c != "합계"]
        pv = pv[cols]
        
        # 인덱스를 컬럼으로 변환
        pv = pv.reset_index()
        
        # ======================================
        # 5단계: 날짜 데이터 및 컬럼 포맷팅
        # ======================================
        # 인덱스 컬럼(첫 번째 컬럼)이 날짜인 경우 문자열로 변환
        first_col = pv.columns[0]
        if first_col == index_col and len(pv) > 0:
            first_val = pv[first_col].iloc[0]
            if isinstance(first_val, (dt.date, dt.datetime)):
                # 날짜 데이터를 YYYY/MM/DD 문자열로 변환
                pv[first_col] = pv[first_col].apply(
                    lambda x: f"{x.year}/{str(x.month).zfill(2)}/{str(x.day).zfill(2)}" 
                    if isinstance(x, (dt.date, dt.datetime)) else x
                )
        
        # 컬럼 헤더가 날짜인 경우 이름 변경
        new_cols = {}
        for c in pv.columns:
            if isinstance(c, (dt.date, dt.datetime)):
                # 날짜 컬럼은 "YYYY/MM/DD" 형식으로
                new_cols[c] = f"{c.year}/{str(c.month).zfill(2)}/{str(c.day).zfill(2)}"
        
        if new_cols:
            pv = pv.rename(columns=new_cols)
        
        # ======================================
        # 6단계: 합계 행 추가
        # ======================================
        sum_row = {}
        for c in pv.columns:
            if c == index_col:
                sum_row[c] = "합계"
            elif pv[c].dtype in ['int64', 'float64', 'int32', 'float32']:
                sum_row[c] = pv[c].sum()
            else:
                sum_row[c] = ""
        
        # 합계 행을 맨 위에 추가
        pv = pd.concat([pd.DataFrame([sum_row]), pv], ignore_index=True)
        
        return pv
    
    def analyze_all(
        self,
        start_date: dt.date,
        end_date: dt.date,
        malls: Optional[List[str]] = None
    ) -> Dict[str, pd.DataFrame]:
        """
        4가지 피벗 테이블을 모두 생성합니다.
        
        Args:
            start_date: 시작일
            end_date: 종료일
            malls: 쇼핑몰 필터 (None이면 전체)
            
        Returns:
            피벗 테이블 딕셔너리
            {
                "product_daily": 상품별 일자 분석,
                "product_mall": 상품별 쇼핑몰 분석,
                "daily_mall": 일자별 쇼핑몰 분석,
                "yearly_mall": 연도별 쇼핑몰 분석
            }
        """
        if self.loader.colmap is None:
            return {}
        
        colmap = self.loader.colmap
        
        # ======================================
        # 필터링된 데이터 가져오기
        # ======================================
        df_filtered = self.loader.filter_data(start_date, end_date, malls)
        
        if df_filtered.empty:
            return {
                "product_daily": pd.DataFrame(),
                "product_mall": pd.DataFrame(),
                "daily_mall": pd.DataFrame(),
                "yearly_mall": pd.DataFrame()
            }
        
        # ======================================
        # 1. 상품별 일자 분석 (상품코드 × 일자)
        # ======================================
        pv_product_daily = self.pivot_generic(
            df_filtered,
            index_col=colmap.sku,
            columns_col=colmap.date,
            qty_col=colmap.qty
        )
        
        # ======================================
        # 2. 상품별 쇼핑몰 분석 (상품코드 × 쇼핑몰)
        # ======================================
        pv_product_mall = self.pivot_generic(
            df_filtered,
            index_col=colmap.sku,
            columns_col=colmap.mall,
            qty_col=colmap.qty
        )
        
        # ======================================
        # 3. 일자별 쇼핑몰 분석 (일자 × 쇼핑몰)
        # ======================================
        pv_daily_mall = self.pivot_generic(
            df_filtered,
            index_col=colmap.date,
            columns_col=colmap.mall,
            qty_col=colmap.qty
        )
        
        # ======================================
        # 4. 연도별 쇼핑몰 분석 (년도 × 쇼핑몰)
        # ======================================
        # 전체 원본 데이터 사용 (날짜 필터 무시, 쇼핑몰 필터만 적용)
        df_all = self.loader.df_raw.copy()
        if malls:
            df_all = df_all[df_all[colmap.mall].isin(set(malls))]
        
        # NaN 날짜 제거
        df_all = df_all.dropna(subset=[colmap.date])
        
        # 년도 컬럼 추가
        df_all["년도"] = pd.to_datetime(df_all[colmap.date]).dt.year.astype(int)
        
        pv_yearly_mall = self.pivot_generic(
            df_all,
            index_col="년도",
            columns_col=colmap.mall,
            qty_col=colmap.qty
        )
        
        return {
            "product_daily": pv_product_daily,
            "product_mall": pv_product_mall,
            "daily_mall": pv_daily_mall,
            "yearly_mall": pv_yearly_mall
        }
    
    def get_summary_stats(
        self,
        start_date: dt.date,
        end_date: dt.date,
        malls: Optional[List[str]] = None
    ) -> Dict[str, Any]:
        """
        요약 통계를 반환합니다.
        
        대시보드 상단 카드에 표시할 주요 지표를 계산합니다.
        
        Args:
            start_date: 시작일
            end_date: 종료일
            malls: 쇼핑몰 필터
            
        Returns:
            요약 통계 딕셔너리
            {
                "total_quantity": 총 판매 수량,
                "total_products": 총 상품 수,
                "total_orders": 총 주문 건수,
                "top_product": 상위 상품,
                "top_mall": 상위 쇼핑몰,
                "period": 기간 정보
            }
        """
        if self.loader.colmap is None:
            return {}
        
        colmap = self.loader.colmap
        df = self.loader.filter_data(start_date, end_date, malls)
        
        if df.empty:
            return {
                "total_quantity": 0,
                "total_products": 0,
                "total_orders": 0,
                "top_product": None,
                "top_mall": None,
                "period": {
                    "start": start_date.isoformat(),
                    "end": end_date.isoformat(),
                    "days": (end_date - start_date).days + 1
                }
            }
        
        # 총 판매 수량
        total_qty = int(df[colmap.qty].sum())
        
        # 총 상품 수
        total_products = df[colmap.sku].nunique()
        
        # 총 주문 건수 (행 수)
        total_orders = len(df)
        
        # 상위 상품 (판매량 기준)
        top_product_df = df.groupby(colmap.sku)[colmap.qty].sum().sort_values(ascending=False)
        if len(top_product_df) > 0:
            top_product = {
                "name": top_product_df.index[0],
                "quantity": int(top_product_df.iloc[0])
            }
        else:
            top_product = None
        
        # 상위 쇼핑몰 (판매량 기준)
        top_mall_df = df.groupby(colmap.mall)[colmap.qty].sum().sort_values(ascending=False)
        if len(top_mall_df) > 0:
            top_mall = {
                "name": top_mall_df.index[0],
                "quantity": int(top_mall_df.iloc[0])
            }
        else:
            top_mall = None
        
        return {
            "total_quantity": total_qty,
            "total_products": total_products,
            "total_orders": total_orders,
            "top_product": top_product,
            "top_mall": top_mall,
            "period": {
                "start": start_date.isoformat(),
                "end": end_date.isoformat(),
                "days": (end_date - start_date).days + 1
            }
        }
    
    @staticmethod
    def paginate_dataframe(
        df: pd.DataFrame,
        page: int = 1,
        per_page: int = 20
    ) -> Dict[str, Any]:
        """
        DataFrame을 페이지네이션합니다.
        
        Args:
            df: 페이지네이션할 DataFrame
            page: 페이지 번호 (1부터 시작)
            per_page: 페이지당 행 수 (0이면 전체)
            
        Returns:
            페이지네이션 정보 딕셔너리
            {
                "data": 해당 페이지의 데이터,
                "page": 현재 페이지,
                "per_page": 페이지당 행 수,
                "total": 전체 행 수,
                "total_pages": 전체 페이지 수
            }
        """
        if df.empty:
            return {
                "data": [],
                "columns": [],
                "page": 1,
                "per_page": per_page,
                "total": 0,
                "total_pages": 0
            }
        
        total = len(df)
        
        # per_page가 0이면 전체 데이터 반환
        if per_page <= 0:
            return {
                "data": df.to_dict(orient="records"),
                "columns": df.columns.tolist(),
                "page": 1,
                "per_page": total,
                "total": total,
                "total_pages": 1
            }
        
        # 전체 페이지 수 계산
        total_pages = (total + per_page - 1) // per_page
        
        # 페이지 범위 검증
        if page < 1:
            page = 1
        elif page > total_pages:
            page = total_pages
        
        # 시작/종료 인덱스 계산
        start_idx = (page - 1) * per_page
        end_idx = min(start_idx + per_page, total)
        
        # 해당 페이지 데이터 추출
        df_page = df.iloc[start_idx:end_idx]
        
        return {
            "data": df_page.to_dict(orient="records"),
            "columns": df.columns.tolist(),
            "page": page,
            "per_page": per_page,
            "total": total,
            "total_pages": total_pages
        }
    
    @staticmethod
    def dataframe_to_json(df: pd.DataFrame) -> Dict[str, Any]:
        """
        DataFrame을 JSON 직렬화 가능한 형태로 변환합니다.
        
        NaN, Infinity 등의 특수값을 처리하고
        날짜 형식을 문자열로 변환합니다.
        
        Args:
            df: 변환할 DataFrame
            
        Returns:
            JSON 직렬화 가능한 딕셔너리
        """
        if df.empty:
            return {
                "columns": [],
                "data": []
            }
        
        # NaN을 None으로 변환
        df_clean = df.replace({np.nan: None, np.inf: None, -np.inf: None})
        
        # 날짜 타입을 문자열로 변환
        for col in df_clean.columns:
            if df_clean[col].dtype == 'object':
                # object 타입 중 날짜인 것만 변환
                sample = df_clean[col].dropna().iloc[0] if len(df_clean[col].dropna()) > 0 else None
                if isinstance(sample, (dt.date, dt.datetime)):
                    df_clean[col] = df_clean[col].apply(
                        lambda x: x.isoformat() if x is not None else None
                    )
        
        return {
            "columns": df_clean.columns.tolist(),
            "data": df_clean.to_dict(orient="records")
        }
