"""
===========================================
데이터 로더 모듈 (data_loader.py)
===========================================

판매 데이터 Excel 파일을 로드하고 캐싱하는 모듈입니다.
기존 Tkinter 애플리케이션의 데이터 로딩 로직을 웹 서비스용으로 이식했습니다.

주요 기능:
1. Excel 파일 자동 로드
2. Pickle 캐싱으로 빠른 재로드
3. 컬럼 자동 매핑 (날짜, 상품코드, 쇼핑몰, 수량 등)
4. 데이터 전처리 (날짜 변환, 수량 형변환)
5. 필터링 (날짜 범위, 쇼핑몰)

작성자: AI Assistant
작성일: 2026-02-09
"""

import os
import pickle
import datetime as dt
from typing import Optional, List, Dict, Tuple
from dataclasses import dataclass

import pandas as pd

# ===========================================
# 설정 상수
# ===========================================

# 기본 데이터 파일 경로 (NAS 서버)
DEFAULT_EXCEL_PATH = r"\\NAS451\team451\DB\통합매출데이터.xlsx"

# 컬럼 인식 후보군 (여러 가지 이름 형식을 지원)
DATE_COL_CANDIDATES = ["주문일자", "일자", "날짜", "order_date", "date"]
SKU_COL_CANDIDATES = ["상품코드", "상품", "product_code", "sku"]
MALL_COL_CANDIDATES = ["쇼핑몰", "몰명", "채널", "mall", "channel"]
QTY_COL_CANDIDATES = ["수량", "판매수량", "qty", "quantity"]
COLOR_COL_CANDIDATES = ["컬러", "색상", "색깔", "color", "colour"]
SIZE_COL_CANDIDATES = ["사이즈", "크기", "size"]


# ===========================================
# 데이터 클래스
# ===========================================

@dataclass
class ColumnMap:
    """
    컬럼 매핑 정보를 저장하는 데이터 클래스
    
    Excel 파일의 실제 컬럼명을 저장하여
    다양한 형식의 파일을 처리할 수 있도록 합니다.
    """
    date: str              # 날짜 컬럼명
    sku: str               # 상품코드 컬럼명
    mall: str              # 쇼핑몰 컬럼명
    qty: str               # 수량 컬럼명
    color: Optional[str] = None  # 컬러 컬럼명 (선택)
    size: Optional[str] = None   # 사이즈 컬럼명 (선택)


# ===========================================
# 데이터 로더 클래스
# ===========================================

class DataLoader:
    """
    판매 데이터 로딩 및 관리 클래스
    
    Excel 파일을 로드하고, 캐싱하고, 필터링하는 기능을 제공합니다.
    싱글톤 패턴으로 구현되어 애플리케이션 전체에서 하나의 인스턴스만 사용합니다.
    """
    
    def __init__(self, excel_path: str = DEFAULT_EXCEL_PATH):
        """
        데이터 로더 초기화
        
        Args:
            excel_path: Excel 파일 경로 (기본값: NAS 서버 경로)
        """
        self.excel_path = excel_path
        self.df_raw: Optional[pd.DataFrame] = None      # 원본 데이터
        self.colmap: Optional[ColumnMap] = None         # 컬럼 매핑 정보
        self.is_dummy: bool = False                     # 더미 데이터 사용 여부
        
    def _auto_map_columns(self, df: pd.DataFrame) -> ColumnMap:
        """
        DataFrame의 컬럼명을 자동으로 매핑합니다.
        
        여러 가지 컬럼명 형식을 지원하기 위해 후보군 리스트를 순회하며
        일치하는 컬럼을 찾습니다. 대소문자를 구분하지 않습니다.
        
        Args:
            df: 매핑할 DataFrame
            
        Returns:
            ColumnMap: 매핑된 컬럼 정보
            
        Raises:
            ValueError: 필수 컬럼을 찾지 못한 경우
        """
        # 대소문자 무시를 위한 컬럼 매핑 딕셔너리 생성
        columns_lower = {str(c).lower(): c for c in df.columns}
        
        def find_column(candidates: List[str]) -> Optional[str]:
            """
            후보 컬럼명 리스트에서 실제 컬럼을 찾습니다.
            
            1. 정확히 일치하는 컬럼 찾기
            2. 대소문자 무시하고 일치하는 컬럼 찾기
            
            Args:
                candidates: 후보 컬럼명 리스트
                
            Returns:
                찾은 컬럼명 또는 None
            """
            # 1단계: 정확히 일치하는 컬럼 찾기
            for c in df.columns:
                if c in candidates:
                    return c
            
            # 2단계: 대소문자 무시하고 찾기
            for cand in candidates:
                if cand.lower() in columns_lower:
                    return columns_lower[cand.lower()]
            
            return None
        
        # 각 필수 컬럼 찾기
        date_col = find_column(DATE_COL_CANDIDATES)
        sku_col = find_column(SKU_COL_CANDIDATES)
        mall_col = find_column(MALL_COL_CANDIDATES)
        qty_col = find_column(QTY_COL_CANDIDATES)
        
        # 선택 컬럼 찾기
        color_col = find_column(COLOR_COL_CANDIDATES)
        size_col = find_column(SIZE_COL_CANDIDATES)
        
        # 필수 컬럼 누락 검사
        missing = []
        if date_col is None:
            missing.append("날짜")
        if sku_col is None:
            missing.append("상품코드")
        if mall_col is None:
            missing.append("쇼핑몰")
        if qty_col is None:
            missing.append("수량")
            
        if missing:
            raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {', '.join(missing)}")
        
        return ColumnMap(
            date=date_col,
            sku=sku_col,
            mall=mall_col,
            qty=qty_col,
            color=color_col,
            size=size_col
        )
    
    def _make_dummy_data(self) -> pd.DataFrame:
        """
        테스트용 더미 데이터를 생성합니다.
        
        실제 파일이 없거나 로드에 실패한 경우
        데모용으로 사용할 수 있는 샘플 데이터를 만듭니다.
        
        Returns:
            더미 데이터가 담긴 DataFrame
        """
        import random
        
        # 최근 90일 날짜 범위 생성
        rng = pd.date_range(end=dt.date.today(), periods=90, freq="D")
        
        # 샘플 상품코드, 쇼핑몰, 옵션 정의
        skus = ["TPX-001", "TPX-002", "TPW-010", "TPX-003", "TPW-020"]
        malls = ["스마트스토어", "쿠팡", "11번가", "지마켓"]
        colors = ["빨강", "파랑", "검정", "흰색", "회색"]
        sizes = ["S", "M", "L", "XL"]
        
        rows = []
        
        # 랜덤 데이터 생성 (약 30%의 날짜-상품-쇼핑몰 조합만 생성)
        for d in rng:
            for sku in skus:
                for mall in malls:
                    # 30% 확률로 데이터 생성 (너무 많지 않게)
                    if random.random() > 0.3:
                        continue
                    
                    color = random.choice(colors)
                    size = random.choice(sizes)
                    qty = random.randint(1, 10)  # 1~10개 판매
                    
                    rows.append([d.date(), sku, mall, qty, color, size])
        
        return pd.DataFrame(
            rows,
            columns=["주문일자", "상품코드", "쇼핑몰", "수량", "컬러", "사이즈"]
        )
    
    def load_data(self, path: Optional[str] = None) -> Tuple[bool, str]:
        """
        Excel 파일을 로드합니다. 캐시가 있으면 캐시를 우선 사용합니다.
        
        로딩 프로세스:
        1. Pickle 캐시 파일 확인 (있고 최신이면 로드)
        2. Excel 파일 로드
        3. 컬럼 자동 매핑
        4. 데이터 전처리 (날짜 변환, 수량 형변환)
        5. Pickle 캐시 저장
        
        Args:
            path: Excel 파일 경로 (None이면 기본 경로 사용)
            
        Returns:
            (성공여부, 메시지) 튜플
        """
        if path is None:
            path = self.excel_path
        
        try:
            # Pickle 캐시 파일 경로 생성
            pickle_path = os.path.splitext(path)[0] + ".pkl"
            
            # ======================================
            # 1단계: Pickle 캐시 로드 시도
            # ======================================
            if os.path.exists(pickle_path) and os.path.exists(path):
                # 캐시가 원본보다 최신인지 확인
                if os.path.getmtime(pickle_path) >= os.path.getmtime(path):
                    try:
                        with open(pickle_path, 'rb') as f:
                            data = pickle.load(f)
                            self.df_raw = data['df']
                            self.colmap = data['colmap']
                            self.is_dummy = False
                            return True, f"캐시 로드 성공 ({len(self.df_raw):,}행)"
                    except Exception as e:
                        # 피클 로드 실패 시 Excel 로드로 진행
                        print(f"캐시 로드 실패: {e}")
            
            # ======================================
            # 2단계: Excel 파일 로드
            # ======================================
            if not os.path.exists(path):
                raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
            
            # Excel 읽기 (openpyxl 엔진 사용)
            df = pd.read_excel(path, engine="openpyxl")
            
            # ======================================
            # 3단계: 컬럼 자동 매핑
            # ======================================
            colmap = self._auto_map_columns(df)
            
            # ======================================
            # 4단계: 데이터 전처리
            # ======================================
            
            # 날짜 컬럼 변환 (문자열 또는 datetime을 date로 통일)
            if df[colmap.date].dtype == 'object':
                df[colmap.date] = pd.to_datetime(
                    df[colmap.date], 
                    errors='coerce'
                ).dt.date
            else:
                df[colmap.date] = pd.to_datetime(df[colmap.date]).dt.date
            
            # 수량 컬럼 정수형으로 변환 (에러는 0으로 처리)
            df[colmap.qty] = pd.to_numeric(
                df[colmap.qty], 
                errors="coerce"
            ).fillna(0).astype('int32')
            
            # 필요한 컬럼만 선택
            cols = [colmap.date, colmap.sku, colmap.mall, colmap.qty]
            if colmap.color:
                cols.append(colmap.color)
            if colmap.size:
                cols.append(colmap.size)
            
            self.df_raw = df[cols].copy()
            self.colmap = colmap
            self.is_dummy = False
            
            # ======================================
            # 5단계: Pickle 캐시 저장
            # ======================================
            try:
                with open(pickle_path, 'wb') as f:
                    pickle.dump(
                        {'df': self.df_raw, 'colmap': self.colmap},
                        f,
                        protocol=pickle.HIGHEST_PROTOCOL
                    )
            except Exception as e:
                print(f"캐시 저장 실패 (무시): {e}")
            
            return True, f"Excel 로드 성공 ({len(self.df_raw):,}행)"
            
        except Exception as e:
            # ======================================
            # 오류 시 더미 데이터 생성
            # ======================================
            print(f"데이터 로드 실패, 더미 데이터 사용: {e}")
            self.df_raw = self._make_dummy_data()
            self.colmap = self._auto_map_columns(self.df_raw)
            self.is_dummy = True
            return False, f"[더미 데이터] 오류: {e}"
    
    def get_mall_list(self) -> List[str]:
        """
        데이터에 포함된 쇼핑몰 목록을 반환합니다.
        
        Returns:
            쇼핑몰명 리스트 (정렬됨)
        """
        if self.df_raw is None or self.colmap is None:
            return []
        
        return sorted(self.df_raw[self.colmap.mall].unique().astype(str))
    
    def filter_data(
        self,
        start_date: dt.date,
        end_date: dt.date,
        malls: Optional[List[str]] = None
    ) -> pd.DataFrame:
        """
        날짜 범위와 쇼핑몰로 데이터를 필터링합니다.
        
        Args:
            start_date: 시작일
            end_date: 종료일
            malls: 쇼핑몰 리스트 (None이면 전체)
            
        Returns:
            필터링된 DataFrame
        """
        if self.df_raw is None or self.colmap is None:
            return pd.DataFrame()
        
        # 날짜 범위 필터
        mask = (
            (self.df_raw[self.colmap.date] >= start_date) &
            (self.df_raw[self.colmap.date] <= end_date)
        )
        
        # 쇼핑몰 필터 (지정된 경우)
        if malls:
            mask &= self.df_raw[self.colmap.mall].isin(set(malls))
        
        return self.df_raw.loc[mask].copy()
    
    def get_data_info(self) -> Dict:
        """
        현재 로드된 데이터의 정보를 반환합니다.
        
        Returns:
            데이터 정보 딕셔너리
        """
        if self.df_raw is None:
            return {
                "loaded": False,
                "message": "데이터가 로드되지 않았습니다."
            }
        
        return {
            "loaded": True,
            "is_dummy": self.is_dummy,
            "row_count": len(self.df_raw),
            "date_range": {
                "min": self.df_raw[self.colmap.date].min().isoformat(),
                "max": self.df_raw[self.colmap.date].max().isoformat()
            },
            "malls": self.get_mall_list(),
            "columns": {
                "date": self.colmap.date,
                "sku": self.colmap.sku,
                "mall": self.colmap.mall,
                "qty": self.colmap.qty,
                "color": self.colmap.color,
                "size": self.colmap.size
            }
        }
