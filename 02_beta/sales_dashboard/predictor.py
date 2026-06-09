"""
===========================================
판매 예측 모듈 (predictor.py)
===========================================

판매 데이터를 기반으로 향후 판매량을 예측하는 모듈입니다.

예측 방법:
1. 이동 평균 (Moving Average)
   - 7일 이동 평균
   - 14일 이동 평균
   - 가중 이동 평균 (최근 데이터에 더 높은 가중치)

2. 선형 회귀 (Linear Regression)
   - 최근 30일 데이터로 추세선 계산
   - 증가/감소 추세 반영

3. 요일별 계절성 (Weekday Seasonality)
   - 같은 요일의 과거 판매 패턴 분석
   - 주말/평일 차이 반영

4. 복합 예측
   - 이동평균 40% + 선형회귀 30% + 요일패턴 30%

작성자: AI Assistant
작성일: 2026-02-09
"""

import datetime as dt
from typing import Dict, List, Optional, Any, Tuple
import pandas as pd
import numpy as np
from sklearn.linear_model import LinearRegression

from data_loader import DataLoader, ColumnMap


class SalesPredictor:
    """
    판매 예측 클래스
    
    쇼핑몰별 및 상품별 판매량을 예측합니다.
    """
    
    def __init__(self, data_loader: DataLoader):
        """
        예측기 초기화
        
        Args:
            data_loader: 데이터 로더 인스턴스
        """
        self.loader = data_loader
    
    def _calculate_moving_average(
        self,
        series: pd.Series,
        window: int = 7
    ) -> float:
        """
        이동 평균을 계산합니다.
        
        Args:
            series: 시계열 데이터
            window: 윈도우 크기 (일)
            
        Returns:
            이동 평균값
        """
        if len(series) == 0:
            return 0.0
        
        # 윈도우보다 데이터가 적으면 전체 평균 사용
        if len(series) < window:
            return float(series.mean())
        
        # 최근 window일의 평균
        return float(series.tail(window).mean())
    
    def _calculate_weighted_moving_average(
        self,
        series: pd.Series,
        window: int = 7
    ) -> float:
        """
        가중 이동 평균을 계산합니다.
        
        최근 데이터에 더 높은 가중치를 부여합니다.
        
        Args:
            series: 시계열 데이터
            window: 윈도우 크기 (일)
            
        Returns:
            가중 이동 평균값
        """
        if len(series) == 0:
            return 0.0
        
        # 윈도우보다 데이터가 적으면 전체 평균 사용
        if len(series) < window:
            return float(series.mean())
        
        # 최근 window일 데이터 추출
        recent = series.tail(window).values
        
        # 선형 가중치 생성 (1, 2, 3, ..., window)
        weights = np.arange(1, len(recent) + 1)
        
        # 가중 평균 계산
        weighted_avg = np.average(recent, weights=weights)
        
        return float(weighted_avg)
    
    def _calculate_linear_trend(
        self,
        series: pd.Series,
        days_ahead: int = 1
    ) -> float:
        """
        선형 회귀로 추세를 계산하고 미래값을 예측합니다.
        
        Args:
            series: 시계열 데이터
            days_ahead: 몇 일 후를 예측할지
            
        Returns:
            예측값
        """
        if len(series) < 3:
            # 데이터가 너무 적으면 평균 반환
            return float(series.mean()) if len(series) > 0 else 0.0
        
        # X: 일자 인덱스 (0, 1, 2, ...)
        X = np.arange(len(series)).reshape(-1, 1)
        # Y: 판매량
        y = series.values
        
        # 선형 회귀 모델 학습
        model = LinearRegression()
        model.fit(X, y)
        
        # 미래 시점 예측
        future_X = np.array([[len(series) + days_ahead - 1]])
        prediction = model.predict(future_X)[0]
        
        # 음수 방지
        return max(0.0, float(prediction))
    
    def _calculate_weekday_pattern(
        self,
        df: pd.DataFrame,
        date_col: str,
        qty_col: str,
        target_weekday: int
    ) -> float:
        """
        특정 요일의 평균 판매 패턴을 계산합니다.
        
        Args:
            df: 데이터프레임
            date_col: 날짜 컬럼명
            qty_col: 수량 컬럼명
            target_weekday: 목표 요일 (0=월요일, 6=일요일)
            
        Returns:
            해당 요일의 평균 판매량
        """
        if df.empty:
            return 0.0
        
        # 날짜를 datetime으로 변환하여 요일 추출
        df_temp = df.copy()
        df_temp['weekday'] = pd.to_datetime(df_temp[date_col]).dt.weekday
        
        # 해당 요일 데이터만 추출
        weekday_data = df_temp[df_temp['weekday'] == target_weekday]
        
        if weekday_data.empty:
            # 해당 요일 데이터가 없으면 전체 평균 사용
            daily_avg = df_temp.groupby(date_col)[qty_col].sum().mean()
            return float(daily_avg)
        
        # 해당 요일의 일별 합계 평균
        daily_sum = weekday_data.groupby(date_col)[qty_col].sum()
        return float(daily_sum.mean())
    
    def predict_daily(
        self,
        df: pd.DataFrame,
        date_col: str,
        qty_col: str,
        days: int = 14,
        use_weekday_pattern: bool = True
    ) -> List[Dict[str, Any]]:
        """
        일별 판매량을 예측합니다.
        
        Args:
            df: 과거 데이터
            date_col: 날짜 컬럼명
            qty_col: 수량 컬럼명
            days: 예측 일수
            use_weekday_pattern: 요일 패턴 사용 여부
            
        Returns:
            예측 결과 리스트
            [
                {
                    "date": "2026/02/10",
                    "predicted": 예측값,
                    "lower_bound": 하한값,
                    "upper_bound": 상한값
                },
                ...
            ]
        """
        if df.empty:
            # 데이터가 없으면 빈 리스트 반환
            return []
        
        # 일별 집계
        daily = df.groupby(date_col)[qty_col].sum().sort_index()
        
        if len(daily) == 0:
            return []
        
        # 예측 시작 날짜 (마지막 데이터 다음 날)
        last_date = daily.index[-1]
        if isinstance(last_date, str):
            last_date = dt.datetime.strptime(last_date, "%Y/%m/%d").date()
        
        start_date = last_date + dt.timedelta(days=1)
        
        predictions = []
        
        for i in range(days):
            target_date = start_date + dt.timedelta(days=i)
            target_weekday = target_date.weekday()
            
            # ======================================
            # 1. 이동 평균 계산
            # ======================================
            ma_7 = self._calculate_moving_average(daily, window=7)
            ma_14 = self._calculate_moving_average(daily, window=14)
            wma_7 = self._calculate_weighted_moving_average(daily, window=7)
            
            # 이동 평균 종합 (가중 이동 평균에 더 높은 비중)
            ma_prediction = wma_7 * 0.5 + ma_7 * 0.3 + ma_14 * 0.2
            
            # ======================================
            # 2. 선형 회귀 예측
            # ======================================
            trend_prediction = self._calculate_linear_trend(daily, days_ahead=i+1)
            
            # ======================================
            # 3. 요일 패턴 예측
            # ======================================
            if use_weekday_pattern:
                weekday_prediction = self._calculate_weekday_pattern(
                    df, date_col, qty_col, target_weekday
                )
            else:
                weekday_prediction = ma_prediction
            
            # ======================================
            # 4. 복합 예측
            # ======================================
            # 가중치: 이동평균 40%, 선형회귀 30%, 요일패턴 30%
            final_prediction = (
                ma_prediction * 0.4 +
                trend_prediction * 0.3 +
                weekday_prediction * 0.3
            )
            
            # ======================================
            # 5. 신뢰구간 계산 (표준편차 기반)
            # ======================================
            std = daily.std()
            lower_bound = max(0, final_prediction - std)
            upper_bound = final_prediction + std
            
            predictions.append({
                "date": target_date.isoformat(),
                "predicted": round(final_prediction, 1),
                "lower_bound": round(lower_bound, 1),
                "upper_bound": round(upper_bound, 1),
                "weekday": target_weekday,
                "weekday_name": ["월", "화", "수", "목", "금", "토", "일"][target_weekday]
            })
        
        return predictions
    
    def predict_by_mall(
        self,
        start_date: dt.date,
        end_date: dt.date,
        days: int = 14
    ) -> Dict[str, Any]:
        """
        쇼핑몰별 판매량을 예측합니다.
        
        Args:
            start_date: 학습 데이터 시작일
            end_date: 학습 데이터 종료일 (오늘)
            days: 예측 일수
            
        Returns:
            쇼핑몰별 예측 결과 딕셔너리
        """
        if self.loader.colmap is None or self.loader.df_raw is None:
            return {}
        
        colmap = self.loader.colmap
        df = self.loader.filter_data(start_date, end_date, malls=None)
        
        if df.empty:
            return {"predictions": {}}
        
        # 쇼핑몰 목록
        malls = sorted(df[colmap.mall].unique())
        
        predictions = {}
        
        for mall in malls:
            # 해당 쇼핑몰 데이터만 추출
            df_mall = df[df[colmap.mall] == mall]
            
            # 예측 실행
            mall_predictions = self.predict_daily(
                df_mall,
                date_col=colmap.date,
                qty_col=colmap.qty,
                days=days,
                use_weekday_pattern=True
            )
            
            predictions[mall] = {
                "predictions": mall_predictions,
                "total_predicted": sum(p["predicted"] for p in mall_predictions)
            }
        
        # 전체 합계 예측
        total_predictions = []
        for i in range(days):
            date_str = mall_predictions[i]["date"] if mall_predictions else None
            total = sum(
                predictions[mall]["predictions"][i]["predicted"]
                for mall in malls
            )
            total_predictions.append({
                "date": date_str,
                "predicted": round(total, 1)
            })
        
        return {
            "by_mall": predictions,
            "total": total_predictions,
            "total_sum": sum(p["predicted"] for p in total_predictions)
        }
    
    def predict_by_product(
        self,
        start_date: dt.date,
        end_date: dt.date,
        days: int = 14,
        top_n: int = 10
    ) -> Dict[str, Any]:
        """
        상품별 판매량을 예측합니다.
        
        상위 N개 상품에 대해서만 예측을 수행합니다.
        
        Args:
            start_date: 학습 데이터 시작일
            end_date: 학습 데이터 종료일
            days: 예측 일수
            top_n: 상위 N개 상품
            
        Returns:
            상품별 예측 결과 딕셔너리
        """
        if self.loader.colmap is None or self.loader.df_raw is None:
            return {}
        
        colmap = self.loader.colmap
        df = self.loader.filter_data(start_date, end_date, malls=None)
        
        if df.empty:
            return {"predictions": {}}
        
        # 상위 N개 상품 선정 (판매량 기준)
        top_products = (
            df.groupby(colmap.sku)[colmap.qty]
            .sum()
            .sort_values(ascending=False)
            .head(top_n)
            .index
            .tolist()
        )
        
        predictions = {}
        
        for product in top_products:
            # 해당 상품 데이터만 추출
            df_product = df[df[colmap.sku] == product]
            
            # 예측 실행
            product_predictions = self.predict_daily(
                df_product,
                date_col=colmap.date,
                qty_col=colmap.qty,
                days=days,
                use_weekday_pattern=True
            )
            
            # 과거 판매량 (참고용)
            historical_total = int(df_product[colmap.qty].sum())
            
            predictions[product] = {
                "predictions": product_predictions,
                "total_predicted": sum(p["predicted"] for p in product_predictions),
                "historical_total": historical_total
            }
        
        return {
            "by_product": predictions,
            "top_n": top_n
        }
