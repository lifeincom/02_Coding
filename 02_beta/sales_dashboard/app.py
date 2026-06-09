"""
===========================================
Flask 웹 애플리케이션 (app.py)
===========================================

판매 데이터 분석 대시보드 웹 서버 메인 파일입니다.

주요 기능:
1. 정적 파일 서빙 (HTML, CSS, JS)
2. REST API 엔드포인트
   - GET  /api/info          : 데이터 정보 조회
   - GET  /api/malls         : 쇼핑몰 목록 조회
   - POST /api/analyze       : 분석 실행
   - POST /api/predict-mall  : 쇼핑몰별 예측
   - POST /api/predict-product : 상품별 예측
   - POST /api/export        : Excel 내보내기
3. CORS 지원

실행 방법:
    python app.py

접속 주소:
    http://localhost:5000

작성자: AI Assistant
작성일: 2026-02-09
"""

import os
import datetime as dt
from io import BytesIO
from typing import Dict, Any

from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
import pandas as pd

from data_loader import DataLoader
from analyzer import SalesAnalyzer
from predictor import SalesPredictor


# ===========================================
# Flask 애플리케이션 초기화
# ===========================================

app = Flask(__name__, static_folder='static', static_url_path='')
CORS(app)  # CORS 활성화 (프론트엔드에서 API 접근 허용)

# ===========================================
# 전역 인스턴스 초기화
# ===========================================

# 데이터 로더 초기화 및 자동 로드
data_loader = DataLoader()
success, message = data_loader.load_data()
print(f"[초기 로드] {message}")

# 분석기 및 예측기 초기화
analyzer = SalesAnalyzer(data_loader)
predictor = SalesPredictor(data_loader)


# ===========================================
# 유틸리티 함수
# ===========================================

def parse_date(date_str: str) -> dt.date:
    """
    날짜 문자열을 date 객체로 변환합니다.
    
    Args:
        date_str: "YYYY/MM/DD" 형식의 문자열
        
    Returns:
        date 객체
    """
    return dt.datetime.strptime(date_str, "%Y/%m/%d").date()


def create_error_response(message: str, status_code: int = 400) -> tuple:
    """
    에러 응답을 생성합니다.
    
    Args:
        message: 에러 메시지
        status_code: HTTP 상태 코드
        
    Returns:
        (JSON 응답, 상태 코드) 튜플
    """
    return jsonify({"error": message}), status_code


# ===========================================
# 라우트: 정적 파일
# ===========================================

@app.route('/')
def index():
    """
    메인 페이지를 반환합니다.
    
    Returns:
        index.html 파일
    """
    return send_from_directory('static', 'index.html')


# ===========================================
# API 라우트: 데이터 정보
# ===========================================

@app.route('/api/info', methods=['GET'])
def get_info():
    """
    현재 로드된 데이터의 정보를 반환합니다.
    
    Returns:
        JSON 응답
        {
            "loaded": bool,
            "is_dummy": bool,
            "row_count": int,
            "date_range": {"min": str, "max": str},
            "malls": [str, ...],
            "columns": {...}
        }
    """
    try:
        info = data_loader.get_data_info()
        return jsonify(info)
    except Exception as e:
        return create_error_response(f"정보 조회 실패: {str(e)}", 500)


@app.route('/api/malls', methods=['GET'])
def get_malls():
    """
    쇼핑몰 목록을 반환합니다.
    
    Returns:
        JSON 응답
        {
            "malls": [str, ...]
        }
    """
    try:
        malls = data_loader.get_mall_list()
        return jsonify({"malls": malls})
    except Exception as e:
        return create_error_response(f"쇼핑몰 목록 조회 실패: {str(e)}", 500)


# ===========================================
# API 라우트: 분석
# ===========================================

@app.route('/api/analyze', methods=['POST'])
def analyze():
    """
    판매 데이터 분석을 실행합니다.
    
    요청 본문 (JSON):
    {
        "start_date": "YYYY/MM/DD",
        "end_date": "YYYY/MM/DD",
        "malls": [str, ...],  // 선택 사항
        "page": int,          // 페이지 번호 (기본: 1)
        "per_page": int       // 페이지당 행 수 (기본: 20, 0이면 전체)
    }
    
    Returns:
        JSON 응답
        {
            "summary": {...},
            "pivots": {
                "product_daily": {...},
                "product_mall": {...},
                "daily_mall": {...},
                "yearly_mall": {...}
            }
        }
    """
    try:
        # 요청 데이터 파싱
        data = request.get_json()
        
        if not data:
            return create_error_response("요청 본문이 비어있습니다.")
        
        # 필수 파라미터 검증
        if 'start_date' not in data or 'end_date' not in data:
            return create_error_response("start_date와 end_date는 필수입니다.")
        
        # 날짜 파싱
        start_date = parse_date(data['start_date'])
        end_date = parse_date(data['end_date'])
        
        # 선택 파라미터
        malls = data.get('malls', None)
        page = data.get('page', 1)
        per_page = data.get('per_page', 20)
        
        # ======================================
        # 1. 요약 통계 생성
        # ======================================
        summary = analyzer.get_summary_stats(start_date, end_date, malls)
        
        # ======================================
        # 2. 피벗 테이블 생성
        # ======================================
        pivots_raw = analyzer.analyze_all(start_date, end_date, malls)
        
        # ======================================
        # 3. 페이지네이션 적용
        # ======================================
        pivots_paginated = {}
        for key, df in pivots_raw.items():
            pivots_paginated[key] = analyzer.paginate_dataframe(df, page, per_page)
        
        # 응답 생성
        response = {
            "summary": summary,
            "pivots": pivots_paginated
        }
        
        return jsonify(response)
        
    except ValueError as e:
        return create_error_response(f"잘못된 입력값: {str(e)}")
    except Exception as e:
        return create_error_response(f"분석 실패: {str(e)}", 500)


# ===========================================
# API 라우트: 예측
# ===========================================

@app.route('/api/predict-mall', methods=['POST'])
def predict_mall():
    """
    쇼핑몰별 판매량을 예측합니다.
    
    요청 본문 (JSON):
    {
        "start_date": "YYYY/MM/DD",
        "end_date": "YYYY/MM/DD",
        "days": int  // 예측 일수 (기본: 14)
    }
    
    Returns:
        JSON 응답
        {
            "by_mall": {
                "쇼핑몰명": {
                    "predictions": [...],
                    "total_predicted": float
                },
                ...
            },
            "total": [...],
            "total_sum": float
        }
    """
    try:
        # 요청 데이터 파싱
        data = request.get_json()
        
        if not data:
            return create_error_response("요청 본문이 비어있습니다.")
        
        # 필수 파라미터 검증
        if 'start_date' not in data or 'end_date' not in data:
            return create_error_response("start_date와 end_date는 필수입니다.")
        
        # 날짜 파싱
        start_date = parse_date(data['start_date'])
        end_date = parse_date(data['end_date'])
        
        # 선택 파라미터
        days = data.get('days', 14)
        
        # 예측 실행
        predictions = predictor.predict_by_mall(start_date, end_date, days)
        
        return jsonify(predictions)
        
    except ValueError as e:
        return create_error_response(f"잘못된 입력값: {str(e)}")
    except Exception as e:
        return create_error_response(f"예측 실패: {str(e)}", 500)


@app.route('/api/predict-product', methods=['POST'])
def predict_product():
    """
    상품별 판매량을 예측합니다.
    
    요청 본문 (JSON):
    {
        "start_date": "YYYY/MM/DD",
        "end_date": "YYYY/MM/DD",
        "days": int,     // 예측 일수 (기본: 14)
        "top_n": int     // 상위 N개 상품 (기본: 10)
    }
    
    Returns:
        JSON 응답
        {
            "by_product": {
                "상품코드": {
                    "predictions": [...],
                    "total_predicted": float,
                    "historical_total": int
                },
                ...
            },
            "top_n": int
        }
    """
    try:
        # 요청 데이터 파싱
        data = request.get_json()
        
        if not data:
            return create_error_response("요청 본문이 비어있습니다.")
        
        # 필수 파라미터 검증
        if 'start_date' not in data or 'end_date' not in data:
            return create_error_response("start_date와 end_date는 필수입니다.")
        
        # 날짜 파싱
        start_date = parse_date(data['start_date'])
        end_date = parse_date(data['end_date'])
        
        # 선택 파라미터
        days = data.get('days', 14)
        top_n = data.get('top_n', 10)
        
        # 예측 실행
        predictions = predictor.predict_by_product(start_date, end_date, days, top_n)
        
        return jsonify(predictions)
        
    except ValueError as e:
        return create_error_response(f"잘못된 입력값: {str(e)}")
    except Exception as e:
        return create_error_response(f"예측 실패: {str(e)}", 500)


# ===========================================
# API 라우트: Excel 내보내기
# ===========================================

@app.route('/api/export', methods=['POST'])
def export_excel():
    """
    분석 결과를 Excel 파일로 내보냅니다.
    
    요청 본문 (JSON):
    {
        "start_date": "YYYY/MM/DD",
        "end_date": "YYYY/MM/DD",
        "malls": [str, ...]  // 선택 사항
    }
    
    Returns:
        Excel 파일 (application/vnd.openxmlformats-officedocument.spreadsheetml.sheet)
    """
    try:
        # 요청 데이터 파싱
        data = request.get_json()
        
        if not data:
            return create_error_response("요청 본문이 비어있습니다.")
        
        # 필수 파라미터 검증
        if 'start_date' not in data or 'end_date' not in data:
            return create_error_response("start_date와 end_date는 필수입니다.")
        
        # 날짜 파싱
        start_date = parse_date(data['start_date'])
        end_date = parse_date(data['end_date'])
        
        # 선택 파라미터
        malls = data.get('malls', None)
        
        # 피벗 테이블 생성
        pivots = analyzer.analyze_all(start_date, end_date, malls)
        
        # ======================================
        # Excel 파일 생성 (메모리 버퍼 사용)
        # ======================================
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 각 시트 작성
            pivots['product_daily'].to_excel(writer, sheet_name='상품별_일자', index=False)
            pivots['product_mall'].to_excel(writer, sheet_name='상품별_쇼핑몰', index=False)
            pivots['daily_mall'].to_excel(writer, sheet_name='일자별_쇼핑몰', index=False)
            pivots['yearly_mall'].to_excel(writer, sheet_name='연도별_쇼핑몰', index=False)
            
            # 각 시트에 틀 고정 적용 (C3 기준)
            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]
                ws.freeze_panes = "C3"
        
        output.seek(0)
        
        # 파일명 생성 (날짜 포함)
        filename = f"판매분석_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # 파일 전송
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=filename
        )
        
    except ValueError as e:
        return create_error_response(f"잘못된 입력값: {str(e)}")
    except Exception as e:
        return create_error_response(f"내보내기 실패: {str(e)}", 500)


# ===========================================
# 애플리케이션 실행
# ===========================================

if __name__ == '__main__':
    print("\n" + "="*50)
    print("판매 데이터 인사이트 대시보드 웹 서버")
    print("="*50)
    print(f"접속 주소: http://localhost:5000")
    print(f"데이터 상태: {'더미 데이터' if data_loader.is_dummy else '실제 데이터'}")
    print("="*50 + "\n")
    
    # 개발 서버 실행
    # 프로덕션 환경에서는 gunicorn, uWSGI 등 사용 권장
    app.run(
        host='0.0.0.0',  # 외부 접속 허용
        port=5000,
        debug=True       # 개발 모드 (자동 재시작)
    )
