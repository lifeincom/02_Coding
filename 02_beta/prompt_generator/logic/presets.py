# -*- coding: utf-8 -*-
import os
import json

# ==============================
# 프리셋 (기본값)
# ==============================
DEFAULT_PRESETS = {
    "Python GUI (tkinter + tksheet)": {
        "persona": "전문 개발자",
        "project_goal": "첨부한 엑셀 파일을 기반으로 상품 판매데이터 분석 Python GUI 프로그램 제작",
        "tech_stack": "Python\nTkinter\ntksheet\nopenpyxl\ntkcalendar",
        "inputs": "검색기간, 상품명, 단가, 수량, 배송비, 수수료율",
        "outputs": "검색기간의 판매수량 분석\n1.인덱스 상품코드 컬럼 일자\n2. 인덱스 상품코드 컬럼 쇼핑몰\n3. 인덱스 쇼핑몰 컬럼 년도 월\n4. 인덱스 년도 컬럼 쇼핑몰",
        "features": "결과를 tksheet 표에 출력\n결과를 시각화한 그래프 출력\nExcel 저장 기능",
        "ui_layout": "창 크기 최대화\n최소창 크기 1024x800\n상단 입력필드, 중앙 결과표, 하단 버튼 배치",
        "extras": {
            "detailed_comments": True,
            "intuitive_names": True,
            "error_handling": True,
            "full_executable": True,
            "modularization": True,
        },
        "output_lang": "한국어",
        "comments_lang": "한국어",
        "code_block": "하나의 완전한 코드 블록",
        "file_format": "단일 파일",
        "include_examples": True,
        "include_tests": False,
        "strict_mode": True,
    },
    "웹(HTML/CSS/JS) 계산기": {
        "persona": "프론트엔드 개발자",
        "project_goal": "상품 가격 자동 계산 웹페이지 제작",
        "tech_stack": "HTML\nCSS\nJavaScript",
        "inputs": "단가, 수량, 수수료율",
        "outputs": "총액, 이익금, 이익률",
        "features": "입력 즉시 계산\n반응형 레이아웃\n접근성 고려(레이블/키보드)",
        "ui_layout": "상단 입력, 중단 결과, 하단 저장/복사 버튼",
        "extras": {
            "detailed_comments": True,
            "intuitive_names": True,
            "error_handling": True,
            "full_executable": True,
            "modularization": False,
        },
        "output_lang": "한국어",
        "comments_lang": "한국어",
        "code_block": "하나의 완전한 코드 블록",
        "file_format": "단일 파일",
        "include_examples": False,
        "include_tests": False,
        "strict_mode": True,
    },
}

# ==============================
# 파일 경로 설정
# ==============================
# logic/presets.py 위치 기준: ../resources/presets.json
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PRESET_FILE = os.path.join(BASE_DIR, "resources", "presets.json")

def load_presets_from_file() -> dict:
    """프리셋 파일에서 불러오기"""
    if os.path.exists(PRESET_FILE):
        try:
            with open(PRESET_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"프리셋 파일 로드 오류: {e}")
    return {}

def save_presets_to_file(presets_dict: dict) -> bool:
    """프리셋을 파일에 저장"""
    try:
        # resources 폴더가 없으면 생성
        os.makedirs(os.path.dirname(PRESET_FILE), exist_ok=True)
        
        with open(PRESET_FILE, "w", encoding="utf-8") as f:
            json.dump(presets_dict, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"프리셋 파일 저장 오류: {e}")
        return False

def get_all_presets() -> dict:
    """기본 프리셋과 파일 저장된 프리셋 병합 반환"""
    return {**DEFAULT_PRESETS, **load_presets_from_file()}
