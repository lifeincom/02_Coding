#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
코딩 프롬프트 자동 생성기 (고급형)
- UI: Flet (Modern Cross-Platform)
- 기능: 입력 → 템플릿 형식의 '코딩 요청 프롬프트' 자동 생성
- 추가: 미리보기, 클립보드 복사, 파일 저장(.txt), 프리셋 불러오기/저장
"""

import sys
import os
import json
from datetime import datetime

import flet as ft

# ==============================
# 공통 유틸
# ==============================
def split_lines(s):
    if s is None:
        return []
    return [line.strip() for line in s.splitlines() if line.strip()]

def normalize_csv_list(s):
    # 콤마/개행 혼합 입력을 일괄 처리
    if s is None:
        return []
    parts = []
    for line in s.splitlines():
        for p in line.split(","):
            p = p.strip()
            if p:
                parts.append(p)
    return parts

def build_prompt(data: dict) -> str:
    """입력 데이터(dict)로 프롬프트 텍스트 생성"""
    if data is None:
        data = {}
    persona = data.get("persona", "전문 개발자") or "전문 개발자"
    project_goal = (data.get("project_goal") or "").strip()
    stack = split_lines(data.get("tech_stack", ""))
    inputs_ = normalize_csv_list(data.get("inputs", ""))
    outputs_ = normalize_csv_list(data.get("outputs", ""))
    features = split_lines(data.get("features", ""))
    ui_layout = split_lines(data.get("ui_layout", ""))
    extras = data.get("extras", {}) or {}
    output_lang = data.get("output_lang", "한국어") or "한국어"
    code_block = data.get("code_block", "하나의 완전한 코드 블록") or "하나의 완전한 코드 블록"
    comments_lang = data.get("comments_lang", "한국어") or "한국어"
    file_format = data.get("file_format", "단일 파일") or "단일 파일"
    include_examples = data.get("include_examples", False)
    include_tests = data.get("include_tests", False)
    strict_mode = data.get("strict_mode", True)

    # 언어 설정 확인
    is_english = output_lang == "English" or output_lang == "영어"
    
    # 언어별 텍스트 정의
    if is_english:
        # English translations
        texts = {
            "intro": f"You are a {persona}.",
            "request": "Please write code according to the requirements below.",
            "strict": "Reflect all requirements without omission and implement everything.",
            "project_goal": "[Project Goal]",
            "project_goal_placeholder": "(Please specify the project goal here)",
            "tech_stack": "[Technology Stack]",
            "inputs": "[Input Items]",
            "outputs": "[Output Results]",
            "features": "[Feature Requirements]",
            "ui_layout": "[UI Layout]",
            "additional": "[Additional Conditions]",
            "output_format": "[Output Format]",
            "code_block": f"- Provide as {code_block}",
            "comments_lang": f"- Write comments in {comments_lang}",
            "output_lang": f"- Write descriptions and result guidance in {output_lang}",
            "notes": "[Notes]",
            "note1": "- If external dependencies are required, explain installation methods in code comments",
            "note2": "- Recommend separating constants/environment settings instead of hardcoding",
            "note3": "- Maintain balance between performance and readability",
            "extras": {
                "detailed_comments": "Add detailed comments to each function block and major logic",
                "intuitive_names": "Write intuitive variable/function names",
                "error_handling": "User-friendly error handling for empty input and exception situations",
                "full_executable": "Provide complete code that can be executed immediately after copying",
                "modularization": "Modularize functions into function units",
                "examples": "Include example inputs/outputs or dummy data",
                "tests": "Include simple self-tests (e.g., main function or test code)",
                "file_format": f"Provide output code as {file_format} structure"
            }
        }
    else:
        # Korean (기본값)
        texts = {
            "intro": f"당신은 {persona}입니다.",
            "request": "아래 요구사항에 맞는 코드를 작성해주세요.",
            "strict": "요구사항을 빠짐없이 반영하고, 누락 없이 구현하세요.",
            "project_goal": "[프로젝트 목적]",
            "project_goal_placeholder": "(여기에 프로젝트 목적을 명시하세요)",
            "tech_stack": "[기술 스택]",
            "inputs": "[입력 항목]",
            "outputs": "[출력 결과]",
            "features": "[기능 요구사항]",
            "ui_layout": "[UI 구성]",
            "additional": "[추가 조건]",
            "output_format": "[출력 형태]",
            "code_block": f"- {code_block}으로 제공",
            "comments_lang": f"- 주석은 {comments_lang}로 작성",
            "output_lang": f"- 설명 및 결과 안내는 {output_lang}로 작성",
            "notes": "[유의 사항]",
            "note1": "- 외부 종속 라이브러리가 필요하면 설치 방법을 코드 주석에 설명",
            "note2": "- 하드코딩 대신 상수/환경설정 분리 권장",
            "note3": "- 성능과 가독성의 균형 유지",
            "extras": {
                "detailed_comments": "각 기능 블록과 주요 로직에 상세 주석을 추가",
                "intuitive_names": "변수/함수명을 직관적으로 작성",
                "error_handling": "빈 입력 및 예외 상황에 대한 사용자 친화적 오류 처리",
                "full_executable": "복사 후 즉시 실행 가능한 완전한 코드 제공",
                "modularization": "기능을 함수 단위로 모듈화",
                "examples": "예시 입력/출력 또는 더미 데이터 포함",
                "tests": "간단한 자체 테스트(예: main 함수 또는 테스트 코드) 포함",
                "file_format": f"출력 코드는 {file_format} 구조로 제공"
            }
        }

    # 체크박스 → 문장화
    extras_lines = []
    if extras and extras.get("detailed_comments"): 
        extras_lines.append(texts["extras"]["detailed_comments"])
    if extras and extras.get("intuitive_names"): 
        extras_lines.append(texts["extras"]["intuitive_names"])
    if extras and extras.get("error_handling"): 
        extras_lines.append(texts["extras"]["error_handling"])
    if extras and extras.get("full_executable"): 
        extras_lines.append(texts["extras"]["full_executable"])
    if extras and extras.get("modularization"): 
        extras_lines.append(texts["extras"]["modularization"])
    if include_examples: 
        extras_lines.append(texts["extras"]["examples"])
    if include_tests: 
        extras_lines.append(texts["extras"]["tests"])
    if file_format != ("Single file" if is_english else "단일 파일"): 
        extras_lines.append(texts["extras"]["file_format"])

    extras_text = "- " + "\n- ".join(extras_lines) if extras_lines else ""

    # 본문 템플릿
    lines = []
    lines.append(texts["intro"])
    lines.append(texts["request"])
    if strict_mode:
        lines.append(texts["strict"])
    lines.append("")
    lines.append(texts["project_goal"])
    lines.append(project_goal if project_goal else texts["project_goal_placeholder"])
    lines.append("")
    if stack:
        lines.append(texts["tech_stack"])
        lines.extend([f"- {s}" for s in stack])
        lines.append("")
    if inputs_:
        lines.append(texts["inputs"])
        lines.extend([f"- {s}" for s in inputs_])
        lines.append("")
    if outputs_:
        lines.append(texts["outputs"])
        lines.extend([f"- {s}" for s in outputs_])
        lines.append("")
    if features:
        lines.append(texts["features"])
        for i, f in enumerate(features, 1):
            lines.append(f"{i}. {f}")
        lines.append("")
    if ui_layout:
        lines.append(texts["ui_layout"])
        for item in ui_layout:
            lines.append(f"- {item}")
        lines.append("")
    if extras_text:
        lines.append(texts["additional"])
        lines.append(extras_text)
        lines.append("")

    lines.append(texts["output_format"])
    lines.append(texts["code_block"])
    lines.append(texts["comments_lang"])
    lines.append(texts["output_lang"])
    lines.append("")

    # 마무리
    lines.append(texts["notes"])
    lines.append(texts["note1"])
    lines.append(texts["note2"])
    lines.append(texts["note3"])
    lines.append("")
    lines.append(f"(Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})" if is_english else f"(생성 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")

    return "\n".join(lines)

# ==============================
# 프리셋 파일 경로
# ==============================
PRESET_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "code_generation_prompt_presets.json")

# ==============================
# 프리셋 저장/로드 함수
# ==============================
def load_presets_from_file():
    """프리셋 파일에서 불러오기"""
    if os.path.exists(PRESET_FILE):
        try:
            with open(PRESET_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"프리셋 파일 로드 오류: {e}")
    return {}

def save_presets_to_file(presets_dict):
    """프리셋을 파일에 저장"""
    try:
        with open(PRESET_FILE, "w", encoding="utf-8") as f:
            json.dump(presets_dict, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"프리셋 파일 저장 오류: {e}")
        return False

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

# 전역 PRESETS 변수 초기화 (파일에서 불러온 값과 기본값 병합)
PRESETS = {**DEFAULT_PRESETS, **load_presets_from_file()}


# ==============================
# Flet UI 메인 함수
# ==============================
def main(page: ft.Page):
    global PRESETS
    
    # 페이지 설정
    page.title = "코딩 프롬프트 자동 생성기 - Flet"
    page.theme_mode = ft.ThemeMode.DARK
    page.padding = 20
    page.window.width = 1600
    page.window.height = 900
    page.window.min_width = 1200
    page.window.min_height = 700
    page.window.maximized = True
    
    # 색상 테마
    primary_color = ft.Colors.BLUE_700
    success_color = ft.Colors.GREEN_700
    danger_color = ft.Colors.RED_700
    
    # ==============================
    # 상태 관리 (입력 필드 참조)
    # ==============================
    persona_field = ft.TextField(
        value="전문 개발자",
        label="페르소나 (역할)",
        border_radius=8,
    )
    
    project_goal_field = ft.TextField(
        label="프로젝트 목적",
        multiline=True,
        min_lines=3,
        max_lines=5,
        border_radius=8,
    )
    
    tech_stack_field = ft.TextField(
        label="기술 스택 (줄바꿈 구분)",
        multiline=True,
        min_lines=4,
        max_lines=6,
        border_radius=8,
    )
    
    inputs_field = ft.TextField(
        label="입력 항목",
        multiline=True,
        min_lines=4,
        max_lines=6,
        border_radius=8,
    )
    
    outputs_field = ft.TextField(
        label="출력 결과",
        multiline=True,
        min_lines=4,
        max_lines=6,
        border_radius=8,
    )
    
    features_field = ft.TextField(
        label="기능 요구사항 (줄바꿈 구분)",
        multiline=True,
        min_lines=5,
        max_lines=8,
        border_radius=8,
    )
    
    ui_layout_field = ft.TextField(
        label="UI 구성 아이디어",
        multiline=True,
        min_lines=5,
        max_lines=8,
        border_radius=8,
    )
    
    # 체크박스들
    cb_detailed_comments = ft.Checkbox(label="상세 주석 추가", value=True)
    cb_intuitive_names = ft.Checkbox(label="직관적 변수명", value=True)
    cb_error_handling = ft.Checkbox(label="예외 처리 강화", value=True)
    cb_full_executable = ft.Checkbox(label="완전 실행 코드", value=True)
    cb_modularization = ft.Checkbox(label="모듈화(함수 분리)", value=True)
    cb_include_examples = ft.Checkbox(label="예시 데이터 포함", value=False)
    cb_include_tests = ft.Checkbox(label="간단 테스트 코드", value=False)
    cb_strict_mode = ft.Checkbox(label="요구사항 엄수(Strict)", value=True)
    
    # 드롭다운
    output_lang_dropdown = ft.Dropdown(
        label="출력 언어",
        value="한국어",
        options=[
            ft.dropdown.Option("한국어"),
            ft.dropdown.Option("English"),
        ],
        width=150,
        border_radius=8,
    )
    
    comments_lang_dropdown = ft.Dropdown(
        label="주석 언어",
        value="한국어",
        options=[
            ft.dropdown.Option("한국어"),
            ft.dropdown.Option("English"),
        ],
        width=150,
        border_radius=8,
    )
    
    code_block_field = ft.TextField(
        label="코드 블록",
        value="하나의 완전한 코드 블록",
        width=200,
        border_radius=8,
    )
    
    file_format_field = ft.TextField(
        label="파일 구조",
        value="단일 파일",
        width=200,
        border_radius=8,
    )
    
    # 프리셋 드롭다운
    preset_dropdown = ft.Dropdown(
        label="프리셋 선택",
        options=[ft.dropdown.Option(key) for key in PRESETS.keys()],
        width=300,
        border_radius=8,
    )
    
    # 프롬프트 미리보기
    preview_field = ft.TextField(
        label="프롬프트 미리보기",
        multiline=True,
        read_only=False,
        min_lines=20,
        max_lines=30,
        border_radius=8,
        text_style=ft.TextStyle(font_family="Consolas", size=13),
    )
    
    # 파일 피커
    file_picker = ft.FilePicker()
    page.overlay.append(file_picker)
    
    # ==============================
    # 데이터 수집
    # ==============================
    def collect_data():
        return {
            "persona": persona_field.value or "전문 개발자",
            "project_goal": project_goal_field.value or "",
            "tech_stack": tech_stack_field.value or "",
            "inputs": inputs_field.value or "",
            "outputs": outputs_field.value or "",
            "features": features_field.value or "",
            "ui_layout": ui_layout_field.value or "",
            "output_lang": output_lang_dropdown.value or "한국어",
            "comments_lang": comments_lang_dropdown.value or "한국어",
            "code_block": code_block_field.value or "하나의 완전한 코드 블록",
            "file_format": file_format_field.value or "단일 파일",
            "include_examples": cb_include_examples.value,
            "include_tests": cb_include_tests.value,
            "strict_mode": cb_strict_mode.value,
            "extras": {
                "detailed_comments": cb_detailed_comments.value,
                "intuitive_names": cb_intuitive_names.value,
                "error_handling": cb_error_handling.value,
                "full_executable": cb_full_executable.value,
                "modularization": cb_modularization.value,
            },
        }
    
    # ==============================
    # 이벤트 핸들러
    # ==============================
    def show_snackbar(message: str, color=ft.Colors.GREEN):
        page.snack_bar = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE),
            bgcolor=color,
        )
        page.snack_bar.open = True
        page.update()
    
    def on_generate_click(e):
        """프롬프트 생성"""
        data = collect_data()
        prompt_text = build_prompt(data)
        preview_field.value = prompt_text
        page.update()
        show_snackbar("프롬프트가 생성되었습니다!", ft.Colors.BLUE_700)
    
    def on_copy_click(e):
        """클립보드에 복사"""
        if not preview_field.value:
            show_snackbar("먼저 '프롬프트 생성' 버튼을 눌러주세요.", ft.Colors.ORANGE_700)
            return
        page.set_clipboard(preview_field.value)
        show_snackbar("클립보드에 복사되었습니다!", ft.Colors.GREEN_700)
    
    def on_save_file_result(e: ft.FilePickerResultEvent):
        """파일 저장 결과 처리"""
        if e.path:
            try:
                with open(e.path, "w", encoding="utf-8") as f:
                    f.write(preview_field.value or "")
                show_snackbar(f"파일 저장 완료: {e.path}", ft.Colors.GREEN_700)
            except Exception as ex:
                show_snackbar(f"파일 저장 오류: {str(ex)}", ft.Colors.RED_700)
    
    file_picker.on_result = on_save_file_result
    
    def on_save_click(e):
        """파일 저장"""
        if not preview_field.value:
            show_snackbar("먼저 '프롬프트 생성' 버튼을 눌러주세요.", ft.Colors.ORANGE_700)
            return
        file_picker.save_file(
            dialog_title="프롬프트 저장",
            file_name="coding_prompt.txt",
            allowed_extensions=["txt"],
        )
    
    def on_load_preset_click(e):
        """프리셋 불러오기"""
        name = preset_dropdown.value
        if not name or name not in PRESETS:
            show_snackbar("프리셋을 선택해주세요.", ft.Colors.ORANGE_700)
            return
        
        p = PRESETS[name]
        persona_field.value = p.get("persona", "전문 개발자")
        project_goal_field.value = p.get("project_goal", "")
        tech_stack_field.value = p.get("tech_stack", "")
        inputs_field.value = p.get("inputs", "")
        outputs_field.value = p.get("outputs", "")
        features_field.value = p.get("features", "")
        ui_layout_field.value = p.get("ui_layout", "")
        output_lang_dropdown.value = p.get("output_lang", "한국어")
        comments_lang_dropdown.value = p.get("comments_lang", "한국어")
        code_block_field.value = p.get("code_block", "하나의 완전한 코드 블록")
        file_format_field.value = p.get("file_format", "단일 파일")
        cb_include_examples.value = bool(p.get("include_examples", False))
        cb_include_tests.value = bool(p.get("include_tests", False))
        cb_strict_mode.value = bool(p.get("strict_mode", True))
        
        extras = p.get("extras", {})
        cb_detailed_comments.value = bool(extras.get("detailed_comments", True))
        cb_intuitive_names.value = bool(extras.get("intuitive_names", True))
        cb_error_handling.value = bool(extras.get("error_handling", True))
        cb_full_executable.value = bool(extras.get("full_executable", True))
        cb_modularization.value = bool(extras.get("modularization", True))
        
        page.update()
        show_snackbar(f"'{name}' 프리셋을 불러왔습니다.", ft.Colors.BLUE_700)
    
    def on_save_preset_click(e):
        """현재 상태를 프리셋으로 저장"""
        global PRESETS
        
        preset_name_field = ft.TextField(label="프리셋 이름", autofocus=True, width=300)
        
        def close_dialog(e):
            dialog.open = False
            page.update()
        
        def save_preset(e):
            name = preset_name_field.value
            if not name or not name.strip():
                show_snackbar("프리셋 이름을 입력해주세요.", ft.Colors.ORANGE_700)
                return
            
            name = name.strip()
            PRESETS[name] = collect_data()
            
            if save_presets_to_file(PRESETS):
                # 드롭다운 업데이트
                preset_dropdown.options = [ft.dropdown.Option(key) for key in PRESETS.keys()]
                preset_dropdown.value = name
                dialog.open = False
                page.update()
                show_snackbar(f"'{name}' 프리셋이 저장되었습니다.", ft.Colors.GREEN_700)
            else:
                show_snackbar("프리셋 저장에 실패했습니다.", ft.Colors.RED_700)
        
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("프리셋 저장"),
            content=ft.Column([
                ft.Text("현재 입력 내용을 새 프리셋으로 저장합니다."),
                preset_name_field,
            ], tight=True, spacing=10),
            actions=[
                ft.TextButton("취소", on_click=close_dialog),
                ft.FilledButton("저장", on_click=save_preset),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.overlay.append(dialog)
        dialog.open = True
        page.update()
    
    def on_reset_click(e):
        """모든 입력 필드 초기화"""
        def close_dialog(e):
            confirm_dialog.open = False
            page.update()
        
        def do_reset(e):
            persona_field.value = "전문 개발자"
            project_goal_field.value = ""
            tech_stack_field.value = ""
            inputs_field.value = ""
            outputs_field.value = ""
            features_field.value = ""
            ui_layout_field.value = ""
            output_lang_dropdown.value = "한국어"
            comments_lang_dropdown.value = "한국어"
            code_block_field.value = "하나의 완전한 코드 블록"
            file_format_field.value = "단일 파일"
            cb_include_examples.value = False
            cb_include_tests.value = False
            cb_strict_mode.value = True
            cb_detailed_comments.value = True
            cb_intuitive_names.value = True
            cb_error_handling.value = True
            cb_full_executable.value = True
            cb_modularization.value = True
            preset_dropdown.value = None
            preview_field.value = ""
            
            confirm_dialog.open = False
            page.update()
            show_snackbar("모든 입력 내용이 초기화되었습니다.", ft.Colors.BLUE_700)
        
        confirm_dialog = ft.AlertDialog(
            modal=True,
            title=ft.Text("초기화 확인"),
            content=ft.Text("모든 입력 내용을 초기화하시겠습니까?"),
            actions=[
                ft.TextButton("취소", on_click=close_dialog),
                ft.FilledButton("초기화", on_click=do_reset, color=ft.Colors.WHITE, bgcolor=danger_color),
            ],
            actions_alignment=ft.MainAxisAlignment.END,
        )
        page.overlay.append(confirm_dialog)
        confirm_dialog.open = True
        page.update()
    
    # ==============================
    # UI 레이아웃 구성
    # ==============================
    
    # 상단 바 (프리셋)
    top_bar = ft.Container(
        content=ft.Row([
            ft.Text("프리셋", size=16, weight=ft.FontWeight.BOLD),
            preset_dropdown,
            ft.ElevatedButton("가져오기", icon=ft.Icons.DOWNLOAD, on_click=on_load_preset_click),
            ft.ElevatedButton("현재 상태 저장", icon=ft.Icons.SAVE, on_click=on_save_preset_click, bgcolor=ft.Colors.GREY_700),
            ft.Container(expand=True),  # Spacer
            ft.ElevatedButton("모두 초기화", icon=ft.Icons.REFRESH, on_click=on_reset_click, bgcolor=danger_color, color=ft.Colors.WHITE),
        ], alignment=ft.MainAxisAlignment.START, spacing=10),
        padding=ft.padding.only(bottom=20),
    )
    
    # 좌측 패널 - 입력 폼
    def create_section_header(text):
        return ft.Text(text, size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_200)
    
    left_panel = ft.Container(
        content=ft.Column([
            # 섹션 1: 기본 정보
            create_section_header("1. 기본 정보"),
            persona_field,
            project_goal_field,
            tech_stack_field,
            
            ft.Divider(height=20),
            
            # 섹션 2: 상세 요구사항
            create_section_header("2. 상세 요구사항"),
            ft.Row([
                ft.Container(inputs_field, expand=1),
                ft.Container(outputs_field, expand=1),
            ], spacing=10),
            features_field,
            ui_layout_field,
            
            ft.Divider(height=20),
            
            # 섹션 3: 설정 및 옵션
            create_section_header("3. 설정 및 옵션"),
            ft.Container(
                content=ft.Column([
                    ft.Row([cb_detailed_comments, cb_intuitive_names], spacing=20),
                    ft.Row([cb_error_handling, cb_full_executable], spacing=20),
                    ft.Row([cb_modularization, cb_include_examples], spacing=20),
                    ft.Row([cb_include_tests, cb_strict_mode], spacing=20),
                ]),
                bgcolor=ft.Colors.with_opacity(0.1, ft.Colors.WHITE),
                border_radius=8,
                padding=15,
            ),
            
            ft.Divider(height=20),
            
            # 섹션 4: 출력 형식
            create_section_header("4. 출력 형식"),
            ft.Row([
                output_lang_dropdown,
                comments_lang_dropdown,
                code_block_field,
                file_format_field,
            ], spacing=10, wrap=True),
        ], 
        scroll=ft.ScrollMode.AUTO,
        spacing=10,
        ),
        expand=2,
        padding=ft.padding.only(right=20),
    )
    
    # 우측 패널 - 미리보기 및 액션
    right_panel = ft.Container(
        content=ft.Column([
            ft.Text("프롬프트 미리보기", size=18, weight=ft.FontWeight.BOLD),
            ft.Container(
                content=preview_field,
                expand=True,
            ),
            ft.Container(
                content=ft.Row([
                    ft.FilledButton(
                        "▶ 프롬프트 생성",
                        icon=ft.Icons.PLAY_ARROW,
                        on_click=on_generate_click,
                        style=ft.ButtonStyle(
                            bgcolor=primary_color,
                            color=ft.Colors.WHITE,
                            padding=ft.padding.symmetric(horizontal=30, vertical=15),
                        ),
                    ),
                    ft.Container(expand=True),  # Spacer
                    ft.ElevatedButton(
                        "파일 저장",
                        icon=ft.Icons.SAVE_ALT,
                        on_click=on_save_click,
                        bgcolor=ft.Colors.GREY_700,
                    ),
                    ft.FilledButton(
                        "복사",
                        icon=ft.Icons.COPY,
                        on_click=on_copy_click,
                        style=ft.ButtonStyle(
                            bgcolor=success_color,
                            color=ft.Colors.WHITE,
                        ),
                    ),
                ], spacing=10),
                padding=ft.padding.only(top=15),
            ),
        ], expand=True),
        expand=2,
    )
    
    # 메인 레이아웃
    main_layout = ft.Column([
        top_bar,
        ft.Row([
            left_panel,
            ft.VerticalDivider(width=1),
            right_panel,
        ], expand=True, spacing=0),
    ], expand=True)
    
    page.add(main_layout)


# ==============================
# 진입점
# ==============================
if __name__ == "__main__":
    ft.app(target=main)
