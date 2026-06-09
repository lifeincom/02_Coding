# -*- coding: utf-8 -*-
from datetime import datetime

def split_lines(s: str) -> list[str]:
    """문자열을 줄 단위로 분리하고 공백을 제거합니다."""
    if s is None:
        return []
    return [line.strip() for line in s.splitlines() if line.strip()]

def normalize_csv_list(s: str) -> list[str]:
    """콤마와 개행이 섞인 문자열을 파싱하여 리스트로 반환합니다."""
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
    """입력 데이터(dict)를 기반으로 프롬프트 텍스트를 생성합니다."""
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

    lines.append(texts["notes"])
    lines.append(texts["note1"])
    lines.append(texts["note2"])
    lines.append(texts["note3"])
    lines.append("")
    lines.append(f"(Generated at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})" if is_english else f"(생성 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')})")

    return "\n".join(lines)
