# -*- coding: utf-8 -*-
import customtkinter as ctk
import tkinter.messagebox as mb
import tkinter.filedialog as fd
from logic.builder import build_prompt
from logic.presets import get_all_presets, save_presets_to_file

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("코딩 프롬프트 자동 생성기 - CustomTkinter (Rich UI)")
        self.geometry("1600x900")
        self.minsize(1200, 700)
        
        # Windows 최대화
        self.after(0, lambda: self.state("zoomed"))

        # 상태 보관
        self.app_state = {
            "persona": ctk.StringVar(value="전문 개발자"),
            "project_goal": ctk.StringVar(value=""),
            "tech_stack": ctk.StringVar(value=""),
            "inputs": ctk.StringVar(value=""),
            "outputs": ctk.StringVar(value=""),
            "features": ctk.StringVar(value=""),
            "ui_layout": ctk.StringVar(value=""),
            "output_lang": ctk.StringVar(value="한국어"),
            "comments_lang": ctk.StringVar(value="한국어"),
            "code_block": ctk.StringVar(value="하나의 완전한 코드 블록"),
            "file_format": ctk.StringVar(value="단일 파일"),
            "include_examples": ctk.BooleanVar(value=False),
            "include_tests": ctk.BooleanVar(value=False),
            "strict_mode": ctk.BooleanVar(value=True),
            "extras": {
                "detailed_comments": ctk.BooleanVar(value=True),
                "intuitive_names": ctk.BooleanVar(value=True),
                "error_handling": ctk.BooleanVar(value=True),
                "full_executable": ctk.BooleanVar(value=True),
                "modularization": ctk.BooleanVar(value=True),
            }
        }
        
        self.current_presets = get_all_presets()
        self._build_ui()

    def _build_ui(self):
        # 전체 컨테이너
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        
        main_container = ctk.CTkFrame(self, fg_color="transparent")
        main_container.grid(row=0, column=0, sticky="nsew", padx=20, pady=20)
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=2) # Left (Inputs)
        main_container.grid_columnconfigure(1, weight=2) # Right (Preview)

        # ==============================
        # 1. Top Bar (프리셋)
        # ==============================
        top_bar = ctk.CTkFrame(main_container, fg_color="transparent")
        top_bar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 20))
        
        ctk.CTkLabel(top_bar, text="프리셋 선택", font=("Arial", 14, "bold")).pack(side="left", padx=(0, 10))
        self.preset_combo = ctk.CTkComboBox(top_bar, values=list(self.current_presets.keys()), width=250)
        self.preset_combo.pack(side="left", padx=5)
        
        btn_color = "transparent"
        ctk.CTkButton(top_bar, text="가져오기", command=self.load_preset_ctk, width=100).pack(side="left", padx=5)
        ctk.CTkButton(top_bar, text="현재 상태 저장", command=self.save_preset_ctk, width=120, fg_color="gray").pack(side="left", padx=5)
        ctk.CTkButton(top_bar, text="모두 초기화", command=self.reset_all_ctk, width=100, fg_color="darkred", hover_color="red").pack(side="right", padx=5)

        # ==============================
        # 2. Left Panel (입력 폼)
        # ==============================
        left_panel = ctk.CTkScrollableFrame(main_container, label_text="입력 항목", label_font=("Arial", 16, "bold"))
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 15))
        left_panel.grid_columnconfigure(0, weight=1)

        # Helper functions for UI
        def add_header(parent, text):
            ctk.CTkLabel(parent, text=text, font=("Arial", 15, "bold"), anchor="w").pack(fill="x", pady=(20, 5))
        
        def add_field_label(parent, text):
            ctk.CTkLabel(parent, text=text, font=("Arial", 12), anchor="w", text_color="gray").pack(fill="x", pady=(2, 0))

        # --- 섹션 1: 기본 정보 ---
        add_header(left_panel, "1. 기본 정보")
        
        add_field_label(left_panel, "페르소나 (역할)")
        ctk.CTkEntry(left_panel, textvariable=self.app_state["persona"], height=35).pack(fill="x", pady=(0, 10))

        add_field_label(left_panel, "프로젝트 목적")
        self.project_goal_txt = ctk.CTkTextbox(left_panel, height=90)
        self.project_goal_txt.pack(fill="x", pady=(0, 10))

        add_field_label(left_panel, "기술 스택 (줄바꿈 구분)")
        self.tech_stack_txt = ctk.CTkTextbox(left_panel, height=120)
        self.tech_stack_txt.pack(fill="x", pady=(0, 10))

        # --- 섹션 2: 상세 내용 ---
        add_header(left_panel, "2. 상세 요구사항")
        
        split_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        split_frame.pack(fill="x", pady=(0, 10))
        split_frame.grid_columnconfigure(0, weight=1)
        split_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(split_frame, text="입력 항목", anchor="w").grid(row=0, column=0, sticky="w")
        ctk.CTkLabel(split_frame, text="출력 결과", anchor="w").grid(row=0, column=1, sticky="w", padx=(10,0))

        self.inputs_txt = ctk.CTkTextbox(split_frame, height=120)
        self.inputs_txt.grid(row=1, column=0, sticky="ew")
        
        self.outputs_txt = ctk.CTkTextbox(split_frame, height=120)
        self.outputs_txt.grid(row=1, column=1, sticky="ew", padx=(10,0))

        add_field_label(left_panel, "기능 요구사항 (줄바꿈 구분)")
        self.features_txt = ctk.CTkTextbox(left_panel, height=150)
        self.features_txt.pack(fill="x", pady=(0, 10))

        add_field_label(left_panel, "UI 구성 아이디어")
        self.ui_layout_txt = ctk.CTkTextbox(left_panel, height=150)
        self.ui_layout_txt.pack(fill="x", pady=(0, 10))

        # --- 섹션 3: 옵션 설정 ---
        add_header(left_panel, "3. 설정 및 옵션")
        
        opt_frame = ctk.CTkFrame(left_panel)
        opt_frame.pack(fill="x", pady=(0, 10))
        
        checks = [
            ("상세 주석 추가", "detailed_comments"),
            ("직관적 변수명", "intuitive_names"),
            ("예외 처리 강화", "error_handling"),
            ("완전 실행 코드", "full_executable"),
            ("모듈화(함수 분리)", "modularization")
        ]
        
        for i, (txt, key) in enumerate(checks):
            r, c = divmod(i, 2)
            ctk.CTkCheckBox(opt_frame, text=txt, variable=self.app_state["extras"][key]).grid(row=r, column=c, sticky="w", padx=10, pady=5)
        
        ctk.CTkCheckBox(opt_frame, text="예시 데이터 포함", variable=self.app_state["include_examples"]).grid(row=3, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkCheckBox(opt_frame, text="간단 테스트 코드", variable=self.app_state["include_tests"]).grid(row=3, column=1, sticky="w", padx=10, pady=5)
        ctk.CTkCheckBox(opt_frame, text="요구사항 엄수(Strict)", variable=self.app_state["strict_mode"]).grid(row=4, column=0, sticky="w", padx=10, pady=5)

        # --- 섹션 4: 출력 포맷 ---
        add_header(left_panel, "4. 출력 형식")
        format_frame = ctk.CTkFrame(left_panel, fg_color="transparent")
        format_frame.pack(fill="x")
        
        ctk.CTkLabel(format_frame, text="출력 언어:").grid(row=0, column=0, sticky="e", padx=5)
        self.output_lang = ctk.CTkComboBox(format_frame, values=["한국어", "English"], variable=self.app_state["output_lang"], width=120)
        self.output_lang.grid(row=0, column=1, sticky="w", pady=5)
        
        ctk.CTkLabel(format_frame, text="주석 언어:").grid(row=0, column=2, sticky="e", padx=5)
        self.comments_lang = ctk.CTkComboBox(format_frame, values=["한국어", "English"], variable=self.app_state["comments_lang"], width=120)
        self.comments_lang.grid(row=0, column=3, sticky="w", pady=5)

        ctk.CTkLabel(format_frame, text="코드 블록:").grid(row=1, column=0, sticky="e", padx=5)
        ctk.CTkEntry(format_frame, textvariable=self.app_state["code_block"], width=120).grid(row=1, column=1, sticky="w", pady=5)

        ctk.CTkLabel(format_frame, text="파일 구조:").grid(row=1, column=2, sticky="e", padx=5)
        ctk.CTkEntry(format_frame, textvariable=self.app_state["file_format"], width=120).grid(row=1, column=3, sticky="w", pady=5)


        # ==============================
        # 3. Right Panel (결과 & 액션)
        # ==============================
        right_panel = ctk.CTkFrame(main_container, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew")
        right_panel.grid_rowconfigure(1, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)

        # Header
        ctk.CTkLabel(right_panel, text="프롬프트 미리보기", font=("Arial", 16, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))

        # Preview Text
        self.preview = ctk.CTkTextbox(right_panel, font=("Consolas", 13), wrap="word")
        self.preview.grid(row=1, column=0, sticky="nsew", pady=(0, 20))

        # Action Buttons
        action_bar = ctk.CTkFrame(right_panel, height=60)
        action_bar.grid(row=2, column=0, sticky="ew")
        action_bar.grid_columnconfigure(0, weight=1) # Spacer

        # Big Generate Button
        self.btn_gen = ctk.CTkButton(action_bar, text="▶ 프롬프트 생성 (Generate)", command=self.generate_preview_ctk, 
                                    font=("Arial", 15, "bold"), height=45, fg_color="#106EBE", hover_color="#2B579A")
        self.btn_gen.pack(side="left", padx=10, pady=10, fill="x", expand=True)

        # Other Actions
        ctk.CTkButton(action_bar, text="복사 (Copy)", command=self.copy_clipboard_ctk, 
                      height=45, fg_color="green", hover_color="darkgreen").pack(side="right", padx=(5, 10), pady=10)
        
        ctk.CTkButton(action_bar, text="파일 저장", command=self.save_to_file_ctk, 
                      height=45, fg_color="#444", hover_color="#666").pack(side="right", padx=5, pady=10)


    def collect_data_ctk(self):
        data = {
            "persona": self.app_state["persona"].get(),
            "project_goal": self.project_goal_txt.get("1.0", "end").strip(),
            "tech_stack": self.tech_stack_txt.get("1.0", "end").strip(),
            "inputs": self.inputs_txt.get("1.0", "end").strip(),
            "outputs": self.outputs_txt.get("1.0", "end").strip(),
            "features": self.features_txt.get("1.0", "end").strip(),
            "ui_layout": self.ui_layout_txt.get("1.0", "end").strip(),
            "output_lang": self.app_state["output_lang"].get(),
            "comments_lang": self.app_state["comments_lang"].get(),
            "code_block": self.app_state["code_block"].get(),
            "file_format": self.app_state["file_format"].get(),
            "include_examples": self.app_state["include_examples"].get(),
            "include_tests": self.app_state["include_tests"].get(),
            "strict_mode": self.app_state["strict_mode"].get(),
            "extras": {k: v.get() for k, v in self.app_state["extras"].items()},
        }
        return data

    def generate_preview_ctk(self):
        text = build_prompt(self.collect_data_ctk())
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)

    def copy_clipboard_ctk(self):
        text = self.preview.get("1.0", "end").strip()
        if not text:
            self._msg("경고", "먼저 '미리보기 생성'을 눌러주세요.")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()  # Windows에서 클립보드 동기화를 위해 필요
            self._msg("완료", "클립보드에 복사되었습니다.")
        except Exception as e:
            self._msg("오류", f"클립보드 복사 실패: {str(e)}")

    def save_to_file_ctk(self):
        text = self.preview.get("1.0", "end").strip()
        if not text:
            self._msg("경고", "먼저 '미리보기 생성'을 눌러주세요.")
            return
        
        fname = fd.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files","*.txt"), ("All Files","*.*")],
            initialfile="coding_prompt.txt",
            title="프롬프트 저장 위치 선택"
        )
        if not fname:
            return
        try:
            with open(fname, "w", encoding="utf-8") as f:
                f.write(text)
            self._msg("저장 완료", f"파일로 저장했습니다:\n{fname}")
        except PermissionError:
            self._msg("오류", f"파일 저장 권한이 없습니다:\n{fname}")
        except OSError as e:
            self._msg("오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")
        except Exception as e:
            self._msg("오류", f"예상치 못한 오류가 발생했습니다:\n{str(e)}")

    def _msg(self, title, text):
        try:
            mb.showinfo(title, text)
        except Exception:
            pass

    def load_preset_ctk(self):
        name = self.preset_combo.get()
        if not name or name not in self.current_presets:
            self._msg("프리셋", "프리셋을 선택해주세요.")
            return
        p = self.current_presets[name]
        self.app_state["persona"].set(p.get("persona", "전문 개발자"))
        self.project_goal_txt.delete("1.0","end"); self.project_goal_txt.insert("1.0", p.get("project_goal",""))
        self.tech_stack_txt.delete("1.0","end"); self.tech_stack_txt.insert("1.0", p.get("tech_stack",""))
        self.inputs_txt.delete("1.0","end"); self.inputs_txt.insert("1.0", p.get("inputs",""))
        self.outputs_txt.delete("1.0","end"); self.outputs_txt.insert("1.0", p.get("outputs",""))
        self.features_txt.delete("1.0","end"); self.features_txt.insert("1.0", p.get("features",""))
        self.ui_layout_txt.delete("1.0","end"); self.ui_layout_txt.insert("1.0", p.get("ui_layout",""))
        for k, v in self.app_state["extras"].items():
            v.set(bool(p.get("extras",{}).get(k, False)))
        self.app_state["output_lang"].set(p.get("output_lang","한국어"))
        self.app_state["comments_lang"].set(p.get("comments_lang","한국어"))
        self.app_state["code_block"].set(p.get("code_block","하나의 완전한 코드 블록"))
        self.app_state["file_format"].set(p.get("file_format","단일 파일"))
        self.app_state["strict_mode"].set(bool(p.get("strict_mode", True)))
        self.app_state["include_examples"].set(bool(p.get("include_examples", False)))
        self.app_state["include_tests"].set(bool(p.get("include_tests", False)))
        self._msg("프리셋", f"'{name}' 프리셋을 불러왔습니다.")

    def save_preset_ctk(self):
        """현재 입력 내용을 프리셋으로 저장"""
        # 프리셋 이름 입력 다이얼로그
        dialog = ctk.CTkToplevel(self)
        dialog.title("프리셋 저장")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        ctk.CTkLabel(dialog, text="프리셋 이름:").pack(pady=10)
        name_entry = ctk.CTkEntry(dialog, width=300)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        result = {"confirmed": False, "name": ""}
        
        def on_ok():
            name = name_entry.get().strip()
            if not name:
                self._msg("경고", "프리셋 이름을 입력해주세요.")
                return
            result["confirmed"] = True
            result["name"] = name
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = ctk.CTkFrame(dialog)
        btn_frame.pack(pady=10)
        ctk.CTkButton(btn_frame, text="저장", command=on_ok, width=100).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="취소", command=on_cancel, width=100).pack(side="left", padx=5)
        
        dialog.wait_window()
        
        if not result["confirmed"]:
            return
        
        preset_name = result["name"]
        
        # 현재 입력 데이터 수집
        data = self.collect_data_ctk()
        
        # 프리셋에 추가/업데이트
        self.current_presets[preset_name] = data
        
        # 파일에 저장
        if save_presets_to_file(self.current_presets):
            # 콤보박스 업데이트
            self.preset_combo.configure(values=list(self.current_presets.keys()))
            self.preset_combo.set(preset_name)
            self._msg("저장 완료", f"'{preset_name}' 프리셋이 저장되었습니다.")
        else:
            self._msg("오류", "프리셋 저장에 실패했습니다.")

    def reset_all_ctk(self):
        """모든 입력 필드를 기본값으로 초기화"""
        # 확인 다이얼로그
        if not mb.askyesno("초기화", "모든 입력 내용을 초기화하시겠습니까?"):
            return
        
        # 모든 필드 초기화
        self.app_state["persona"].set("전문 개발자")
        self.project_goal_txt.delete("1.0", "end")
        self.tech_stack_txt.delete("1.0", "end")
        self.inputs_txt.delete("1.0", "end")
        self.outputs_txt.delete("1.0", "end")
        self.features_txt.delete("1.0", "end")
        self.ui_layout_txt.delete("1.0", "end")
        self.app_state["output_lang"].set("한국어")
        self.app_state["comments_lang"].set("한국어")
        self.app_state["code_block"].set("하나의 완전한 코드 블록")
        self.app_state["file_format"].set("단일 파일")
        self.app_state["include_examples"].set(False)
        self.app_state["include_tests"].set(False)
        self.app_state["strict_mode"].set(True)
        
        # 체크박스 초기화
        for k, v in self.app_state["extras"].items():
            v.set(True if k in ["detailed_comments", "intuitive_names", "error_handling", "full_executable", "modularization"] else False)
        
        # 프리셋 콤보박스 초기화
        self.preset_combo.set("")
        
        # 미리보기 초기화
        self.preview.delete("1.0", "end")
        
        self._msg("완료", "모든 입력 내용이 초기화되었습니다.")

if __name__ == "__main__":
    app = App()
    app.mainloop()
