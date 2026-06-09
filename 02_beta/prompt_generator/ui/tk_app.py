# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from logic.builder import build_prompt
from logic.presets import get_all_presets, save_presets_to_file

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("코딩 프롬프트 자동 생성기 - Tkinter Fallback")
        self.geometry("1600x900")
        self.minsize(1024, 768)

        # Windows 최대화
        try:
            self.state("zoomed")
        except:
            pass

        # 상태 보관
        self.app_state = {
            "persona": tk.StringVar(value="전문 개발자"),
            "project_goal": tk.StringVar(value=""),
            "tech_stack": tk.StringVar(value=""),
            "inputs": tk.StringVar(value=""),
            "outputs": tk.StringVar(value=""),
            "features": tk.StringVar(value=""),
            "ui_layout": tk.StringVar(value=""),
            "output_lang": tk.StringVar(value="한국어"),
            "comments_lang": tk.StringVar(value="한국어"),
            "code_block": tk.StringVar(value="하나의 완전한 코드 블록"),
            "file_format": tk.StringVar(value="단일 파일"),
            "include_examples": tk.BooleanVar(value=False),
            "include_tests": tk.BooleanVar(value=False),
            "strict_mode": tk.BooleanVar(value=True),
            "extras": {
                "detailed_comments": tk.BooleanVar(value=True),
                "intuitive_names": tk.BooleanVar(value=True),
                "error_handling": tk.BooleanVar(value=True),
                "full_executable": tk.BooleanVar(value=True),
                "modularization": tk.BooleanVar(value=True),
            }
        }
        
        self.current_presets = get_all_presets()
        self._build_widgets_tk()

    def _label(self, parent, text, font=("Arial", 10)):
        lbl = ttk.Label(parent, text=text, font=font)
        lbl.pack(anchor="w", pady=(6,2))
        return lbl

    def _header(self, parent, text):
        lbl = ttk.Label(parent, text=text, font=("Arial", 12, "bold"), foreground="#333")
        lbl.pack(anchor="w", pady=(15,5))
        return lbl

    def _multiline(self, parent, var, height=5):
        txt = tk.Text(parent, height=height)
        if var.get():
            txt.insert("1.0", var.get())
        txt.pack(fill="x", padx=0, pady=2)
        return txt

    def _get_text(self, widget:tk.Text):
        return widget.get("1.0", "end").strip()

    def _build_widgets_tk(self):
        # 전체 컨테이너
        main_container = ttk.Frame(self)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)

        # ==============================
        # 1. Top Bar (프리셋)
        # ==============================
        top_bar = ttk.Frame(main_container)
        top_bar.pack(fill="x", pady=(0, 10))
        
        ttk.Label(top_bar, text="프리셋 선택:", font=("Arial", 11, "bold")).pack(side="left", padx=(0,5))
        self.preset_combo_tk = ttk.Combobox(top_bar, values=list(self.current_presets.keys()) , state="readonly", width=30)
        self.preset_combo_tk.pack(side="left", padx=5)
        
        ttk.Button(top_bar, text="불러오기", command=lambda:self.load_preset_tk(self.preset_combo_tk.get())).pack(side="left", padx=2)
        ttk.Button(top_bar, text="저장", command=self.save_preset_tk).pack(side="left", padx=2)
        ttk.Button(top_bar, text="초기화", command=self.reset_all_tk).pack(side="left", padx=10)

        # ==============================
        # 2. Main Split (Left/Right)
        # ==============================
        paned = ttk.PanedWindow(main_container, orient="horizontal")
        paned.pack(fill="both", expand=True)

        # --- Left Panel (Inputs) ---
        left_frame = ttk.Frame(paned, padding=10)
        paned.add(left_frame, weight=2)

        # 섹션 1: 기본 정보
        self._header(left_frame, "1. 기본 정보")
        self._label(left_frame, "페르소나")
        ttk.Entry(left_frame, textvariable=self.app_state["persona"]).pack(fill="x")
        
        self._label(left_frame, "프로젝트 목적")
        self.project_goal_txt = self._multiline(left_frame, self.app_state["project_goal"], 3)

        self._label(left_frame, "기술 스택")
        self.tech_stack_txt = self._multiline(left_frame, self.app_state["tech_stack"], 3)

        # 섹션 2: 상세
        self._header(left_frame, "2. 상세 요구사항")
        split = ttk.Frame(left_frame)
        split.pack(fill="x")
        
        f_in = ttk.Frame(split); f_in.pack(side="left", fill="both", expand=True, padx=(0,5))
        f_out = ttk.Frame(split); f_out.pack(side="left", fill="both", expand=True, padx=(5,0))
        
        self._label(f_in, "입력 항목")
        self.inputs_txt = self._multiline(f_in, self.app_state["inputs"], 3)
        
        self._label(f_out, "출력 결과")
        self.outputs_txt = self._multiline(f_out, self.app_state["outputs"], 3)

        self._label(left_frame, "기능 요구사항")
        self.features_txt = self._multiline(left_frame, self.app_state["features"], 4)

        self._label(left_frame, "UI 구성")
        self.ui_layout_txt = self._multiline(left_frame, self.app_state["ui_layout"], 3)

        # 섹션 3: 옵션
        self._header(left_frame, "3. 설정 및 옵션")
        opts = ttk.Frame(left_frame)
        opts.pack(fill="x")
        
        checks = [
            ("상세 주석", "detailed_comments"),
            ("직관적 네이밍", "intuitive_names"),
            ("오류 처리", "error_handling"),
            ("완전 실행 코드", "full_executable"),
            ("모듈화", "modularization")
        ]
        for i, (txt, key) in enumerate(checks):
            r, c = divmod(i, 3)
            ttk.Checkbutton(opts, text=txt, variable=self.app_state["extras"][key]).grid(row=r, column=c, sticky="w", padx=5)

        ttk.Checkbutton(opts, text="예시 포함", variable=self.app_state["include_examples"]).grid(row=2, column=0, sticky="w", padx=5)
        ttk.Checkbutton(opts, text="테스트 포함", variable=self.app_state["include_tests"]).grid(row=2, column=1, sticky="w", padx=5)
        ttk.Checkbutton(opts, text="Strict 모드", variable=self.app_state["strict_mode"]).grid(row=2, column=2, sticky="w", padx=5)

        # 출력 포맷
        fmt = ttk.Frame(left_frame)
        fmt.pack(fill="x", pady=5)
        # 간단히 배치
        ttk.Label(fmt, text="언어:").pack(side="left")
        ttk.Combobox(fmt, textvariable=self.app_state["output_lang"], values=["한국어", "English"], width=10).pack(side="left")
        ttk.Label(fmt, text="주석:").pack(side="left", padx=(10,0))
        ttk.Combobox(fmt, textvariable=self.app_state["comments_lang"], values=["한국어", "English"], width=10).pack(side="left")

        # --- Right Panel (Preview) ---
        right_frame = ttk.Frame(paned, padding=10)
        paned.add(right_frame, weight=2)

        ttk.Label(right_frame, text="프롬프트 미리보기", font=("Arial", 14, "bold")).pack(anchor="w", pady=(0,10))
        
        self.preview = tk.Text(right_frame, wrap="word", font=("Consolas", 11))
        self.preview.pack(fill="both", expand=True, pady=(0,10))

        # Action Buttons
        btn_bar = ttk.Frame(right_frame)
        btn_bar.pack(fill="x")
        
        ttk.Button(btn_bar, text="▶ 미리보기 생성", command=self.generate_preview_tk).pack(side="left", fill="x", expand=True, padx=(0,5))
        ttk.Button(btn_bar, text="복사", command=self.copy_clipboard_tk).pack(side="right", padx=5)
        ttk.Button(btn_bar, text="파일 저장", command=self.save_to_file_tk).pack(side="right", padx=5)


    # 데이터 수집/생성
    def collect_data_tk(self):
        data = {
            "persona": self.app_state["persona"].get(),
            "project_goal": self._get_text(self.project_goal_txt),
            "tech_stack": self._get_text(self.tech_stack_txt),
            "inputs": self._get_text(self.inputs_txt),
            "outputs": self._get_text(self.outputs_txt),
            "features": self._get_text(self.features_txt),
            "ui_layout": self._get_text(self.ui_layout_txt),
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

    def generate_preview_tk(self):
        text = build_prompt(self.collect_data_tk())
        self.preview.delete("1.0", "end")
        self.preview.insert("1.0", text)

    def copy_clipboard_tk(self):
        text = self.preview.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("경고", "먼저 '미리보기 생성'을 눌러주세요.")
            return
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.update()  # Windows에서 클립보드 동기화를 위해 필요
            messagebox.showinfo("완료", "클립보드에 복사되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"클립보드 복사 실패: {str(e)}")

    def save_to_file_tk(self):
        text = self.preview.get("1.0", "end").strip()
        if not text:
            messagebox.showwarning("경고", "먼저 '미리보기 생성'을 눌러주세요.")
            return
        fname = filedialog.asksaveasfilename(
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
            messagebox.showinfo("저장 완료", f"파일로 저장했습니다:\n{fname}")
        except PermissionError:
            messagebox.showerror("오류", f"파일 저장 권한이 없습니다:\n{fname}")
        except OSError as e:
            messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다:\n{str(e)}")
        except Exception as e:
            messagebox.showerror("오류", f"예상치 못한 오류가 발생했습니다:\n{str(e)}")

    def load_preset_tk(self, name):
        if not name or name not in self.current_presets:
            messagebox.showwarning("프리셋", "프리셋을 선택해주세요.")
            return
        p = self.current_presets[name]
        self.app_state["persona"].set(p.get("persona", "전문 개발자"))
        self.project_goal_txt.delete("1.0","end")
        self.project_goal_txt.insert("1.0", p.get("project_goal",""))
        self.tech_stack_txt.delete("1.0","end")
        self.tech_stack_txt.insert("1.0", p.get("tech_stack",""))
        self.inputs_txt.delete("1.0","end")
        self.inputs_txt.insert("1.0", p.get("inputs",""))
        self.outputs_txt.delete("1.0","end")
        self.outputs_txt.insert("1.0", p.get("outputs",""))
        self.features_txt.delete("1.0","end")
        self.features_txt.insert("1.0", p.get("features",""))
        self.ui_layout_txt.delete("1.0","end")
        self.ui_layout_txt.insert("1.0", p.get("ui_layout",""))
        for k, v in self.app_state["extras"].items():
            v.set(bool(p.get("extras",{}).get(k, False)))
        self.app_state["output_lang"].set(p.get("output_lang","한국어"))
        self.app_state["comments_lang"].set(p.get("comments_lang","한국어"))
        self.app_state["code_block"].set(p.get("code_block","하나의 완전한 코드 블록"))
        self.app_state["file_format"].set(p.get("file_format","단일 파일"))
        self.app_state["strict_mode"].set(bool(p.get("strict_mode", True)))
        self.app_state["include_examples"].set(bool(p.get("include_examples", False)))
        self.app_state["include_tests"].set(bool(p.get("include_tests", False)))
        messagebox.showinfo("프리셋", f"'{name}' 프리셋을 불러왔습니다.")

    def save_preset_tk(self):
        """현재 입력 내용을 프리셋으로 저장"""
        # 프리셋 이름 입력 다이얼로그
        dialog = tk.Toplevel(self)
        dialog.title("프리셋 저장")
        dialog.geometry("400x150")
        dialog.transient(self)
        dialog.grab_set()
        
        tk.Label(dialog, text="프리셋 이름:").pack(pady=10)
        name_entry = tk.Entry(dialog, width=30)
        name_entry.pack(pady=5)
        name_entry.focus()
        
        result = {"confirmed": False, "name": ""}
        
        def on_ok():
            name = name_entry.get().strip()
            if not name:
                messagebox.showwarning("경고", "프리셋 이름을 입력해주세요.")
                return
            result["confirmed"] = True
            result["name"] = name
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(pady=10)
        tk.Button(btn_frame, text="저장", command=on_ok).pack(side="left", padx=5)
        tk.Button(btn_frame, text="취소", command=on_cancel).pack(side="left", padx=5)
        
        dialog.wait_window()
        
        if not result["confirmed"]:
            return
        
        preset_name = result["name"]
        
        # 현재 입력 데이터 수집
        data = self.collect_data_tk()
        
        # 프리셋에 추가/업데이트
        self.current_presets[preset_name] = data
        
        # 파일에 저장
        if save_presets_to_file(self.current_presets):
            # 콤보박스 업데이트
            self.preset_combo_tk["values"] = list(self.current_presets.keys())
            self.preset_combo_tk.set(preset_name)
            messagebox.showinfo("저장 완료", f"'{preset_name}' 프리셋이 저장되었습니다.")
        else:
            messagebox.showerror("오류", "프리셋 저장에 실패했습니다.")

    def reset_all_tk(self):
        """모든 입력 필드를 기본값으로 초기화"""
        # 확인 다이얼로그
        if not messagebox.askyesno("초기화", "모든 입력 내용을 초기화하시겠습니까?"):
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
        self.preset_combo_tk.set("")
        
        # 미리보기 초기화
        self.preview.delete("1.0", "end")
        
        messagebox.showinfo("완료", "모든 입력 내용이 초기화되었습니다.")
