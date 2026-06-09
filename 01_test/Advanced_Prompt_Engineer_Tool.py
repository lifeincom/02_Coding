import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import datetime

class PromptBuilderApp:
    """
    프롬프트 생성기 애플리케이션 클래스입니다.
    사용자로부터 단계별 입력을 받아 구조화된 프롬프트를 생성합니다.
    """
    def __init__(self, root):
        # --- 기본 윈도우 설정 (Basic Window Setup) ---
        self.root = root
        self.root.title("Advanced Prompt Engineer Tool") # 윈도우 제목 설정
        self.root.geometry("600x850") # 창의 초기 크기 설정 (가로x세로)
        self.root.resizable(True, True) # 창 크기 조절 가능 여부 설정
        self.root.minsize(600, 850) # 최소 크기 설정

        # --- 스타일 설정 (Styling) ---s
        # ttk 스타일 객체를 생성하여 위젯들의 디자인을 관리합니다.
        style = ttk.Style()
        style.theme_use('clam') # 깔끔한 느낌의 'clam' 테마 사용

        # --- 메인 프레임 (Main Frame) ---
        # 모든 위젯을 담을 메인 컨테이너입니다. 여백(padding)을 주어 가독성을 높입니다.
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # --- 입력 필드 구성 (Input Fields Configuration) ---
        # 프롬프트의 5가지 핵심 요소를 입력받기 위한 위젯들입니다.
        # 기존 StringVar 대신 Text 위젯을 직접 제어합니다.

        # 반복되는 입력 필드 생성을 위해 헬퍼 함수를 사용하여 UI를 구성합니다.
        # 1. 역할 (Role) 입력 섹션
        self.role_entry = self.create_label_entry(main_frame, "1. 역할 (Persona): 당신은 누구입니까?", 
                                "예: 10년차 파이썬 개발자, 마케팅 전문가")
        
        # 2. 맥락 (Context) 입력 섹션
        self.context_entry = self.create_label_entry(main_frame, "2. 맥락 (Context): 배경 상황은 무엇입니까?", 
                                "예: 초보자를 위한 코딩 강의를 준비 중이다.")

        # 3. 작업 (Task) 입력 섹션
        self.task_entry = self.create_label_entry(main_frame, "3. 작업 (Task): AI가 수행해야 할 일은?", 
                                "예: 파이썬 기초 커리큘럼을 작성해줘.")

        # 4. 제약사항 (Constraints) 입력 섹션
        self.constraints_entry = self.create_label_entry(main_frame, "4. 제약사항 (Constraints): 하지 말아야 할 것은?", 
                                "예: 너무 어려운 용어 사용 금지, 300자 이내.")

        # 5. 출력 형식 (Format) 입력 섹션
        self.format_entry = self.create_label_entry(main_frame, "5. 출력 형식 (Output Format): 결과물 형태는?", 
                                "예: 마크다운 표 형식, 블머머리 기호 목록.")

        # --- 구분선 (Separator) ---
        ttk.Separator(main_frame, orient='horizontal').pack(fill='x', pady=10)

        # --- 버튼 영역 (Button Area) ---
        # 버튼들을 가로로 배치하기 위해 별도의 프레임을 만듭니다.
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=5)

        # 프롬프트 생성 버튼
        generate_btn = ttk.Button(btn_frame, text="✨ 프롬프트 생성 (Generate)", command=self.generate_prompt)
        generate_btn.pack(side=tk.LEFT, padx=5, expand=True, fill='x')

        # 초기화 버튼
        clear_btn = ttk.Button(btn_frame, text="초기화 (Clear)", command=self.clear_fields)
        clear_btn.pack(side=tk.LEFT, padx=5)

        # --- 결과 출력 영역 (Result Area) ---
        # 결과 라벨
        ttk.Label(main_frame, text="[ 생성된 프롬프트 ]", font=("Arial", 10, "bold")).pack(anchor='w', pady=(10, 0))
        
        # 텍스트 박스와 스크롤바를 묶는 프레임
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        # 스크롤바 생성
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill='y')

        # 텍스트 위젯 (결과가 표시될 곳)
        self.result_text = tk.Text(text_frame, height=10, yscrollcommand=scrollbar.set, font=("Consolas", 10))
        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 스크롤바와 텍스트 위젯 연결
        scrollbar.config(command=self.result_text.yview)

        # --- 하단 기능 버튼 (Bottom Action Buttons) ---
        # 복사 및 저장 버튼을 위한 하단 프레임
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(fill='x', pady=10)

        # 클립보드 복사 버튼
        copy_btn = ttk.Button(bottom_frame, text="📋 클립보드 복사", command=self.copy_to_clipboard)
        copy_btn.pack(side=tk.LEFT, fill='x', expand=True, padx=5)

        # 파일 저장 버튼
        save_btn = ttk.Button(bottom_frame, text="💾 파일로 저장", command=self.save_to_file)
        save_btn.pack(side=tk.LEFT, fill='x', expand=True, padx=5)

    def create_label_entry(self, parent, label_text, placeholder="", height=3):
        """
        UI 구성을 돕는 헬퍼 함수입니다. 라벨과 입력창(Text)을 세트로 생성합니다.
        
        Args:
            parent: 위젯이 배치될 부모 프레임
            label_text: 라벨에 표시될 텍스트
            placeholder: (구현 편의상 라벨 옆에 힌트로 표시)
            height: 입력창의 높이 (줄 수)
        """
        # 개별 입력 구역을 감싸는 프레임
        frame = ttk.Frame(parent)
        frame.pack(fill='x', pady=2)

        # 라벨 표시
        lbl = ttk.Label(frame, text=label_text, font=("Arial", 9, "bold"))
        lbl.pack(anchor='w')
        
        # 힌트 텍스트 (작게 표시)
        if placeholder:
            hint = ttk.Label(frame, text=f"  ({placeholder})", font=("Arial", 8), foreground="gray")
            hint.pack(anchor='w')

        # 입력창 (Text) - 여러 줄 입력 가능
        text_widget = tk.Text(frame, height=height, font=("Arial", 10))
        text_widget.pack(fill='x', pady=(2, 5))
        
        return text_widget

    def generate_prompt(self):
        """
        입력된 5가지 요소를 조합하여 완성된 프롬프트를 생성하는 함수입니다.
        명확한 구분을 위해 마크다운 헤더 형식(#)을 사용합니다.
        """
        # 각 요소의 값을 가져옵니다. Text 위젯은 get("1.0", "end-1c") 사용
        role = self.role_entry.get("1.0", "end-1c").strip()
        context = self.context_entry.get("1.0", "end-1c").strip()
        task = self.task_entry.get("1.0", "end-1c").strip()
        constraints = self.constraints_entry.get("1.0", "end-1c").strip()
        format_spec = self.format_entry.get("1.0", "end-1c").strip()

        # 필수 요소(Task)가 비어있으면 경고
        if not task:
            messagebox.showwarning("입력 오류", "최소한 '3. 작업(Task)' 내용은 입력해야 합니다.")
            return

        # 프롬프트 조합 (f-string 사용)
        full_prompt = (
            f"# Role (역할)\n{role if role else '도우미 AI'}\n\n"
            f"# Context (배경)\n{context if context else '없음'}\n\n"
            f"# Task (임무)\n{task}\n\n"
            f"# Constraints (제약사항)\n{constraints if constraints else '없음'}\n\n"
            f"# Output Format (출력 형식)\n{format_spec if format_spec else '자유 형식'}"
        )

        # 결과창 초기화 후 새 텍스트 삽입
        self.result_text.delete(1.0, tk.END) # 기존 텍스트 삭제 (첫 줄부터 끝까지)
        self.result_text.insert(tk.END, full_prompt) # 새 텍스트 삽입

    def copy_to_clipboard(self):
        """
        생성된 프롬프트를 시스템 클립보드에 복사합니다.
        """
        content = self.result_text.get(1.0, tk.END).strip()
        if content:
            self.root.clipboard_clear() # 클립보드 비우기
            self.root.clipboard_append(content) # 내용 추가
            self.root.update() # 시스템에 변경사항 알림
            messagebox.showinfo("성공", "프롬프트가 클립보드에 복사되었습니다!")
        else:
            messagebox.showwarning("경고", "복사할 내용이 없습니다. 먼저 프롬프트를 생성해주세요.")

    def save_to_file(self):
        """
        생성된 프롬프트를 텍스트 파일(.txt)로 저장합니다.
        """
        content = self.result_text.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("경고", "저장할 내용이 없습니다.")
            return

        # 현재 시간을 파일명에 포함시키기 위해 포맷팅
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"prompt_{timestamp}.txt"

        # 파일 저장 대화상자 열기
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            initialfile=default_filename,
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            title="프롬프트 저장"
        )

        # 사용자가 경로를 선택하고 저장을 눌렀을 경우
        if file_path:
            try:
                # UTF-8 인코딩으로 파일 쓰기
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(content)
                messagebox.showinfo("저장 완료", f"파일이 성공적으로 저장되었습니다:\n{file_path}")
            except Exception as e:
                messagebox.showerror("오류", f"파일 저장 중 오류가 발생했습니다:\n{e}")

    def clear_fields(self):
        """
        모든 입력 필드와 결과창을 초기화합니다.
        """
        self.role_entry.delete("1.0", tk.END)
        self.context_entry.delete("1.0", tk.END)
        self.task_entry.delete("1.0", tk.END)
        self.constraints_entry.delete("1.0", tk.END)
        self.format_entry.delete("1.0", tk.END)
        self.result_text.delete(1.0, tk.END)

# --- 메인 실행 블록 (Main Execution Block) ---
if __name__ == "__main__":
    # Tkinter 루트 윈도우 생성
    root = tk.Tk()
    # 애플리케이션 인스턴스 생성
    app = PromptBuilderApp(root)
    # 이벤트 루프 시작 (창이 닫힐 때까지 대기)
    root.mainloop()