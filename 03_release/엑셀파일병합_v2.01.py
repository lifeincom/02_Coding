import os
import threading
import tkinter as tk
import tkinter.ttk as ttk
import tkinter.messagebox as msgbox
from tkinter import filedialog
import pandas as pd

class ExcelMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("엑셀파일 병합 v2.01")
        self.root.geometry("1000x600+20+20")
        self.root.minsize(1000, 600)
        self.root.resizable(False, False)

        # Modern Color Palette
        self.colors = {
            'bg_main': '#F5F7FA',           # Light gray background
            'bg_header': '#2C3E50',         # Dark blue-gray header
            'bg_frame': '#FFFFFF',          # White for frames
            'primary': '#3498DB',           # Blue accent
            'success': '#27AE60',           # Green for success action
            'danger': '#E74C3C',            # Red for delete/close
            'warning': '#F39C12',           # Orange for reset
            'text_dark': '#222222',         # Dark text
            'text_light': '#FFFFFF',        # Light text
            'border': '#BDC3C7',            # Border color
            'input_bg': '#ECF0F1',          # Input background
        }
        
        # Modern Fonts
        self.fonts = {
            'title': ('NanumGothic', 11, 'bold'),
            'button': ('NanumGothic', 10, 'bold'),
            'label': ('NanumGothic', 10),
            'input': ('NanumGothic', 10),
            'desc': ('NanumGothic', 10),
        }
        
        # Configure root background
        self.root.configure(bg=self.colors['bg_main'])
        
        # State variables
        self.load_files = []
        self.output_path = ""

        self._init_ui()

    def _init_ui(self):
        # 1. Top Control Frame (Action Buttons)
        top_frame = tk.Frame(self.root, bg=self.colors['bg_header'], height=60)
        top_frame.pack(fill="x", padx=0, pady=0)
        top_frame.pack_propagate(False)

        # Guide Label (Left side)
        txt = "📊 엑셀파일 병합 도구"
        title_label = tk.Label(
            top_frame, 
            text=txt, 
            font=self.fonts['title'],
            bg=self.colors['bg_header'],
            fg=self.colors['text_light'],
            anchor="w"
        )
        title_label.pack(side="left", padx=15, pady=5)

        # Right side buttons (Action buttons with modern styling)
        btn_close = tk.Button(
            top_frame, 
            text="✕ 닫기", 
            width=10,
            font=self.fonts['button'],
            relief="flat",
            bg=self.colors['danger'],
            fg=self.colors['text_light'],
            activebackground='#C0392B',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.root.quit
        )
        btn_close.pack(side="right", padx=8, pady=12)

        btn_reset = tk.Button(
            top_frame, 
            text="↻ 초기화", 
            width=10,
            font=self.fonts['button'],
            relief="flat",
            bg=self.colors['warning'],
            fg=self.colors['text_light'],
            activebackground='#D68910',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.reset_files
        )
        btn_reset.pack(side="right", padx=5, pady=12)

        btn_del_file = tk.Button(
            top_frame, 
            text="🗑 선택삭제", 
            width=10,
            font=self.fonts['button'],
            relief="flat",
            bg='#E67E22',
            fg=self.colors['text_light'],
            activebackground='#CA6F1E',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.del_file
        )
        btn_del_file.pack(side="right", padx=5, pady=12)

        btn_start = tk.Button(
            top_frame, 
            text="▶ 4.병합 실행", 
            width=12,
            font=self.fonts['button'],
            relief="flat",
            bg=self.colors['success'],
            fg=self.colors['text_light'],
            activebackground='#229954',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.start_merge_thread
        )
        btn_start.pack(side="right", padx=5, pady=12)

        btn_add_file = tk.Button(
            top_frame, 
            text="📁 1.파일 추가", 
            width=12,
            font=self.fonts['button'],
            relief="flat",
            bg=self.colors['primary'],
            fg=self.colors['text_light'],
            activebackground='#2980B9',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.add_file
        )
        btn_add_file.pack(side="right", padx=5, pady=12)

        # 2. Description Frame
        desc_frame = tk.Frame(self.root, bg=self.colors['bg_main'])
        desc_frame.pack(fill="x", padx=15, pady=(10, 5))
        
        txt = "※ 동일한 데이터 구조의 엑셀파일을 하나의 파일로 병합하는 프로그램입니다."
        desc_label = tk.Label(
            desc_frame, 
            text=txt, 
            font=self.fonts['desc'],
            bg=self.colors['bg_main'],
            fg=self.colors['text_dark'],
            anchor="w"
        )
        desc_label.pack(side="left", padx=0, pady=2)

        # 3. Configuration Frame (Save Path & Options)
        config_frame = tk.Frame(self.root, bg=self.colors['bg_frame'], relief="solid", bd=1)
        config_frame.pack(fill="x", padx=15, pady=10)

        # Options (Header Row)
        opt_frame = tk.Frame(config_frame, bg=self.colors['bg_frame'])
        opt_frame.pack(fill="x", padx=15, pady=10)

        tk.Label(
            opt_frame, 
            text="2.제목행 지정 (Starts at 1):", 
            font=self.fonts['label'],
            bg=self.colors['bg_frame'],
            fg=self.colors['text_dark']
        ).pack(side="left", padx=(0, 10))
        
        self.ent_header = tk.Entry(
            opt_frame, 
            width=8,
            font=self.fonts['input'],
            bg=self.colors['input_bg'],
            fg=self.colors['text_dark'],
            relief="flat",
            bd=1
        )
        self.ent_header.insert(0, "1")
        self.ent_header.pack(side="left", padx=5, pady=5, ipady=5)

        txt = "※ 엑셀시트의 제목행을 지정합니다."
        opt_label = tk.Label(
            opt_frame, 
            text=txt, 
            font=self.fonts['desc'],
            bg=self.colors['bg_frame'],
            fg=self.colors['text_dark'],
            anchor="w"
        )
        opt_label.pack(side="left", padx=15)

        # Separator
        sep1 = tk.Frame(config_frame, height=1, bg=self.colors['border'])
        sep1.pack(fill="x", padx=15)

        # Save Path
        path_frame = tk.Frame(config_frame, bg=self.colors['bg_frame'])
        path_frame.pack(fill="x", padx=15, pady=10)
        
        tk.Label(
            path_frame, 
            text="저장 경로:", 
            font=self.fonts['label'],
            bg=self.colors['bg_frame'],
            fg=self.colors['text_dark']
        ).pack(side="left", padx=(0, 10))
        
        self.txt_save_path = tk.Entry(
            path_frame,
            font=self.fonts['input'],
            bg=self.colors['input_bg'],
            fg=self.colors['text_dark'],
            relief="flat",
            bd=1
        )
        self.txt_save_path.pack(side="left", fill="x", expand=True, padx=5, ipady=5)
        
        btn_dest_path = tk.Button(
            path_frame, 
            text="📂 3.저장위치 & 파일명 지정", 
            width=30,
            font=self.fonts['button'],
            relief="flat",
            bg=self.colors['primary'],
            fg=self.colors['text_light'],
            activebackground='#2980B9',
            activeforeground=self.colors['text_light'],
            cursor="hand2",
            command=self.select_save_path
        )
        btn_dest_path.pack(side="right", padx=5)

        # 4. List Frame
        list_frame = tk.LabelFrame(
            self.root, 
            text=" 📋 병합 파일 리스트 ",
            font=self.fonts['title'],
            bg=self.colors['bg_frame'],
            fg=self.colors['text_dark'],
            relief="solid",
            bd=1
        )
        list_frame.pack(fill="both", expand=True, padx=15, pady=(5, 15))

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side="right", fill="y")

        self.list_file = tk.Listbox(
            list_frame, 
            selectmode="extended", 
            height=15,
            font=self.fonts['input'],
            bg=self.colors['input_bg'],
            fg=self.colors['text_dark'],
            selectbackground=self.colors['primary'],
            selectforeground=self.colors['text_light'],
            relief="flat",
            bd=0,
            highlightthickness=0,
            yscrollcommand=scrollbar.set
        )
        self.list_file.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        scrollbar.config(command=self.list_file.yview)

    def add_file(self):
        files = filedialog.askopenfilenames(
            title="병합 할 파일을 선택하세요",
            filetypes=(("XLSX 파일", "*.xlsx"), ("XLS 파일", "*.xls"), ("모든 파일", "*.*")),
            initialdir=r"다운로드"
        )
        for file in files:
            self.list_file.insert(tk.END, file)
            self.load_files.append(file) # Keep track in list

    def del_file(self):
        # Reverse sort to delete correctly by index
        start_indexes = sorted(self.list_file.curselection(), reverse=True)
        for index in start_indexes:
            self.list_file.delete(index)
            if index < len(self.load_files):
                 del self.load_files[index]

    def reset_files(self):
        self.list_file.delete(0, tk.END)
        self.load_files = []
        self.txt_save_path.delete(0, tk.END)
        self.output_path = ""

    def select_save_path(self):
        self.output_path = filedialog.asksaveasfilename(
            title="저장 할 파일명을 입력하세요",
            defaultextension=".xlsx",
            filetypes=[("XLSX 파일", "*.xlsx"), ("XLS 파일", "*.xls")]
        )
        if self.output_path:
            self.txt_save_path.delete(0, tk.END)
            self.txt_save_path.insert(0, self.output_path)

    def start_merge_thread(self):
        # Validation
        if self.list_file.size() == 0:
            msgbox.showwarning("경고", "병합 할 엑셀 파일을 추가하세요.")
            return

        if len(self.txt_save_path.get()) == 0:
            msgbox.showwarning("경고", "저장 경로를 선택하세요")
            return

        # Header Validation
        try:
            header_val = int(self.ent_header.get())
            if header_val < 1:
                raise ValueError
        except ValueError:
             msgbox.showwarning("경고", "헤더 행은 1 이상의 정수여야 합니다.")
             return

        # Start Thread
        t = threading.Thread(target=self.merge_files, args=(header_val - 1,)) # 0-indexed for pandas
        t.start()
        self.show_loading()

    def show_loading(self):
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("처리중")
        self.loading_win.geometry("350x120")
        self.loading_win.resizable(False, False)
        self.loading_win.configure(bg=self.colors['bg_frame'])
        
        # Center the window
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 175
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 60
        self.loading_win.geometry(f"+{x}+{y}")

        self.loading_win.transient(self.root) # Stay on top of main window
        self.loading_win.grab_set() # Modal

        # Loading content frame
        content_frame = tk.Frame(self.loading_win, bg=self.colors['bg_frame'])
        content_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Icon/Title
        title_label = tk.Label(
            content_frame,
            text="⏳ 파일 병합중",
            font=('Segoe UI', 12, 'bold'),
            bg=self.colors['bg_frame'],
            fg=self.colors['primary']
        )
        title_label.pack(pady=(5, 10))

        # Description
        desc_label = tk.Label(
            content_frame,
            text="잠시만 기다려주세요...",
            font=self.fonts['label'],
            bg=self.colors['bg_frame'],
            fg=self.colors['text_dark']
        )
        desc_label.pack(pady=5)
        
        self.root.update()

    def hide_loading(self):
        if hasattr(self, 'loading_win'):
            self.loading_win.destroy()

    def merge_files(self, header_idx):
        try:
            # Re-read files from listbox to be safe (in case load_files state is out of sync, though simpler to use load_files if maintained well)
            # Let's rely on listbox content to be 100% WYSIWYG
            files_to_merge = self.list_file.get(0, tk.END)
            
            dfs = []
            for file in files_to_merge:
                # Auto-detect engine based on file extension
                df = pd.read_excel(file, header=header_idx, engine=None)
                dfs.append(df)

            merged_df = pd.concat(dfs, ignore_index=True)
            
            save_dest = self.txt_save_path.get()
            # Determine engine based on file extension
            if save_dest.endswith('.xls'):
                merged_df.to_excel(save_dest, index=False, engine='xlwt')
            else:
                merged_df.to_excel(save_dest, index=False)

            self.root.after(0, self.finish_merge, True)
        except Exception as e:
            self.root.after(0, self.finish_merge, False, str(e))

    def finish_merge(self, success, error_msg=None):
        self.hide_loading()
        if success:
            msgbox.showinfo("알림", "파일 병합이 완료되었습니다.")
        else:
            msgbox.showerror("에러", f"오류가 발생했습니다: {error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelMergerApp(root)
    root.mainloop()
