#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extcode 함수 테스트 UI
- 입력: 상품 코드
- 출력: 변환된 상품 코드
- 기능: 단일/일괄 테스트, 결과 복사, 히스토리 저장
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from extcode import extcode
import json
import os
from datetime import datetime


class ExtcodeTestUI:
    def __init__(self, root):
        self.root = root
        self.root.title("🔧 상품 코드 변환 테스트")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        
        # 배경색 설정
        self.root.configure(bg='#f5f5f5')
        
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        
        # 커스텀 스타일 설정
        style.configure('Title.TLabel', font=("맑은 고딕", 16, "bold"), foreground='#2c3e50')
        style.configure('Success.TLabel', foreground='#27ae60', font=("맑은 고딕", 9, "bold"))
        style.configure('Fail.TLabel', foreground='#e74c3c', font=("맑은 고딕", 9, "bold"))
        style.configure('Stats.TLabel', font=("맑은 고딕", 10), foreground='#34495e')
        
        # 버튼 스타일
        style.configure('Action.TButton', font=("맑은 고딕", 10))
        style.configure('Primary.TButton', font=("맑은 고딕", 10, "bold"))
        
        # 메인 프레임
        main_frame = ttk.Frame(root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # 제목
        title_label = ttk.Label(
            main_frame, 
            text="🔧 상품 코드 변환 테스트", 
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 20))
        
        # 입력 섹션
        input_frame = ttk.LabelFrame(main_frame, text="입력", padding="10")
        input_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 단일 입력
        single_input_frame = ttk.Frame(input_frame)
        single_input_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(single_input_frame, text="상품 코드:", font=("맑은 고딕", 10)).pack(side=tk.LEFT, padx=(0, 10))
        self.code_entry = ttk.Entry(single_input_frame, width=35, font=("맑은 고딕", 11))
        self.code_entry.pack(side=tk.LEFT, padx=(0, 10))
        self.code_entry.bind('<Return>', lambda e: self.test_single())
        
        test_btn = ttk.Button(single_input_frame, text="▶ 변환", command=self.test_single, style='Primary.TButton')
        test_btn.pack(side=tk.LEFT)
        
        # 일괄 입력
        batch_label = ttk.Label(input_frame, text="📋 일괄 테스트 (한 줄에 하나씩):", font=("맑은 고딕", 10))
        batch_label.pack(anchor=tk.W, pady=(10, 5))
        
        self.batch_text = scrolledtext.ScrolledText(
            input_frame, 
            height=5, 
            width=50,
            font=("맑은 고딕", 10)
        )
        self.batch_text.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        batch_btn_frame = ttk.Frame(input_frame)
        batch_btn_frame.pack(fill=tk.X)
        
        ttk.Button(
            batch_btn_frame, 
            text="▶ 일괄 변환", 
            command=self.test_batch,
            style='Primary.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            batch_btn_frame, 
            text="🔄 전체 초기화", 
            command=self.clear_inputs,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        # 결과 섹션
        result_frame = ttk.LabelFrame(main_frame, text="결과", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 결과 표시 영역
        result_container = ttk.Frame(result_frame)
        result_container.pack(fill=tk.BOTH, expand=True)
        
        # 트리뷰 생성
        columns = ("입력 코드", "변환 결과", "상태")
        self.result_tree = ttk.Treeview(result_container, columns=columns, show="headings", height=15)
        
        # 컬럼 설정
        self.result_tree.heading("입력 코드", text="입력 코드")
        self.result_tree.heading("변환 결과", text="변환 결과")
        self.result_tree.heading("상태", text="상태")
        
        self.result_tree.column("입력 코드", width=280, anchor=tk.W)
        self.result_tree.column("변환 결과", width=280, anchor=tk.W)
        self.result_tree.column("상태", width=120, anchor=tk.CENTER)
        
        # 태그 설정 (색상 구분)
        self.result_tree.tag_configure("success", background="#d4edda", foreground="#155724")
        self.result_tree.tag_configure("fail", background="#f8d7da", foreground="#721c24")
        
        # 스크롤바
        scrollbar = ttk.Scrollbar(result_container, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscrollcommand=scrollbar.set)
        
        self.result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 우클릭 메뉴
        self.context_menu = tk.Menu(root, tearoff=0)
        self.context_menu.add_command(label="결과 복사", command=self.copy_selected_result)
        self.context_menu.add_command(label="전체 결과 복사", command=self.copy_all_results)
        self.result_tree.bind("<Button-3>", self.show_context_menu)
        
        # 하단 버튼
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame, 
            text="🗑️ 결과 초기화", 
            command=self.clear_results,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame, 
            text="📋 전체 결과 복사", 
            command=self.copy_all_results,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame, 
            text="💾 히스토리 저장", 
            command=self.save_history,
            style='Action.TButton'
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame, 
            text="📂 히스토리 불러오기", 
            command=self.load_history,
            style='Action.TButton'
        ).pack(side=tk.LEFT)
        
        # 통계 정보
        stats_frame = ttk.Frame(main_frame)
        stats_frame.pack(pady=(10, 0), fill=tk.X)
        
        self.stats_label = ttk.Label(
            stats_frame, 
            text="📊 총 테스트: 0 | ✅ 성공: 0 | ❌ 실패: 0",
            style='Stats.TLabel'
        )
        self.stats_label.pack()
        
        # 통계 변수
        self.total_count = 0
        self.success_count = 0
        self.fail_count = 0
        
    def test_single(self):
        """단일 코드 테스트"""
        code = self.code_entry.get().strip()
        if not code:
            messagebox.showwarning("경고", "상품 코드를 입력해주세요.")
            return
        
        result = extcode(code)
        status = "✅ 성공" if result else "❌ 실패"
        tag = "success" if result else "fail"
        
        # 결과 추가
        self.result_tree.insert("", tk.END, values=(code, result, status), tags=(tag,))
        
        # 통계 업데이트
        self.total_count += 1
        if result:
            self.success_count += 1
        else:
            self.fail_count += 1
        self.update_stats()
        
        # 입력 필드 초기화
        self.code_entry.delete(0, tk.END)
        self.code_entry.focus()
    
    def test_batch(self):
        """일괄 코드 테스트"""
        text = self.batch_text.get("1.0", tk.END).strip()
        if not text:
            messagebox.showwarning("경고", "테스트할 코드를 입력해주세요.")
            return
        
        codes = [line.strip() for line in text.split('\n') if line.strip()]
        if not codes:
            messagebox.showwarning("경고", "유효한 코드가 없습니다.")
            return
        
        # 결과 추가
        for code in codes:
            result = extcode(code)
            status = "✅ 성공" if result else "❌ 실패"
            tag = "success" if result else "fail"
            self.result_tree.insert("", tk.END, values=(code, result, status), tags=(tag,))
            
            # 통계 업데이트
            self.total_count += 1
            if result:
                self.success_count += 1
            else:
                self.fail_count += 1
        
        self.update_stats()
        messagebox.showinfo("완료", f"{len(codes)}개의 코드를 테스트했습니다.")
    
    def clear_inputs(self):
        """입력 필드 및 결과 초기화"""
        self.code_entry.delete(0, tk.END)
        self.batch_text.delete("1.0", tk.END)
        # 결과창도 함께 초기화
        self.clear_results()
        self.code_entry.focus()
    
    def clear_results(self):
        """결과 초기화"""
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.total_count = 0
        self.success_count = 0
        self.fail_count = 0
        self.update_stats()
    
    def update_stats(self):
        """통계 업데이트"""
        self.stats_label.config(
            text=f"📊 총 테스트: {self.total_count} | ✅ 성공: {self.success_count} | ❌ 실패: {self.fail_count}"
        )
    
    def show_context_menu(self, event):
        """우클릭 메뉴 표시"""
        item = self.result_tree.identify_row(event.y)
        if item:
            self.result_tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def copy_selected_result(self):
        """선택된 결과 복사"""
        selection = self.result_tree.selection()
        if not selection:
            messagebox.showinfo("알림", "복사할 항목을 선택해주세요.")
            return
        
        item = selection[0]
        values = self.result_tree.item(item, "values")
        if values:
            code, result, _ = values
            text = f"{code} → {result}"
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("완료", "클립보드에 복사되었습니다.")
    
    def copy_all_results(self):
        """전체 결과 복사"""
        items = self.result_tree.get_children()
        if not items:
            messagebox.showinfo("알림", "복사할 결과가 없습니다.")
            return
        
        lines = []
        for item in items:
            values = self.result_tree.item(item, "values")
            if values:
                code, result, _ = values
                lines.append(f"{code}\t{result}")
        
        text = "\n".join(lines)
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        messagebox.showinfo("완료", f"{len(lines)}개의 결과가 클립보드에 복사되었습니다.")
    
    def save_history(self):
        """히스토리 저장"""
        items = self.result_tree.get_children()
        if not items:
            messagebox.showinfo("알림", "저장할 결과가 없습니다.")
            return
        
        history = []
        for item in items:
            values = self.result_tree.item(item, "values")
            if values:
                code, result, status = values
                history.append({
                    "code": code,
                    "result": result,
                    "status": status
                })
        
        # 파일 저장
        filename = f"extcode_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        filepath = os.path.join(os.path.dirname(__file__), filename)
        
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump({
                    "timestamp": datetime.now().isoformat(),
                    "total": len(history),
                    "success": self.success_count,
                    "fail": self.fail_count,
                    "results": history
                }, f, ensure_ascii=False, indent=2)
            
            messagebox.showinfo("완료", f"히스토리가 저장되었습니다.\n{filename}")
        except Exception as e:
            messagebox.showerror("오류", f"저장 중 오류가 발생했습니다.\n{str(e)}")
    
    def load_history(self):
        """히스토리 불러오기"""
        from tkinter import filedialog
        
        filepath = filedialog.askopenfilename(
            title="히스토리 파일 선택",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
            initialdir=os.path.dirname(__file__)
        )
        
        if not filepath:
            return
        
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 결과 초기화
            self.clear_results()
            
            # 결과 추가
            for item in data.get("results", []):
                code = item.get("code", "")
                result = item.get("result", "")
                status_raw = item.get("status", "성공")
                # 상태 포맷팅
                if "성공" in status_raw or "✅" in status_raw:
                    status = "✅ 성공"
                    tag = "success"
                else:
                    status = "❌ 실패"
                    tag = "fail"
                self.result_tree.insert("", tk.END, values=(code, result, status), tags=(tag,))
            
            # 통계 업데이트
            self.total_count = data.get("total", 0)
            self.success_count = data.get("success", 0)
            self.fail_count = data.get("fail", 0)
            self.update_stats()
            
            messagebox.showinfo("완료", f"히스토리를 불러왔습니다.\n{len(data.get('results', []))}개의 결과")
        except Exception as e:
            messagebox.showerror("오류", f"불러오기 중 오류가 발생했습니다.\n{str(e)}")


def main():
    root = tk.Tk()
    app = ExtcodeTestUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
