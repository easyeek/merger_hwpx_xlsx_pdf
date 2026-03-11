import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import threading
import os
import subprocess
from typing import List

class ExcelMergerPremium:
    def __init__(self, root):
        self.root = root
        self.root.geometry('700x700')
        self.root.title("Excel Premium Merger v2.0")
        self.root.configure(bg="#1e293b")
        
        self.file_paths: List[str] = []
        self.merge_mode = tk.StringVar(value='file') 

        self.setup_ui()

    def setup_ui(self):
        # 스타일 설정
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Premium.TLabelframe", background="#1e293b", foreground="#D4AF37")
        style.configure("Premium.TLabelframe.Label", background="#1e293b", foreground="#D4AF37", font=("Pretendard", 10, "bold"))
        
        # 상단 타이틀
        title_label = tk.Label(self.root, text="엑셀 데이터 통합 프리미엄 도구", font=("Pretendard", 18, "bold"), 
                              bg="#1e293b", fg="#D4AF37", pady=20)
        title_label.pack(fill=tk.X)

        # 0. 모드 선택 섹션
        mode_frame = tk.LabelFrame(self.root, text=" ⚙️ 결합 모드 설정 ", bg="#1e293b", fg="#D4AF37", font=("Pretendard", 10, "bold"), padx=10, pady=10)
        mode_frame.pack(fill=tk.X, padx=20, pady=5)

        tk.Radiobutton(mode_frame, text="파일별 첫 시트 결합", variable=self.merge_mode, value='file', 
                      bg="#1e293b", fg="#f1f5f9", selectcolor="#0f172a", activebackground="#1e293b", activeforeground="#D4AF37").pack(side=tk.LEFT, padx=20)
        tk.Radiobutton(mode_frame, text="모든 파일의 모든 시트 결합", variable=self.merge_mode, value='sheet', 
                      bg="#1e293b", fg="#f1f5f9", selectcolor="#0f172a", activebackground="#1e293b", activeforeground="#D4AF37").pack(side=tk.LEFT, padx=20)
        
        # 1. 파일 목록 섹션
        file_frame = tk.LabelFrame(self.root, text=" 📂 결합 대상 파일 목록 ", bg="#1e293b", fg="#D4AF37", font=("Pretendard", 10, "bold"), padx=10, pady=10)
        file_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        list_container = tk.Frame(file_frame, bg="#1e293b")
        list_container.pack(fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(list_container)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(list_container, selectmode=tk.SINGLE, height=8,
                                     bg="#0f172a", fg="#f1f5f9", borderwidth=0, highlightthickness=1, 
                                     highlightbackground="#334155", selectbackground="#D4AF37", selectforeground="black",
                                     yscrollcommand=self.scrollbar.set)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.file_listbox.yview)

        # 파일 제어 버튼
        btn_frame = tk.Frame(file_frame, bg="#1e293b")
        btn_frame.pack(fill=tk.X, pady=10)
        
        self.add_btn = tk.Button(btn_frame, text="+ 파일 추가", command=self.select_files, bg="#D4AF37", fg="black", font=("Pretendard", 9, "bold"), width=12, relief=tk.FLAT)
        self.add_btn.pack(side=tk.LEFT, padx=5)
        
        self.up_btn = tk.Button(btn_frame, text="▲ 위로", command=self.move_file_up, bg="#334155", fg="white", width=8, relief=tk.FLAT)
        self.up_btn.pack(side=tk.LEFT, padx=5)
        
        self.down_btn = tk.Button(btn_frame, text="▼ 아래로", command=self.move_file_down, bg="#334155", fg="white", width=8, relief=tk.FLAT)
        self.down_btn.pack(side=tk.LEFT, padx=5)

        self.del_btn = tk.Button(btn_frame, text="선택 삭제", command=self.remove_file, bg="#334155", fg="white", width=10, relief=tk.FLAT)
        self.del_btn.pack(side=tk.LEFT, padx=5)

        self.clear_btn = tk.Button(btn_frame, text="전체 비우기", command=self.clear_all, bg="#e11d48", fg="white", width=10, relief=tk.FLAT)
        self.clear_btn.pack(side=tk.RIGHT, padx=5)

        # 2. 데이터 미리보기 섹션 (Row=2)
        preview_frame = tk.LabelFrame(self.root, text=" 👀 데이터 미리보기 (첫 번째 파일 상위 5행) ", bg="#1e293b", fg="#D4AF37", font=("Pretendard", 10, "bold"), padx=10, pady=10)
        preview_frame.pack(fill=tk.X, padx=20, pady=5)

        self.preview_tree = ttk.Treeview(preview_frame, columns=("No Data"), show='headings', height=5)
        self.preview_tree.pack(fill=tk.X)

        # 3. 진행 섹션
        progress_frame = tk.Frame(self.root, bg="#1e293b")
        progress_frame.pack(fill=tk.X, padx=20, pady=10)

        self.status_var = tk.StringVar(value="대기 중...")
        self.status_label = tk.Label(progress_frame, textvariable=self.status_var, bg="#1e293b", fg="#94a3b8", font=("Pretendard", 9))
        self.status_label.pack(side=tk.LEFT)

        self.progress = ttk.Progressbar(progress_frame, orient='horizontal', mode='determinate')
        self.progress.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=(10, 0))

        # 4. 결합 시작 버튼
        self.merge_button = tk.Button(self.root, text="통합 엑셀 파일 생성 시작", command=self.start_merge_thread,
                                     bg="#D4AF37", fg="black", font=("Pretendard", 12, "bold"), height=2, relief=tk.RAISED)
        self.merge_button.pack(fill=tk.X, padx=20, pady=20)

    def set_status(self, msg):
        self.status_var.set(msg)
        self.root.update_idletasks()

    def update_preview(self):
        # 트리뷰 초기화
        for i in self.preview_tree.get_children():
            self.preview_tree.delete(i)
        
        if not self.file_paths:
            return

        try:
            first_file = self.file_paths[0]
            ext = os.path.splitext(first_file)[1].lower()
            if ext == '.csv':
                try: df = pd.read_csv(first_file, nrows=5, encoding='utf-8')
                except: df = pd.read_csv(first_file, nrows=5, encoding='cp949')
            else:
                df = pd.read_excel(first_file, nrows=5)
            
            self.preview_tree["columns"] = list(df.columns)
            for col in df.columns:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=100, anchor=tk.W)
            
            for index, row in df.iterrows():
                self.preview_tree.insert("", tk.END, values=list(row))
        except Exception as e:
            self.set_status(f"미리보기 로드 실패: {str(e)}")

    def select_files(self):
        new_paths = filedialog.askopenfilenames(
            filetypes=[("Excel & CSV", "*.xls *.xlsx *.csv"), ("모든 파일", "*.*")]
        )
        if new_paths:
            for path in new_paths:
                if path not in self.file_paths:
                    self.file_paths.append(path)
            self.update_file_listbox()
            self.update_preview()
            self.set_status(f"{len(new_paths)}개 파일 추가됨")

    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for path in self.file_paths:
            self.file_listbox.insert(tk.END, os.path.basename(path))

    def remove_file(self):
        sel = self.file_listbox.curselection()
        if sel:
            del self.file_paths[sel[0]]
            self.update_file_listbox()
            self.update_preview()

    def clear_all(self):
        if messagebox.askyesno("확인", "모든 목록을 삭제할까요?"):
            self.file_paths = []
            self.update_file_listbox()
            self.update_preview()

    def move_file_up(self):
        sel = self.file_listbox.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            self.file_paths[idx], self.file_paths[idx-1] = self.file_paths[idx-1], self.file_paths[idx]
            self.update_file_listbox()
            self.file_listbox.select_set(idx-1)

    def move_file_down(self):
        sel = self.file_listbox.curselection()
        if sel and sel[0] < len(self.file_paths) - 1:
            idx = sel[0]
            self.file_paths[idx], self.file_paths[idx+1] = self.file_paths[idx+1], self.file_paths[idx]
            self.update_file_listbox()
            self.file_listbox.select_set(idx+1)

    def start_merge_thread(self):
        if not self.file_paths:
            messagebox.showwarning("경고", "파일을 먼저 추가해주세요.")
            return
        self.merge_button.config(state=tk.DISABLED)
        threading.Thread(target=self.merge_process, daemon=True).start()

    def merge_process(self):
        try:
            self.set_status("데이터 분석 및 통합 중...")
            all_dfs = []
            mode = self.merge_mode.get()
            total = len(self.file_paths)
            
            for i, fp in enumerate(self.file_paths, 1):
                self.set_status(f"처리 중 ({i}/{total}): {os.path.basename(fp)}")
                ext = os.path.splitext(fp)[1].lower()
                
                if ext in ['.xls', '.xlsx']:
                    engine = 'xlrd' if ext == '.xls' else 'openpyxl'
                    xls = pd.ExcelFile(fp, engine=engine)
                    sheets = xls.sheet_names if mode == 'sheet' else [xls.sheet_names[0]]
                    for s in sheets:
                        all_dfs.append(pd.read_excel(xls, sheet_name=s))
                elif ext == '.csv':
                    try: df = pd.read_csv(fp, encoding='utf-8')
                    except: df = pd.read_csv(fp, encoding='cp949')
                    all_dfs.append(df)
                
                self.progress['value'] = (i / total) * 100
                self.root.update_idletasks()

            merged_df = pd.concat(all_dfs, ignore_index=True)
            
            self.set_status("파일 저장 대기 중...")
            save_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                    filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
                                                    initialfile="데이터통합결과.xlsx")
            if save_path:
                if save_path.endswith('.csv'):
                    merged_df.to_csv(save_path, index=False, encoding='utf-8-sig')
                else:
                    merged_df.to_excel(save_path, index=False)
                
                self.set_status("작업 성공!")
                if messagebox.askyesno("완료", "통합 작업이 완료되었습니다!\n결과 폴더를 열까요?"):
                    subprocess.run(['explorer', f'/select,{os.path.normpath(save_path)}'])
            else:
                self.set_status("취소됨")

        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.set_status("실패")
        finally:
            self.merge_button.config(state=tk.NORMAL)
            self.progress['value'] = 0

if __name__ == "__main__":
    root = tk.Tk()
    # 화면 중앙 배치
    w, h = 700, 700
    root.geometry(f"{w}x{h}+{(root.winfo_screenwidth()-w)//2}+{(root.winfo_screenheight()-h)//2}")
    app = ExcelMergerPremium(root)
    root.mainloop()
