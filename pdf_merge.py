import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from pypdf import PdfWriter

class PdfMergerGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("PDF Premium Merger v2.0")
        self.master.configure(bg="#0f172a")
        
        self.file_list = []
        self.setup_ui()

    def setup_ui(self):
        # 상단 타이틀
        title_label = tk.Label(self.master, text="PDF 문서 병합 프리미엄 도구", font=("Pretendard", 18, "bold"), 
                              bg="#0f172a", fg="#D4AF37", pady=25)
        title_label.pack()

        # 버튼 프레임
        top_btn_frame = tk.Frame(self.master, bg="#0f172a")
        top_btn_frame.pack(pady=5)

        self.select_button = tk.Button(top_btn_frame, text="+ PDF 추가", command=self.select_files,
                                      bg="#D4AF37", fg="#0f172a", font=("Pretendard", 10, "bold"),
                                      padx=20, pady=8, relief=tk.FLAT, cursor="hand2")
        self.select_button.pack(side=tk.LEFT, padx=10)

        self.clear_button = tk.Button(top_btn_frame, text="모두 비우기", command=self.clear_all,
                                     bg="#1e293b", fg="#f1f5f9", font=("Pretendard", 10),
                                     padx=20, pady=8, relief=tk.FLAT, cursor="hand2")
        self.clear_button.pack(side=tk.LEFT, padx=10)

        # 리스트 영역
        list_container = tk.Frame(self.master, bg="#0f172a")
        list_container.pack(padx=30, pady=15, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(list_container, bg="#1e293b", troughcolor="#0f172a")
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(list_container, width=60, height=12, 
                                 bg="#1e293b", fg="#f1f5f9", font=("Pretendard", 10),
                                 selectbackground="#D4AF37", selectforeground="#0f172a",
                                 borderwidth=0, highlightthickness=1, highlightbackground="#334155",
                                 yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)

        # 컨트롤 버튼
        ctrl_frame = tk.Frame(self.master, bg="#0f172a")
        ctrl_frame.pack(pady=10)
        btn_style = {"bg": "#334155", "fg": "#f1f5f9", "relief": tk.FLAT, "width": 10, "font": ("Pretendard", 9), "cursor": "hand2"}
        
        tk.Button(ctrl_frame, text="▲ 위로", command=self.move_up, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_frame, text="▼ 아래로", command=self.move_down, **btn_style).pack(side=tk.LEFT, padx=5)
        tk.Button(ctrl_frame, text="선택 삭제", command=self.delete_file, **btn_style).pack(side=tk.LEFT, padx=5)

        # 병합 버튼
        self.merge_button = tk.Button(self.master, text="PDF 병합 시작", command=self.merge_files,
                                     bg="#D4AF37", fg="#0f172a", font=("Pretendard", 13, "bold"),
                                     width=35, height=2, relief=tk.FLAT, cursor="hand2")
        self.merge_button.pack(pady=25)

        # 상태바
        self.status_var = tk.StringVar(value="준비됨")
        tk.Label(self.master, textvariable=self.status_var, bd=0, relief=tk.FLAT, anchor=tk.W,
                 bg="#1e293b", fg="#94a3b8", font=("Pretendard", 9), padx=15, pady=8).pack(side=tk.BOTTOM, fill=tk.X)

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF 문서", "*.pdf")])
        if files:
            self.file_list.extend(files)
            self.update_listbox()
            self.status_var.set(f"{len(files)}개 파일 추가됨")

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for f in self.file_list:
            self.listbox.insert(tk.END, f"  📕 {os.path.basename(f)}")

    def clear_all(self):
        if self.file_list and messagebox.askyesno("확인", "목록을 비우시겠습니까?"):
            self.file_list = []
            self.update_listbox()

    def move_up(self):
        sel = self.listbox.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            self.file_list[idx-1], self.file_list[idx] = self.file_list[idx], self.file_list[idx-1]
            self.update_listbox()
            self.listbox.selection_set(idx-1)

    def move_down(self):
        sel = self.listbox.curselection()
        if sel and sel[0] < len(self.file_list)-1:
            idx = sel[0]
            self.file_list[idx], self.file_list[idx+1] = self.file_list[idx+1], self.file_list[idx]
            self.update_listbox()
            self.listbox.selection_set(idx+1)

    def delete_file(self):
        sel = self.listbox.curselection()
        if sel:
            del self.file_list[sel[0]]
            self.update_listbox()

    def merge_files(self):
        if not self.file_list: return
        out_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], initialfile="통합결과.pdf")
        if not out_file: return
        
        try:
            merger = PdfWriter()
            for f in self.file_list:
                merger.append(f)
            merger.write(out_file)
            merger.close()
            messagebox.showinfo("완료", "PDF 병합이 완료되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    PdfMergerGUI(root)
    root.mainloop()
