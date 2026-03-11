import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import win32com.client as win32
import zipfile
import xml.etree.ElementTree as ET
import time
import subprocess

class HwpMergerGUI:
    def __init__(self, master):
        self.master = master
        master.title("HWP/HWPX Premium Merger v2.0")
        master.geometry("600x650")
        master.configure(bg="#1e293b")
        
        self.file_list = []

        # 상단 타이틀
        title_label = tk.Label(master, text="한글 문서 병합 프리미엄 도구", font=("Pretendard", 16, "bold"), 
                              bg="#1e293b", fg="#D4AF37", pady=20)
        title_label.pack()

        # 파일 선택 및 전체 삭제 버튼 프레임
        top_btn_frame = tk.Frame(master, bg="#1e293b")
        top_btn_frame.pack(pady=5)

        self.select_button = tk.Button(top_btn_frame, text="+ 파일 추가", command=self.select_files,
                                      bg="#D4AF37", fg="black", font=("Pretendard", 10, "bold"),
                                      padx=15, pady=5, relief=tk.FLAT)
        self.select_button.pack(side=tk.LEFT, padx=5)

        self.clear_button = tk.Button(top_btn_frame, text="모두 비우기", command=self.clear_all,
                                     bg="#e11d48", fg="white", font=("Pretendard", 10),
                                     padx=15, pady=5, relief=tk.FLAT)
        self.clear_button.pack(side=tk.LEFT, padx=5)

        # 리스트 영역 (스크롤바 포함)
        list_container = tk.Frame(master, bg="#1e293b")
        list_container.pack(padx=20, pady=10, fill=tk.BOTH, expand=True)

        self.scrollbar = tk.Scrollbar(list_container)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(list_container, width=60, height=15, 
                                 bg="#0f172a", fg="#f1f5f9", font=("Pretendard", 10),
                                 selectbackground="#D4AF37", selectforeground="black",
                                 borderwidth=0, highlightthickness=1, highlightbackground="#334155",
                                 yscrollcommand=self.scrollbar.set)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.listbox.yview)

        # 더블 클릭으로 삭제 기능 추가
        self.listbox.bind('<Double-Button-1>', lambda e: self.delete_file())

        # 안내 문구
        guide_label = tk.Label(master, text="* 파일을 위/아래로 드래그하여 순서를 변경하거나, 더블 클릭하여 삭제하세요.", 
                              font=("Pretendard", 9), bg="#1e293b", fg="#94a3b8")
        guide_label.pack(pady=5)

        # 컨트롤 버튼 박스
        ctrl_frame = tk.Frame(master, bg="#1e293b")
        ctrl_frame.pack(pady=10)

        self.up_button = tk.Button(ctrl_frame, text="▲ 위로", command=self.move_up,
                                  bg="#334155", fg="white", width=10, relief=tk.FLAT)
        self.up_button.pack(side=tk.LEFT, padx=5)

        self.down_button = tk.Button(ctrl_frame, text="▼ 아래로", command=self.move_down,
                                    bg="#334155", fg="white", width=10, relief=tk.FLAT)
        self.down_button.pack(side=tk.LEFT, padx=5)

        self.delete_button = tk.Button(ctrl_frame, text="선택 삭제", command=self.delete_file,
                                      bg="#334155", fg="white", width=10, relief=tk.FLAT)
        self.delete_button.pack(side=tk.LEFT, padx=5)

        # 병합 실행 버튼
        self.merge_button = tk.Button(master, text="병합 프로세스 시작", command=self.merge_files,
                                     bg="#D4AF37", fg="black", font=("Pretendard", 12, "bold"),
                                     width=30, height=2, relief=tk.RAISED)
        self.merge_button.pack(pady=20)

        # 상태 표시줄
        self.status_var = tk.StringVar()
        self.status_var.set("준비됨")
        self.status_bar = tk.Label(master, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W,
                                  bg="#0f172a", fg="#94a3b8", font=("Pretendard", 9), padx=10, pady=5)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def set_status(self, msg):
        self.status_var.set(msg)
        self.master.update_idletasks()

    def clear_all(self):
        if messagebox.askyesno("확인", "목록을 모두 비우시겠습니까?"):
            self.file_list = []
            self.update_listbox()
            self.set_status("목록이 비워졌습니다.")

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("한글 문서", "*.hwp;*.hwpx")])
        if files:
            self.file_list.extend(files)
            self.update_listbox()
            self.set_status(f"{len(files)}개 파일이 추가되었습니다.")

    def update_listbox(self):
        self.listbox.delete(0, tk.END)
        for file in self.file_list:
            self.listbox.insert(tk.END, os.path.basename(file))

    def move_up(self):
        selected = self.listbox.curselection()
        if selected and selected[0] > 0:
            idx = selected[0]
            self.file_list[idx-1], self.file_list[idx] = self.file_list[idx], self.file_list[idx-1]
            self.update_listbox()
            self.listbox.selection_set(idx-1)

    def move_down(self):
        selected = self.listbox.curselection()
        if selected and selected[0] < len(self.file_list) - 1:
            idx = selected[0]
            self.file_list[idx], self.file_list[idx+1] = self.file_list[idx+1], self.file_list[idx]
            self.update_listbox()
            self.listbox.selection_set(idx+1)

    def delete_file(self):
        selected = self.listbox.curselection()
        if selected:
            idx = selected[0]
            name = os.path.basename(self.file_list[idx])
            del self.file_list[idx]
            self.update_listbox()
            self.set_status(f"'{name}' 파일이 제거되었습니다.")

    def merge_files(self):
        if not self.file_list:
            messagebox.showerror("오류", "병합할 파일이 없습니다.")
            return

        total = len(self.file_list)
        output_file = filedialog.asksaveasfilename(defaultextension=".hwp", 
                                                 filetypes=[("HWP 통합 문서", "*.hwp")],
                                                 initialfile="통합결과_문서.hwp")
        if not output_file:
            return

        try:
            self.set_status("한글 엔진(OLE) 연결 중...")
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
            hwp.XHwpWindows.Item(0).Visible = True
            hwp.Open(self.file_list[0])

            for i, file in enumerate(self.file_list[1:], 2):
                self.set_status(f"병합 중... ({i}/{total}): {os.path.basename(file)}")
                hwp.MovePos(3) # 문서 끝
                
                if file.lower().endswith('.hwp'):
                    hwp.HAction.GetDefault("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
                    hwp.HParameterSet.HInsertFile.filename = file
                    hwp.HParameterSet.HInsertFile.KeepSection = 1
                    hwp.HAction.Execute("InsertFile", hwp.HParameterSet.HInsertFile.HSet)
                elif file.lower().endswith('.hwpx'):
                    content = self.extract_hwpx_content(file)
                    hwp.HAction.GetDefault("InsertText", hwp.HParameterSet.HInsertText.HSet)
                    hwp.HParameterSet.HInsertText.Text = content
                    hwp.HAction.Execute("InsertText", hwp.HParameterSet.HInsertText.HSet)

            hwp.SaveAs(output_file)
            time.sleep(1)
            hwp.Quit()
            
            self.set_status("병합 완료!")
            if messagebox.askyesno("완료", "병합이 성공적으로 완료되었습니다.\n저장된 폴더를 여시겠습니까?"):
                folder = os.path.dirname(os.path.abspath(output_file))
                subprocess.run(['explorer', f'/select,{os.path.normpath(output_file)}'])

        except Exception as e:
            self.set_status("오류 발생")
            messagebox.showerror("오류", f"병합 도중 문제가 발생했습니다:\n{str(e)}")

    def extract_hwpx_content(self, hwpx_file):
        try:
            with zipfile.ZipFile(hwpx_file, 'r') as zip_ref:
                with zip_ref.open('Contents/section0.xml') as section_file:
                    tree = ET.parse(section_file)
                    root = tree.getroot()
                    content = ""
                    ns = {'hml': 'http://www.hancom.co.kr/schema/2011/hml'}
                    for paragraph in root.findall('.//hml:p', ns):
                        for text in paragraph.findall('.//hml:t', ns):
                            if text.text:
                                content += text.text
                        content += "\n"
                    return content
        except Exception as e:
            print(f"HWPX 파싱 오류: {e}")
            return ""

if __name__ == "__main__":
    root = tk.Tk()
    # 윈도우 중앙 배치
    window_width = 600
    window_height = 650
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    app = HwpMergerGUI(root)
    root.mainloop()
