import tkinter as tk
from tkinter import messagebox
from hwpx_merge import HwpMergerGUI
from xlsx_merge import ExcelMergerPremium
from pdf_merge import PdfMergerGUI
import os

class MainDashboard:
    def __init__(self, root):
        self.root = root
        self.root.title("Premium Data Merger Hub v2.0")
        self.root.geometry("1100x550")
        self.root.configure(bg="#0f172a")
        self.root.resizable(False, False)

        self.title_font = ("Pretendard", 24, "bold")
        self.desc_font = ("Pretendard", 11)
        self.btn_font = ("Pretendard", 14, "bold")

        self.setup_ui()

    def setup_ui(self):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 1100) // 2
        y = (screen_height - 550) // 4
        self.root.geometry(f"1100x550+{x}+{y}")

        header_frame = tk.Frame(self.root, bg="#0f172a", pady=40)
        header_frame.pack(fill=tk.X)

        tk.Label(header_frame, text="PREMIUM MERGER HUB", font=self.title_font, 
                 bg="#0f172a", fg="#D4AF37").pack()
        tk.Label(header_frame, text="한글, 엑셀, PDF 문서를 스마트하게 통합하세요", font=self.desc_font, 
                 bg="#0f172a", fg="#94a3b8").pack(pady=5)

        card_frame = tk.Frame(self.root, bg="#0f172a")
        card_frame.pack(pady=10)

        def create_card(parent, title, icon, desc, command):
            card = tk.Frame(parent, bg="#1e293b", padx=2, pady=2)
            card.pack(side=tk.LEFT, padx=15)
            inner = tk.Frame(card, bg="#1e293b", padx=25, pady=30, width=320, height=240)
            inner.pack_propagate(False)
            inner.pack()
            tk.Label(inner, text=icon, font=("Pretendard", 40), bg="#1e293b", fg="#D4AF37").pack()
            tk.Label(inner, text=title, font=self.btn_font, bg="#1e293b", fg="#f1f5f9").pack(pady=10)
            tk.Label(inner, text=desc, font=("Pretendard", 9), bg="#1e293b", fg="#94a3b8").pack()
            tk.Button(inner, text="시작하기", command=command, bg="#D4AF37", fg="#0f172a", 
                      font=("Pretendard", 10, "bold"), relief=tk.FLAT, padx=20, pady=8, cursor="hand2").pack(side=tk.BOTTOM)
            return card

        create_card(card_frame, "한글 문서 병합", "📄", "HWP / HWPX 파일들을\n하나의 문서로 통합합니다", self.open_hwpx_merger)
        create_card(card_frame, "엑셀 데이터 통합", "📊", "여러 개의 XLSX / CSV 시트를\n하나의 파일로 결합합니다", self.open_xlsx_merger)
        create_card(card_frame, "PDF 문서 합치기", "📕", "기본 PDF 파일들을\n순서대로 병합합니다", self.open_pdf_merger)

        # 제작자 표기 수정 (USER 중심)
        footer = tk.Label(self.root, text="v2.0 Premium Edition | Created by USER & Antigravity Collaborative", 
                          bg="#0f172a", fg="#334155", font=("Pretendard", 8))
        footer.pack(side=tk.BOTTOM, pady=30)

    def open_hwpx_merger(self):
        new_win = tk.Toplevel(self.root)
        HwpMergerGUI(new_win)

    def open_xlsx_merger(self):
        new_win = tk.Toplevel(self.root)
        ExcelMergerPremium(new_win)

    def open_pdf_merger(self):
        new_win = tk.Toplevel(self.root)
        PdfMergerGUI(new_win)

if __name__ == "__main__":
    root = tk.Tk()
    app = MainDashboard(root)
    root.mainloop()
