"""
DART 보고서 → 엑셀 추출 GUI

여러 개의 DART 원문 zip을 한꺼번에 선택해서
목차를 확인하고, 원하는 섹션의 표를 엑셀로 추출하는 프로그램.

사용법:
    pip install openpyxl
    python dart_gui.py

  1) [파일 추가] 로 zip 여러 개 선택 (Ctrl/Shift 로 다중 선택)
     또는 [폴더 추가] 로 zip 이 든 폴더째 선택
  2) [목차 보기] 로 어떤 섹션이 있는지 확인
  3) 섹션 키워드 입력 (쉼표로 여러 개, 예: 요약재무정보, 매출 및 수주상황)
  4) [엑셀로 추출] → 저장 위치 지정
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

from dart_report_to_excel import load_reports, toc_lines, extract_to_workbook


class DartGui:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("DART 보고서 → 엑셀 추출")
        root.geometry("760x600")
        self.files: list[Path] = []
        self.reports = None  # 파싱 결과 캐시

        # ── 파일 목록 영역 ──
        top = ttk.Frame(root, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="파일 추가", command=self.add_files).pack(side="left")
        ttk.Button(top, text="폴더 추가", command=self.add_folder).pack(side="left", padx=4)
        ttk.Button(top, text="선택 제거", command=self.remove_selected).pack(side="left")
        ttk.Button(top, text="전체 비우기", command=self.clear_files).pack(side="left", padx=4)

        self.listbox = tk.Listbox(root, height=7, selectmode="extended")
        self.listbox.pack(fill="x", padx=8)

        # ── 섹션 입력 + 실행 ──
        mid = ttk.Frame(root, padding=8)
        mid.pack(fill="x")
        ttk.Label(mid, text="추출할 섹션 (쉼표로 구분):").pack(side="left")
        self.section_var = tk.StringVar(value="요약재무정보")
        ttk.Entry(mid, textvariable=self.section_var, width=40).pack(
            side="left", padx=6, fill="x", expand=True)
        self.with_text_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(mid, text="본문 텍스트 포함", variable=self.with_text_var).pack(side="left")

        btns = ttk.Frame(root, padding=(8, 0))
        btns.pack(fill="x")
        self.toc_btn = ttk.Button(btns, text="목차 보기", command=self.show_toc)
        self.toc_btn.pack(side="left")
        self.run_btn = ttk.Button(btns, text="엑셀로 추출", command=self.run_extract)
        self.run_btn.pack(side="left", padx=6)

        # ── 로그/목차 출력 영역 ──
        self.text = tk.Text(root, wrap="none", font=("Consolas", 10))
        self.text.pack(fill="both", expand=True, padx=8, pady=8)
        sy = ttk.Scrollbar(self.text, command=self.text.yview)
        sy.pack(side="right", fill="y")
        self.text.config(yscrollcommand=sy.set)
        self.log("① [파일 추가]/[폴더 추가]로 DART zip을 넣고 ② [목차 보기]로 섹션명을 확인한 뒤 "
                 "③ 섹션 키워드를 입력하고 [엑셀로 추출]을 누르세요.")

    # ── 파일 관리 ──
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="DART 원문 파일 선택 (여러 개 선택 가능)",
            filetypes=[("DART 원문", "*.zip *.xml"), ("모든 파일", "*.*")])
        self._add(paths)

    def add_folder(self):
        d = filedialog.askdirectory(title="zip이 들어있는 폴더 선택")
        if d:
            found = sorted(Path(d).glob("*.zip")) + sorted(Path(d).glob("*.xml"))
            if not found:
                messagebox.showwarning("알림", "폴더에 zip/xml 파일이 없습니다.")
            self._add(found)

    def _add(self, paths):
        for p in paths:
            p = Path(p)
            if p not in self.files:
                self.files.append(p)
                self.listbox.insert("end", p.name)
        self.reports = None

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self.reports = None

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.files.clear()
        self.reports = None

    # ── 출력 ──
    def log(self, msg):
        self.text.insert("end", str(msg) + "\n")
        self.text.see("end")
        self.root.update_idletasks()

    def clear_log(self):
        self.text.delete("1.0", "end")

    # ── 동작 ──
    def _ensure_reports(self):
        if not self.files:
            messagebox.showwarning("알림", "먼저 파일을 추가하세요.")
            return None
        if self.reports is None:
            self.log("보고서 읽는 중... (파일이 크면 몇 초 걸립니다)")
            self.reports = load_reports(self.files, log=self.log)
            self.log(f"→ {len(self.reports)}개 보고서 로드 완료\n")
        if not self.reports:
            messagebox.showerror("오류", "읽을 수 있는 보고서가 없습니다.")
            return None
        return self.reports

    def _busy(self, on):
        state = "disabled" if on else "normal"
        self.toc_btn.config(state=state)
        self.run_btn.config(state=state)

    def show_toc(self):
        self._busy(True)
        def work():
            try:
                reports = self._ensure_reports()
                if reports:
                    self.clear_log()
                    self.log("\n".join(toc_lines(reports)))
            finally:
                self._busy(False)
        threading.Thread(target=work, daemon=True).start()

    def run_extract(self):
        keywords = [s.strip() for s in self.section_var.get().split(",") if s.strip()]
        if not keywords:
            messagebox.showwarning("알림", "추출할 섹션 키워드를 입력하세요.")
            return
        out = filedialog.asksaveasfilename(
            title="엑셀 저장 위치",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile="dart_extract.xlsx")
        if not out:
            return
        self._busy(True)
        def work():
            try:
                reports = self._ensure_reports()
                if not reports:
                    return
                self.log(f"\n추출 시작: {', '.join(keywords)}")
                wb, n = extract_to_workbook(
                    reports, keywords, self.with_text_var.get(), log=self.log)
                if n == 0:
                    messagebox.showwarning(
                        "알림", "추출된 섹션이 없습니다.\n[목차 보기]로 정확한 섹션명을 확인하세요.")
                    return
                wb.save(out)
                self.log(f"\n완료: {out} ({n}개 시트)")
                messagebox.showinfo("완료", f"저장했습니다.\n{out}\n({n}개 시트)")
            except Exception as e:
                self.log(f"[오류] {e}")
                messagebox.showerror("오류", str(e))
            finally:
                self._busy(False)
        threading.Thread(target=work, daemon=True).start()


def main():
    root = tk.Tk()
    DartGui(root)
    root.mainloop()


if __name__ == "__main__":
    main()
