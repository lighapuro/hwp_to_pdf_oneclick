import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk
import tkinter.messagebox as mb

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except Exception:
    DND_AVAILABLE = False

import pythoncom
import win32com.client

VERSION = "1.0.0"

# ────────────────────────────────────────────────────────────────────────────
# 변환 로직
# ────────────────────────────────────────────────────────────────────────────

def convert_hwp_to_pdf(files, log_callback, done_callback):
    if not files:
        log_callback("변환할 파일이 없습니다.")
        done_callback(0, 0)
        return

    pythoncom.CoInitialize()
    hwp = None
    success, fail = 0, 0
    try:
        hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        hwp.XHwpWindows.Item(0).Visible = False

        for abs_path in files:
            filename = os.path.basename(abs_path)
            pdf_path = os.path.splitext(abs_path)[0] + ".pdf"
            try:
                hwp.Open(abs_path, "HWP", "ForceOpen:1")
                hwp.SaveAs(pdf_path, "PDF", "")
                hwp.Clear(1)
                log_callback(f"[OK] {filename}")
                success += 1
            except Exception as e:
                log_callback(f"[FAIL] {filename}: {e}")
                fail += 1

    except Exception as e:
        log_callback(f"[오류] HWP 초기화 실패: {e}")
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
        done_callback(success, fail)


# ────────────────────────────────────────────────────────────────────────────
# GUI
# ────────────────────────────────────────────────────────────────────────────

_BaseApp: type[tk.Tk] = TkinterDnD.Tk if DND_AVAILABLE else tk.Tk  # type: ignore[assignment]


class App(_BaseApp):  # type: ignore[misc]
    def __init__(self):
        super().__init__()
        self.title(f"HWP → PDF 변환기 v{VERSION}")
        self.resizable(True, True)
        self._file_set = []          # 중복 방지용 절대경로 리스트
        self._running = False
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()

    # ── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── 0행: 버전 표시 ──────────────────────────────────────────────────
        tk.Label(
            self, text=f"v{VERSION}",
            fg="#888888", font=("Malgun Gothic", 8), anchor="e"
        ).pack(fill="x", padx=10, pady=(4, 0))

        # ── 1행: 폴더 선택 ──────────────────────────────────────────────────
        row1 = tk.Frame(self)
        row1.pack(fill="x", padx=10, pady=(10, 2))

        tk.Label(row1, text="폴더 선택:", width=9, anchor="w").pack(side="left")
        self.folder_var = tk.StringVar()
        folder_entry = tk.Entry(row1, textvariable=self.folder_var, width=50)
        folder_entry.pack(side="left", padx=(2, 4), fill="x", expand=True)
        folder_entry.bind("<Return>", lambda _e: self._load_folder_from_entry())
        folder_entry.bind("<FocusOut>", lambda _e: self._load_folder_from_entry())
        tk.Button(row1, text="찾아보기", command=self._browse_folder, width=9).pack(side="left")

        # ── 2행: 파일 선택 ──────────────────────────────────────────────────
        row2 = tk.Frame(self)
        row2.pack(fill="x", padx=10, pady=(2, 6))

        tk.Label(row2, text="파일 선택:", width=9, anchor="w").pack(side="left")
        self.file_count_var = tk.StringVar(value="선택된 파일 없음")
        tk.Label(
            row2, textvariable=self.file_count_var,
            width=38, anchor="w", relief="sunken", bg="white", padx=4
        ).pack(side="left", padx=(2, 4))
        tk.Button(row2, text="파일 추가", command=self._browse_files, width=9).pack(side="left", padx=(0, 4))
        tk.Button(row2, text="목록 초기화", command=self._clear_files, width=9).pack(side="left")

        # ── 3행: 드래그앤드롭 파일 목록 ─────────────────────────────────────
        dnd_hint = (
            "▼ 파일 목록 — HWP 파일을 여기에 끌어다 놓거나, 위 버튼으로 추가하세요"
            if DND_AVAILABLE
            else "▼ 파일 목록  (드래그앤드롭 사용하려면 tkinterdnd2 설치 필요)"
        )
        tk.Label(self, text=dnd_hint, fg="#555555", font=("Malgun Gothic", 9)).pack(
            padx=10, pady=(0, 2), anchor="w"
        )

        list_frame = tk.Frame(self, relief="groove", bd=2, bg="#f0f4f8")
        list_frame.pack(fill="both", expand=False, padx=10, pady=(0, 8))

        sb = tk.Scrollbar(list_frame)
        sb.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            list_frame,
            height=7,
            yscrollcommand=sb.set,
            selectmode=tk.EXTENDED,
            font=("Consolas", 9),
            activestyle="dotbox",
            bg="#f8fafc",
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.config(command=self.listbox.yview)

        # 드래그앤드롭 등록
        if DND_AVAILABLE:
            self.listbox.drop_target_register(DND_FILES)  # type: ignore[attr-defined]
            self.listbox.dnd_bind("<<Drop>>", self._on_drop)  # type: ignore[attr-defined]

        # 우클릭 → 선택 항목 삭제
        ctx = tk.Menu(self, tearoff=0)
        ctx.add_command(label="선택 항목 삭제 (Del)", command=self._delete_selected)
        self.listbox.bind("<Button-3>", lambda e: ctx.tk_popup(e.x_root, e.y_root))
        self.listbox.bind("<Delete>", lambda e: self._delete_selected())

        # ── 4행: 로그 ───────────────────────────────────────────────────────
        self.log = scrolledtext.ScrolledText(
            self, width=72, height=10, state="disabled", font=("Consolas", 9)
        )
        self.log.pack(fill="both", expand=True, padx=10, pady=4)

        # ── 5행: 진행 바 ────────────────────────────────────────────────────
        self.progress = ttk.Progressbar(self, mode="indeterminate")
        self.progress.pack(fill="x", padx=10, pady=(0, 4))

        # ── 6행: 실행 버튼 ──────────────────────────────────────────────────
        self.btn_run = tk.Button(
            self, text="변환 시작", width=16, font=("Malgun Gothic", 10, "bold"),
            command=self._run
        )
        self.btn_run.pack(pady=(0, 12))

    # ── 파일 관리 ─────────────────────────────────────────────────────────────

    def _add_file(self, path):
        """절대경로로 정규화 후 중복 없이 추가"""
        path = os.path.normpath(path)
        if not path.lower().endswith(".hwp"):
            return
        if path in self._file_set:
            return
        self._file_set.append(path)
        self.listbox.insert("end", path)
        self._update_count()

    def _update_count(self):
        n = len(self._file_set)
        self.file_count_var.set(f"{n}개 파일 선택됨" if n else "선택된 파일 없음")

    def _clear_files(self):
        self._file_set.clear()
        self.listbox.delete(0, "end")
        self.folder_var.set("")
        self._update_count()

    def _delete_selected(self):
        indices = list(self.listbox.curselection())
        for i in reversed(indices):
            self._file_set.pop(i)
            self.listbox.delete(i)
        self._update_count()

    # ── 이벤트 핸들러 ─────────────────────────────────────────────────────────

    def _load_folder_from_entry(self):
        """폴더 입력창에 직접 입력/붙여넣기한 경로로 HWP 파일을 스캔"""
        folder = self.folder_var.get().strip()
        if not folder or not os.path.isdir(folder):
            return
        added = 0
        for fname in sorted(os.listdir(folder)):
            if fname.lower().endswith(".hwp"):
                self._add_file(os.path.join(folder, fname))
                added += 1
        if added == 0:
            self._log(f"[알림] 폴더에 HWP 파일이 없습니다: {folder}")

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if not folder:
            return
        self.folder_var.set(folder)
        added = 0
        for fname in sorted(os.listdir(folder)):
            if fname.lower().endswith(".hwp"):
                self._add_file(os.path.join(folder, fname))
                added += 1
        if added == 0:
            self._log(f"[알림] 폴더에 HWP 파일이 없습니다: {folder}")

    def _browse_files(self):
        files = filedialog.askopenfilenames(
            title="HWP 파일 선택",
            filetypes=[("한글 파일", "*.hwp"), ("모든 파일", "*.*")],
        )
        for f in files:
            self._add_file(f)

    def _on_drop(self, event):
        """tkinterdnd2 drop 이벤트 — 공백 포함 경로 파싱"""
        paths = self._parse_dnd_data(event.data)
        skipped = 0
        for p in paths:
            if p.lower().endswith(".hwp"):
                self._add_file(p)
            else:
                skipped += 1
        if skipped:
            self._log(f"[건너뜀] HWP 파일이 아닌 항목 {skipped}개는 무시됐습니다.")

    @staticmethod
    def _parse_dnd_data(raw: str) -> list:
        """'{경로1} {경로2}' 또는 '경로1 경로2' 형식을 모두 처리"""
        paths = re.findall(r'\{([^}]+)\}', raw)
        remainder = re.sub(r'\{[^}]+\}', '', raw).strip()
        if remainder:
            paths.extend(remainder.split())
        return [p.strip() for p in paths if p.strip()]

    # ── 변환 실행 ─────────────────────────────────────────────────────────────

    def _on_close(self):
        if self._running:
            if not mb.askokcancel(
                "변환 중",
                "변환이 진행 중입니다. 종료하시겠습니까?\n"
                "(HWP 프로세스가 백그라운드에 남을 수 있습니다.)"
            ):
                return
        self.destroy()

    def _run(self):
        # 폴더 입력창에 경로가 있으면 먼저 스캔 (붙여넣기 후 버튼 클릭 시에도 동작)
        folder = self.folder_var.get().strip()
        if folder and os.path.isdir(folder):
            self._load_folder_from_entry()

        if not self._file_set:
            mb.showwarning("파일 없음", "변환할 HWP 파일을 먼저 추가해주세요.")
            return

        self._running = True
        self.btn_run.configure(state="disabled")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._log(f"총 {len(self._file_set)}개 파일 변환 시작...\n")
        self.progress.start(10)

        threading.Thread(
            target=convert_hwp_to_pdf,
            args=(list(self._file_set), self._safe_log, self._done),
            daemon=True,
        ).start()

    # ── 로그 헬퍼 ─────────────────────────────────────────────────────────────

    def _log(self, msg):
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _safe_log(self, msg):
        self.after(0, self._log, msg)

    def _done(self, success, fail):
        def _finish():
            self._running = False
            self.progress.stop()
            self._log(f"\n완료: 성공 {success}개, 실패 {fail}개")
            self.btn_run.configure(state="normal")

        self.after(0, _finish)


# ────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    try:
        app = App()
        app.mainloop()
    except Exception as e:
        mb.showerror("오류", str(e))
