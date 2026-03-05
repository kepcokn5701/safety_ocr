"""
산업안전보건관리비 진위여부 확인 시스템
================================================
tkinter 기반 통합 데스크톱 애플리케이션

실행: python app_gui.py
"""

import os, sys, re, threading, queue, time, random, logging
from pathlib import Path
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# ──────────────────────────────────────────────
# 로깅 → 큐로 보내기 (GUI 스레드 안전)
# ──────────────────────────────────────────────
log_queue: queue.Queue = queue.Queue()


class QueueHandler(logging.Handler):
    """로그 메시지를 queue에 넣는 핸들러"""
    def emit(self, record):
        log_queue.put(self.format(record))


_qh = QueueHandler()
_qh.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S"))
logging.root.addHandler(_qh)
logging.root.setLevel(logging.INFO)

log = logging.getLogger("app_gui")

# ──────────────────────────────────────────────
# 지연 임포트용 래퍼
# ──────────────────────────────────────────────
_qr_module = None
_verify_module = None


def _import_qr():
    global _qr_module
    if _qr_module is None:
        import df_qr_batch as m
        _qr_module = m
    return _qr_module


def _import_verify():
    global _verify_module
    if _verify_module is None:
        import hometax_verify as m
        _verify_module = m
    return _verify_module


# ══════════════════════════════════════════════
# 메인 애플리케이션
# ══════════════════════════════════════════════
class App(tk.Tk):
    """
    ┌─────────────────────────────────────────────────────┐
    │  전자세금계산서  QR 인식 & 진위 확인 시스템           │
    ├─────────────────────────────────────────────────────┤
    │  PDF 폴더: [____________________________] [선택]     │
    │                                                     │
    │  [1단계: QR 추출]  [2단계: 진위 확인]  [Excel 저장]  │
    ├─────────────────────────────────────────────────────┤
    │  ┌ 결과 테이블 ──────────────────────────────────┐  │
    │  │ # | 파일명 | 승인번호 | 공급자 | 일자 | 금액 | 진위 │  │
    │  │ 1 | ...    | ...     | ...   | ... | ... | O  │  │
    │  └───────────────────────────────────────────────┘  │
    ├─────────────────────────────────────────────────────┤
    │  진행: ████████░░░░  60%   (3/5건)                   │
    ├─────────────────────────────────────────────────────┤
    │  로그:                                               │
    │  15:30:01 [INFO] safety.pdf QR 인식 성공             │
    │  15:30:05 [INFO] 모바일 API → O                      │
    └─────────────────────────────────────────────────────┘
    """

    COL_DEF = [
        ("#",         40,  tk.CENTER),
        ("파일명",     140, tk.W),
        ("페이지",     50,  tk.CENTER),
        ("승인번호",   200, tk.W),
        ("공급자번호", 100, tk.CENTER),
        ("작성일자",   90,  tk.CENTER),
        ("합계금액",   100, tk.E),
        ("진위여부",   65,  tk.CENTER),
    ]

    def __init__(self):
        super().__init__()
        self.title("산업안전보건관리비 진위여부 확인 시스템")
        self.geometry("960x680")
        self.minsize(800, 560)
        self.configure(bg="#f0f0f0")

        self.df: pd.DataFrame | None = None
        self.folder_path = tk.StringVar()
        self._running = False

        self._build_ui()
        self._poll_log_queue()

    # ── UI 빌드 ────────────────────────────────────
    def _build_ui(self):
        # === 상단: 타이틀 ===
        hdr = tk.Frame(self, bg="#1a237e", height=48)
        hdr.pack(fill=tk.X)
        hdr.pack_propagate(False)
        tk.Label(
            hdr, text="  산업안전보건관리비 진위여부 확인 시스템",
            bg="#1a237e", fg="white", font=("맑은 고딕", 14, "bold"),
            anchor=tk.W,
        ).pack(fill=tk.BOTH, expand=True, padx=8)

        # === 폴더 선택 ===
        frm_path = tk.Frame(self, bg="#f0f0f0", pady=8)
        frm_path.pack(fill=tk.X, padx=12)

        tk.Label(frm_path, text="PDF 폴더:", bg="#f0f0f0",
                 font=("맑은 고딕", 10)).pack(side=tk.LEFT)
        ent = ttk.Entry(frm_path, textvariable=self.folder_path, width=60)
        ent.pack(side=tk.LEFT, padx=(6, 4), fill=tk.X, expand=True)
        ttk.Button(frm_path, text="찾아보기...", command=self._browse_folder,
                   width=11).pack(side=tk.LEFT)

        # === 버튼 바 ===
        frm_btn = tk.Frame(self, bg="#f0f0f0")
        frm_btn.pack(fill=tk.X, padx=12, pady=(0, 6))

        self.btn_qr = ttk.Button(
            frm_btn, text="1단계: QR 추출", command=self._run_qr_extract, width=18)
        self.btn_qr.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_verify = ttk.Button(
            frm_btn, text="2단계: 진위 확인", command=self._run_verify, width=18)
        self.btn_verify.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_save = ttk.Button(
            frm_btn, text="Excel 저장", command=self._save_excel, width=12)
        self.btn_save.pack(side=tk.LEFT, padx=(0, 6))

        self.btn_open = ttk.Button(
            frm_btn, text="Excel 열기", command=self._open_excel, width=12)
        self.btn_open.pack(side=tk.LEFT, padx=(0, 6))

        # 통계 라벨 (우측)
        self.lbl_stats = tk.Label(
            frm_btn, text="", bg="#f0f0f0", font=("맑은 고딕", 9), fg="#555")
        self.lbl_stats.pack(side=tk.RIGHT)

        # === 결과 테이블 ===
        frm_tree = tk.Frame(self, bg="#f0f0f0")
        frm_tree.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 4))

        cols = [c[0] for c in self.COL_DEF]
        self.tree = ttk.Treeview(frm_tree, columns=cols, show="headings",
                                 selectmode="browse", height=12)
        for cname, cw, anchor in self.COL_DEF:
            self.tree.heading(cname, text=cname)
            self.tree.column(cname, width=cw, anchor=anchor, minwidth=40)

        vsb = ttk.Scrollbar(frm_tree, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(frm_tree, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        frm_tree.grid_rowconfigure(0, weight=1)
        frm_tree.grid_columnconfigure(0, weight=1)

        # 테이블 행 교대 색상
        self.tree.tag_configure("ok",   background="#e8f5e9")   # 연초록
        self.tree.tag_configure("fail", background="#ffebee")   # 연빨강
        self.tree.tag_configure("unk",  background="#fff8e1")   # 연노랑
        self.tree.tag_configure("even", background="#fafafa")

        # === 진행 바 ===
        frm_prog = tk.Frame(self, bg="#f0f0f0")
        frm_prog.pack(fill=tk.X, padx=12, pady=(0, 2))

        self.progress = ttk.Progressbar(frm_prog, length=400, mode="determinate")
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.lbl_progress = tk.Label(
            frm_prog, text="대기 중", bg="#f0f0f0", font=("맑은 고딕", 9),
            width=22, anchor=tk.W)
        self.lbl_progress.pack(side=tk.LEFT, padx=(8, 0))

        # === 로그 영역 ===
        frm_log = tk.LabelFrame(self, text=" 로그 ", bg="#f0f0f0",
                                font=("맑은 고딕", 9))
        frm_log.pack(fill=tk.X, padx=12, pady=(0, 8))

        self.txt_log = tk.Text(frm_log, height=6, wrap=tk.WORD,
                               font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4",
                               insertbackground="#d4d4d4", state=tk.DISABLED)
        log_vsb = ttk.Scrollbar(frm_log, orient=tk.VERTICAL, command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_vsb.set)
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0), pady=4)
        log_vsb.pack(side=tk.RIGHT, fill=tk.Y, pady=4, padx=(0, 4))

        # 로그 색상 태그
        self.txt_log.tag_configure("INFO",    foreground="#9cdcfe")
        self.txt_log.tag_configure("WARNING", foreground="#dcdcaa")
        self.txt_log.tag_configure("ERROR",   foreground="#f48771")

    # ── 폴더 선택 ──────────────────────────────────
    def _browse_folder(self):
        d = filedialog.askdirectory(title="PDF 파일이 있는 폴더 선택")
        if d:
            self.folder_path.set(d)

    # ── 로그 큐 폴링 ──────────────────────────────
    def _poll_log_queue(self):
        try:
            while not log_queue.empty():
                msg = log_queue.get_nowait()
                self._append_log(msg)
            self.after(100, self._poll_log_queue)
        except tk.TclError:
            pass   # 윈도우 종료 후 after 호출 시 무시

    def _append_log(self, text: str):
        self.txt_log.configure(state=tk.NORMAL)
        tag = "INFO"
        if "[WARNING]" in text:
            tag = "WARNING"
        elif "[ERROR]" in text:
            tag = "ERROR"
        self.txt_log.insert(tk.END, text + "\n", tag)
        self.txt_log.see(tk.END)
        self.txt_log.configure(state=tk.DISABLED)

    # ── 테이블 갱신 ────────────────────────────────
    def _refresh_table(self):
        """self.df → Treeview 동기화"""
        self.tree.delete(*self.tree.get_children())
        if self.df is None or self.df.empty:
            return

        for i, (_, row) in enumerate(self.df.iterrows()):
            verdict = str(row.get("진위여부", ""))
            if verdict == "O":
                tag = "ok"
            elif verdict == "X":
                tag = "fail"
            elif verdict == "확인불가":
                tag = "unk"
            elif i % 2 == 1:
                tag = "even"
            else:
                tag = ""

            # 합계금액 포맷
            amt = row.get("합계금액", "")
            try:
                amt = f"{int(str(amt).replace(',', '')):,}"
            except (ValueError, TypeError):
                amt = str(amt)

            self.tree.insert("", tk.END, values=(
                i + 1,
                str(row.get("파일명", "")),
                str(row.get("인식페이지", "")),
                str(row.get("승인번호", "")),
                str(row.get("공급자번호", "")),
                str(row.get("작성일자", "")),
                amt,
                verdict,
            ), tags=(tag,))

        self._update_stats()

    def _update_stats(self):
        """통계 라벨 갱신"""
        if self.df is None:
            self.lbl_stats.config(text="")
            return
        total = len(self.df)
        ok   = (self.df.get("진위여부") == "O").sum()   if "진위여부" in self.df.columns else 0
        fail = (self.df.get("진위여부") == "X").sum()   if "진위여부" in self.df.columns else 0
        unk  = (self.df.get("진위여부") == "확인불가").sum() if "진위여부" in self.df.columns else 0
        qr_ok = total - (self.df.get("승인번호") == "인식 불가").sum() if "승인번호" in self.df.columns else 0
        self.lbl_stats.config(
            text=f"전체 {total}건  |  QR인식 {qr_ok}건  |  O: {ok}  X: {fail}  ?: {unk}"
        )

    # ── UI 잠금/해제 ───────────────────────────────
    def _lock_ui(self):
        self._running = True
        self.btn_qr.config(state=tk.DISABLED)
        self.btn_verify.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)

    def _unlock_ui(self):
        self._running = False
        self.btn_qr.config(state=tk.NORMAL)
        self.btn_verify.config(state=tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)

    def _set_progress(self, current: int, total: int, text: str = ""):
        pct = int(current / total * 100) if total else 0
        self.progress["maximum"] = total
        self.progress["value"]   = current
        self.lbl_progress.config(text=text or f"{current}/{total}건  ({pct}%)")

    # ══════════════════════════════════════════════
    # 1단계: QR 추출
    # ══════════════════════════════════════════════
    def _run_qr_extract(self):
        folder = self.folder_path.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning("경고", "유효한 PDF 폴더를 선택해 주세요.")
            return

        pdf_files = sorted(Path(folder).glob("*.pdf"))
        if not pdf_files:
            messagebox.showinfo("알림", "선택한 폴더에 PDF 파일이 없습니다.")
            return

        self._lock_ui()
        self._set_progress(0, len(pdf_files), "QR 추출 시작...")
        log.info(f"QR 추출 시작: {folder}  ({len(pdf_files)}개 PDF)")

        threading.Thread(
            target=self._qr_worker, args=(folder, pdf_files), daemon=True
        ).start()

    def _qr_worker(self, folder: str, pdf_files: list):
        """백그라운드 스레드: QR 일괄 추출"""
        try:
            qr = _import_qr()
            records = []
            total = len(pdf_files)

            for i, pdf_path in enumerate(pdf_files):
                log.info(f"[{i+1}/{total}] {pdf_path.name}")
                self.after(0, self._set_progress, i, total,
                           f"QR 스캔 중... {i+1}/{total}")
                record = qr.process_pdf(str(pdf_path))
                records.append(record)

            col_order = ["파일명", "인식페이지", "승인번호",
                         "공급자번호", "작성일자", "합계금액", "원본QR"]
            self.df = pd.DataFrame(records, columns=col_order)

            # 자동 저장
            out_path = str(Path(folder) / "results.xlsx")
            self.df.to_excel(out_path, index=False, engine="openpyxl")
            log.info(f"results.xlsx 저장 완료 → {out_path}")

            success = (self.df["승인번호"].notna() &
                       ~self.df["승인번호"].isin(["인식 불가"])).sum()
            log.info(f"QR 추출 완료: 성공 {success} / 전체 {total}")

            self.after(0, self._set_progress, total, total, "QR 추출 완료!")
            self.after(0, self._refresh_table)
            self.after(0, self._unlock_ui)

        except Exception as e:
            log.error(f"QR 추출 오류: {e}")
            self.after(0, self._unlock_ui)
            self.after(0, messagebox.showerror, "오류", str(e))

    # ══════════════════════════════════════════════
    # 2단계: 진위 확인
    # ══════════════════════════════════════════════
    def _run_verify(self):
        if self.df is None or self.df.empty:
            messagebox.showwarning("경고",
                "먼저 1단계 QR 추출을 실행하거나\n"
                "기존 results.xlsx가 있는 폴더를 선택하세요.")
            return

        # results.xlsx에서 로드 (QR 추출 없이 바로 진위 확인 시)
        if "승인번호" not in self.df.columns:
            messagebox.showwarning("경고", "승인번호 컬럼이 없습니다.")
            return

        self._lock_ui()
        total = len(self.df)
        self._set_progress(0, total, "진위 확인 시작...")
        log.info(f"홈택스 진위 확인 시작: {total}건")

        threading.Thread(target=self._verify_worker, daemon=True).start()

    def _verify_worker(self):
        """백그라운드 스레드: 홈택스 진위 확인"""
        try:
            import urllib3
            urllib3.disable_warnings()
            verify = _import_verify()

            if "진위여부" not in self.df.columns:
                self.df["진위여부"] = ""

            total = len(self.df)

            for idx, row in self.df.iterrows():
                i = idx + 1
                existing = str(row.get("진위여부", "")).strip()

                if existing in ("O", "X"):
                    log.info(f"[{i}/{total}] {row['파일명']}: 이미 처리됨({existing})")
                    self.after(0, self._set_progress, i, total,
                               f"진위 확인 중... {i}/{total}")
                    continue

                # 인식 불가 건은 건너뜀
                if str(row.get("승인번호", "")).startswith("인식"):
                    self.df.at[idx, "진위여부"] = "확인불가"
                    self.after(0, self._set_progress, i, total,
                               f"진위 확인 중... {i}/{total}")
                    self.after(0, self._refresh_table)
                    continue

                log.info(f"[{i}/{total}] {row['파일명']}")
                self.after(0, self._set_progress, i - 1, total,
                           f"진위 확인 중... {i}/{total}")

                # 데이터 정제
                row_data = {
                    "승인번호_원본": str(row.get("승인번호", "")),
                    "승인번호_정제": verify.clean_approval_number(str(row.get("승인번호", ""))),
                    "공급자번호":    verify.clean_number(row.get("공급자번호", "")),
                    "작성일자_raw":  verify.clean_date(str(row.get("작성일자", ""))),
                    "합계금액":      verify.clean_number(row.get("합계금액", "")),
                    "수급자번호":    "",
                }

                # 전략 2: 모바일 API (가장 빠르고 안정적)
                result = verify.verify_via_api(row_data)
                if not result:
                    result = "확인불가"

                self.df.at[idx, "진위여부"] = result
                log.info(f"  판정: {result}")

                self.after(0, self._refresh_table)

                # 랜덤 대기
                wait = random.uniform(2, 5)
                time.sleep(wait)

            self.after(0, self._set_progress, total, total, "진위 확인 완료!")
            self.after(0, self._refresh_table)
            self.after(0, self._unlock_ui)

            # 자동 저장
            folder = self.folder_path.get().strip()
            if folder:
                out = str(Path(folder) / "verification_final.xlsx")
                self.df.to_excel(out, index=False, engine="openpyxl")
                log.info(f"verification_final.xlsx 저장 → {out}")

            ok  = (self.df["진위여부"] == "O").sum()
            ng  = (self.df["진위여부"] == "X").sum()
            unk = (self.df["진위여부"] == "확인불가").sum()
            log.info(f"진위 확인 완료! O:{ok} X:{ng} ?:{unk}")
            self.after(0, messagebox.showinfo, "완료",
                       f"진위 확인이 완료되었습니다.\n\n"
                       f"O (정상): {ok}건\nX (불일치): {ng}건\n확인불가: {unk}건")

        except Exception as e:
            log.error(f"진위 확인 오류: {e}")
            self.after(0, self._unlock_ui)
            self.after(0, messagebox.showerror, "오류", str(e))

    # ══════════════════════════════════════════════
    # Excel 저장 / 열기
    # ══════════════════════════════════════════════
    def _save_excel(self):
        if self.df is None or self.df.empty:
            messagebox.showinfo("알림", "저장할 데이터가 없습니다.")
            return

        path = filedialog.asksaveasfilename(
            title="Excel 저장",
            defaultextension=".xlsx",
            filetypes=[("Excel 파일", "*.xlsx")],
            initialfile="verification_final.xlsx",
        )
        if not path:
            return

        self.df.to_excel(path, index=False, engine="openpyxl")

        # 열 너비 자동 조정
        try:
            from openpyxl import load_workbook
            wb = load_workbook(path)
            ws = wb.active
            for col in ws.columns:
                ml = max((len(str(c.value)) for c in col if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(ml + 4, 60)
            wb.save(path)
        except Exception:
            pass

        log.info(f"Excel 저장 → {path}")
        messagebox.showinfo("저장 완료", f"저장되었습니다.\n{path}")

    def _open_excel(self):
        """기존 results.xlsx 또는 verification_final.xlsx 불러오기"""
        path = filedialog.askopenfilename(
            title="Excel 파일 열기",
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if not path:
            return

        try:
            self.df = pd.read_excel(path)
            self.folder_path.set(str(Path(path).parent))
            self._refresh_table()
            log.info(f"Excel 로드 → {path}  ({len(self.df)}건)")
        except Exception as e:
            messagebox.showerror("오류", f"파일 열기 실패:\n{e}")


# ══════════════════════════════════════════════
# 진입점
# ══════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
