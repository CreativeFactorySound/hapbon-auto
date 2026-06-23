"""합본 자동화 GUI 실행기."""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import ctypes
import json
import os
import io
import re
import subprocess
import sys
import threading
from pathlib import Path

# Windows 프로세스 일시정지/재개
def _suspend_process(proc):
    try:
        ctypes.windll.ntdll.NtSuspendProcess(int(proc._handle))
    except Exception:
        pass

def _resume_process(proc):
    try:
        ctypes.windll.ntdll.NtResumeProcess(int(proc._handle))
    except Exception:
        pass

CONFIG_FILE = Path(__file__).parent / ".hapbon_gui_config.json"


def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def save_config(cfg: dict):
    CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")


# ── exe 내부 CLI 모드 ──────────────────────────────────────────────────────
def _run_as_cli():
    sys.argv.remove("--_cli")
    if sys.stdout.encoding and sys.stdout.encoding.lower() not in ("utf-8", "utf8"):
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    if sys.stderr.encoding and sys.stderr.encoding.lower() not in ("utf-8", "utf8"):
        sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    from main import main
    main()


class HapbonGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("합본 자동화 도구")
        self.root.resizable(False, False)

        # 진행률 추적용 상태
        self._total_sheets  = 0
        self._done_sheets   = 0
        self._proc          = None   # 실행 중인 subprocess
        self._paused        = False  # 일시정지 상태
        self._start_time    = None   # 실행 시작 시각
        self._timer_job     = None   # after() 타이머 job ID

        cfg = load_config()
        self._build_ui(cfg)

    def _build_ui(self, cfg: dict):
        pad = {"padx": 10, "pady": 5}

        # ── 입력 프레임 ──────────────────────────────────────────
        frame = ttk.LabelFrame(self.root, text="설정", padding=10)
        frame.grid(row=0, column=0, sticky="ew", **pad)

        labels = ["프로젝트명", "차수", "옵티컬 언어", "녹음 언어", "원본 대본 폴더", "합본 출력 경로"]
        self.vars: dict[str, tk.StringVar] = {}

        for i, label in enumerate(labels):
            ttk.Label(frame, text=label, width=14, anchor="e").grid(row=i, column=0, sticky="e", pady=3)

        self.vars["project"] = tk.StringVar(value=cfg.get("project", ""))
        ttk.Entry(frame, textvariable=self.vars["project"], width=36).grid(row=0, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        self.vars["round"] = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.vars["round"], width=36).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        self.vars["optical"] = tk.StringVar(value=cfg.get("optical", "EN"))
        ttk.Combobox(frame, textvariable=self.vars["optical"], values=["EN", "CN", "NONE"],
                     width=10, state="readonly").grid(row=2, column=1, sticky="w", padx=(6, 0))

        self.vars["record"] = tk.StringVar(value=cfg.get("record", "KR"))
        ttk.Combobox(frame, textvariable=self.vars["record"], values=["KR", "EN", "JP"],
                     width=10, state="readonly").grid(row=3, column=1, sticky="w", padx=(6, 0))

        self.vars["source"] = tk.StringVar(value=cfg.get("source", ""))
        ttk.Entry(frame, textvariable=self.vars["source"], width=30).grid(row=4, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frame, text="찾기", width=6,
                   command=self._browse_source).grid(row=4, column=2)

        self.vars["output"] = tk.StringVar(value=cfg.get("output", ""))
        ttk.Entry(frame, textvariable=self.vars["output"], width=30).grid(row=5, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frame, text="찾기", width=6,
                   command=self._browse_output).grid(row=5, column=2)

        # ── 실행 / 일시정지 / 정지 버튼 ───────────────────────────
        btn_frame = tk.Frame(self.root)
        btn_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(4, 2))
        btn_frame.columnconfigure(0, weight=3)
        btn_frame.columnconfigure(1, weight=1)
        btn_frame.columnconfigure(2, weight=1)

        self.run_btn = ttk.Button(btn_frame, text="▶  합본 생성 시작", command=self._run)
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 4))

        self.pause_btn = ttk.Button(btn_frame, text="⏸  일시정지", command=self._toggle_pause, state="disabled")
        self.pause_btn.grid(row=0, column=1, sticky="ew", padx=(0, 4))

        self.stop_btn = ttk.Button(btn_frame, text="⏹  정지", command=self._stop, state="disabled")
        self.stop_btn.grid(row=0, column=2, sticky="ew")

        # ── 진행률 바 ──────────────────────────────────────────
        prog_frame = tk.Frame(self.root)
        prog_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 4))
        prog_frame.columnconfigure(0, weight=1)
        prog_frame.columnconfigure(1, weight=0)

        self._progress_var = tk.DoubleVar(value=0)
        self._progress_bar = ttk.Progressbar(
            prog_frame, variable=self._progress_var,
            maximum=100, length=400, mode="determinate"
        )
        self._progress_bar.grid(row=0, column=0, columnspan=2, sticky="ew")

        self._status_var = tk.StringVar(value="대기 중")
        ttk.Label(prog_frame, textvariable=self._status_var,
                  anchor="w", foreground="#555555").grid(row=1, column=0, sticky="w", pady=(2, 0))

        self._timer_var = tk.StringVar(value="")
        ttk.Label(prog_frame, textvariable=self._timer_var,
                  anchor="e", foreground="#888888").grid(row=1, column=1, sticky="e", pady=(2, 0))

        # ── 로그 창 ──────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="진행 로그", padding=5)
        log_frame.grid(row=3, column=0, sticky="nsew", **pad)

        self.log = scrolledtext.ScrolledText(log_frame, width=62, height=16,
                                             font=("Consolas", 9), state="disabled",
                                             bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True)

    # ── 파일 탐색 ─────────────────────────────────────────────────────────

    def _browse_source(self):
        path = filedialog.askdirectory(title="원본 대본 폴더 선택")
        if path:
            self.vars["source"].set(path)
            if not self.vars["output"].get():
                proj = self.vars["project"].get() or "합본"
                rnd  = self.vars["round"].get() or ""
                stem = f"{proj}_{rnd}_합본".strip("_")
                self.vars["output"].set(str(Path(path) / f"{stem}.xlsx"))

    def _browse_output(self):
        proj    = self.vars["project"].get() or "합본"
        rnd     = self.vars["round"].get() or ""
        default = f"{proj}_{rnd}_합본.xlsx".strip("_")
        path = filedialog.asksaveasfilename(
            title="합본 저장 경로",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if path:
            self.vars["output"].set(path)

    # ── 로그 & 진행률 ────────────────────────────────────────────────────

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")
        self._parse_progress(text)

    def _set_progress(self, pct: float, status: str):
        self._progress_var.set(min(pct, 100))
        self._status_var.set(status)

    def _parse_progress(self, line: str):
        """로그 라인을 파싱해서 진행률·상태 업데이트."""
        line = line.strip()

        # 파일 수집 완료
        if line.startswith("[수집]") and "개 파일" in line:
            self._set_progress(5, "📂 " + line.lstrip("[수집]").strip())

        # 분류 중 (개별 시트)
        elif "분류 중..." in line:
            cur = self._progress_var.get()
            self._set_progress(min(cur + 1.5, 38), "🔍 분류 중: " + re.sub(r"\[|\]", "", line.split("]")[0].split("[")[-1]))

        # 분류 완료 → 총 시트 수 파악
        elif line.startswith("[분류 완료]"):
            m = re.search(r"(\d+)개 시트", line)
            if m:
                self._total_sheets = int(m.group(1))
                self._done_sheets  = 0
            self._set_progress(40, f"✅ 분류 완료 — 총 {self._total_sheets}개 시트 처리 예정")

        # 처리 중 (개별 시트)
        elif line.startswith("[처리]"):
            self._done_sheets += 1
            total = self._total_sheets or 1
            pct   = 40 + (self._done_sheets / total) * 50
            # 시트명 추출
            m = re.search(r"\[처리\] .+? / \[(.+?)\]", line)
            sheet = m.group(1) if m else ""
            self._set_progress(pct, f"⚙️  처리 중 ({self._done_sheets}/{total}): {sheet}")

        # 조립
        elif line.startswith("[조립]"):
            self._set_progress(93, "📦 합본 조립 중...")

        # 완료
        elif line.startswith("[완료]") or "✅ 완료" in line:
            self._set_progress(100, "🎉 완료!")

        # 오류
        elif "❌" in line or "ERROR" in line:
            self._status_var.set("❌ 오류 발생 — 로그를 확인하세요")

    # ── 타이머 ───────────────────────────────────────────────────────────

    def _start_timer(self):
        import time
        self._start_time = time.time()
        self._tick_timer()

    def _stop_timer(self):
        if self._timer_job:
            self.root.after_cancel(self._timer_job)
            self._timer_job = None

    def _tick_timer(self):
        import time
        if self._start_time is None:
            return
        elapsed = int(time.time() - self._start_time)
        e_str = f"{elapsed // 60:02d}:{elapsed % 60:02d}"

        pct = self._progress_var.get()
        if pct > 5:
            total_est = elapsed / (pct / 100)
            remaining = max(0, int(total_est - elapsed))
            r_str = f"{remaining // 60:02d}:{remaining % 60:02d}"
            self._timer_var.set(f"경과 {e_str}  |  예상 잔여 {r_str}")
        else:
            self._timer_var.set(f"경과 {e_str}")

        self._timer_job = self.root.after(1000, self._tick_timer)

    # ── 일시정지 / 정지 ──────────────────────────────────────────────────

    def _toggle_pause(self):
        if not self._proc:
            return
        if self._paused:
            _resume_process(self._proc)
            self._paused = False
            self.pause_btn.configure(text="⏸  일시정지")
            self._status_var.set(self._status_var.get().replace("⏸ 일시정지됨", "").strip() or "실행 재개")
            self._log("\n[재개]\n")
        else:
            _suspend_process(self._proc)
            self._paused = True
            self.pause_btn.configure(text="▶  재개")
            self._status_var.set("⏸ 일시정지됨")
            self._log("\n[일시정지]\n")

    def _stop(self):
        if not self._proc:
            return
        if self._paused:
            _resume_process(self._proc)
            self._paused = False
        try:
            self._proc.terminate()
        except Exception:
            pass
        self._log("\n[정지] 사용자가 중단했습니다.\n")
        self._status_var.set("🛑 정지됨")

    # ── 실행 ─────────────────────────────────────────────────────────────

    def _run(self):
        project = self.vars["project"].get().strip()
        source  = self.vars["source"].get().strip()
        output  = self.vars["output"].get().strip()

        if not project:
            tk.messagebox.showwarning("입력 오류", "프로젝트명을 입력하세요.")
            return
        if not source or not Path(source).is_dir():
            tk.messagebox.showwarning("입력 오류", "원본 대본 폴더 경로가 올바르지 않습니다.")
            return
        if not output:
            tk.messagebox.showwarning("입력 오류", "합본 출력 경로를 입력하세요.")
            return
        if Path(output).is_dir() or not output.lower().endswith(".xlsx"):
            rnd    = self.vars["round"].get().strip()
            stem   = f"{project}_{rnd}_합본.xlsx".strip("_")
            output = str(Path(output) / stem) if Path(output).is_dir() else output + ".xlsx"
            self.vars["output"].set(output)

        save_config({k: v.get() for k, v in self.vars.items() if k != "round"})

        api_key = os.environ.get("GEMINI_API_KEY", "")
        if not api_key:
            api_key = tk.simpledialog.askstring("API 키", "GEMINI_API_KEY를 입력하세요:", show="*")
            if not api_key:
                return

        if getattr(sys, "frozen", False):
            cmd = [sys.executable, "--_cli"]
        else:
            cmd = [sys.executable, "-X", "utf8", str(Path(__file__).parent / "main.py")]

        cmd += [
            "--source",  source,
            "--output",  output,
            "--project", project,
            "--round",   self.vars["round"].get().strip(),
            "--optical", self.vars["optical"].get(),
            "--record",  self.vars["record"].get(),
            "--api-key", api_key,
        ]

        # UI 초기화
        self.run_btn.configure(state="disabled", text="실행 중...")
        self.pause_btn.configure(state="normal")
        self.stop_btn.configure(state="normal")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self._progress_var.set(0)
        self._status_var.set("시작 중...")
        self._timer_var.set("")
        self._total_sheets = 0
        self._done_sheets  = 0
        self._paused       = False
        self._stop_timer()
        self._start_timer()

        def worker():
            try:
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    encoding="utf-8",
                    errors="replace",
                )
                self._proc = proc
                for line in proc.stdout:
                    self.root.after(0, self._log, line)
                proc.wait()
                if proc.returncode == 0:
                    self.root.after(0, self._log, "\n✅ 완료!\n")
                    self.root.after(0, self._set_progress, 100, "🎉 완료!")
                elif proc.returncode == 1 and self._paused is False:
                    self.root.after(0, self._log, f"\n❌ 오류 발생 (종료코드 {proc.returncode})\n")
                    self.root.after(0, self._status_var.set, "❌ 오류 발생")
            except Exception as e:
                self.root.after(0, self._log, f"\n❌ 실행 오류: {e}\n")
            finally:
                self._proc = None
                self.root.after(0, self._stop_timer)
                self.root.after(0, lambda: self.run_btn.configure(state="normal", text="▶  합본 생성 시작"))
                self.root.after(0, lambda: self.pause_btn.configure(state="disabled", text="⏸  일시정지"))
                self.root.after(0, lambda: self.stop_btn.configure(state="disabled"))

        threading.Thread(target=worker, daemon=True).start()


def main():
    import tkinter.messagebox
    import tkinter.simpledialog
    root = tk.Tk()
    app = HapbonGUI(root)
    root.mainloop()


if __name__ == "__main__":
    if getattr(sys, "frozen", False) and "--_cli" in sys.argv:
        _run_as_cli()
    else:
        main()
