"""합본 자동화 GUI 실행기."""

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import json
import os
import io
import subprocess
import sys
import threading
from pathlib import Path

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


# ── exe 내부 CLI 모드 (GUI가 subprocess로 자기 자신을 호출할 때 진입) ──────
# PyInstaller exe에서 sys.executable은 exe 자신이므로,
# --_cli 플래그로 GUI 모드와 처리 모드를 구분한다.
def _run_as_cli():
    """--_cli 플래그로 실행됐을 때: GUI 없이 main 파이프라인만 실행."""
    sys.argv.remove("--_cli")
    # UTF-8 출력 보장
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

        # 프로젝트명
        self.vars["project"] = tk.StringVar(value=cfg.get("project", ""))
        ttk.Entry(frame, textvariable=self.vars["project"], width=36).grid(row=0, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        # 차수 (저장 안 함 — 매번 직접 입력)
        self.vars["round"] = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.vars["round"], width=36).grid(row=1, column=1, columnspan=2, sticky="ew", padx=(6, 0))

        # 옵티컬 언어
        self.vars["optical"] = tk.StringVar(value=cfg.get("optical", "EN"))
        ttk.Combobox(frame, textvariable=self.vars["optical"], values=["EN", "CN", "NONE"],
                     width=10, state="readonly").grid(row=2, column=1, sticky="w", padx=(6, 0))

        # 녹음 언어
        self.vars["record"] = tk.StringVar(value=cfg.get("record", "KR"))
        ttk.Combobox(frame, textvariable=self.vars["record"], values=["KR", "EN", "JP"],
                     width=10, state="readonly").grid(row=3, column=1, sticky="w", padx=(6, 0))

        # 원본 대본 폴더
        self.vars["source"] = tk.StringVar(value=cfg.get("source", ""))
        ttk.Entry(frame, textvariable=self.vars["source"], width=30).grid(row=4, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frame, text="찾기", width=6,
                   command=self._browse_source).grid(row=4, column=2)

        # 합본 출력 경로
        self.vars["output"] = tk.StringVar(value=cfg.get("output", ""))
        ttk.Entry(frame, textvariable=self.vars["output"], width=30).grid(row=5, column=1, sticky="ew", padx=(6, 4))
        ttk.Button(frame, text="찾기", width=6,
                   command=self._browse_output).grid(row=5, column=2)

        # ── 실행 버튼 ──────────────────────────────────────────
        self.run_btn = ttk.Button(self.root, text="▶  합본 생성 시작", command=self._run)
        self.run_btn.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 5))

        # ── 로그 창 ──────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="진행 로그", padding=5)
        log_frame.grid(row=2, column=0, sticky="nsew", **pad)

        self.log = scrolledtext.ScrolledText(log_frame, width=62, height=18,
                                             font=("Consolas", 9), state="disabled",
                                             bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="both", expand=True)

    def _browse_source(self):
        path = filedialog.askdirectory(title="원본 대본 폴더 선택")
        if path:
            self.vars["source"].set(path)
            if not self.vars["output"].get():
                proj = self.vars["project"].get() or "합본"
                rnd = self.vars["round"].get() or ""
                stem = f"{proj}_{rnd}_합본".strip("_")
                self.vars["output"].set(str(Path(path) / f"{stem}.xlsx"))

    def _browse_output(self):
        proj = self.vars["project"].get() or "합본"
        rnd = self.vars["round"].get() or ""
        default = f"{proj}_{rnd}_합본.xlsx".strip("_")
        path = filedialog.asksaveasfilename(
            title="합본 저장 경로",
            defaultextension=".xlsx",
            initialfile=default,
            filetypes=[("Excel 파일", "*.xlsx")],
        )
        if path:
            self.vars["output"].set(path)

    def _log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert("end", text)
        self.log.see("end")
        self.log.configure(state="disabled")

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
        # 폴더 경로만 입력된 경우 자동으로 파일명 보정
        if Path(output).is_dir() or not output.lower().endswith(".xlsx"):
            proj = project
            rnd  = self.vars["round"].get().strip()
            stem = f"{proj}_{rnd}_합본.xlsx".strip("_")
            output = str(Path(output) / stem) if Path(output).is_dir() else output + ".xlsx"
            self.vars["output"].set(output)

        # 설정 저장 (차수 제외)
        save_config({k: v.get() for k, v in self.vars.items() if k != "round"})

        api_key = os.environ.get("GEMINI_API_KEY", "")
        if not api_key:
            api_key = tk.simpledialog.askstring("API 키", "GEMINI_API_KEY를 입력하세요:", show="*")
            if not api_key:
                return

        # exe 여부에 따라 실행 커맨드 구성
        # - 일반 Python: python -X utf8 main.py ...
        # - exe: 합본자동화.exe --_cli ... (GUI 재진입 방지)
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

        self.run_btn.configure(state="disabled", text="실행 중...")
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

        def worker():
            try:
                proc = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    encoding="utf-8",
                    errors="replace",
                )
                for line in proc.stdout:
                    self.root.after(0, self._log, line)
                proc.wait()
                if proc.returncode == 0:
                    self.root.after(0, self._log, "\n✅ 완료!\n")
                else:
                    self.root.after(0, self._log, f"\n❌ 오류 발생 (종료코드 {proc.returncode})\n")
            except Exception as e:
                self.root.after(0, self._log, f"\n❌ 실행 오류: {e}\n")
            finally:
                self.root.after(0, lambda: self.run_btn.configure(state="normal", text="▶  합본 생성 시작"))

        threading.Thread(target=worker, daemon=True).start()


def main():
    import tkinter.messagebox
    import tkinter.simpledialog
    root = tk.Tk()
    app = HapbonGUI(root)
    root.mainloop()


if __name__ == "__main__":
    # exe로 실행됐고 --_cli 플래그가 있으면 → GUI 없이 파이프라인만 실행
    if getattr(sys, "frozen", False) and "--_cli" in sys.argv:
        _run_as_cli()
    else:
        main()
