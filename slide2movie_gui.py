import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess
import sys
import threading
import os


# -----------------------------------------------------------------------
# 定数
# -----------------------------------------------------------------------
SCRIPT_PATH = os.path.join(os.path.dirname(__file__), "slide2movie.py")


# -----------------------------------------------------------------------
# モジュール
# -----------------------------------------------------------------------
# UTF-8 → CP932 → latin-1 の順でフォールバックデコード
def _decode_auto(raw: bytes) -> str:
    for enc in ("utf-8", "cp932", "latin-1"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")

# -----------------------------------------------------------------------
# GUI本体
# -----------------------------------------------------------------------
class Slide2MovieGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("slide2movie GUI")
        self.iconbitmap(os.path.join(os.path.dirname(__file__), "icon.ico"))
        self.resizable(True, True)
        self.minsize(620, 700)

        # --- 変数定義 ---
        self.var_file       = tk.StringVar()
        self.var_output     = tk.StringVar()
        self.var_dpi        = tk.IntVar(value=150)
        self.var_quality    = tk.IntVar(value=5)
        self.var_lang       = tk.StringVar(value="ja")
        self.var_voicevox   = tk.BooleanVar(value=False)
        self.var_voicevoxid = tk.IntVar(value=3)
        self.var_creditimg  = tk.StringVar()
        self.var_creditbg   = tk.StringVar(value="#000000")
        self.var_creditcolor= tk.StringVar(value="#ffffff")
        self.var_debug      = tk.BooleanVar(value=False)

        self._build_ui()

    # -------------------------------------------------------------------
    # UI構築
    # -------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}

        # ---- ファイル選択 ----
        frame_files = ttk.LabelFrame(self, text="ファイル設定", padding=8)
        frame_files.pack(fill="x", **pad)

        self._file_row(frame_files, "入力 PPTX *", self.var_file,
                       self._browse_pptx, row=0)
        self._file_row(frame_files, "出力 MP4  *", self.var_output,
                       self._browse_output, row=1)

        # ---- 基本設定 ----
        frame_basic = ttk.LabelFrame(self, text="基本設定", padding=8)
        frame_basic.pack(fill="x", **pad)

        # DPI
        ttk.Label(frame_basic, text="DPI").grid(
            row=0, column=0, sticky="w", pady=3)
        ttk.Spinbox(frame_basic, from_=72, to=600, increment=1,
                    textvariable=self.var_dpi, width=8).grid(
            row=0, column=1, sticky="w", padx=6)

        # Quality
        ttk.Label(frame_basic, text="Quality (1〜31, 低いほど高品質)").grid(
            row=1, column=0, sticky="w", pady=3)
        ttk.Spinbox(frame_basic, from_=1, to=31, increment=1,
                    textvariable=self.var_quality, width=8).grid(
            row=1, column=1, sticky="w", padx=6)

        # Lang
        ttk.Label(frame_basic, text="言語 (gTTS用)").grid(
            row=2, column=0, sticky="w", pady=3)
        ttk.Entry(frame_basic, textvariable=self.var_lang, width=10).grid(
            row=2, column=1, sticky="w", padx=6)

        # ---- VOICEVOX設定 ----
        frame_vv = ttk.LabelFrame(self, text="VOICEVOX設定", padding=8)
        frame_vv.pack(fill="x", **pad)

        ttk.Checkbutton(frame_vv, text="VOICEVOX を使用する",
                        variable=self.var_voicevox).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=3)

        ttk.Label(frame_vv, text="VOICEVOX 話者ID").grid(
            row=1, column=0, sticky="w", pady=3)
        ttk.Spinbox(frame_vv, from_=0, to=9999, increment=1,
                    textvariable=self.var_voicevoxid, width=8).grid(
            row=1, column=1, sticky="w", padx=6)

        # ---- クレジット設定 ----
        frame_credit = ttk.LabelFrame(self, text="クレジット設定", padding=8)
        frame_credit.pack(fill="x", **pad)

        self._file_row(frame_credit, "クレジット画像", self.var_creditimg,
                       self._browse_creditimg, row=0)

        ttk.Label(frame_credit, text="背景色 (例: #000000)").grid(
            row=1, column=0, sticky="w", pady=3)
        ttk.Entry(frame_credit, textvariable=self.var_creditbg, width=12).grid(
            row=1, column=1, sticky="w", padx=6)

        ttk.Label(frame_credit, text="文字色 (例: #ffffff)").grid(
            row=2, column=0, sticky="w", pady=3)
        ttk.Entry(frame_credit, textvariable=self.var_creditcolor, width=12).grid(
            row=2, column=1, sticky="w", padx=6)

        # ---- その他 ----
        frame_misc = ttk.LabelFrame(self, text="その他", padding=8)
        frame_misc.pack(fill="x", **pad)

        ttk.Checkbutton(frame_misc, text="デバッグモード (中間ファイルを残す)",
                        variable=self.var_debug).pack(anchor="w")

        # ---- 実行ボタン ----
        self.btn_run = ttk.Button(self, text="▶ 実行", command=self._on_run)
        self.btn_run.pack(pady=8)

        # ---- ログ出力 ----
        frame_log = ttk.LabelFrame(self, text="実行ログ", padding=8)
        frame_log.pack(fill="both", expand=True, **pad)

        self.log_area = scrolledtext.ScrolledText(
            frame_log, state="disabled", wrap="word",
            font=("Courier", 9), height=12)
        self.log_area.pack(fill="both", expand=True)

    # -------------------------------------------------------------------
    # ファイル行ウィジェット（共通）
    # -------------------------------------------------------------------
    def _file_row(self, parent, label, var, cmd, row):
        ttk.Label(parent, text=label).grid(
            row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=var, width=44).grid(
            row=row, column=1, sticky="ew", padx=6)
        ttk.Button(parent, text="参照…", command=cmd, width=7).grid(
            row=row, column=2, padx=2)
        parent.columnconfigure(1, weight=1)

    # -------------------------------------------------------------------
    # ファイルダイアログ
    # -------------------------------------------------------------------
    def _browse_pptx(self):
        path = filedialog.askopenfilename(
            title="入力PPTXを選択",
            filetypes=[("PowerPoint", "*.pptx *.ppt"), ("すべて", "*.*")])
        if path:
            self.var_file.set(path)
            # 出力パスを自動補完（未入力の場合）
            if not self.var_output.get():
                base = os.path.splitext(path)[0]
                self.var_output.set(base + ".mp4")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="出力MP4のパスを指定",
            defaultextension=".mp4",
            filetypes=[("MP4動画", "*.mp4"), ("すべて", "*.*")])
        if path:
            self.var_output.set(path)

    def _browse_creditimg(self):
        path = filedialog.askopenfilename(
            title="クレジット画像を選択",
            filetypes=[("画像", "*.png *.jpg *.jpeg"), ("すべて", "*.*")])
        if path:
            self.var_creditimg.set(path)

    # -------------------------------------------------------------------
    # 実行
    # -------------------------------------------------------------------
    def _on_run(self):
        # バリデーション
        if not self.var_file.get():
            messagebox.showwarning("入力エラー", "入力PPTXを指定してください。")
            return
        if not self.var_output.get():
            messagebox.showwarning("入力エラー", "出力MP4のパスを指定してください。")
            return

        self.btn_run.config(state="disabled")
        self._log_clear()
        self._log("実行開始...\n")

        # 別スレッドで実行（GUIフリーズ防止）
        thread = threading.Thread(target=self._run_script, daemon=True)
        thread.start()

    def _run_script(self):
        cmd = self._build_command()
        self._log(f"コマンド: {' '.join(cmd)}\n{'─'*60}\n")

        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                bufsize=0, # バイナリ受け取り
                env={**os.environ, "PYTHONUNBUFFERED": "1"}
            )
            for raw_line in proc.stdout:
                line = _decode_auto(raw_line)
                self._log(line)
            proc.wait()

            if proc.returncode == 0:
                self._log(f"\n{'─'*60}\n✅ 完了しました。\n出力: {self.var_output.get()}\n")
            else:
                self._log(f"\n{'─'*60}\n❌ エラーで終了しました（終了コード: {proc.returncode}）\n")

        except FileNotFoundError:
            self._log(f"\n❌ スクリプトが見つかりません: {SCRIPT_PATH}\n")
        except Exception as e:
            self._log(f"\n❌ 予期しないエラー: {e}\n")
        finally:
            # ボタンをメインスレッドで再有効化
            self.after(0, lambda: self.btn_run.config(state="normal"))

    def _build_command(self):
        cmd = [sys.executable, SCRIPT_PATH]

        cmd += ["--file",    self.var_file.get()]
        cmd += ["--output",  self.var_output.get()]
        cmd += ["--dpi",     str(self.var_dpi.get())]
        cmd += ["--quality", str(self.var_quality.get())]
        cmd += ["--lang",    self.var_lang.get()]

        if self.var_voicevox.get():
            cmd += ["--voicevox"]
            cmd += ["--voicevoxid", str(self.var_voicevoxid.get())]

        if self.var_creditimg.get():
            cmd += ["--creditimg", self.var_creditimg.get()]

        if self.var_creditbg.get():
            cmd += ["--creditbg", self.var_creditbg.get()]

        if self.var_creditcolor.get():
            cmd += ["--creditcolor", self.var_creditcolor.get()]

        if self.var_debug.get():
            cmd += ["--debug"]

        return cmd

    # -------------------------------------------------------------------
    # ログ操作
    # -------------------------------------------------------------------
    def _log(self, text: str):
        """スレッドセーフなログ追記"""
        def _append():
            self.log_area.config(state="normal")
            self.log_area.insert("end", text)
            self.log_area.see("end")
            self.log_area.config(state="disabled")
        self.after(0, _append)

    def _log_clear(self):
        def _clear():
            self.log_area.config(state="normal")
            self.log_area.delete("1.0", "end")
            self.log_area.config(state="disabled")
        self.after(0, _clear)


# -----------------------------------------------------------------------
# エントリーポイント
# -----------------------------------------------------------------------
if __name__ == "__main__":
    app = Slide2MovieGUI()
    app.mainloop()