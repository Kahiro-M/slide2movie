import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import subprocess
import sys
import threading
import os
import platform
import sys

# ──────────────
# オプション定義
# ──────────────
from external_define import OPTION_DEFS, CONFIG_DEFAULT, COMMON_TEXT


# -----------------------------------------------------------------------
# モジュール
# -----------------------------------------------------------------------
from external_define import load_ini, save_ini

# PyInstaller実行時と通常実行時の両方でリソースにアクセスできるようにする関数
def _get_base_dir() -> str:
    if hasattr(sys, "_MEIPASS"):
        return sys._MEIPASS          # PyInstaller実行時の一時展開フォルダ
    return os.path.dirname(__file__) # 通常実行時

# ランチャー（slide2movie_gui）と同階層のディレクトリを返す
def _get_launcher_dir() -> str:
    if hasattr(sys, "_MEIPASS"):
        # PyInstaller実行時 → .exe のあるフォルダ
        return os.path.dirname(sys.executable)
    # 通常実行時 → スクリプトのあるフォルダ
    return os.path.dirname(os.path.abspath(__file__))
SCRIPT_PATH = os.path.join(_get_launcher_dir(), "slide2movie.py")

# -----------------------------------------------------------------------
# 定数
# -----------------------------------------------------------------------
CONFIG_PATH = os.path.join(_get_launcher_dir(), CONFIG_DEFAULT)

# UTF-8 → CP932 → latin-1 の順でフォールバックデコード
def _decode_auto(raw: bytes) -> str:
    for enc in ("utf-8", "cp932", "latin-1"):
        try:
            return raw.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return raw.decode("utf-8", errors="replace")

# -----------------------------------------------------------------------
# print() の出力を GUI ログに転送するラッパー
# -----------------------------------------------------------------------
class _GuiWriter:
    def __init__(self, log_fn, file_path=None):
        self._log = log_fn
        self._file = open(file_path, "w", encoding="utf-8", buffering=1) if file_path else None
    def write(self, text):
        if text:
            self._log(text)
            if self._file:
                self._file.write(text)
    def flush(self):
        if self._file:
            self._file.flush()
    def close(self):
        if self._file:
            self._file.close()
    def __getattr__(self, name):
        return getattr(sys.__stdout__, name)

# -----------------------------------------------------------------------
# GUI本体
# -----------------------------------------------------------------------
class Slide2MovieGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{COMMON_TEXT['PROGRAM_NAME']} GUI   {COMMON_TEXT['VERSION_NUMBER']}")
        self.iconbitmap(os.path.join(_get_base_dir(), "icon.ico"))
        self.resizable(True, True)
        self.minsize(620, 700)

        # OPTION_DEFS から tk変数を自動生成
        self._vars: dict[str, tk.Variable] = {}
        for opt in OPTION_DEFS:
            name    = opt["name"]
            default = opt["default"]
            if opt["store_true"] or opt["type"] == bool:
                self._vars[name] = tk.BooleanVar(value=bool(default))
            elif opt["type"] == int:
                self._vars[name] = tk.IntVar(value=int(default) if default is not None else 0)
            else:
                self._vars[name] = tk.StringVar(value=str(default) if default is not None else "")
        # UI構築
        self._build_ui()
        # INIファイルから設定を読み込む
        self._load_config()

    # -------------------------------------------------------------------
    # UI構築
    # -------------------------------------------------------------------
    def _build_ui(self):
        pad = {"padx": 10, "pady": 4}
        # ファイル参照が必要なキーワード
        FILE_KEYWORDS = ("file", "img", "image", "path", "output")

        # frame名の順序を保ちながらグループ化
        from collections import OrderedDict
        groups: OrderedDict[str, list] = OrderedDict()
        for opt in OPTION_DEFS:
            frame_name = opt.get("frame", "その他")
            groups.setdefault(frame_name, []).append(opt)

        # グループごとに LabelFrame を生成
        for frame_name, opts in groups.items():
            lf = ttk.LabelFrame(self, text=frame_name, padding=8)
            lf.pack(fill="x", **pad)

            row = 0
            for opt in opts:
                name  = opt["name"]
                label = opt["help"] or name
                var   = self._vars[name]

                if opt["store_true"] or opt["type"] == bool:
                    ttk.Checkbutton(lf, text=label, variable=var).grid(
                        row=row, column=0, columnspan=3, sticky="w", pady=3, padx=4)

                elif opt["type"] == int:
                    ttk.Label(lf, text=label).grid(row=row, column=0, sticky="w", pady=3)
                    ttk.Spinbox(lf, from_=0, to=99999, increment=1,
                                textvariable=var, width=10).grid(
                        row=row, column=1, sticky="w", padx=6)

                elif any(kw in name.lower() for kw in FILE_KEYWORDS):
                    ttk.Label(lf, text=label).grid(row=row, column=0, sticky="w", pady=3)
                    ttk.Entry(lf, textvariable=var, width=40).grid(
                        row=row, column=1, sticky="ew", padx=6)
                    ttk.Button(lf, text="参照…", width=7,
                            command=lambda n=name: self._browse_file(n)).grid(
                        row=row, column=2, padx=2)
                    lf.columnconfigure(1, weight=1)

                else:
                    ttk.Label(lf, text=label).grid(row=row, column=0, sticky="w", pady=3)
                    ttk.Entry(lf, textvariable=var, width=20).grid(
                        row=row, column=1, sticky="w", padx=6)

                row += 1

        # 実行ボタン
        self.btn_run = ttk.Button(self, text="▶ 実行", command=self._on_run)
        self.btn_run.pack(pady=8)

        # ログ出力
        frame_log = ttk.LabelFrame(self, text="実行ログ", padding=8)
        frame_log.pack(fill="both", expand=True, **pad)
        self.log_area = scrolledtext.ScrolledText(
            frame_log, state="disabled", wrap="word",
            font=("Courier", 9), height=12)
        self.log_area.pack(fill="both", expand=True)

    # config.ini が存在すれば読み込んで各ウィジェットの初期値に反映する
    def _load_config(self):
        ini = load_ini(CONFIG_PATH)
        if not ini:
            return

        for opt in OPTION_DEFS:
            name = opt["name"]
            if name not in ini:
                continue
            value = ini[name]
            var = self._vars[name]
            try:
                if isinstance(var, tk.BooleanVar):
                    var.set(bool(value))
                elif isinstance(var, tk.IntVar):
                    var.set(int(value))
                else:
                    var.set(str(value) if value is not None else "")
            except (ValueError, tk.TclError):
                pass  # 型変換失敗時はデフォルト値のまま


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
    def _browse_file(self, name: str):
        var = self._vars[name]

        if name == "output":
            path = filedialog.asksaveasfilename(
                title="出力ファイルを指定",
                defaultextension=".mp4",
                filetypes=[("MP4動画", "*.mp4"), ("すべて", "*.*")])
        elif "img" in name or "image" in name:
            path = filedialog.askopenfilename(
                title="画像ファイルを選択",
                filetypes=[("画像", "*.png *.jpg *.jpeg"), ("すべて", "*.*")])
        else:
            path = filedialog.askopenfilename(
                title=f"{name} を選択",
                filetypes=[("すべて", "*.*")])

        if path:
            var.set(path)
            # 入力PPTXが選ばれたとき出力パスを自動補完
            if name == "file" and not self._vars.get("output", tk.StringVar()).get():
                self._vars["output"].set(os.path.splitext(path)[0] + ".mp4")

    # -------------------------------------------------------------------
    # 実行
    # -------------------------------------------------------------------
    def _on_run(self):
        # バリデーション
        if not self._vars["file"].get():
            messagebox.showwarning("入力エラー", "入力PPTXを指定してください。")
            return
        if not self._vars["output"].get():
            messagebox.showwarning("入力エラー", "出力MP4のパスを指定してください。")
            return

        self.btn_run.config(state="disabled")
        self._log_clear()
        self._log("実行開始...\n")

        # 別スレッドで実行（GUIフリーズ防止）
        thread = threading.Thread(target=self._run_script, daemon=True)
        thread.start()

    def _run_script(self):
        from slide2movie import pptx_to_video
        from mkdir_datetime import mkdir_dt,get_today_date,get_now_time
        from pathlib import Path
        import sys

        # 引数整理
        args = self._vars
        args_val = {key: var.get() for key, var in args.items()}

        if args['debug'].get():
            dbg_dir_path = Path(mkdir_dt())  # デバッグ用にフォルダ作成
            log_path = dbg_dir_path / (Path(args['output'].get()).stem + "_debug.log")
        else:
            dbg_dir_path = None
            log_path = None

        # ラッパー差し替え前の標準出力を退避
        _orig_stdout = sys.stdout
        sys.stdout = _GuiWriter(self._log,file_path=log_path)

        print(f'指定された引数: {args}', flush=True)

        try:
            if args['debug'].get():
                dbg_dir_path = Path(mkdir_dt())  # デバッグ用にフォルダ作成
            else:
                dbg_dir_path = None

            if args['debug'].get():
                print('--- デバッグモード', flush=True)
                save_ini(dbg_dir_path / 'config.ini', args_val)  # デバッグ用に現在の設定値をiniとして保存
                for key, value in args_val.items():
                    if isinstance(value, str):
                        if os.path.exists(value):
                            # デバッグ用にファイルをコピー
                            import shutil
                            dest_path = dbg_dir_path / os.path.basename(value)
                            shutil.copy2(value, dest_path)
                            print(f'  - {key}: {value} を {dest_path} にコピーしました。', flush=True)

            pptx_to_video(
                pptx_path   = args["file"].get(),
                output_mp4  = args["output"].get(),
                dpi         = args["dpi"].get(),
                quality     = args["quality"].get(),
                lang        = args["lang"].get(),
                voicevox    = args["voicevox"].get(),
                voicevoxid  = args["voicevoxid"].get(),
                creditimg   = args["creditimg"].get() or None,
                creditbg    = args["creditbg"].get() or None,
                creditcolor = args["creditcolor"].get() or None,
                debug       = args["debug"].get(),
                dbg_dir_path = dbg_dir_path,
            )
            self._log(f"\n{'─'*60}\n✅ 完了しました。\n出力: {args['output'].get()}\n")

        except Exception as e:
            self._log(f"\n❌ エラーが発生しました: {e}\n")

        finally:
            # 退避していたラッパー差し替え前の標準出力を戻す
            sys.stdout = _orig_stdout
            self.after(0, lambda: self.btn_run.config(state="normal"))

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