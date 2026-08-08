from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox as msgbox
import json
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
import zipfile
import xml.etree.ElementTree as ET
import jaconv
from sudachipy import tokenizer, dictionary
import win32com.client
import glob
from pathlib import Path
import tkinter.filedialog as filedialog
import time

# ✅ バージョン情報（サポート対応時にユーザーへ確認してもらうため、GUI上にも表示する）
APP_VERSION = "2.0"


# ✅ ログ設定（実行フォルダに rubigui.log として出力。
#    ユーザーから不具合報告を受けた際、このファイルを送ってもらえば調査しやすくなる）
# ★重要：以前は logging.basicConfig(filename=...) で単一ファイルに無制限に
#    追記し続けており、長期間使い続けるとログファイルが際限なく肥大化してしまう
#    問題があった。RotatingFileHandler に変更し、1ファイルあたり最大5MBまでとし、
#    上限に達したら rubigui.log.1, .2, .3 ...（最大3世代）にローテーションして
#    古いログから自動的に破棄されるようにした。
_log_handler = RotatingFileHandler(
    "rubigui.log",
    maxBytes=5 * 1024 * 1024,  # 5MBごとにローテーション
    backupCount=3,             # 直近3世代（＋現行分で最大4ファイル、合計20MB程度）まで保持
    encoding="utf-8",
)
_log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

logging.basicConfig(level=logging.DEBUG, handlers=[_log_handler])
logging.getLogger().addHandler(logging.StreamHandler())  # コンソールにも出す（従来のprintの代わり）

# ✅ Sudachi初期化
try:
    tokenizer_obj = dictionary.Dictionary(config_path="sudachi.json").create()
    mode = tokenizer.Tokenizer.SplitMode.C
except Exception as e:
    logging.error(f"Sudachi辞書の初期化に失敗しました: {e}")
    try:
        root = tk.Tk()
        root.withdraw()
        msgbox.showerror(
            "起動エラー",
            "Sudachi辞書の読み込みに失敗しました。\n"
            "sudachi.json の systemDict パスが正しいか、辞書ファイルが\n"
            "指定した場所に展開されているか確認してください。\n\n"
            f"詳細: {e}"
        )
    except Exception:
        pass
    sys.exit(1)


def to_hiragana(katakana):
    return jaconv.kata2hira(katakana)

def is_katakana_only(surface):
    """全角・半角カタカナ（と長音符）のみで構成されているか判定"""
    def is_kana_char(ch):
        return ('\u30A0' <= ch <= '\u30FF') or ('\uFF66' <= ch <= '\uFF9D') or ch == 'ー'
    return all(is_kana_char(ch) for ch in surface)

def is_number_only(surface):
    """半角・全角の数字と区切り記号（.,-：/など）のみで構成されているか判定"""
    number_chars = set("0123456789０１２３４５６７８９.,-:：/／")
    return all(ch in number_chars for ch in surface)

def is_latin_only(surface):
    """半角英字（と半角数字・一般的な区切り記号）のみで構成されているか判定。
    英単語・アルファベット略語などにルビ（カタカナ読み）を振ってしまう不具合の対策。
    Sudachiは英単語に対してもカタカナの読みを返すことがあり、これがsurfaceと
    完全一致しないため従来の「surface == reading」チェックだけではすり抜けて
    しまっていた。半角英字のみで構成される語句は、そもそもルビ付与の対象外とする。"""
    if not surface:
        return False
    allowed_symbols = set(" .,-_'&/()")
    has_alpha = False
    for ch in surface:
        if ch.isascii() and ch.isalpha():
            has_alpha = True
        elif ch.isascii() and ch.isdigit():
            continue
        elif ch in allowed_symbols:
            continue
        else:
            return False
    return has_alpha

def get_actual_desktop_path():
    """OneDriveでデスクトップがリダイレクトされている環境でも、
    エクスプローラー上の実際の「デスクトップ」フォルダを取得する。
    Path.home() / "Desktop" はリダイレクトを考慮しないため、
    レジストリの User Shell Folders から正しいパスを引く。"""
    try:
        import winreg
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        )
        raw_path, _ = winreg.QueryValueEx(key, "Desktop")
        desktop_path = Path(os.path.expandvars(raw_path))
        if desktop_path.exists():
            return desktop_path
    except Exception as e:
        logging.warning(f"レジストリからDesktopパス取得失敗、フォールバックします: {e}")
    return Path.home() / "Desktop"

def get_ruby_project_dirs():
    desktop = get_actual_desktop_path()
    base_dir = desktop / "ルビ振り"
    ruby_dir = base_dir / "ルビデータ"
    output_dir = base_dir / "出力（ルビ付き）"
    ruby_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return ruby_dir, output_dir

def cleanup_temp_file(word_file_path, retries=5, delay=0.5):
    """VBAマクロが %TEMP% フォルダに作る _temp_〇〇.docx を後片付けする。
    【重要】Module1.bas側の修正により、テンポラリファイルは元ファイルと同じ
    フォルダ（OneDrive同期対象）ではなく、ローカルの%TEMP%フォルダに作成される
    ようになった。以前は元ファイルと同じフォルダを見ていたが、それだと
    OneDriveの同期タイミングによっては一度削除したファイルが復活してしまう
    レース条件を避けられないため、そもそも同期対象外の場所を見るように変更。
    Word側のファイルロック解放に少し時間がかかることがあるためリトライもしておく。"""
    word_file_path = Path(word_file_path)
    temp_folder = Path(os.environ.get("TEMP", os.environ.get("TMP", "")))
    temp_path = temp_folder / f"_temp_{word_file_path.name}"
    if not temp_path.exists():
        return
    for attempt in range(retries):
        try:
            temp_path.unlink()
            logging.debug(f"一時ファイル削除完了: {temp_path}")
            return
        except Exception as e:
            logging.debug(f"一時ファイル削除リトライ({attempt + 1}/{retries}): {e}")
            time.sleep(delay)
    logging.warning(f"一時ファイルを削除できませんでした（手動で削除してください）: {temp_path}")

def wait_for_ruby_file(file_path, timeout=15):
    # ルビ付きファイル名を構築
    ruby_filename = f"{Path(file_path).stem}（ルビ）.docx"

    # 保存先の出力ディレクトリを取得
    _, output_dir = get_ruby_project_dirs()
    ruby_path = output_dir / ruby_filename

    # 指定時間内でファイルの出現を監視
    # 0.1秒間隔 = 1秒あたり10回チェックなので、timeout秒待つには timeout*10 回ループする
    # （旧コードは timeout*5 になっており、実際には指定の半分の時間しか待っていなかった）
    for _ in range(timeout * 10):
        if ruby_path.exists():
            return ruby_path
        time.sleep(0.1)

    return None

def save_ruby_word(file_path):
    base_name = Path(file_path).stem
    ext = Path(file_path).suffix
    ruby_path = Path(file_path).with_name(f"{base_name}（ルビ）{ext}")
    _, output_dir = get_ruby_project_dirs()
    final_path = output_dir / Path(file_path).name
    logging.debug("ファイル移動準備")
    logging.debug(f"元ファイル存在チェック: {Path(file_path).exists()}")
    if Path(file_path).exists():
        Path(file_path).replace(final_path)
        logging.debug(f"ファイル移動完了: {file_path} → {final_path}")
    else:
        logging.debug(f"ルビ付きファイルが見つかりません: {ruby_path}")
class RubyEditorApp:
    def center_main_window(self, width=800, height=600):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width // 2) - (width // 2)
        y = (screen_height // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def center_window_auto(self, win):
        win.update_idletasks()
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"+{x}+{y}")

    def __init__(self, root):
        self.root = root
        self.root.title(f"ルビ編集ツール v{APP_VERSION}")
        self.root.geometry("800x600")
        self.center_main_window(800,600)
        self.data = []
        self.override_dict = {}
        self.file_path = None
        self.current_file_path = None
        self.docx_files = []
        self.current_index = 0
        self.setup_ui()
        self.load_override_dict()
        self.ruby_dir, self.output_dir = get_ruby_project_dirs()

    def batch_process(self, event):
        logging.debug(f"batch_process() 呼び出し元データ: {event.data}")
        paths = self.root.tk.splitlist(event.data)
        file_paths = [Path(p.strip()) for p in paths]
        docx_files = [f for f in file_paths if f.suffix.lower() == ".docx" and not f.name.startswith("~$")]

        if not docx_files:
            msgbox.showinfo("一括処理", ".docx ファイルが見つかりません。")
            return

        if len(docx_files) == 1:
            # ✅ 1ファイルだけ → 抽出＋表示だけに留める
            logging.debug("1ファイルのみ → 語句抽出＋表示のみ処理")
            self.extract_words(docx_files[0])
            return

        # ✅ 複数ファイル → フル一括処理へ
        self._process_batch_files(docx_files)

    def process_docx_batch(self, file_paths):
        logging.debug(f"process_docx_batch 呼び出し: {[str(f) for f in file_paths]}")
        docx_files = [f for f in file_paths if f.suffix.lower() == ".docx" and not f.name.startswith("~$")]
        if not docx_files:
            msgbox.showinfo("一括処理", ".docx ファイルが見つかりません。")
            return
        self._process_batch_files(docx_files)

    def _update_batch_progress(self, text):
        if hasattr(self, "progress_label"):
            self.progress_label.config(text=text)
            self.root.update_idletasks()

    def _process_batch_files(self, docx_files):
        """複数ファイルの一括処理本体（ドラッグ＆ドロップ／フォルダ選択どちらからも使う共通処理）。
        Wordは1つだけ起動して使い回し、ファイルごとの起動・終了コストを省く。"""
        total = len(docx_files)
        success_count = 0
        failures = []

        word = None
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
        except Exception as e:
            msgbox.showerror("Word起動エラー", f"Wordの起動に失敗しました。\n{e}")
            return

        try:
            for i, file_path in enumerate(docx_files, start=1):
                file_path = Path(file_path)
                self._update_batch_progress(f"処理中: {i}/{total} - {file_path.name}")
                try:
                    logging.info(f"処理中: {file_path}")
                    terms = extract_terms(file_path, self.override_dict)
                    logging.debug(f"{file_path.name} 語句抽出数: {len(terms)}")
                    if not terms:
                        logging.warning(f"{file_path} → 語句ゼロ（抽出失敗の可能性）")
                    # extract_terms は [{"word":..,"reading":..}, ...] を返すので
                    # save_tsv / self.data が期待する (word, reading) タプル形式に変換する
                    term_pairs = [(t["word"], t["reading"]) for t in terms]
                    self.data = term_pairs
                    self.save_tsv(term_pairs, file_path)
                    self.run_vba_macro(file_path, word_app=word)
                    ruby_file = wait_for_ruby_file(file_path)
                    logging.debug(f"ルビ付きファイル候補: {ruby_file}")
                    ruby_filename = f"{file_path.stem}（ルビ）.docx"
                    if ruby_file:
                        save_ruby_word(ruby_file)
                        success_count += 1
                    else:
                        failures.append(f"{ruby_filename}（ルビ付きファイルが見つかりません）")
                except Exception as e:
                    logging.error(f"{file_path} → {e}")
                    failures.append(f"{file_path.name}（{e}）")
        finally:
            try:
                word.Quit()
            except Exception:
                pass
            self._update_batch_progress("")

        summary = f"{total} 件中 {success_count} 件のファイルを正常に処理しました。"
        if failures:
            summary += "\n\n失敗したファイル:\n" + "\n".join(failures)
        msgbox.showinfo("一括処理", summary)

    def start_batch_review(self):
        if not hasattr(self, "target_folder"):
            msgbox.showwarning("警告", "先に「フォルダ選択」で処理対象フォルダを選んでください。")
            return
        self.docx_files = [f for f in glob.glob(f"{self.target_folder}/*.docx") if not Path(f).name.startswith("~$")]
        self.current_index = 0
        if not self.docx_files:
            msgbox.showinfo("一括処理", f"{self.target_folder} に .docx ファイルが見つかりません。")
            return
        self.process_next_file()

    def process_next_file(self):
        if self.current_index >= len(self.docx_files):
            msgbox.showinfo("完了", "すべてのファイルを処理しました")
            return
        file_path = self.docx_files[self.current_index]
        self.file_path = file_path
        terms = extract_terms(file_path, self.override_dict)
        self.data = [(t["word"], t["reading"]) for t in terms]
        self.tree.delete(*self.tree.get_children())
        for i, (word, reading) in enumerate(self.data):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(word, reading), tags=(tag,))
        msgbox.showinfo("確認", f"{Path(file_path).name} の語句を確認・修正してください")

    def confirm_and_continue(self,advance=True):
        logging.debug(f"confirm_and_continue 呼び出し")
        logging.debug(f"self.file_path: {self.file_path}")
        logging.debug(f"self.data: {self.data}")

        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return
        self.save_tsv(self.data, self.file_path)
        self.run_vba_macro(self.file_path)
        ruby_file = wait_for_ruby_file(self.file_path)
        ruby_filename = f"{Path(self.file_path).stem}（ルビ）.docx"
        if ruby_file:
            save_ruby_word(ruby_file)
        else:
            msgbox.showerror("保存失敗", f"{ruby_filename} のルビ付きファイルが見つかりませんでした。")

        if advance:
            self.current_index += 1
            self.process_next_file()

    def select_folder(self):
        folder = filedialog.askdirectory(title="処理対象フォルダを選択")
        if folder:
            self.target_folder = folder
            self.files = list(Path(folder).glob("*.docx"))  # ← .docxだけを対象にしたリスト化
            self.current_index = 0  # ← 現在処理中のインデックスを初期化
            msgbox.showinfo("フォルダ選択", f"選択されたフォルダ:\n{folder}")

    def setup_ui(self):
        # ✅ バージョン表記（画面下部に常時表示。サポート対応時に問い合わせてもらいやすくするため）
        version_label = tk.Label(
            self.root,
            text=f"RubiGUI v{APP_VERSION}",
            fg="gray50",
            anchor="e",
            font=("", 9),
        )
        version_label.pack(side="bottom", fill="x", padx=8, pady=(0, 4))

        # メインフレーム（2列構成）
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True)

        # グリッド構成：左2/3、右1/3
        main_frame.columnconfigure(0, weight=2)
        main_frame.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        # 左側フレーム（ドロップエリア＋語句テーブル＋辞書ボタン）
        left_frame = tk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

        # ドロップエリア
        self.drop_label = tk.Label(left_frame, text="ここにWordファイルをドロップ", relief="ridge", height=4)
        self.drop_label.pack(fill="x", pady=(0, 10))
        self.drop_label.drop_target_register(DND_FILES)
        self.drop_label.dnd_bind('<<Drop>>', lambda e: self.batch_process(e))

        # --- スクロールバー太くする ---
        style = ttk.Style()
        style.layout("Thick.Vertical.TScrollbar",
            [('Vertical.Scrollbar.trough',
            {'children': [('Vertical.Scrollbar.thumb', {'expand': '1'})],
            'sticky': 'nswe'})]
        )
        style.configure("Thick.Vertical.TScrollbar", arrowsize=18, width=18)

        # --- Treeview の罫線を確実に表示するため classic テーマに変更 ---
        style = ttk.Style()
        style.theme_use("classic")

        # --- 罫線を見やすく強調 ---
        style.configure("Treeview",
            background="white",
            fieldbackground="white",
            bordercolor="gray50",
            borderwidth=1,
            relief="solid",
            rowheight=26,
            font=("", 10)
        )
        style.configure("Treeview.Heading",
            bordercolor="gray50",
            borderwidth=1,
            relief="raised",
            font=("", 10, "bold")
        )
        style.map("Treeview",
            background=[("selected", "#3a7ebf")],
            foreground=[("selected", "white")]
        )

        # --- 語句テーブル（スクロールバー付き） ---
        tree_frame = tk.Frame(left_frame)
        tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        # --- 太いスクロールバーのスタイル定義（これが無いと起動しない） ---
        style.layout("Thick.Vertical.TScrollbar",
            [('Vertical.Scrollbar.trough',
            {'children': [('Vertical.Scrollbar.thumb', {'expand': '1'})],
            'sticky': 'nswe'})]
        )
        style.configure("Thick.Vertical.TScrollbar", arrowsize=18, width=18)

        # --- スクロールバー ---
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", style="Thick.Vertical.TScrollbar")
        tree_scrollbar.pack(side="right", fill="y")

        # --- Treeview ---
        self.tree = ttk.Treeview(
            tree_frame,
            columns=("word", "reading"),
            show="headings",
            yscrollcommand=tree_scrollbar.set
        )
        self.tree.heading("word", text="語句")
        self.tree.heading("reading", text="読み")
        self.tree.pack(fill="both", expand=True)
        # --- 縞模様タグの定義 ---
        self.tree.tag_configure("oddrow", background="#f2f2f2")
        self.tree.tag_configure("evenrow", background="white")

        tree_scrollbar.config(command=self.tree.yview)

        self.tree.bind("<Double-1>", self.edit_item)

        # 語句テーブルの下に辞書ボタン
        dict_btn_frame = tk.Frame(left_frame)
        dict_btn_frame.pack(fill="x", pady=(0, 10))
        tk.Button(dict_btn_frame, text="辞書編集", command=self.edit_override_dict).pack(side="left", padx=5)
        tk.Button(dict_btn_frame, text="辞書再適用", command=self.reapply_ruby).pack(side="right", padx=5)

        # 右側フレーム
        right_frame = tk.Frame(main_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

        # １ファイル処理
        single_file_frame = tk.LabelFrame(right_frame, text="１ファイル処理", padx=10, pady=5)
        single_file_frame.pack(fill="x", pady=(0, 50))
        tk.Button(single_file_frame, text="TSV保存", command=self.save_only_tsv).pack(fill="x", pady=2)
        tk.Button(single_file_frame, text="ルビ付きWord出力", command=self.run_single_file_output).pack(fill="x", pady=2)

        # 一括処理フレーム（ラベルなし）
        batch_frame = tk.Frame(right_frame, relief="groove", bd=2, padx=10, pady=10)
        batch_frame.pack(fill="x")

        # フォルダ選択（ラベルのような位置に配置）
        tk.Button(batch_frame, text="フォルダ選択", command=self.select_folder, width=15).pack(pady=(0, 10))

        # 一括処理ボタン群
        tk.Button(batch_frame, text="📁 一括処理", command=self.on_batch_button_click).pack(fill="x", pady=2)
        tk.Button(batch_frame, text="一括処理(確認あり)", command=self.start_batch_review).pack(fill="x", pady=2)
        tk.Button(batch_frame, text="▶ 次へ", command=self.advance_to_next).pack(fill="x", pady=2)

        # 一括処理の進捗表示
        self.progress_label = tk.Label(batch_frame, text="", fg="gray30", anchor="w")
        self.progress_label.pack(fill="x", pady=(8, 0))

    def run_single_file_output(self):
        if not self.current_file_path:
            msgbox.showwarning("エラー", "先にWordファイルをドロップまたは選択してください。")
            return
        success = self.run_vba_macro(self.current_file_path)
        if not success:
            return  # エラー内容は run_vba_macro 内で表示済み
        ruby_file = wait_for_ruby_file(self.current_file_path)
        if ruby_file:
            msgbox.showinfo("ルビ付きWord出力", f"{Path(ruby_file).name} を保存しました。")
        else:
            msgbox.showwarning(
                "確認",
                "マクロは実行されましたが、出力ファイルの生成を確認できませんでした。\nWord側でエラーが出ていないか確認してください。"
            )

    def select_file_for_processing(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("Wordファイル", "*.docx")])
        if file_path:
            self.process_single_file(Path(file_path))

    def batch_process_from_folder(self):
        folder = filedialog.askdirectory(title="処理対象フォルダを選択")
        if folder:
            docx_files = [Path(f) for f in glob.glob(f"{folder}/*.docx") if not Path(f).name.startswith("~$")]
            dummy_event = type("DummyEvent", (), {"data": " ".join(str(f) for f in docx_files)})
            self.batch_process(dummy_event)

    def process_single_file(self, file_path):
        try:
            terms = extract_terms(file_path, self.override_dict)
            self.save_tsv(self.data, file_path)
            self.run_vba_macro(file_path)
            ruby_file = wait_for_ruby_file(file_path)
            ruby_filename = f"{Path(file_path).stem}（ルビ）.docx"
            if ruby_file:
                save_ruby_word(ruby_file)
                msgbox.showinfo("処理完了", f"{Path(file_path).name} を正常に処理しました。")
            else:
                msgbox.showerror("保存失敗", f"{ruby_filename} のルビ付きファイルが見つかりませんでした。")
        except Exception as e:
            msgbox.showerror("エラー", f"{Path(file_path).name} の処理中にエラーが発生しました。\n{e}")

    def select_file_for_review(self, event=None):
        path = filedialog.askopenfilename(filetypes=[("Wordファイル", "*.docx")])
        if path:
            self.file_path = path
            self.extract_words(path)

    def on_drop(self, event):
        logging.debug(f"self.file_path: {self.file_path}")
        logging.debug(f"self.data: {self.data}")
        logging.debug(f"event.data: {event.data}")
        paths = self.root.tk.splitlist(event.data)
        docx_files = [p for p in paths if p.lower().endswith(".docx") and not Path(p).name.startswith("~$")]

        if not docx_files:
            msgbox.showinfo("ドロップされたファイル", ".docx ファイルが見つかりませんでした。")
            return
        elif len(docx_files) == 1:
            logging.debug(f"1ファイルドロップ: {docx_files[0]}")
            self.file_path = docx_files[0]
            self.extract_words(docx_files[0])
        else:
            logging.debug(f"複数ファイルドロップ: {docx_files}")
            logging.debug(f"batch_process 呼び出し開始 from on_drop()")
            class DummyEvent:
                def __init__(self, data):
                    self.data = data
            self.batch_process(DummyEvent(" ".join(docx_files)))

    def on_batch_button_click(self):
        logging.debug("on_batch_button_click 呼び出し")
        if hasattr(self, "target_folder"):  # ←ここを self.selected_folder → self.target_folder に修正
            docx_files = list(Path(self.target_folder).glob("*.docx"))
            if docx_files:
                self.process_docx_batch(docx_files)
            else:
                msgbox.showinfo("情報", "フォルダ内に Word ファイルが見つかりませんでした。")
        else:
            msgbox.showwarning("警告", "フォルダが選択されていません。")

    def extract_words(self, path):
        logging.debug(f"self.file_path: {self.file_path}")
        logging.debug(f"self.data: {self.data}")
        self.current_file_path = path
        self.file_path = path
        logging.debug("語句抽出開始")
        self.data.clear()
        self.tree.delete(*self.tree.get_children())

        # 抽出ロジックは extract_terms に一本化（バッチ処理と同じ結果になるようにする）
        terms = extract_terms(path, self.override_dict)
        for i, t in enumerate(terms):
            word = t["word"]
            reading = t["reading"]

            self.data.append((word, reading))

            # 偶数行・奇数行でタグを切り替え（縞模様）
            tag = "evenrow" if i % 2 == 0 else "oddrow"

            self.tree.insert("", "end", values=(word, reading), tags=(tag,))

        logging.debug(f"self.current_file_path: {self.current_file_path}")
        logging.debug(f"語句抽出結果 self.data: {self.data}")
        logging.debug(f"語句数: {len(self.data)}")

    def save_only_tsv(self):
        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return
        terms = [(w, r) for w, r in self.data]
        self.save_tsv(terms, self.file_path)
        msgbox.showinfo("TSV保存", f"{len(terms)}件の語句をTSV形式で保存しました。")

    def advance_to_next(self):
        if not self.files or self.current_index >= len(self.files):
            msgbox.showinfo("完了", "すべてのファイルを処理しました")
            return
        current_file = self.files[self.current_index]
        if not self.data:
            msgbox.showwarning("エラー", "処理対象の語句がありません")
            return
        self.save_tsv(self.data, current_file)
        self.run_vba_macro(current_file)
        ruby_file = wait_for_ruby_file(current_file)
        ruby_filename = f"{Path(current_file).stem}（ルビ）.docx"
        if ruby_file:
            save_ruby_word(ruby_file)
        else:
            msgbox.showerror("保存失敗", f"{ruby_filename} のルビ付きファイルが見つかりませんでした。")
        self.current_index += 1
        self.process_next_file()

    def edit_item(self, event):
        item_id = self.tree.focus()
        if not item_id:
            return
        word, reading = self.tree.item(item_id, "values")
        edit_win = tk.Toplevel(self.root)
        edit_win.title("編集")
        edit_win.geometry("300x200")
        tk.Label(edit_win, text="語句").pack()
        word_entry = tk.Entry(edit_win)
        word_entry.insert(0, word)
        word_entry.pack()
        tk.Label(edit_win, text="読み").pack()
        reading_entry = tk.Entry(edit_win)
        reading_entry.insert(0, reading)
        reading_entry.pack()
    
        def save_edit():
            new_word = word_entry.get().strip()
            new_reading = reading_entry.get().strip()

            # ★重要：語句・読みを空欄にして「保存」すると、以前はそのまま
            # self.data / Treeview に空文字の行が残ってしまい、TSV出力にも
            # 空欄行が混ざる不具合があった。削除したい場合は下の「削除」ボタンを
            # 使ってもらうよう案内し、空欄のままの保存はブロックする。
            if not new_word or not new_reading:
                msgbox.showwarning(
                    "エラー",
                    "語句・読みは空欄のまま保存できません。\n"
                    "この語句を削除したい場合は「削除」ボタンを使用してください。"
                )
                return

            # 行番号を取得してタグを付け直す
            index = self.tree.index(item_id)
            tag = "evenrow" if index % 2 == 0 else "oddrow"

            # Treeview の値とタグを更新
            self.tree.item(item_id, values=(new_word, new_reading), tags=(tag,))

            # self.data の更新
            for i, (w, r) in enumerate(self.data):
                if w == word and r == reading:
                    self.data[i] = (new_word, new_reading)
                    break
            edit_win.destroy()

        def delete_edit():
            # 確認なしで即削除すると誤操作が怖いので一言確認する
            if not msgbox.askyesno("削除確認", f"「{word}」をリストから削除しますか？"):
                return
            for i, (w, r) in enumerate(self.data):
                if w == word and r == reading:
                    del self.data[i]
                    break
            self.tree.delete(item_id)
            self._refresh_row_tags()
            edit_win.destroy()

        btn_frame = tk.Frame(edit_win)
        btn_frame.pack(pady=5)
        tk.Button(btn_frame, text="保存", command=save_edit).pack(side="left", padx=5)
        tk.Button(btn_frame, text="削除", command=delete_edit, fg="red").pack(side="left", padx=5)
        self.center_window_auto(edit_win)  # ← ここで中央配置！

    def _refresh_row_tags(self):
        """語句一覧Treeviewの縞模様タグを、現在の並び順で振り直す（行削除後のズレ防止）"""
        for i, item_id in enumerate(self.tree.get_children()):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.item(item_id, tags=(tag,))

    def save_tsv(self, terms, file_path):
        _, output_dir = get_ruby_project_dirs()
        logging.debug("TSV保存開始")
        logging.debug(f"保存対象データ: {terms[:3]}")
        logging.debug(f"保存対象語句数: {len(terms)}")

        base_name = Path(file_path).stem
        tsv_path = self.ruby_dir / f"{base_name}.tsv"

        # ★重要：VBAマクロ側の都合でTSVはcp932（Shift-JIS）で保存する必要があるが、
        # 長文になるほど「cp932に存在しない文字（機種依存文字・一部の記号など）」を
        # 含む語句が出現する確率が上がる。従来はファイル全体を1回のwrite()に近い形で
        # 書き込んでおり、その途中でUnicodeEncodeErrorが発生すると、それ以降の行が
        # 一切書き込まれないまま例外だけがどこかで握りつぶされ、
        # 「TSVが途中までしか出力されずルビ振りが完了しない」不具合の原因になっていた。
        # → 1行ずつ事前にcp932へのエンコード可否をチェックし、
        #    エンコードできない語句だけをスキップして、残りは最後まで書き込む。
        tsv_lines = []
        skipped_terms = []
        for term in terms:
            tsv_line = f"{term[0]}\t{term[1]}"
            try:
                tsv_line.encode("cp932")
            except UnicodeEncodeError as e:
                skipped_terms.append(term[0])
                logging.warning(
                    f"TSV書き込みをスキップ（cp932非対応文字を含むためルビ付与対象外）: "
                    f"{term[0]} → {term[1]}（{e}）"
                )
                continue
            tsv_lines.append(tsv_line)

        logging.debug(f"file_path: {file_path}")
        logging.debug(f"base_name: {base_name}")
        logging.debug(f"output_dir: {output_dir}")
        logging.debug(f"tsv_path: {tsv_path}")
        logging.debug(f"TSVファイル内容（先頭3行）: {tsv_lines[:3]}")

        with open(tsv_path, "w", encoding="cp932") as f:
            for line in tsv_lines:
                f.write(line + "\n")

        logging.debug(f"TSV保存完了: {tsv_path}（{len(tsv_lines)}件、スキップ{len(skipped_terms)}件）")

        if skipped_terms:
            msgbox.showwarning(
                "TSV保存",
                f"{len(skipped_terms)}件の語句は特殊文字を含むためルビ付与対象から除外されました。\n"
                f"（詳細は rubigui.log を確認してください）\n\n"
                + "、".join(skipped_terms[:10])
                + ("　他" if len(skipped_terms) > 10 else "")
            )

    def load_override_dict(self):
        if not os.path.exists("override.json"):
            self.override_dict = {}
            return
        try:
            with open("override.json", "r", encoding="utf-8") as f:
                self.override_dict = json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
            logging.error(f"override.json の読み込みに失敗しました: {e}")
            msgbox.showwarning(
                "辞書読み込みエラー",
                f"override.json の読み込みに失敗したため、強制読み指定なしで続行します。\n{e}"
            )
            self.override_dict = {}

    def edit_override_dict(self):
        edit_win = tk.Toplevel(self.root)
        edit_win.title("辞書編集")
        edit_win.geometry("500x500")
        self.center_window_auto(edit_win)
        tree = ttk.Treeview(edit_win, columns=("word", "reading"), show="headings", selectmode="browse")
        tree.heading("word", text="語句")
        tree.heading("reading", text="読み")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for word, reading in self.override_dict.items():
            tree.insert("", "end", values=(word, reading))

        # 編集ポップアップ
        def edit_item(event=None):
            item_id = tree.focus()
            if not item_id:
                return
            word, reading = tree.item(item_id, "values")
            popup = tk.Toplevel(edit_win)
            popup.title("編集")
            popup.geometry("300x200")
            self.center_window_auto(popup)
            tk.Label(popup, text="語句").pack()
            word_entry = tk.Entry(popup)
            word_entry.insert(0, word)
            word_entry.pack()
            tk.Label(popup, text="読み").pack()
            reading_entry = tk.Entry(popup)
            reading_entry.insert(0, reading)
            reading_entry.pack()

            def save():
                new_word = word_entry.get().strip()
                new_reading = reading_entry.get().strip()
                if new_word and new_reading:
                    self.override_dict.pop(word, None)
                    self.override_dict[new_word] = new_reading
                    tree.item(item_id, values=(new_word, new_reading))
                popup.destroy()
            tk.Button(popup, text="保存", command=save).pack(pady=5)
        # 追加フォーム
        tk.Label(edit_win, text="新規追加").pack(pady=(10, 0))
        add_frame = tk.Frame(edit_win)
        add_frame.pack(pady=5)
        tk.Label(add_frame, text="語句").grid(row=0, column=0)
        new_word_entry = tk.Entry(add_frame)
        new_word_entry.grid(row=0, column=1)
        tk.Label(add_frame, text="読み").grid(row=1, column=0)
        new_reading_entry = tk.Entry(add_frame)
        new_reading_entry.grid(row=1, column=1)

        def add_entry():
            word = new_word_entry.get().strip()
            reading = new_reading_entry.get().strip()
            if word and reading and word not in self.override_dict:
                self.override_dict[word] = reading
                tree.insert("", "end", values=(word, reading))
                new_word_entry.delete(0, "end")
                new_reading_entry.delete(0, "end")
        tk.Button(edit_win, text="追加", command=add_entry).pack(pady=5)
        tree.bind("<Double-1>", edit_item)

        # 削除ボタン
        def delete_entry():
            selected = tree.focus()
            if selected:
                word, _ = tree.item(selected, "values")
                self.override_dict.pop(word, None)
                tree.delete(selected)
        tk.Button(edit_win, text="選択した語句を削除", command=delete_entry).pack(pady=5)

        def save_dict():
            with open("override.json", "w", encoding="utf-8") as f:
                json.dump(self.override_dict, f, ensure_ascii=False, indent=2)
            msgbox.showinfo("保存完了", "辞書を保存しました")
            edit_win.destroy()
        tk.Button(edit_win, text="保存して閉じる", command=save_dict).pack(pady=10)

    def apply_ruby(self):
        if not self.data:
            msgbox.showwarning("エラー", "語句一覧が空です")
            return
        updated = 0
        for i, (word, reading) in enumerate(self.data):
            if word in self.override_dict:
                new_reading = self.override_dict[word]
                self.data[i] = (word, new_reading)
                updated += 1
        self.tree.delete(*self.tree.get_children())
        for word, reading in self.data:
            self.tree.insert("", "end", values=(word, reading))
        msgbox.showinfo("再適用完了", f"辞書の読みを {updated} 件適用しました")

    def reapply_ruby(self):
        if not self.file_path:
            msgbox.showwarning("ファイル未選択", "先にファイルを開いてください。")
            return
        self.load_override_dict()  # override.json を再読み込み
        self.extract_words(self.file_path)  # 語句抽出のみ
        msgbox.showinfo("完了", "辞書を反映して語句を再抽出しました")

    
    def run_vba_macro(self, word_file_path, word_app=None):
        """Wordマクロを実行する。
        word_app を渡すと既存のWordインスタンスを使い回す（一括処理用、高速化のため）。
        渡さない場合は従来どおり関数内でWordを起動・終了する。"""
        logging.debug(f"run_vba_macro に渡された Word path: {word_file_path}")

        # Wordファイル名からTSVファイル名を生成
        base_name = Path(word_file_path).stem
        tsv_path = self.ruby_dir / f"{base_name}.tsv"
        logging.debug(f"対応するTSVファイル: {tsv_path}")
        if not Path(word_file_path).exists():
            msgbox.showerror("エラー", "対象のWordファイルが存在しません")
            return False

        macro_name = "InsertFuriganaFromTSV_SaveToNewFile_Stable"
        owns_word_instance = word_app is None
        word = word_app
        doc = None
        success = False
        actual_dir = None  # Wordが実際に認識しているフォルダパス（OneDriveリダイレクト対策）
        try:
            import win32com.client
            if word is None:
                word = win32com.client.Dispatch("Word.Application")
                word.Visible = False
            doc = word.Documents.Open(str(Path(word_file_path).resolve()))
            # ★重要：PythonがPath.resolve()で計算したフォルダと、
            # Wordが内部的に認識しているフォルダ（doc.Path）は、
            # OneDriveでDesktop等がリダイレクトされている環境では
            # 文字列として一致しないことがある（例：
            # "C:\Users\xxx\Desktop" vs "C:\Users\xxx\OneDrive\Desktop"）。
            # VBA側のtempPathはdocOriginal.Path（＝doc.Path）を基準に
            # 作られるため、後片付けもこちらを基準にしないと
            # 「そんなファイルは無い」と誤判定して削除を試みすらしない。
            try:
                actual_dir = Path(doc.Path)
                logging.debug(f"Wordが認識している実際のフォルダ: {actual_dir}")
            except Exception as e:
                logging.debug(f"doc.Path取得に失敗、Python側のパスにフォールバック: {e}")
            word.Run(macro_name)
            success = True
        except Exception as e:
            msgbox.showerror("VBA実行エラー", str(e))
        finally:
            # 途中で例外が起きても後始末は必ず行う
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass

            # 保険：VBAマクロ側の実装ミスなどでdoc以外にドキュメントが
            # 開いたまま残っていた場合に備え、_temp_〇〇ファイルを参照している
            # ドキュメントが他に残っていないか確認し、あれば強制的に閉じる。
            # （本来はModule1.bas側の修正で発生しなくなっているはずだが、
            #   一括処理でWordインスタンスを使い回す都合上、ここで漏れを
            #   吸収できるようにしておく）
            if word is not None:
                try:
                    temp_name = f"_temp_{Path(word_file_path).name}"
                    # word.Documentsはインデックスが1始まりで、逆順に走査しないと
                    # Close中にインデックスがずれて一部を見逃すことがある
                    for i in range(word.Documents.Count, 0, -1):
                        stray_doc = word.Documents(i)
                        try:
                            stray_name = stray_doc.Name
                        except Exception:
                            continue
                        if stray_name == temp_name:
                            logging.warning(
                                f"未クローズのテンポラリドキュメントを検出、強制的に閉じます: {stray_name}"
                            )
                            try:
                                stray_doc.Close(False)
                            except Exception as e:
                                logging.warning(f"強制Closeに失敗: {e}")
                except Exception as e:
                    logging.debug(f"未クローズドキュメントのチェック中にエラー（無視して続行）: {e}")

            # 自分でWordを起動した場合のみ終了する（使い回し時は呼び出し元が管理）
            if owns_word_instance:
                try:
                    if word is not None:
                        word.Quit()
                except Exception:
                    pass
            # _temp_〇〇.docx の後片付け（Wordがファイルを解放した後に行う）
            cleanup_temp_file(word_file_path)
        return success

def extract_terms(file_path, override_dict):
    import zipfile
    import xml.etree.ElementTree as ET
    import jaconv
    # 起動時に一度だけ作成した tokenizer_obj / mode を再利用する
    # （ファイルごとにフル辞書を読み直すと非常に重くなるため）
    with zipfile.ZipFile(file_path, "r") as docx:
        with docx.open("word/document.xml") as file:
            tree = ET.parse(file)
            root = tree.getroot()
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            texts = [node.text for node in root.findall(".//w:t", ns) if node.text]
    full_text = "".join(texts)
    results = []          # ← dict ではなく list にする
    seen = set()          # ← 重複チェック用

    for m in tokenizer_obj.tokenize(full_text, mode):
        surface = m.surface()

        if len(surface) <= 1 or surface in seen:
            continue
        if all('\u3040' <= ch <= '\u309F' for ch in surface):
            continue

        if surface in override_dict:
            reading = override_dict[surface]
        elif is_katakana_only(surface) or is_number_only(surface) or is_latin_only(surface):
            continue
        else:
            reading = jaconv.kata2hira(m.reading_form())

        if surface == reading:
            continue

        seen.add(surface)
        results.append({"word": surface, "reading": reading})  # ← 出現順で追加

    return results

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = RubyEditorApp(root)
    root.mainloop()

