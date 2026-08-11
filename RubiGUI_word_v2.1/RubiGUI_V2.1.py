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
from pathlib import Path

# ✅ バージョン情報（サポート対応時にユーザーへ確認してもらうため、GUI上にも表示する）
APP_VERSION = "2.1"


def hide_console_window():
    """起動時に出る黒いコンソール画面を隠す。

    ★重要：既に開いているコマンドプロンプトやターミナルから実行された
    場合は隠さない。そのウィンドウはユーザーのものなので、隠すと
    作業中の画面ごと消えてしまう。GetConsoleProcessList が 1 を返す
    （＝このアプリのためだけに作られたコンソール）ときだけ隠す。

    pythonw.exe での起動や、PyInstallerの --noconsole で作ったexeでは
    そもそもコンソールが無いので、何もせずに戻る。"""
    try:
        import ctypes
        from ctypes import wintypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        user32 = ctypes.WinDLL("user32", use_last_error=True)

        hwnd = kernel32.GetConsoleWindow()
        if not hwnd:
            return  # コンソールを持っていない

        # 自分だけがぶら下がっているコンソールかどうかを調べる
        buf = (wintypes.DWORD * 8)()
        attached = kernel32.GetConsoleProcessList(buf, len(buf))
        if attached != 1:
            return  # 別のシェルから起動された → 触らない

        user32.ShowWindow(hwnd, 0)  # SW_HIDE
    except Exception:
        # 隠せなくても動作には支障がないので、黙って続行する
        pass


hide_console_window()


def get_app_dir():
    """設定ファイル・辞書・ログを置くフォルダ（exe／スクリプトと同じ場所）を返す。

    ★重要：v2.0 は "override.json" のように相対パスで開いていたため、
    カレントディレクトリがアプリのフォルダと違う状態（ショートカットの
    「作業フォルダ」が空、別フォルダからの起動など）だと、辞書が
    1件も読み込まれないまま静かに起動していた。「辞書編集を開いても
    一覧が空」という症状の原因になるので、必ずアプリのフォルダを基準にする。
    """
    if getattr(sys, "frozen", False):
        # PyInstallerでexe化した場合。readmeの通り設定ファイルはexeと同じ
        # フォルダに置く運用なので、展開先(_MEIPASS)ではなくexeの場所を使う。
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
OVERRIDE_PATH = APP_DIR / "override.json"
SETTINGS_PATH = APP_DIR / "ruby_settings.json"
SUDACHI_CONFIG_PATH = APP_DIR / "sudachi.json"
LOG_PATH = APP_DIR / "rubigui.log"

# Word側マクロ（Module1.bas）のマクロ名。
# ★重要：v2.0 の "InsertFuriganaFromTSV_SaveToNewFile_Stable" から改名している。
# v2.1 はTSVの書式（UTF-8＋設定行）とマクロの引数が変わっており、旧マクロに
# 新しいTSVを渡すと「何も起きない（ルビが振られないだけ）」という分かりにくい
# 失敗になる。名前を変えておけば word.Run が確実に失敗するので、
# 「Module1.bas を入れ直してください」と明示的に案内できる。
MACRO_NAME = "InsertFuriganaFromTSV_V21"

# TSVの設定行（先頭の「#」行）の終わりを示すマーカー。
# これ以降の行は「#」で始まっていても語句として扱われる。
TSV_HEADER_END = "#DATA"

# 「Wordが忙しくて呼び出しを受け付けられない」ことを表すCOMのエラーコード。
# マクロの有無とは無関係なので、これらは別の案内にする。
BUSY_HRESULTS = (
    -2147418111,  # RPC_E_CALL_REJECTED         呼び出し先によって拒否されました
    -2147417846,  # RPC_E_SERVERCALL_RETRYLATER アプリケーションがビジー状態です
    -2147417851,  # RPC_E_SERVERFAULT
)

# 「Wordとの接続が切れた」ことを表すCOMのエラーコード。
# 処理中にユーザーがWordを閉じた／Wordが強制終了した／Word自身が
# クラッシュした場合に返る。マクロの有無とは無関係。
DISCONNECTED_HRESULTS = (
    -2147417848,  # RPC_E_DISCONNECTED        呼び出し先が切断されました
    -2147023174,  # RPC_S_SERVER_UNAVAILABLE  RPCサーバーを利用できません
    -2147023170,  # RPC_S_CALL_FAILED         リモートプロシージャコールに失敗しました
)


# ✅ ログ設定（アプリのフォルダに rubigui.log として出力。
#    ユーザーから不具合報告を受けた際、このファイルを送ってもらえば調査しやすくなる）
# ★重要：以前は logging.basicConfig(filename=...) で単一ファイルに無制限に
#    追記し続けており、長期間使い続けるとログファイルが際限なく肥大化してしまう
#    問題があった。RotatingFileHandler に変更し、1ファイルあたり最大5MBまでとし、
#    上限に達したら rubigui.log.1, .2, .3 ...（最大3世代）にローテーションして
#    古いログから自動的に破棄されるようにした。
_log_handler = RotatingFileHandler(
    str(LOG_PATH),
    maxBytes=5 * 1024 * 1024,  # 5MBごとにローテーション
    backupCount=3,             # 直近3世代（＋現行分で最大4ファイル、合計20MB程度）まで保持
    encoding="utf-8",
)
_log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

logging.basicConfig(level=logging.DEBUG, handlers=[_log_handler])
# コンソールにも出す（従来のprintの代わり）。
# ★コンソールを隠した場合や pythonw / --noconsole のexeでは sys.stderr が
# None になる。そのまま StreamHandler を足すと、ログを出すたびに内部で
# 例外が起きて握りつぶされる（無駄な処理になる）ので、あるときだけ足す。
if sys.stderr is not None:
    logging.getLogger().addHandler(logging.StreamHandler())

# ✅ Sudachi初期化
try:
    tokenizer_obj = dictionary.Dictionary(config_path=str(SUDACHI_CONFIG_PATH)).create()
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
            f"設定ファイル: {SUDACHI_CONFIG_PATH}\n"
            f"詳細: {e}"
        )
    except Exception:
        pass
    sys.exit(1)


# ============================================================
# ルビ設定
# ============================================================
DEFAULT_SETTINGS = {
    # "first" = 最初に出てくる箇所だけ / "all" = すべての出現箇所
    "ruby_mode": "first",
    "ruby_ratio": 50,     # ルビのフォントサイズ（親文字に対する％）
    "ruby_offset": 0.0,   # ルビと親文字の間隔（pt。0で標準）
}
RUBY_MODE_FIRST = "first"
RUBY_MODE_ALL = "all"


def load_settings():
    """ルビ設定を ruby_settings.json から読み込む（無ければ既定値）"""
    settings = dict(DEFAULT_SETTINGS)
    if not SETTINGS_PATH.exists():
        return settings
    try:
        with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        for key in DEFAULT_SETTINGS:
            if key in loaded:
                settings[key] = loaded[key]
    except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
        logging.error(f"{SETTINGS_PATH.name} の読み込みに失敗しました: {e}")
    if settings["ruby_mode"] not in (RUBY_MODE_FIRST, RUBY_MODE_ALL):
        settings["ruby_mode"] = DEFAULT_SETTINGS["ruby_mode"]
    return settings


def save_settings(settings):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
    except OSError as e:
        logging.error(f"{SETTINGS_PATH.name} の保存に失敗しました: {e}")


# ============================================================
# 語句フィルタ
# ============================================================
def to_hiragana(katakana):
    return jaconv.kata2hira(katakana)


def is_katakana_only(surface):
    """全角・半角カタカナ（と長音符）のみで構成されているか判定"""
    def is_kana_char(ch):
        return ('゠' <= ch <= 'ヿ') or ('ｦ' <= ch <= 'ﾝ') or ch == 'ー'
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


# ============================================================
# 出力先フォルダ
# ============================================================
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


def get_word():
    """Wordを取得する。戻り値は (アプリ, 自分で起動したか)。

    ★重要：win32com の Dispatch は「起動中のWordがあればそれにアタッチする」
    仕様のため、無条件に Visible = False にしたり、処理の最後に Quit() を
    呼んだりすると、ユーザーが別の文書を編集中だった場合にそのウィンドウを
    いきなり消したり、インスタンスごと落としたりしてしまう。
    先に GetActiveObject で既存インスタンスを探し、見つかった場合は
    「自分で起動したのではない」と覚えておいて Visible も Quit も触らない。
    （PowerPoint版の get_powerpoint() と同じ考え方）"""
    try:
        return win32com.client.GetActiveObject("Word.Application"), False
    except Exception:
        word = win32com.client.Dispatch("Word.Application")
        try:
            word.Visible = False
        except Exception:
            pass
        return word, True


def get_ruby_project_dirs():
    desktop = get_actual_desktop_path()
    base_dir = desktop / "ルビ振り"
    ruby_dir = base_dir / "ルビデータ"
    output_dir = base_dir / "出力（ルビ付き）"
    ruby_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return ruby_dir, output_dir


# ============================================================
# 本文の取り出しと形態素解析
# ============================================================
NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def extract_docx_text(file_path):
    """.docx の本文を段落ごとに取り出して改行で連結する。

    ★重要：v2.0 は w:t（テキスト断片）を全部 "" で繋いでいたため、
    前の段落の末尾と次の段落の先頭がひと続きの語として解析され、
    実在しない語句（「会社」＋「員研修」→「会社員研修」など）が
    一覧に載ることがあった。必ず段落単位で改行を挟む。

    ★重要：テキストボックス（w:txbxContent）内の文字は、Wordマクロ側の
    docNew.Content には含まれずルビを振れない。一覧に出しても振られない
    語句が増えるだけなので、解析前に取り除いておく。"""
    with zipfile.ZipFile(file_path) as z:
        xml_bytes = z.read("word/document.xml")
    root = ET.fromstring(xml_bytes)

    # ElementTree には親参照が無いので、削除用に親子関係を作ってから外す
    parents = {child: parent for parent in root.iter() for child in parent}
    for textbox in list(root.iter(f"{NS_W}txbxContent")):
        parent = parents.get(textbox)
        if parent is not None:
            try:
                parent.remove(textbox)
            except ValueError:
                # 入れ子のテキストボックスで、外側と一緒に既に外れている場合
                pass

    lines = []
    for para in root.iter(f"{NS_W}p"):
        # 段落内は書式の切れ目で w:t が分かれるだけなので区切らずに連結する
        text = "".join(t.text or "" for t in para.iter(f"{NS_W}t"))
        if text:
            lines.append(text)
    return "\n".join(lines)


# SudachiPy(0.6系)は入力が49149バイトを超えると例外を投げる。余裕をみて分割する。
MAX_TOKENIZE_BYTES = 40000


def _split_for_tokenizer(text, limit=MAX_TOKENIZE_BYTES):
    """SudachiPyの入力長制限に収まるよう分割する。
    戻り値は [(元テキストでの開始位置, 断片), ...]。
    ★重要：ページ数の多い文書では本文全体が軽く数万バイトを超え、
    分割しないと解析そのものが例外で落ちる。切れ目はできるだけ改行や
    句読点に合わせ、語の途中で切れて誤った読みが出るのを避ける。"""
    chunks = []
    start = 0
    length = len(text)
    while start < length:
        if len(text[start:].encode("utf-8")) <= limit:
            chunks.append((start, text[start:]))
            break
        # limitバイト以内に収まる最大の文字位置を二分探索で求める
        lo, hi = start + 1, length
        while lo < hi:
            mid = (lo + hi + 1) // 2
            if len(text[start:mid].encode("utf-8")) <= limit:
                lo = mid
            else:
                hi = mid - 1
        cut = lo
        for sep in ("\n", "\r", "。", "、", "　", " "):
            pos = text.rfind(sep, start + 1, cut)
            if pos > start:
                cut = pos + 1
                break
        chunks.append((start, text[start:cut]))
        start = cut
    return chunks


def tokenize_all(text):
    """入力長制限を気にせず形態素解析する。
    (表層形, 読み, 元テキストでの開始位置) を順に返す。"""
    for offset, chunk in _split_for_tokenizer(text):
        for m in tokenizer_obj.tokenize(chunk, mode):
            yield m.surface(), m.reading_form(), offset + m.begin()


def extract_terms(file_path, override_dict):
    """Wordファイルから語句と読みを抽出する（出現順・重複排除）"""
    full_text = extract_docx_text(file_path)

    results = []
    seen = set()

    for surface, reading_form, _begin in tokenize_all(full_text):
        if surface in seen:
            continue

        # ★重要：辞書（override.json）の判定は文字数チェックより先に行う。
        # 後ろに置くと「私→わたし」のような1文字の強制読み指定が
        # 一生効かず、辞書に登録しても無反応になってしまう。
        if surface in override_dict:
            reading = override_dict[surface]
        else:
            if len(surface) <= 1:
                continue
            # ひらがなのみは対象外
            if all('぀' <= ch <= 'ゟ' for ch in surface):
                continue
            if is_katakana_only(surface) or is_number_only(surface) or is_latin_only(surface):
                continue
            reading = to_hiragana(reading_form)

        # ★読みが語句とまったく同じものは、辞書に登録されていても対象外にする。
        # 辞書の判定を文字数チェックより先に出した副作用で、
        # {"ひつまぶし": "ひつまぶし"} のような登録が素通りしてしまい、
        # 「ひつまぶし」の上に「ひつまぶし」が重なって振られていた。
        if surface == reading:
            continue

        seen.add(surface)
        results.append({"word": surface, "reading": reading})

    return results


# ============================================================
# GUI
# ============================================================
class RubyEditorApp:
    def center_main_window(self, width=860, height=640):
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

    def center_window(self, win, width, height):
        """サイズを指定して画面中央へ置く（geometryで大きさも決めたい窓用）"""
        win.update_idletasks()
        x = (win.winfo_screenwidth() // 2) - (width // 2)
        y = (win.winfo_screenheight() // 2) - (height // 2)
        win.geometry(f"{width}x{height}+{x}+{y}")

    def __init__(self, root):
        self.root = root
        self.root.title(f"ルビ編集ツール v{APP_VERSION}")
        self.center_main_window(860, 640)
        self.data = []
        self.override_dict = {}
        self.settings = load_settings()
        self.file_path = None
        self.current_file_path = None
        self.docx_files = []
        self.current_index = 0
        self.review_ready = False  # 「一括処理(確認あり)」で語句を表示済みか
        self.ruby_dir, self.output_dir = get_ruby_project_dirs()
        self.setup_ui()
        self.load_override_dict()

    # ---------- 設定 ----------
    def read_settings_from_ui(self):
        """入力欄の値を検証して self.settings へ取り込む。不正なら既定値に戻す。"""
        def to_float(entry, key, minimum, maximum):
            raw = entry.get().strip()
            fallback = float(DEFAULT_SETTINGS[key])
            try:
                value = float(raw)
            except ValueError:
                msgbox.showwarning(
                    "ルビ設定",
                    f"「{raw}」は数値として読み取れないため、既定値 {DEFAULT_SETTINGS[key]} を使用します。"
                )
                value = None
            if value is not None and not (minimum <= value <= maximum):
                msgbox.showwarning(
                    "ルビ設定",
                    f"{value} は指定できる範囲（{minimum}〜{maximum}）外のため、"
                    f"既定値 {DEFAULT_SETTINGS[key]} を使用します。"
                )
                value = None
            if value is None:
                # 入力欄も直しておかないと、次回も同じ警告が出続けてしまう
                entry.delete(0, "end")
                entry.insert(0, str(DEFAULT_SETTINGS[key]))
                return fallback
            return value

        self.settings["ruby_ratio"] = to_float(self.ratio_entry, "ruby_ratio", 10, 100)
        # ★PhoneticGuide の Raise は「親文字からの距離」なので負値は指定できない
        self.settings["ruby_offset"] = to_float(self.offset_entry, "ruby_offset", 0, 50)
        self.settings["ruby_mode"] = self.ruby_mode_var.get()
        save_settings(self.settings)
        return self.settings

    def on_ruby_mode_changed(self):
        """ルビ範囲は語句一覧の中身に影響しないので、設定を保存するだけでよい"""
        self.settings["ruby_mode"] = self.ruby_mode_var.get()
        save_settings(self.settings)

    # ---------- 画面 ----------
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

        # --- 太いスクロールバーのスタイル定義（辞書編集画面でも使う） ---
        style.layout("Thick.Vertical.TScrollbar",
            [('Vertical.Scrollbar.trough',
            {'children': [('Vertical.Scrollbar.thumb', {'expand': '1'})],
            'sticky': 'nswe'})]
        )
        style.configure("Thick.Vertical.TScrollbar", arrowsize=18, width=18)

        # --- 語句テーブル（スクロールバー付き） ---
        tree_frame = tk.Frame(left_frame)
        tree_frame.pack(fill="both", expand=True, pady=(0, 10))

        tree_scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", style="Thick.Vertical.TScrollbar")
        tree_scrollbar.pack(side="right", fill="y")

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
        single_file_frame.pack(fill="x", pady=(0, 12))
        tk.Button(single_file_frame, text="TSV保存", command=self.save_only_tsv).pack(fill="x", pady=2)
        tk.Button(single_file_frame, text="ルビ付きWord出力", command=self.run_single_file_output).pack(fill="x", pady=2)

        # ルビ設定
        settings_frame = tk.LabelFrame(right_frame, text="ルビ設定", padx=10, pady=5)
        settings_frame.pack(fill="x", pady=(0, 12))

        tk.Label(settings_frame, text="ルビの大きさ（親文字の％）", anchor="w").pack(fill="x")
        self.ratio_entry = tk.Entry(settings_frame)
        self.ratio_entry.insert(0, str(self.settings["ruby_ratio"]))
        self.ratio_entry.pack(fill="x", pady=(0, 4))

        tk.Label(settings_frame, text="ルビの高さ（pt・0で標準）", anchor="w").pack(fill="x")
        self.offset_entry = tk.Entry(settings_frame)
        self.offset_entry.insert(0, str(self.settings["ruby_offset"]))
        self.offset_entry.pack(fill="x", pady=(0, 4))

        # ルビを振る範囲（v2.1で追加）
        mode_frame = tk.LabelFrame(right_frame, text="ルビを振る範囲", padx=10, pady=5)
        mode_frame.pack(fill="x", pady=(0, 12))
        self.ruby_mode_var = tk.StringVar(value=self.settings["ruby_mode"])
        tk.Radiobutton(
            mode_frame,
            text="初回のみ（最初の1箇所）",
            value=RUBY_MODE_FIRST,
            variable=self.ruby_mode_var,
            anchor="w",
            justify="left",
            command=self.on_ruby_mode_changed,
        ).pack(fill="x")
        tk.Radiobutton(
            mode_frame,
            text="すべての出現箇所",
            value=RUBY_MODE_ALL,
            variable=self.ruby_mode_var,
            anchor="w",
            justify="left",
            command=self.on_ruby_mode_changed,
        ).pack(fill="x")

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

    # ---------- ファイル受け取り ----------
    def batch_process(self, event):
        logging.debug(f"batch_process() 呼び出し元データ: {event.data}")
        paths = self.root.tk.splitlist(event.data)
        file_paths = [Path(p.strip()) for p in paths]
        docx_files = [f for f in file_paths
                      if f.suffix.lower() == ".docx" and self._is_processable(f)]

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
        docx_files = [f for f in file_paths
                      if f.suffix.lower() == ".docx" and self._is_processable(f)]
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
        settings = self.read_settings_from_ui()
        total = len(docx_files)
        success_count = 0
        total_ruby = 0
        failures = []

        try:
            word, own_word = get_word()
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
                    term_pairs = [(t["word"], t["reading"]) for t in terms]
                    self.save_tsv(term_pairs, file_path, settings)
                    result = self.run_vba_macro(file_path, word_app=word, settings=settings)
                    success_count += 1
                    total_ruby += result.get("applied") or 0
                except Exception as e:
                    logging.error(f"{file_path} → {e}")
                    failures.append(f"{file_path.name}（{e}）")
        finally:
            # 元から起動していたWordは終了しない（ユーザーの編集中の文書を守るため）
            if own_word:
                try:
                    word.Quit()
                except Exception:
                    pass
            self._update_batch_progress("")
            # ★重要：一括処理中は語句一覧の表示を更新していないので、
            # 処理後の self.data / self.file_path をそのまま残すと
            # 「画面はAのまま、中身は最後に処理したファイルの語句」という
            # ちぐはぐな状態になる。この状態で「ルビ付きWord出力」や
            # 「TSV保存」を押すと、別ファイルの読みでルビが振られてしまう。
            self._clear_current_file()

        summary = f"{total} 件中 {success_count} 件のファイルを正常に処理しました。"
        summary += f"\nルビ {total_ruby} 件を付与しました。"
        if failures:
            summary += "\n\n失敗したファイル:\n" + "\n".join(failures)
        msgbox.showinfo("一括処理", summary)

    def _clear_current_file(self):
        """画面と内部状態をまとめて空にする（別ファイルの語句が混ざるのを防ぐ）"""
        self.data = []
        self.file_path = None
        self.current_file_path = None
        self.review_ready = False
        try:
            self.tree.delete(*self.tree.get_children())
        except Exception:
            pass

    @staticmethod
    def _is_processable(path):
        """処理対象にしてよい .docx か。
        「~」で始まるものは Word のロックファイル（~$〇〇.docx）と、
        マクロが出力フォルダに作る作業用コピー（~rubigui_〇〇.docx）なので除く。"""
        return not Path(path).name.startswith("~")

    def _list_docx(self, folder):
        """フォルダ内の .docx を列挙する。
        ★重要：glob.glob はフォルダ名の [ ] をワイルドカードとして解釈するため、
        「資料[2024]」のような名前のフォルダで0件になる。必ず Path.glob を使う。"""
        return [f for f in Path(folder).glob("*.docx") if self._is_processable(f)]

    def start_batch_review(self):
        if not getattr(self, "target_folder", None):
            msgbox.showwarning("警告", "先に「フォルダ選択」で処理対象フォルダを選んでください。")
            return
        self.docx_files = self._list_docx(self.target_folder)
        self.current_index = 0
        if not self.docx_files:
            msgbox.showinfo("一括処理", f"{self.target_folder} に .docx ファイルが見つかりません。")
            return
        self.process_next_file()

    def process_next_file(self):
        # ★読み込みに失敗したファイルは飛ばして次へ進むが、ここを再帰で
        # 書くと読めないファイルが大量にあるフォルダで RecursionError に
        # なるため、ループで回す。
        while self.current_index < len(self.docx_files):
            file_path = Path(self.docx_files[self.current_index])
            self.extract_words(file_path)
            if self.current_file_path is not None:
                self.review_ready = True
                msgbox.showinfo("確認", f"{file_path.name} の語句を確認・修正してください")
                return
            # 読み込みに失敗（extract_words内でエラー表示済み）
            self.current_index += 1

        msgbox.showinfo("完了", "すべてのファイルを処理しました")
        self._clear_current_file()

    def select_folder(self):
        folder = filedialog.askdirectory(title="処理対象フォルダを選択")
        if folder:
            self.target_folder = folder
            self.docx_files = self._list_docx(folder)
            self.current_index = 0
            # フォルダを選び直した時点では、まだどのファイルも確認していない
            self.review_ready = False
            msgbox.showinfo("フォルダ選択", f"選択されたフォルダ:\n{folder}")

    def on_batch_button_click(self):
        logging.debug("on_batch_button_click 呼び出し")
        if getattr(self, "target_folder", None):
            docx_files = self._list_docx(self.target_folder)
            if docx_files:
                self.process_docx_batch(docx_files)
            else:
                msgbox.showinfo("情報", "フォルダ内に Word ファイルが見つかりませんでした。")
        else:
            msgbox.showwarning("警告", "フォルダが選択されていません。")

    def select_file_for_processing(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("Wordファイル", "*.docx")])
        if file_path:
            self.extract_words(Path(file_path))

    # ---------- 語句一覧 ----------
    def extract_words(self, path):
        logging.debug("語句抽出開始")
        self.data = []
        self.current_file_path = None
        self.file_path = None
        self.tree.delete(*self.tree.get_children())

        try:
            terms = extract_terms(path, self.override_dict)
        except Exception as e:
            logging.error(f"語句抽出に失敗しました: {e}")
            msgbox.showerror(
                "読み込みエラー",
                f"{Path(path).name} を読み込めませんでした。\n"
                f"Wordファイル（.docx）として開けるか確認してください。\n"
                f"（.doc / .rtf / .pdf は非対応です）\n\n{e}"
            )
            return

        self.current_file_path = path
        self.file_path = path

        for i, t in enumerate(terms):
            word = t["word"]
            reading = t["reading"]
            self.data.append((word, reading))
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.insert("", "end", values=(word, reading), tags=(tag,))

        logging.debug(f"語句数: {len(self.data)}")

    def _confirm_discard_edits(self):
        """語句一覧を作り直すと手で直した読みが消えるので、消える前に一言確認する"""
        if not self.data:
            return True
        return msgbox.askyesno(
            "語句一覧の再読み込み",
            "語句一覧を読み込み直します。\n"
            "手動で修正した読みは失われますがよろしいですか？"
        )

    def reapply_ruby(self):
        """辞書を読み直して、現在開いているファイルの語句一覧に反映する"""
        if not self.current_file_path:
            msgbox.showwarning("辞書再適用", "先にWordファイルをドロップまたは選択してください。")
            return
        if not self._confirm_discard_edits():
            return
        self.load_override_dict()
        self.extract_words(self.current_file_path)
        msgbox.showinfo("辞書再適用", "辞書を読み込み直して語句一覧に反映しました。")

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
        self.center_window_auto(edit_win)

    def _refresh_row_tags(self):
        """語句一覧Treeviewの縞模様タグを、現在の並び順で振り直す（行削除後のズレ防止）"""
        for i, item_id in enumerate(self.tree.get_children()):
            tag = "evenrow" if i % 2 == 0 else "oddrow"
            self.tree.item(item_id, tags=(tag,))

    # ---------- 出力 ----------
    def save_tsv(self, terms, file_path, settings=None):
        """語句・読みと、VBAへ渡すルビ設定をTSVへ書き出す。

        ★v2.1でcp932からUTF-8へ変更した。VBAの Open ... For Input は
        ANSI固定だったため、v2.0 は「cp932で表現できない語句（機種依存文字を
        含むものなど）」をPython側で除外するしかなかった。Module1.bas を
        ADODB.Stream でのUTF-8読み込みに変えたので、その制限は無くなった。

        ★語句は「表層形の長い順」に並べる。VBAは上から順にルビを振るので、
        長い語（例：日本人）を先に処理しておけば、短い語（例：日本）が
        ルビ化済みの語の中へ食い込むことがない。"""
        if settings is None:
            settings = self.settings

        base_name = Path(file_path).stem
        tsv_path = self.ruby_dir / f"{base_name}.tsv"

        ordered = sorted(terms, key=lambda t: len(t[0]), reverse=True)
        lines = [
            f"#RUBIGUI\t{APP_VERSION}",
            f"#mode\t{settings.get('ruby_mode', DEFAULT_SETTINGS['ruby_mode'])}",
            f"#ratio\t{settings.get('ruby_ratio', DEFAULT_SETTINGS['ruby_ratio'])}",
            f"#offset\t{settings.get('ruby_offset', DEFAULT_SETTINGS['ruby_offset'])}",
            # ★ここまでが設定。これ以降の行は「#」で始まっていても語句として扱う。
            # 終端マーカーが無いと、「#タグ」のような語句を辞書に登録したときに
            # 設定行と誤認されて黙って捨てられてしまう。
            TSV_HEADER_END,
        ]
        lines.extend(f"{word}\t{reading}" for word, reading in ordered)

        # BOM付きUTF-8にしておくとExcelでダブルクリックしても文字化けしない。
        # 改行は "\n" 固定（VBA側は ADODB.Stream の LineSeparator を LF にしている）。
        with open(tsv_path, "w", encoding="utf-8-sig", newline="\n") as f:
            f.write("\n".join(lines) + "\n")

        logging.debug(f"TSV保存完了: {tsv_path}（{len(ordered)}件 / mode={settings.get('ruby_mode')}）")
        return tsv_path

    def save_only_tsv(self):
        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return
        settings = self.read_settings_from_ui()
        try:
            self.save_tsv(self.data, self.file_path, settings)
        except OSError as e:
            logging.error(f"TSVの保存に失敗しました: {e}")
            msgbox.showerror("TSV保存", f"TSVファイルを保存できませんでした。\n{e}")
            return
        msgbox.showinfo("TSV保存", f"{len(self.data)}件の語句をTSV形式で保存しました。")

    def run_single_file_output(self):
        if not self.current_file_path:
            msgbox.showwarning("エラー", "先にWordファイルをドロップまたは選択してください。")
            return
        if not self.data:
            msgbox.showwarning("エラー", "ルビを振る語句がありません。")
            return
        settings = self.read_settings_from_ui()
        try:
            self.save_tsv(self.data, self.current_file_path, settings)
            result = self.run_vba_macro(self.current_file_path, settings=settings)
        except OSError as e:
            logging.error(f"TSVの保存に失敗しました: {e}")
            msgbox.showerror("TSV保存", f"TSVファイルを保存できませんでした。\n{e}")
            return
        except RuntimeError as e:
            msgbox.showerror("ルビ付きWord出力", str(e))
            return

        message = f"{Path(result['out_path']).name} を保存しました。"
        if result.get("applied") is not None:
            message += f"\nルビ {result['applied']} 件を付与しました。"
            if result["applied"] == 0:
                message += "\n\n※ルビが1件も付いていません。語句一覧を確認してください。"
        msgbox.showinfo("ルビ付きWord出力", message)

    def advance_to_next(self):
        if not self.docx_files or self.current_index >= len(self.docx_files):
            msgbox.showinfo("完了", "すべてのファイルを処理しました")
            return
        # ★重要：「フォルダ選択」直後に押されると、画面に残っている別ファイルの
        # 語句でフォルダ内の1件目にルビを振ってしまう。確認画面を経由したか
        # どうかを見て、経由していなければ何もしない。
        if not self.review_ready:
            msgbox.showwarning(
                "警告",
                "先に「一括処理(確認あり)」を押して、語句を確認してから使ってください。"
            )
            return
        current_file = Path(self.docx_files[self.current_index])
        if not self.data:
            msgbox.showwarning("エラー", "処理対象の語句がありません")
            return
        if str(self.current_file_path) != str(current_file):
            msgbox.showwarning(
                "警告",
                f"画面に表示されている語句は {Path(str(self.current_file_path)).name} のものです。\n"
                f"{current_file.name} の処理は行いません。"
            )
            return

        settings = self.read_settings_from_ui()
        try:
            self.save_tsv(self.data, current_file, settings)
            self.run_vba_macro(current_file, settings=settings)
        except (OSError, RuntimeError) as e:
            logging.error(f"{current_file} の処理に失敗しました: {e}")
            msgbox.showerror("保存失敗", f"{current_file.name} の処理に失敗しました。\n{e}")

        self.review_ready = False
        self.current_index += 1
        self.process_next_file()

    # ---------- 辞書 ----------
    def load_override_dict(self):
        if not OVERRIDE_PATH.exists():
            self.override_dict = {}
            return
        try:
            with open(OVERRIDE_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
            logging.error(f"{OVERRIDE_PATH} の読み込みに失敗しました: {e}")
            msgbox.showwarning(
                "辞書読み込みエラー",
                f"override.json の読み込みに失敗したため、強制読み指定なしで続行します。\n"
                f"{OVERRIDE_PATH}\n{e}"
            )
            self.override_dict = {}
            return
        if not isinstance(loaded, dict):
            logging.error(f"{OVERRIDE_PATH} の形式が不正です（辞書ではありません）")
            msgbox.showwarning(
                "辞書読み込みエラー",
                "override.json の形式が正しくありません。\n"
                '{"名古屋": "なごや"} のような形式で記述してください。'
            )
            self.override_dict = {}
            return
        self.override_dict = {str(k): str(v) for k, v in loaded.items()}
        logging.debug(f"辞書読み込み完了: {OVERRIDE_PATH}（{len(self.override_dict)}件）")

    def edit_override_dict(self):
        """辞書編集画面。
        v2.1で「一覧（左）＋操作パネル（右）」の左右レイアウトに作り替え、
        一覧にスクロールバーを付けた。v2.0 は一覧とボタンが上下に積まれており、
        登録語句が増えると下の方が見られなかった。

        ★編集は作業用コピー(working)に対して行い、「保存して閉じる」を
        押したときだけ self.override_dict とファイルへ反映する。v2.0 は
        self.override_dict を即時書き換えていたため、× で閉じると
        メモリ上の辞書とファイルの内容が食い違ったままになっていた。"""
        edit_win = tk.Toplevel(self.root)
        edit_win.title("辞書編集（読みの強制指定）")
        edit_win.minsize(620, 360)
        edit_win.transient(self.root)
        self.center_window(edit_win, 780, 520)

        working = dict(self.override_dict)
        dirty = {"value": False}

        edit_win.columnconfigure(0, weight=1)
        edit_win.columnconfigure(1, weight=0)
        edit_win.rowconfigure(0, weight=1)

        # --- 左：辞書の一覧 ---
        list_frame = tk.Frame(edit_win)
        list_frame.grid(row=0, column=0, sticky="nsew", padx=(10, 5), pady=10)
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", style="Thick.Vertical.TScrollbar")
        tree = ttk.Treeview(
            list_frame,
            columns=("word", "reading"),
            show="headings",
            selectmode="browse",
            yscrollcommand=scrollbar.set,
        )
        tree.heading("word", text="語句")
        tree.heading("reading", text="読み（五十音順）")
        tree.column("word", width=220, anchor="w")
        tree.column("reading", width=220, anchor="w")
        tree.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        scrollbar.config(command=tree.yview)
        tree.tag_configure("oddrow", background="#f2f2f2")
        tree.tag_configure("evenrow", background="white")

        # --- 右：操作パネル ---
        panel = tk.Frame(edit_win)
        panel.grid(row=0, column=1, sticky="ns", padx=(5, 10), pady=10)

        count_label = tk.Label(panel, text="", anchor="w", fg="gray30")
        count_label.pack(fill="x", pady=(0, 8))

        def refresh_list(select_word=None):
            tree.delete(*tree.get_children())
            # 読みの五十音順に並べる（漢字のコード順よりも探しやすい）
            for i, (word, reading) in enumerate(
                sorted(working.items(), key=lambda kv: (kv[1], kv[0]))
            ):
                tag = "evenrow" if i % 2 == 0 else "oddrow"
                item = tree.insert("", "end", values=(word, reading), tags=(tag,))
                if word == select_word:
                    tree.selection_set(item)
                    tree.focus(item)
                    tree.see(item)
            count_label.config(text=f"登録数: {len(working)} 件")

        # 新規追加フォーム
        add_frame = tk.LabelFrame(panel, text="新規追加", padx=8, pady=6)
        add_frame.pack(fill="x")
        tk.Label(add_frame, text="語句", anchor="w").pack(fill="x")
        new_word_entry = tk.Entry(add_frame, width=18)
        new_word_entry.pack(fill="x")
        tk.Label(add_frame, text="読み", anchor="w").pack(fill="x", pady=(4, 0))
        new_reading_entry = tk.Entry(add_frame, width=18)
        new_reading_entry.pack(fill="x")

        def add_entry(event=None):
            word = new_word_entry.get().strip()
            reading = new_reading_entry.get().strip()
            if not word or not reading:
                msgbox.showwarning(
                    "辞書編集", "語句と読みの両方を入力してください。", parent=edit_win
                )
                return
            if word in working:
                if not msgbox.askyesno(
                    "上書き確認",
                    f"「{word}」は既に登録されています（{working[word]}）。\n"
                    f"「{reading}」で上書きしますか？",
                    parent=edit_win,
                ):
                    return
            working[word] = reading
            dirty["value"] = True
            new_word_entry.delete(0, "end")
            new_reading_entry.delete(0, "end")
            new_word_entry.focus_set()
            refresh_list(select_word=word)

        tk.Button(add_frame, text="追加", command=add_entry).pack(fill="x", pady=(8, 0))
        new_reading_entry.bind("<Return>", add_entry)

        def edit_selected(event=None):
            item_id = tree.focus()
            if not item_id:
                msgbox.showinfo(
                    "辞書編集", "編集する語句を一覧から選んでください。", parent=edit_win
                )
                return
            old_word, old_reading = tree.item(item_id, "values")

            popup = tk.Toplevel(edit_win)
            popup.title("編集")
            popup.transient(edit_win)
            tk.Label(popup, text="語句").pack()
            word_entry = tk.Entry(popup, width=24)
            word_entry.insert(0, old_word)
            word_entry.pack(padx=12)
            tk.Label(popup, text="読み").pack()
            reading_entry = tk.Entry(popup, width=24)
            reading_entry.insert(0, old_reading)
            reading_entry.pack(padx=12)

            def save():
                new_word = word_entry.get().strip()
                new_reading = reading_entry.get().strip()
                if not new_word or not new_reading:
                    msgbox.showwarning(
                        "辞書編集", "語句と読みの両方を入力してください。", parent=popup
                    )
                    return
                if new_word != old_word and new_word in working:
                    if not msgbox.askyesno(
                        "上書き確認",
                        f"「{new_word}」は既に登録されています（{working[new_word]}）。\n"
                        f"「{new_reading}」で上書きしますか？",
                        parent=popup,
                    ):
                        return
                working.pop(old_word, None)
                working[new_word] = new_reading
                dirty["value"] = True
                popup.destroy()
                refresh_list(select_word=new_word)

            btn_frame = tk.Frame(popup)
            btn_frame.pack(pady=8)
            tk.Button(btn_frame, text="保存", command=save).pack(side="left", padx=5)
            tk.Button(btn_frame, text="キャンセル", command=popup.destroy).pack(side="left", padx=5)
            self.center_window_auto(popup)
            popup.grab_set()

        def delete_selected():
            item_id = tree.focus()
            if not item_id:
                msgbox.showinfo(
                    "辞書編集", "削除する語句を一覧から選んでください。", parent=edit_win
                )
                return
            word, _reading = tree.item(item_id, "values")
            if not msgbox.askyesno(
                "削除確認", f"「{word}」を辞書から削除しますか？", parent=edit_win
            ):
                return
            working.pop(word, None)
            dirty["value"] = True
            refresh_list()

        tk.Button(panel, text="選択した語句を編集", command=edit_selected).pack(fill="x", pady=(14, 4))
        tk.Button(panel, text="選択した語句を削除", command=delete_selected, fg="red").pack(fill="x")

        # 下のボタンを底へ寄せるための伸縮スペーサ
        tk.Frame(panel).pack(fill="both", expand=True)

        def save_and_close():
            try:
                with open(OVERRIDE_PATH, "w", encoding="utf-8") as f:
                    json.dump(working, f, indent=2, ensure_ascii=False)
            except OSError as e:
                logging.error(f"{OVERRIDE_PATH} の保存に失敗しました: {e}")
                msgbox.showerror(
                    "辞書保存",
                    f"override.json を保存できませんでした。\n{OVERRIDE_PATH}\n{e}",
                    parent=edit_win,
                )
                return
            self.override_dict = working
            edit_win.destroy()
            msgbox.showinfo(
                "辞書編集",
                f"辞書を保存しました（{len(working)}件）。\n"
                "開いているファイルの語句一覧へ反映するには「辞書再適用」を押してください。"
            )

        def cancel():
            if dirty["value"]:
                if not msgbox.askyesno(
                    "辞書編集",
                    "保存していない変更があります。破棄して閉じますか？",
                    parent=edit_win,
                ):
                    return
            edit_win.destroy()

        tk.Button(panel, text="保存して閉じる", command=save_and_close).pack(fill="x", pady=(0, 4))
        tk.Button(panel, text="キャンセル", command=cancel).pack(fill="x")

        tree.bind("<Double-1>", edit_selected)
        edit_win.protocol("WM_DELETE_WINDOW", cancel)
        refresh_list()

    # ---------- Wordマクロ ----------
    def run_vba_macro(self, word_file_path, word_app=None, settings=None):
        """Wordマクロを実行してルビ付きファイルを作る。

        word_app を渡すと既存のWordインスタンスを使い回す（一括処理用、高速化のため）。
        渡さない場合は関数内でWordを起動・終了する。

        ★v2.1では、Python側で文書を開かずマクロに3つのパス（元ファイル・TSV・
        出力先）を渡すだけにした。v2.0 は Python が文書を開いてから
        ActiveDocument 頼みでマクロを走らせており、さらにマクロ側は
        WScript.Shell からデスクトップのパスを推測していたため、
        OneDriveのリダイレクト環境でPython側と食い違うことがあった。

        戻り値は {"out_path": Path, "applied": int|None, "skipped": int|None}。
        失敗時は RuntimeError を送出する（呼び出し側でまとめて扱う）。"""
        if settings is None:
            settings = self.settings

        src = Path(word_file_path).resolve()
        tsv_path = (self.ruby_dir / f"{src.stem}.tsv").resolve()
        out_path = (self.output_dir / f"{src.stem}（ルビ）.docx").resolve()
        result_path = Path(f"{tsv_path}.result")

        logging.debug(f"run_vba_macro: src={src} tsv={tsv_path} out={out_path}")

        if not src.exists():
            raise RuntimeError(f"対象のWordファイルが存在しません。\n{src}")
        if not tsv_path.exists():
            raise RuntimeError(
                f"TSVファイルが見つかりません。\n{tsv_path}\n"
                "先に「TSV保存」を実行してください。"
            )

        # 前回の残骸を消しておく（読み違えて古い件数を表示しないため）
        try:
            result_path.unlink()
        except OSError:
            pass

        word = word_app
        owns_word_instance = False
        macro_error = None
        try:
            if word is None:
                word, owns_word_instance = get_word()
            word.Run(MACRO_NAME, str(src), str(tsv_path), str(out_path))
        except Exception as e:
            macro_error = e
        finally:
            # 自分でWordを起動した場合のみ終了する（使い回し時は呼び出し元が管理。
            # 元々起動していたWordは、ユーザーの編集中の文書ごと落とさないよう触らない）
            if owns_word_instance and word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass

        # ★成否にかかわらず結果ファイルを読んで消す。
        # 失敗時に読まずに抜けると、ルビデータフォルダに .result が残り続ける。
        info = self._read_macro_result(result_path)

        if macro_error is not None:
            logging.error(f"Wordマクロの実行に失敗しました: {macro_error}")
            raise RuntimeError(
                self._macro_error_message(macro_error)
            ) from macro_error

        # ★マクロがエラーで終わったことを結果ファイルに書き残しているのに、
        # 例外が Python まで届かなかった場合の保険。ここを見ないと
        # 途中で失敗しているのに「保存しました」と出てしまう。
        if info.get("error"):
            raise RuntimeError(
                "Wordマクロがエラーで終了しました。\n"
                "rubigui.log と、Word側にエラーが表示されていないか確認してください。"
            )

        if not out_path.exists():
            raise RuntimeError(
                "マクロは実行されましたが、出力ファイルを確認できませんでした。\n"
                f"{out_path}"
            )
        info["out_path"] = out_path
        logging.info(f"ルビ付きWord出力完了: {out_path}（applied={info.get('applied')}）")
        return info

    @staticmethod
    def _read_macro_result(result_path):
        """マクロが書き出した .result（applied / skipped / error）を読んで消す。
        読めなくても処理自体は成功していることがあるので、件数を None にして続行する。"""
        info = {"applied": None, "skipped": None, "error": False}
        try:
            with open(result_path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    key, _, value = line.strip().partition("\t")
                    if key in ("applied", "skipped"):
                        info[key] = int(value)
                    elif key == "error":
                        info["error"] = value == "1"
        except (OSError, ValueError) as e:
            logging.debug(f"マクロの結果ファイルを読めませんでした（無視して続行）: {e}")
        finally:
            try:
                result_path.unlink()
            except OSError:
                pass
        return info

    @staticmethod
    def _macro_error_message(exc):
        """word.Run の失敗を、利用者が対処できる文言に直す。"""
        detail = str(exc)
        source = ""
        # pywin32 の com_error.excepinfo は
        # (wCode, ソース名, 説明, ヘルプファイル, ヘルプID, scode)。
        # VBA の Err.Raise の第2引数（"RubiGUI"）はソース名の方に入る。
        excepinfo = getattr(exc, "excepinfo", None)
        if excepinfo:
            try:
                source = str(excepinfo[1] or "")
                if excepinfo[2]:
                    detail = str(excepinfo[2])
            except (IndexError, TypeError):
                pass

        # マクロ自身が投げた業務エラー（TSVが無い等）はそのまま見せる。
        # ★ソース名で判定する。説明文には "RubiGUI" が入らないため、
        # 説明文で判定していると「マクロを入れ直してください」という
        # 見当違いの案内に化けてしまう。
        if source == "RubiGUI":
            return detail

        # ★Wordがビジーなど、マクロの有無と関係のないCOMエラーまで
        # 「マクロを入れ直してください」と案内しないようにする。
        # v2.1 は起動中のWordにアタッチするようになったため、ユーザーが
        # ダイアログを開いたままだとこの経路に入りやすい。
        hresult = None
        if getattr(exc, "args", None):
            first = exc.args[0]
            if isinstance(first, int):
                hresult = first
        if hresult in BUSY_HRESULTS:
            return (
                "Wordが応答できる状態ではないため、処理を実行できませんでした。\n\n"
                "Word側でダイアログが開いていないか、入力途中になっていないかを\n"
                "確認して閉じてから、もう一度お試しください。\n\n"
                f"詳細: {detail}"
            )
        if hresult in DISCONNECTED_HRESULTS:
            return (
                "Wordとの接続が切れたため、処理を中断しました。\n\n"
                "処理中にWordが閉じられたか、Word自体が終了した可能性があります。\n"
                "処理が終わるまでWordを閉じずに、もう一度お試しください。\n\n"
                "繰り返し発生する場合は、Module1.bas が最新版か確認してください。\n"
                "（古い版では、この文書でWordが異常終了することがあります）\n\n"
                f"詳細: {detail}"
            )

        return (
            f"Wordマクロ「{MACRO_NAME}」を実行できませんでした。\n\n"
            "v2.1 でマクロの名前と引数が変わっています。\n"
            "Wordを開き、Alt+F11 でVBAエディタを表示して\n"
            "  1) Normal の古い Module1 を右クリック →「Module1 の解放」で削除\n"
            "  2) 右クリック →「ファイルのインポート」で下記を読み込み\n"
            f"     {APP_DIR / 'Module1.bas'}\n"
            "を行ってから、もう一度お試しください。\n"
            "（すでに導入済みの場合は、Word側でダイアログが開いたままに\n"
            "　なっていないかも確認してください）\n\n"
            f"詳細: {detail}"
        )


if __name__ == "__main__":
    # ★コンソールを隠したので、想定外のエラーが起きても画面に何も出ないまま
    # 終了してしまう。最後の受け皿としてログに残し、ダイアログで知らせる。
    try:
        root = TkinterDnD.Tk()
        app = RubyEditorApp(root)
        root.mainloop()
    except Exception as e:
        logging.exception("予期しないエラーで終了しました")
        try:
            msgbox.showerror(
                "エラー",
                f"予期しないエラーが発生したため終了します。\n\n{e}\n\n"
                f"詳細は次のログを確認してください:\n{LOG_PATH}"
            )
        except Exception:
            pass
        raise
