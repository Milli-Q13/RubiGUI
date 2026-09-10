from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox as msgbox
import json
import os
import sys
import logging
from logging.handlers import RotatingFileHandler
import zipfile
import queue
import threading
import xml.etree.ElementTree as ET
import jaconv
from sudachipy import tokenizer, dictionary
import win32com.client
# ★v3.0：重い処理を別スレッドで走らせるために必要。
# COMは「スレッドごとに初期化する」決まりなので、ワーカースレッドの中で
# CoInitialize() を呼ばないと Word を掴めない。
import pythoncom
from pathlib import Path

# ✅ バージョン情報（サポート対応時にユーザーへ確認してもらうため、GUI上にも表示する）
APP_VERSION = "3.1"


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

# Word側マクロ（RubiGUI_V31.bas）のマクロ名。
# ★重要：版が変わるたびに改名している
# （v2.0 "InsertFuriganaFromTSV_SaveToNewFile_Stable" → v2.1 "..._V21" → v3.0 "..._V30"
#   → v3.1 "..._V31"）。
# 旧マクロに新しい呼び出しを渡すと「何も起きない（ルビが振られないだけ）」という
# 分かりにくい失敗になる。名前を変えておけば word.Run が確実に失敗するので、
# 「RubiGUI_V31.bas を入れ直してください」と明示的に案内できる。
# ★v3.0 ではモジュール名も Module1 から RubiGUI_V30 へ変更した。VBEの
# モジュール一覧を見れば、どの版が入っているか一目で分かるようにするため。
# ★v3.1 でモジュール名を RubiGUI_V31 に改めた（版ごとに追随させる方針）。
MACRO_NAME = "InsertFuriganaFromTSV_V31"

# マクロが書き出す進捗ファイルの拡張子（TSVと同じ場所に作られる）。
# word.Run は1回の呼び出しの中で完結して戻ってこないため、
# 途中経過はこのファイル経由でしか受け取れない。
PROGRESS_SUFFIX = ".progress"

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
    """ルビ設定を ruby_settings.json に保存する。

    ★重要：Word版とPPT版は同じフォルダに同居し、同じ ruby_settings.json を
    共有する。両版が持つキーは一致しない（PPT版だけが line_spacing と
    include_title を持つ）ので、自分の設定だけを書き出すと相手版のキーが
    消える。読み込み側は知らないキーを無視するため気づきにくく、
    「設定したはずなのに戻っている」という分かりにくい不具合になる。
    そのため、既存の内容を読んでから自分のキーだけを更新して書き戻す。
    """
    merged = {}
    try:
        if SETTINGS_PATH.exists():
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                merged.update(loaded)
    except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
        # 読み直せない場合は「相手版のキーは救えないが、自分の設定は保存する」方に倒す。
        # ここで諦めると、ファイルが一度壊れたきり設定を保存できなくなる。
        logging.warning(f"{SETTINGS_PATH.name} を読み直せませんでした（上書きします）: {e}")

    merged.update(settings)

    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(merged, f, indent=2, ensure_ascii=False)
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


# ============================================================
# 進捗表示（v3.0で追加）
# ============================================================
# マクロが書き出す phase の値と、画面に出す日本語の対応。
PHASE_LABELS = {
    "search": "語句を検索中",
    "apply": "ルビを付与中",
}


def read_progress(tsv_path):
    """マクロが書き出した .progress を読む。読めなければ None を返す。

    ★重要：読み取り失敗は「異常」ではなく日常。マクロが書き込んでいる
    最中はWordがこのファイルを排他ロックするため PermissionError になるし、
    処理の開始直後はまだファイル自体が無い。呼び出し側は None を
    「今回は読めなかっただけ」として扱い、次の周期で読み直すこと。"""
    try:
        with open(f"{tsv_path}{PROGRESS_SUFFIX}", "r", encoding="ascii", errors="replace") as f:
            info = {}
            for line in f:
                key, _, value = line.strip().partition("\t")
                info[key] = value
        return info.get("phase", ""), int(info["current"]), int(info["total"])
    except (OSError, KeyError, ValueError):
        return None


def describe_progress(tsv_path):
    """進捗ファイルを (表示文, 現在値, 全体数) に整える。読めなければ None。"""
    info = read_progress(tsv_path)
    if info is None:
        return None
    phase, current, total = info
    label = PHASE_LABELS.get(phase, "処理中")
    if total <= 0:
        return label, 0, 0
    return f"{label}　{current}/{total}", current, total


class ProgressDialog:
    """処理中であることを示す小さなウィンドウ。

    ★重要：これを出すだけでは意味が無い。重い処理は必ず別スレッドで
    走らせること。メインスレッドで走らせると、このウィンドウ自体が
    描画されないまま固まり、Windowsに「応答なし」と表示される。
    それでは「実行中か不具合か分からない」という元の問題が解決しない。

    ★閉じるボタンは無効化してある。閉じても処理は止まらないので、
    「閉じたのに裏で動いている」状態を作らないため。"""

    def __init__(self, root, title, message):
        self.win = tk.Toplevel(root)
        self.win.title(title)
        self.win.resizable(False, False)
        self.win.protocol("WM_DELETE_WINDOW", lambda: None)

        frame = tk.Frame(self.win, padx=20, pady=16)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=message, anchor="w", justify="left").pack(fill="x")
        self.detail = tk.Label(frame, text="準備しています…", anchor="w",
                               justify="left", fg="gray30")
        self.detail.pack(fill="x", pady=(6, 8))

        # 件数が分かるまでは不定形（流れるバー）で「生きている」ことを示す
        self.bar = ttk.Progressbar(frame, mode="indeterminate", length=340)
        self.bar.pack(fill="x")
        self.bar.start(12)
        self._determinate = False

        self.win.transient(root)
        # 処理中に他のボタンを押させない（二重実行の防止も兼ねる）
        self.win.grab_set()
        self.win.update_idletasks()
        x = root.winfo_rootx() + (root.winfo_width() // 2) - (self.win.winfo_width() // 2)
        y = root.winfo_rooty() + (root.winfo_height() // 2) - (self.win.winfo_height() // 2)
        self.win.geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def set_detail(self, text):
        try:
            self.detail.config(text=text)
        except tk.TclError:
            pass

    def set_progress(self, text, current, total):
        """件数が分かったら不定形バーから実数バーへ切り替える。"""
        self.set_detail(text)
        try:
            if total > 0:
                if not self._determinate:
                    self.bar.stop()
                    self.bar.config(mode="determinate", maximum=total)
                    self._determinate = True
                self.bar.config(maximum=total, value=current)
        except tk.TclError:
            pass

    def close(self):
        try:
            self.bar.stop()
            self.win.grab_release()
            self.win.destroy()
        except tk.TclError:
            pass


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
        # ★重要：空白だけの表層形は、辞書の判定より先に捨てる。
        # 字下げやレイアウトに全角スペースを使った文書では、Sudachiが
        # 空白の連なり（例："\n　　　　"）を1つのトークンとして返す。
        # これは下の除外条件をすべてすり抜ける：2文字以上あるので
        # 文字数チェックに引っかからず、改行も全角スペースも
        # ひらがな／カタカナ／数字／英字のいずれの範囲にも入らない。
        # 最後の砦である surface == reading も、Sudachiが読みを返すときに
        # 全角スペース(U+3000)を半角スペース(U+0020)へ正規化するため
        # 文字列として一致せず、素通りしてしまう。
        # 結果、語句一覧に「空行」が並び、TSVにも改行入りのレコードとして
        # 書き出されて行数が狂っていた（v2.1で確認）。
        # str.strip() は全角スペース・NBSP・タブ・改行をまとめて落とす。
        if not surface.strip():
            continue
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
        # 別スレッドで処理中かどうか（v3.0で追加）。詳細は _reject_if_busy を参照。
        self._busy = False
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
        if self._reject_if_busy():
            return
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
        if self._reject_if_busy():
            return
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

    # ---------- バックグラウンド実行（v3.0で追加） ----------
    def _reject_if_busy(self):
        """処理中なら操作を断る。断ったときは True を返す。

        ★重要：ProgressDialog の grab_set() だけでは足りない。
        Tkのグラブは「Tkが配送するイベント」にしか効かないが、
        tkinterdnd2 のドラッグ＆ドロップはOS側のドロップターゲット経由で
        ウィジェットへ直接届くため素通りする。処理中に別のファイルを
        投げ込まれると2つ目のワーカースレッドが起動し、同じWordへ
        別スレッドから同時にマクロ実行を要求することになる。
        Wordのマクロ実行は同時呼び出しを想定していないので、
        異常終了や出力の破損につながり得る。
        v2.1 はメインスレッドを塞いでいたので起こり得なかったが、
        v3.0 でバックグラウンド化した副作用として生じた穴なので、
        グラブに頼らず明示的なフラグで塞ぐ。"""
        if getattr(self, "_busy", False):
            msgbox.showinfo(
                "処理中",
                "現在ほかの処理を実行中です。\n終わるまでお待ちください。"
            )
            return True
        return False

    def _run_in_background(self, work, on_success, on_error,
                           title, message, progress_source=None):
        """work() を別スレッドで実行し、その間 進捗ウィンドウを出す。

        ★なぜスレッドが要るのか
          Wordのマクロ呼び出し（word.Run）は、1回の呼び出しの中で
          最後まで処理してから戻ってくる。これをメインスレッドで
          呼ぶと tkinter が画面を描き直せなくなり、ウィンドウが白く
          固まって「応答なし」になる。実行中なのか壊れたのか
          利用者に区別が付かない、というのが v2.1 までの問題だった。

        ★COMの制約
          COMのオブジェクトはスレッドをまたいで使えない。そのため
          ワーカースレッドの中で CoInitialize() を呼び、Wordの取得も
          そのスレッドの中で行う（work の中で get_word() を呼ぶ）。
          メインスレッドで掴んだWordをワーカーへ渡してはいけない。

        ★UIはメインスレッドから
          tkinter はスレッドセーフではないので、ワーカーからは
          画面を触らない。結果はキューに入れ、メインスレッド側の
          poll() が取り出して on_success / on_error を呼ぶ。

        progress_source は「今の進捗を (表示文, 現在値, 全体数) で返す関数」。
        None なら不定形バーだけを回す。"""
        self._busy = True
        dialog = ProgressDialog(self.root, title, message)
        result_queue = queue.Queue()

        def runner():
            pythoncom.CoInitialize()
            try:
                result_queue.put(("ok", work()))
            except Exception as e:
                result_queue.put(("error", e))
            finally:
                pythoncom.CoUninitialize()

        threading.Thread(target=runner, daemon=True).start()

        def poll():
            try:
                status, payload = result_queue.get_nowait()
            except queue.Empty:
                if progress_source is not None:
                    try:
                        info = progress_source()
                    except Exception as e:
                        # 進捗表示のためだけに処理を止めるのは本末転倒
                        logging.debug(f"進捗の取得に失敗しました（無視して続行）: {e}")
                        info = None
                    if info is not None:
                        dialog.set_progress(*info)
                self.root.after(200, poll)
                return

            # ★フラグの解除はコールバックより前。on_success の中から
            #   次の処理を始める経路（advance_to_next）があるため、
            #   後に置くと自分自身を「処理中」と誤判定して止まる。
            self._busy = False
            dialog.close()
            if status == "ok":
                on_success(payload)
            else:
                on_error(payload)

        self.root.after(200, poll)

    def _process_batch_files(self, docx_files):
        """複数ファイルの一括処理本体（ドラッグ＆ドロップ／フォルダ選択どちらからも使う共通処理）。
        Wordは1つだけ起動して使い回し、ファイルごとの起動・終了コストを省く。

        ★v3.0：処理本体を別スレッドへ移した。v2.1 はメインスレッドで
        回していたため、ファイルとファイルの合間にしか画面が更新されず、
        1ファイルの処理中（＝時間のほとんど）は固まって見えていた。"""
        settings = self.read_settings_from_ui()
        total = len(docx_files)
        docx_files = [Path(f) for f in docx_files]

        # ワーカーとメインスレッドの間で受け渡す進捗。単純な代入と読み出し
        # だけなのでロックは要らない（tkinter側は200msごとに読むだけ）。
        self._batch_state = {"index": 0, "name": "", "tsv_path": None}

        def work():
            success_count = 0
            total_ruby = 0
            failures = []
            word, own_word = get_word()
            try:
                for i, file_path in enumerate(docx_files, start=1):
                    self._batch_state = {
                        "index": i, "name": file_path.name, "tsv_path": None,
                    }
                    try:
                        logging.info(f"処理中: {file_path}")
                        terms = extract_terms(file_path, self.override_dict)
                        logging.debug(f"{file_path.name} 語句抽出数: {len(terms)}")
                        if not terms:
                            logging.warning(f"{file_path} → 語句ゼロ（抽出失敗の可能性）")
                        term_pairs = [(t["word"], t["reading"]) for t in terms]
                        tsv_path, _ = self.save_tsv(term_pairs, file_path, settings)
                        self._batch_state = {
                            "index": i, "name": file_path.name, "tsv_path": tsv_path,
                        }
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
            return success_count, total_ruby, failures

        def progress_source():
            state = self._batch_state
            index, name = state["index"], state["name"]
            if index <= 0:
                return None
            head = f"{index}/{total} 件目　{name}"
            # ファイル内の細かい進捗は、マクロが書く .progress から読む
            tsv_path = state["tsv_path"]
            if tsv_path is not None:
                detail = describe_progress(tsv_path)
                if detail is not None:
                    text, current, sub_total = detail
                    return f"{head}\n{text}", index, total
            return head, index, total

        def finish():
            self._update_batch_progress("")
            # ★重要：一括処理中は語句一覧の表示を更新していないので、
            # 処理後の self.data / self.file_path をそのまま残すと
            # 「画面はAのまま、中身は最後に処理したファイルの語句」という
            # ちぐはぐな状態になる。この状態で「ルビ付きWord出力」や
            # 「TSV保存」を押すと、別ファイルの読みでルビが振られてしまう。
            self._clear_current_file()

        def on_success(payload):
            success_count, total_ruby, failures = payload
            finish()
            summary = f"{total} 件中 {success_count} 件のファイルを正常に処理しました。"
            summary += f"\nルビ {total_ruby} 件を付与しました。"
            if failures:
                summary += "\n\n失敗したファイル:\n" + "\n".join(failures)
            msgbox.showinfo("一括処理", summary)

        def on_error(exc):
            finish()
            logging.error(f"一括処理に失敗しました: {exc}")
            msgbox.showerror("一括処理", f"処理を開始できませんでした。\n{exc}")

        self._run_in_background(
            work, on_success, on_error,
            title="一括処理",
            message=f"{total} 件のファイルにルビを付与しています…",
            progress_source=progress_source,
        )

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
        if self._reject_if_busy():
            return
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
        if self._reject_if_busy():
            return
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
        if self._reject_if_busy():
            return
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
        if self._reject_if_busy():
            return
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
        含むものなど）」をPython側で除外するしかなかった。VBA側を
        ADODB.Stream でのUTF-8読み込みに変えたので、その制限は無くなった。

        ★語句は「表層形の長い順」に並べる。VBAは上から順にルビを振るので、
        長い語（例：日本人）を先に処理しておけば、短い語（例：日本）が
        ルビ化済みの語の中へ食い込むことがない。

        ★v3.0：タブ・改行を含む語句は書き出さない。TSVにはエスケープの
        概念が無いため、語句に制御文字が混ざるとレコードが黙って壊れる。
        v2.1 では全角スペースだけの「語句」（Sudachiが空白の連なりを
        1トークンとして返したもの）に改行が含まれており、TSVの行数が
        語句数と合わなくなっていた。抽出側でも弾いているが、辞書編集など
        別経路から入り込んでも壊れないようにここでも守る。"""
        if settings is None:
            settings = self.settings

        base_name = Path(file_path).stem
        tsv_path = self.ruby_dir / f"{base_name}.tsv"

        safe = []
        for word, reading in terms:
            if not word.strip() or not reading.strip():
                logging.warning(f"空白だけの語句をTSVから除外しました: {word!r} -> {reading!r}")
                continue
            if any(ch in word or ch in reading for ch in ("\t", "\n", "\r")):
                logging.warning(f"制御文字を含む語句をTSVから除外しました: {word!r} -> {reading!r}")
                continue
            safe.append((word, reading))

        ordered = sorted(safe, key=lambda t: len(t[0]), reverse=True)
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
        # ★件数も返す。除外された語句があると self.data の件数と食い違うため、
        #   「〇件保存しました」の表示には必ずこちらを使う。
        return tsv_path, len(ordered)

    def save_only_tsv(self):
        if self._reject_if_busy():
            return
        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return
        settings = self.read_settings_from_ui()
        try:
            _, saved_count = self.save_tsv(self.data, self.file_path, settings)
        except OSError as e:
            logging.error(f"TSVの保存に失敗しました: {e}")
            msgbox.showerror("TSV保存", f"TSVファイルを保存できませんでした。\n{e}")
            return
        msgbox.showinfo("TSV保存", f"{saved_count}件の語句をTSV形式で保存しました。")

    def run_single_file_output(self):
        if self._reject_if_busy():
            return
        if not self.current_file_path:
            msgbox.showwarning("エラー", "先にWordファイルをドロップまたは選択してください。")
            return
        if not self.data:
            msgbox.showwarning("エラー", "ルビを振る語句がありません。")
            return
        settings = self.read_settings_from_ui()
        src = Path(self.current_file_path)

        # TSVの書き出しは一瞬なので、スレッドへ移さずここで済ませる。
        # 先に書いておけば、進捗ファイルの置き場所（＝TSVと同じ場所）も確定する。
        try:
            tsv_path, _ = self.save_tsv(self.data, self.current_file_path, settings)
        except OSError as e:
            logging.error(f"TSVの保存に失敗しました: {e}")
            msgbox.showerror("TSV保存", f"TSVファイルを保存できませんでした。\n{e}")
            return

        def work():
            return self.run_vba_macro(self.current_file_path, settings=settings)

        def on_success(result):
            message = f"{Path(result['out_path']).name} を保存しました。"
            if result.get("applied") is not None:
                message += f"\nルビ {result['applied']} 件を付与しました。"
                if result["applied"] == 0:
                    message += "\n\n※ルビが1件も付いていません。語句一覧を確認してください。"
            # ★行間を緩めた段落があるとページ送りが変わることがあるので必ず知らせる。
            #   黙って直すと「なぜかレイアウトが変わった」という不信につながる。
            if result.get("loosened"):
                message += (
                    f"\n\n※行間が「固定値」の段落 {result['loosened']} 件を「最小値」に"
                    "変更しました。\n"
                    "　固定値のままだとルビが行に収まらず隠れてしまうためです。\n"
                    "　行が高くなる分、ページ送りが変わることがあります。"
                )
            msgbox.showinfo("ルビ付きWord出力", message)

        def on_error(exc):
            if not isinstance(exc, RuntimeError):
                logging.error(f"ルビ付きWord出力に失敗しました: {exc}")
            msgbox.showerror("ルビ付きWord出力", str(exc))

        self._run_in_background(
            work, on_success, on_error,
            title="ルビ付きWord出力",
            message=f"{src.name} にルビを付与しています…",
            progress_source=lambda: describe_progress(tsv_path),
        )

    def advance_to_next(self):
        if self._reject_if_busy():
            return
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
            tsv_path, _ = self.save_tsv(self.data, current_file, settings)
        except OSError as e:
            logging.error(f"{current_file} のTSV保存に失敗しました: {e}")
            msgbox.showerror("保存失敗", f"{current_file.name} の処理に失敗しました。\n{e}")
            return

        def work():
            return self.run_vba_macro(current_file, settings=settings)

        def advance():
            """成否によらず次のファイルへ進む（v2.1と同じ挙動）。"""
            self.review_ready = False
            self.current_index += 1
            self.process_next_file()

        def on_success(_result):
            advance()

        def on_error(exc):
            logging.error(f"{current_file} の処理に失敗しました: {exc}")
            msgbox.showerror("保存失敗", f"{current_file.name} の処理に失敗しました。\n{exc}")
            advance()

        self._run_in_background(
            work, on_success, on_error,
            title="ルビ付きWord出力",
            message=f"{current_file.name} にルビを付与しています…",
            progress_source=lambda: describe_progress(tsv_path),
        )

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
        # ★進捗ファイルも同様。消し忘れると、処理を始めた直後に
        #   前回の「81/81」が一瞬表示されてしまう。
        try:
            Path(f"{tsv_path}{PROGRESS_SUFFIX}").unlink()
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
        info = {"applied": None, "skipped": None, "loosened": None, "error": False}
        try:
            with open(result_path, "r", encoding="utf-8", errors="replace") as f:
                for line in f:
                    key, _, value = line.strip().partition("\t")
                    if key in ("applied", "skipped", "loosened"):
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
                "繰り返し発生する場合は、RubiGUI_V31.bas が最新版か確認してください。\n"
                "（古い版では、この文書でWordが異常終了することがあります）\n\n"
                f"詳細: {detail}"
            )

        return (
            f"Wordマクロ「{MACRO_NAME}」を実行できませんでした。\n\n"
            "v3.0 でマクロの名前とモジュール名が変わっています。\n"
            "Wordを開き、Alt+F11 でVBAエディタを表示して\n"
            "  1) Normal に古い Module1（v2.1以前）が残っていれば、\n"
            "     右クリック →「Module1 の解放」で削除\n"
            "  2) Normal を右クリック →「ファイルのインポート」で下記を読み込み\n"
            f"     {APP_DIR / 'RubiGUI_V31.bas'}\n"
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

        # ★v3.0：処理中に「×」で閉じられないようにする。
        # v2.1 は処理がメインスレッドを塞いでいたため、そもそも閉じる操作が
        # できなかった。v3.0 は別スレッドで処理を回すのでメインループが
        # 生きており、閉じられてしまう。ワーカーは daemon なので
        # プロセスと一緒に強制終了され、Wordの終了処理（Quit）が走らない。
        # その結果、非表示のWINWORD.exeが residual として残ったり、
        # 開いたままの文書のロックファイル（~$〇〇.docx）が残って
        # 元ファイルを開けなくなったりする。
        def on_close():
            if getattr(app, "_busy", False):
                msgbox.showinfo(
                    "処理中",
                    "処理が終わるまで閉じないでください。\n"
                    "途中で閉じるとWordが終了しないまま残ることがあります。"
                )
                return
            root.destroy()

        root.protocol("WM_DELETE_WINDOW", on_close)
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
