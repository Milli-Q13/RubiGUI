from tkinterdnd2 import DND_FILES, TkinterDnD
import tkinter as tk
from tkinter import ttk, filedialog, messagebox as msgbox
import json
import os
import re
import sys
import shutil
import logging
from logging.handlers import RotatingFileHandler
import zipfile
import xml.etree.ElementTree as ET
import jaconv
from sudachipy import tokenizer, dictionary
import win32com.client
from pathlib import Path
import tkinter.filedialog as filedialog

# ✅ バージョン情報（サポート対応時にユーザーへ確認してもらうため、GUI上にも表示する）
APP_VERSION = "1.0"


# ✅ ログ設定（実行フォルダに rubigui_ppt.log として出力。
#    ユーザーから不具合報告を受けた際、このファイルを送ってもらえば調査しやすくなる）
#    Word版と同様、RotatingFileHandlerで1ファイル5MB・3世代までに制限しておく。
_log_handler = RotatingFileHandler(
    "rubigui_ppt.log",
    maxBytes=5 * 1024 * 1024,  # 5MBごとにローテーション
    backupCount=3,             # 直近3世代（＋現行分で最大4ファイル）まで保持
    encoding="utf-8",
)
_log_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

logging.basicConfig(level=logging.DEBUG, handlers=[_log_handler])
logging.getLogger().addHandler(logging.StreamHandler())  # コンソールにも出す

# ✅ Sudachi初期化（Word版と同じ設定・同じ辞書をそのまま使う）
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


# ============================================================
# PowerPoint COM の定数（win32com.client.constants は遅延バインドで
# 使えないことがあるため、必要なものだけ明示的に定義しておく）
# ============================================================
MSO_TRUE = -1
MSO_FALSE = 0

MSO_GROUP = 6
MSO_PLACEHOLDER = 14
MSO_TEXT_BOX = 17

# PpPlaceholderType：ルビ対象外にするタイトル系
PP_TITLE_PLACEHOLDERS = (1, 3, 4, 5)  # Title / CenterTitle / Subtitle / VerticalTitle
# 常に対象外にするプレースホルダ。フッターや社名を全スライドに入れている資料で、
# 同じルビが全ページに量産されてしまうのを防ぐ。
PP_ALWAYS_SKIP_PLACEHOLDERS = (13, 14, 15, 16)  # SlideNumber / Header / Footer / Date
XML_TITLE_PH_TYPES = ("title", "ctrTitle", "subTitle")
XML_ALWAYS_SKIP_PH_TYPES = ("sldNum", "hdr", "ftr", "dt")

# MsoTextOrientation
ORIENT_HORIZONTAL = 1
ORIENT_DOWNWARD = 3
ORIENT_VERTICAL_FAREAST = 4
# ★重要：縦書きとして扱えるのは「行が右→左に進む」向きだけ。
#    実測の結果、3(Downward) と 4(VerticalFarEast) が右→左、
#    2(Upward) / 5(Vertical) / 6(HorizontalRotatedFarEast) は左→右だった。
#    左→右の縦組みは日本語のルビの作法が定義できないため対象外にする。
VERTICAL_ORIENTATIONS = (ORIENT_DOWNWARD, ORIENT_VERTICAL_FAREAST)

# MsoAutoSize
MSO_AUTOSIZE_NONE = 0
MSO_AUTOSIZE_TEXT_TO_FIT_SHAPE = 2

PP_ALIGN_CENTER = 2
PP_AUTOSIZE_SHAPE_TO_FIT_TEXT = 1

MSO_ANIM_TRIGGER_WITH_PREVIOUS = 2

# Shape.Tags に入れるキー（PowerPointファイルに保存される非表示のカスタムデータ。
# ユーザーの画面には出ないので、再実行時の状態復元に使うのに都合が良い）
TAG_ROLE = "RUBIGUI_ROLE"
TAG_ORIG_NAME = "RUBIGUI_ORIG_NAME"
TAG_ORIG_SPACING = "RUBIGUI_ORIG_SPACING"  # 段落ごとの [LineRuleWithin, SpaceWithin] のJSON
TAG_ORIG_AUTOSIZE = "RUBIGUI_ORIG_AUTOSIZE"
TAG_ORIG_ANIM = "RUBIGUI_ORIG_ANIM"

ROLE_RUBY = "RUBY"
ROLE_GROUP = "GROUP"
ROLE_PARENT = "PARENT"

SETTINGS_FILE = "ruby_settings.json"
DEFAULT_SETTINGS = {
    "ruby_ratio": 50,      # ルビのフォントサイズ（親文字に対する％）
    "line_spacing": 1.5,   # 2行以上のときに掛ける行間の倍率
    "ruby_offset": 0.0,    # ルビの位置微調整（pt。正で親文字から離れる）
    "include_title": False,  # タイトル枠もルビ対象にするか
}


# ============================================================
# 語句フィルタ（Word版と同じルール。挙動を揃えるため一切変更していない）
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
    英単語・アルファベット略語などにルビ（カタカナ読み）を振ってしまう不具合の対策。"""
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
    エクスプローラー上の実際の「デスクトップ」フォルダを取得する。"""
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
    """Word版と同じフォルダを共用する（出力の拡張子が違うので衝突しない）"""
    desktop = get_actual_desktop_path()
    base_dir = desktop / "ルビ振り"
    ruby_dir = base_dir / "ルビデータ"
    output_dir = base_dir / "出力（ルビ付き）"
    ruby_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    return ruby_dir, output_dir


def get_powerpoint():
    """PowerPointを取得する。戻り値は (アプリ, 自分で起動したか)。

    ★重要：win32com の Dispatch は「起動中のPowerPointがあればそれにアタッチする」
    仕様のため、処理の最後に無条件で Quit() を呼ぶと、ユーザーが別の資料を
    編集中だった場合にそのインスタンスごと落としてしまう。PowerPointの
    Application.Quit は自動化経由だと保存確認を出さないので、未保存の編集が
    警告なしに消える。先に GetActiveObject で既存インスタンスを探し、
    見つかった場合は「自分で起動したのではない」と覚えておいて Quit しない。"""
    try:
        return win32com.client.GetActiveObject("PowerPoint.Application"), False
    except Exception:
        return win32com.client.Dispatch("PowerPoint.Application"), True


def load_settings():
    """ルビ設定を ruby_settings.json から読み込む（無ければ既定値）"""
    settings = dict(DEFAULT_SETTINGS)
    if not os.path.exists(SETTINGS_FILE):
        return settings
    try:
        with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
            loaded = json.load(f)
        for key in DEFAULT_SETTINGS:
            if key in loaded:
                settings[key] = loaded[key]
    except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
        logging.error(f"{SETTINGS_FILE} の読み込みに失敗しました: {e}")
    return settings


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
    except OSError as e:
        logging.error(f"{SETTINGS_FILE} の保存に失敗しました: {e}")


# ============================================================
# 漢字ブロック分割
# ============================================================
# 々〇〻 と CJK統合漢字（拡張A〜F）＋互換漢字。
# 末尾の範囲は𠮟のような拡張漢字（サロゲートペア）用。Pythonの文字列では1文字だが、
# COM側はUTF-16なので位置の変換が別途必要（_com_index 参照）。
_KANJI_RE = re.compile("[々〇〻㐀-䶿一-鿿豈-﫟𠀀-𯨟]")


def _is_kanji(ch):
    return bool(_KANJI_RE.match(ch))


def split_ruby_blocks(surface, reading):
    """語句と読みから「連続漢字ブロックごとのルビ」を求める。
    戻り値は [(surface内の開始オフセット, 文字数, ルビ), ...]。

    連続した漢字をひとかたまり（ブロック）として扱い、ブロックの内部までは
    分割しない。こうすることで連濁（手作り→てづくり）や熟字訓（大人→おとな）で
    読みの配分を推測せずに済み、破綻しない。
    送り仮名はリテラルとして正規表現に埋め込み、その手前後の漢字ブロックへ
    読みを割り当てる。例：取り引き／とりひき → 取→と、引→ひ
    """
    reading = jaconv.kata2hira(reading or "")
    if not surface or not reading:
        return []

    # surface を「漢字ブロック」と「それ以外」に分ける
    segs = []  # (漢字か, 開始オフセット, 文字列)
    i = 0
    while i < len(surface):
        j = i
        kanji = _is_kanji(surface[i])
        while j < len(surface) and _is_kanji(surface[j]) == kanji:
            j += 1
        segs.append((kanji, i, surface[i:j]))
        i = j

    if not any(seg[0] for seg in segs):
        return []  # 漢字がまったく無い語句にはルビを振らない

    # 全体が1つの漢字ブロック → 分割不要
    if len(segs) == 1:
        return [(0, len(surface), reading)]

    # 正規表現を組む（漢字ブロック=キャプチャ／非漢字=平仮名化したリテラル）
    parts = []
    for idx, (kanji, off, text) in enumerate(segs):
        if kanji:
            # 最後のブロックだけ貪欲にしないと、末尾の読みを取りこぼす
            parts.append("(.+)" if idx == len(segs) - 1 else "(.+?)")
        else:
            parts.append(re.escape(jaconv.kata2hira(text)))

    matched = re.fullmatch("".join(parts), reading)
    if not matched:
        # 送り仮名と読みの整合が取れない（数字混じりなど）→ 語句全体に1つ振る
        logging.debug(f"ブロック分割に失敗したため語句全体へ付与: {surface} / {reading}")
        return [(0, len(surface), reading)]

    blocks = []
    group_index = 0
    for kanji, off, text in segs:
        if not kanji:
            continue
        group_index += 1
        ruby = matched.group(group_index)
        if ruby and ruby != text:
            blocks.append((off, len(text), ruby))
    return blocks


# ============================================================
# 語句抽出（pptx を直接読む。PowerPointを起動しないので高速）
# ============================================================
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"


def _slide_sort_key(name):
    m = re.search(r"slide(\d+)\.xml$", name)
    return int(m.group(1)) if m else 0


def _xml_ph_type(sp):
    """p:sp のプレースホルダ種別を返す（プレースホルダでなければ None）。
    type属性が無いプレースホルダは本文扱いなので空文字を返す。"""
    nv = sp.find(f"{NS_P}nvSpPr")
    if nv is None:
        return None
    nv_pr = nv.find(f"{NS_P}nvPr")
    if nv_pr is None:
        return None
    ph = nv_pr.find(f"{NS_P}ph")
    if ph is None:
        return None
    return ph.get("type") or ""


def _xml_is_textbox(sp):
    """p:cNvSpPr の txBox 属性で「通常のテキストボックス」か判定する。
    これが立っていないプレースホルダ以外の p:sp は、四角形や吹き出しなどの
    オートシェイプであり、ルビの対象外。"""
    nv = sp.find(f"{NS_P}nvSpPr")
    if nv is None:
        return False
    c_nv = nv.find(f"{NS_P}cNvSpPr")
    if c_nv is None:
        return False
    return c_nv.get("txBox") in ("1", "true")


def _xml_is_target(sp, include_title):
    """★重要：ここでの判定は is_target_shape（COM側）と必ず一致させること。
    ずれると「語句一覧には出るのにルビが振られない語句」が生まれる。"""
    ph_type = _xml_ph_type(sp)
    if ph_type is None:
        return _xml_is_textbox(sp)
    if ph_type in XML_ALWAYS_SKIP_PH_TYPES:
        return False
    if not include_title and ph_type in XML_TITLE_PH_TYPES:
        return False
    return True


def extract_pptx_text(file_path, include_title=False):
    """ルビ対象になるシェイプの本文だけを取り出して連結する。
    ★重要：p:spTree の直下の p:sp だけを見ることで、グループ内シェイプ
    （p:grpSp）・表（p:graphicFrame）・図（p:pic）が自然に除外される。
    さらに、シェイプ間・スライド間は必ず改行で区切って連結する。区切らずに
    つなぐと、前のシェイプの末尾と次のシェイプの先頭がひと続きの語
    （「会社」＋「員研修」→「会社員研修」など）として解析され、実在しない
    語句が一覧に載ってしまう。"""
    texts = []
    with zipfile.ZipFile(file_path) as z:
        slide_names = sorted(
            (n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=_slide_sort_key,
        )
        for name in slide_names:
            root = ET.fromstring(z.read(name))
            sp_tree = root.find(f"{NS_P}cSld/{NS_P}spTree")
            if sp_tree is None:
                continue
            for sp in sp_tree.findall(f"{NS_P}sp"):
                if not _xml_is_target(sp, include_title):
                    continue
                for para in sp.iter(f"{NS_A}p"):
                    line = "".join(t.text or "" for t in para.iter(f"{NS_A}t"))
                    if line:
                        texts.append(line)
    return "\n".join(texts)


# SudachiPy(0.6系)は入力が49149バイトを超えると例外を投げる。余裕をみて分割する。
MAX_TOKENIZE_BYTES = 40000


def _split_for_tokenizer(text, limit=MAX_TOKENIZE_BYTES):
    """SudachiPyの入力長制限に収まるよう分割する。
    戻り値は [(元テキストでの開始位置, 断片), ...]。
    ★重要：スライド枚数が多い資料では本文全体が軽く数万バイトを超え、
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


def extract_terms(file_path, override_dict, include_title=False):
    """PowerPointファイルから語句と読みを抽出する。
    除外ルール・override優先・出現順の重複排除はWord版と同一。"""
    full_text = extract_pptx_text(file_path, include_title=include_title)

    result = []
    seen = set()
    for surface, reading_form, _ in tokenize_all(full_text):
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
            if surface == reading:
                continue

        seen.add(surface)
        result.append({"word": surface, "reading": reading})
    return result


# ============================================================
# PowerPoint COM ユーティリティ
# ============================================================
def _tag_get(shape, name):
    """Shape.Tags は未設定のキーに対して空文字を返す"""
    try:
        return shape.Tags.Item(name)
    except Exception:
        return ""


def _tag_set(shape, name, value):
    try:
        shape.Tags.Add(name, str(value))
        return True
    except Exception as e:
        logging.warning(f"タグ設定に失敗: {name}={value} ({e})")
        return False


def _tag_delete(shape, name):
    try:
        if shape.Tags.Item(name):
            shape.Tags.Delete(name)
    except Exception:
        pass


def _com_index(text, py_index):
    """PythonのUnicodeインデックスを、COM(UTF-16)基準の1始まりインデックスへ変換する。
    ★重要：VBA/COMの文字列はUTF-16なので、サロゲートペア（拡張漢字や絵文字）が
    含まれると Characters() の位置がPythonの文字数とずれる。事前に変換しておく。"""
    return len(text[:py_index].encode("utf-16-le")) // 2 + 1


def _com_length(text):
    return len(text.encode("utf-16-le")) // 2


def _is_vertical(shape):
    try:
        return int(shape.TextFrame2.Orientation) in VERTICAL_ORIENTATIONS
    except Exception:
        return False


def _orientation(shape):
    try:
        return int(shape.TextFrame2.Orientation)
    except Exception:
        return ORIENT_HORIZONTAL


def is_target_shape(shape, include_title=False):
    """ルビ振りの対象にするシェイプか判定する。
    対象は「通常のテキストボックス」と「タイトル以外のプレースホルダ」のみ。
    オートシェイプ内テキスト・表・グループ内シェイプ・図は対象外。"""
    try:
        if shape.Type not in (MSO_TEXT_BOX, MSO_PLACEHOLDER):
            return False
        if shape.HasTextFrame != MSO_TRUE:
            return False
        if shape.TextFrame.HasText != MSO_TRUE:
            return False
        if shape.Type == MSO_PLACEHOLDER:
            ph_type = shape.PlaceholderFormat.Type
            if ph_type in PP_ALWAYS_SKIP_PLACEHOLDERS:
                return False
            if not include_title and ph_type in PP_TITLE_PLACEHOLDERS:
                return False
        return True
    except Exception as e:
        logging.debug(f"対象判定でエラー（対象外にします）: {e}")
        return False


def collect_effects(slide, shape_id):
    """指定シェイプに付いているアニメーションを、再現に必要な情報だけ控える。
    ★重要：PowerPointはグループ化・グループ解除の際に構成シェイプの
    アニメーションを破棄する（実測でMainSequence.Countが0になることを確認済み）。
    そのため必ずグループ化の前に控え、後からグループへ付け直す。"""
    effects = []
    try:
        seq = slide.TimeLine.MainSequence
        for i in range(1, seq.Count + 1):
            eff = seq.Item(i)
            try:
                if eff.Shape.Id != shape_id:
                    continue
            except Exception:
                continue
            info = {
                "index": i,
                "effect_type": int(eff.EffectType),
                "trigger_type": int(eff.Timing.TriggerType),
                "delay": float(eff.Timing.TriggerDelayTime),
                "duration": float(eff.Timing.Duration),
            }
            # 効果によっては未対応でエラーになるので個別にtryで囲む
            try:
                info["repeat"] = float(eff.Timing.RepeatCount)
            except Exception:
                pass
            try:
                info["direction"] = int(eff.EffectParameters.Direction)
            except Exception:
                pass
            effects.append(info)
    except Exception as e:
        logging.warning(f"アニメーション情報の取得に失敗しました: {e}")
    return effects


def _effect_positions_of(seq, shape_id):
    """MainSequenceの中で、指定シェイプに付いている効果の位置を順番に返す"""
    positions = []
    for i in range(1, seq.Count + 1):
        try:
            if seq.Item(i).Shape.Id == shape_id:
                positions.append(i)
        except Exception:
            continue
    return positions


def apply_effects(slide, shape, effects, force_with_previous=False, anchor_shape_id=None):
    """控えておいたアニメーションをシェイプへ付け直す。
    再生順は MoveTo で元の位置へ戻す（控えたindexの昇順に処理すれば復元できる）。

    anchor_shape_id を渡した場合は、そのシェイプの「対応する順番の効果」の直後へ
    差し込む。
    ★重要：プレースホルダはグループ化できないためルビへ個別に効果を複製するが、
    このとき元のindexへ移動させてしまうと、ルビが親より先に再生されてしまう
    （親がクリック開始でルビが「直前の動作と同時」だと、ルビだけ先に出る）。
    さらに、親が開始と終了の2つの効果を持つ場合に「親の最後の効果の後ろ」へ
    まとめて置くと、ルビの開始が親の終了と同時になってしまう。
    n番目の効果は必ず親のn番目の効果の直後へ置くこと。
    また、先に入れた分がずれないよう後ろの効果から順に差し込む。"""
    if not effects:
        return
    seq = slide.TimeLine.MainSequence
    ordered = list(enumerate(sorted(effects, key=lambda x: x.get("index", 0))))
    if anchor_shape_id is not None:
        ordered.reverse()
    for slot, info in ordered:
        try:
            eff = seq.AddEffect(shape, info["effect_type"])
        except Exception as e:
            logging.warning(
                f"アニメーションの再設定に失敗しました（EffectType={info.get('effect_type')}）: {e}"
            )
            continue
        try:
            eff.Timing.TriggerType = (
                MSO_ANIM_TRIGGER_WITH_PREVIOUS if force_with_previous else info["trigger_type"]
            )
            eff.Timing.TriggerDelayTime = info["delay"]
            if info.get("duration"):
                eff.Timing.Duration = info["duration"]
            if "repeat" in info:
                eff.Timing.RepeatCount = info["repeat"]
        except Exception as e:
            logging.debug(f"アニメーションのタイミング設定を一部適用できませんでした: {e}")
        if "direction" in info:
            try:
                eff.EffectParameters.Direction = info["direction"]
            except Exception:
                pass
        try:
            if anchor_shape_id is not None:
                positions = _effect_positions_of(seq, anchor_shape_id)
                # 親の同じ順番の効果の直後へ。親の効果が足りなければ最後の直後。
                target = (positions[slot] if slot < len(positions)
                          else (positions[-1] if positions else 0)) + 1
            else:
                target = info.get("index")
            if target and 1 <= target <= seq.Count:
                eff.MoveTo(target)
        except Exception as e:
            logging.debug(f"アニメーションの順序復元に失敗しました: {e}")


def remove_effects_for_shape(slide, shape_id):
    """指定シェイプに付いているアニメーションを全部消す（逆順に消さないとindexがずれる）"""
    try:
        seq = slide.TimeLine.MainSequence
        for i in range(seq.Count, 0, -1):
            eff = seq.Item(i)
            try:
                if eff.Shape.Id == shape_id:
                    eff.Delete()
            except Exception:
                continue
    except Exception as e:
        logging.warning(f"アニメーションの削除に失敗しました: {e}")


def _contains_rubigui_marker(shape, depth=0):
    """グループの中（入れ子も含む）に本ツールの目印があるかを調べる。
    ★重要：ユーザーがルビグループをさらに別の図形とグループ化していると、
    最上位のグループには目印が付いていない。中まで見ないと再実行時に
    古いルビを見つけられず、二重にルビが振られてしまう。"""
    if depth > 5:
        return False
    if _tag_get(shape, TAG_ROLE) in (ROLE_GROUP, ROLE_RUBY, ROLE_PARENT):
        return True
    try:
        if shape.Type != MSO_GROUP:
            return False
        items = shape.GroupItems
        for i in range(1, items.Count + 1):
            if _contains_rubigui_marker(items.Item(i), depth + 1):
                return True
    except Exception:
        return False
    return False


def cleanup_slide(slide):
    """前回のルビ振りの痕跡を消して、スライドを元の状態へ戻す（再実行対応）。
    ★重要：グループ解除でもアニメーションは破棄されるため、親シェイプの
    タグに控えておいた元のアニメーションから復元する。"""
    # 1) 目印を含むグループを解除して中身をスライド直下へ戻す。
    #    入れ子になっている場合があるので、解除できるものが無くなるまで繰り返す。
    for _ in range(6):
        ungrouped = False
        for i in range(slide.Shapes.Count, 0, -1):
            try:
                shape = slide.Shapes.Item(i)
            except Exception:
                continue
            if shape.Type != MSO_GROUP or not _contains_rubigui_marker(shape):
                continue
            try:
                shape.Ungroup()
                ungrouped = True
            except Exception as e:
                logging.warning(f"ルビグループの解除に失敗しました: {e}")
        if not ungrouped:
            break

    # 2) ルビ図形を削除
    for i in range(slide.Shapes.Count, 0, -1):
        try:
            shape = slide.Shapes.Item(i)
        except Exception:
            continue
        if _tag_get(shape, TAG_ROLE) == ROLE_RUBY:
            try:
                shape.Delete()
            except Exception as e:
                logging.warning(f"ルビ図形の削除に失敗しました: {e}")

    # 3) 親シェイプの行間・自動調整・図形名・アニメーションを復元
    for i in range(1, slide.Shapes.Count + 1):
        try:
            shape = slide.Shapes.Item(i)
        except Exception:
            continue
        if _tag_get(shape, TAG_ROLE) != ROLE_PARENT:
            continue
        try:
            restore_parent_shape(slide, shape)
        except Exception as e:
            logging.warning(f"親シェイプの復元に失敗しました: {e}")


def restore_parent_shape(slide, shape):
    """タグに退避しておいた元の状態へ親シェイプを戻す"""
    spacing_tag = _tag_get(shape, TAG_ORIG_SPACING)
    if spacing_tag:
        try:
            original = json.loads(spacing_tag)
        except (json.JSONDecodeError, TypeError):
            original = []
        try:
            paragraphs = shape.TextFrame.TextRange.Paragraphs()
            for index, entry in enumerate(original, start=1):
                if index > paragraphs.Count:
                    break
                rule, space = entry
                # 混在（msoTriStateMixed = -2）などの不正値は書き戻せないので飛ばす
                if int(rule) not in (MSO_TRUE, MSO_FALSE):
                    continue
                pf = shape.TextFrame.TextRange.Paragraphs(index).ParagraphFormat
                pf.LineRuleWithin = int(rule)
                pf.SpaceWithin = float(space)
        except Exception as e:
            logging.warning(f"行間の復元に失敗しました: {e}")

    autosize = _tag_get(shape, TAG_ORIG_AUTOSIZE)
    if autosize != "":
        try:
            shape.TextFrame2.AutoSize = int(float(autosize))
        except Exception as e:
            logging.warning(f"自動調整の復元に失敗しました: {e}")

    anim = _tag_get(shape, TAG_ORIG_ANIM)
    if anim != "":
        try:
            effects = json.loads(anim)
        except (json.JSONDecodeError, TypeError):
            effects = []
        # グループ化で失われた分を復元する。二重付与を避けるため一度全部消す。
        remove_effects_for_shape(slide, shape.Id)
        apply_effects(slide, shape, effects)

    orig_name = _tag_get(shape, TAG_ORIG_NAME)
    if orig_name:
        try:
            shape.Name = orig_name
        except Exception:
            pass

    for key in (TAG_ROLE, TAG_ORIG_NAME, TAG_ORIG_SPACING,
                TAG_ORIG_AUTOSIZE, TAG_ORIG_ANIM):
        _tag_delete(shape, key)


class GlyphExtentMeasurer:
    """フォントごとの「行間1.0のときの行の高さ（縦書きなら幅）」を実測して覚えておく。

    ★重要：行の高さはフォントによって違う（MS ゴシックは約1.2倍だが、
    メイリオなどは1.4倍近い）。定数を決め打ちすると配置が狂うので、
    使い捨てのテキストボックスを1つ作って実測し、フォント＋サイズ単位で
    キャッシュする。行間を広げた行ボックスのどこにグリフがあるかは、
    この実測値からしか正しく求められない。"""

    def __init__(self, refresh):
        self.refresh = refresh
        self.cache = {}

    def get(self, slide, font, vertical):
        """★重要：日本語の行の高さを決めるのは Font.NameFarEast の方なので、
        プローブにも Name と NameFarEast の両方を設定する。片方だけだと
        親がメイリオでも既定のMSゴシックの行高を測ってしまい、
        ルビが数pt浮いたり親文字に食い込んだりする。"""
        size = float(font["size"])
        key = (font["name"], font["far_east"], round(size, 2), bool(vertical))
        if key in self.cache:
            return self.cache[key]

        extent = None
        probe = None
        try:
            orient = ORIENT_VERTICAL_FAREAST if vertical else ORIENT_HORIZONTAL
            probe = slide.Shapes.AddTextbox(orient, -2000, -2000, 60, 60)
            tf = probe.TextFrame
            tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
            tf.WordWrap = MSO_FALSE
            tr = tf.TextRange
            tr.Text = "国"
            tr.Font.Size = size
            if font["name"]:
                tr.Font.Name = font["name"]
            if font["far_east"]:
                try:
                    tr.Font.NameFarEast = font["far_east"]
                except Exception:
                    pass
            pf = tr.ParagraphFormat
            pf.LineRuleWithin = MSO_TRUE
            pf.SpaceWithin = 1.0
            self.refresh()
            rng = tr.Characters(1, 1)
            extent = float(rng.BoundWidth if vertical else rng.BoundHeight)
        except Exception as e:
            logging.warning(f"行の高さの実測に失敗しました（推定値を使います）: {e}")
        finally:
            if probe is not None:
                try:
                    probe.Delete()
                except Exception:
                    pass

        if not extent or extent <= 0:
            extent = float(size) * 1.2  # 実測できなかったときの保険
        self.cache[key] = extent
        return extent


def _font_of(text_range):
    """親文字の書式を採取する。範囲内で書式が混在していると PowerPoint は
    -2（混在）や空文字を返すので、必ず先頭1文字から読む。"""
    font = text_range.Font
    info = {"name": "", "far_east": "", "size": 18.0,
            "rgb": None, "bold": MSO_FALSE}
    try:
        name = font.Name
        if name and name != "":
            info["name"] = name
    except Exception:
        pass
    try:
        far_east = font.NameFarEast
        if far_east:
            info["far_east"] = far_east
    except Exception:
        pass
    try:
        size = float(font.Size)
        if size > 0:
            info["size"] = size
    except Exception:
        pass
    try:
        info["rgb"] = int(font.Color.RGB)
    except Exception:
        pass
    try:
        bold = int(font.Bold)
        info["bold"] = bold if bold in (MSO_TRUE, MSO_FALSE) else MSO_FALSE
    except Exception:
        pass
    return info


def build_ruby_blocks(full_text, term_dict):
    """本文全体を形態素解析し、ルビを振るブロックの一覧を作る。
    戻り値は [(Python上の開始位置, 文字数, ルビ), ...]"""
    blocks = []
    for surface, _reading_form, begin in tokenize_all(full_text):
        reading = term_dict.get(surface)
        if not reading:
            continue  # GUIで削除された語句・対象外の語句は振らない
        for offset, length, ruby in split_ruby_blocks(surface, reading):
            blocks.append((begin + offset, length, ruby))
    return blocks


def adjust_line_spacing(shape, multiplier):
    """2行以上のときだけ行間を広げ、ルビを置く余白を作る。
    戻り値は「段落ごとの適用後の行間倍率」のリスト（先頭が1段落目）。

    ★重要：PowerPointの倍数指定の行間は、余白を各行の「上」（縦書きなら「右」）に
    追加する。先頭行にも余白が入るので、1行目のルビもシェイプ内に収まる。
    ★重要：段落ごとに行間が違うシェイプ（箇条書きのレベル別など）では、
    シェイプ全体の ParagraphFormat は「混在」を意味する -2 を返す。これを
    そのまま退避すると復元時に不正値の書き戻しで例外になり、行間が広がった
    まま元に戻せなくなる。必ず段落ごとに読み書きすること。"""
    tr = shape.TextFrame.TextRange
    try:
        line_count = tr.Lines().Count
    except Exception as e:
        logging.debug(f"行数の取得に失敗しました: {e}")
        line_count = 1
    try:
        para_count = int(tr.Paragraphs().Count)
    except Exception as e:
        logging.warning(f"段落数の取得に失敗しました: {e}")
        return [1.0]

    originals = []
    bases = []
    for index in range(1, para_count + 1):
        try:
            pf = tr.Paragraphs(index).ParagraphFormat
            rule = int(pf.LineRuleWithin)
            space = float(pf.SpaceWithin)
        except Exception as e:
            logging.debug(f"段落{index}の行間取得に失敗しました: {e}")
            rule, space = MSO_TRUE, 1.0
        originals.append([rule, space])
        # 元がpt指定・混在の場合は基準を1.0とみなす（倍数へ作り替える）
        bases.append(max(space, 1.0) if rule == MSO_TRUE else 1.0)

    if line_count < 2:
        # 1行だけなら行間を触らない。倍数指定ならその値が実効の行間倍率。
        return bases

    _tag_set(shape, TAG_ORIG_SPACING, json.dumps(originals))
    final = []
    for index in range(1, para_count + 1):
        value = bases[index - 1] * float(multiplier)
        try:
            pf = tr.Paragraphs(index).ParagraphFormat
            pf.LineRuleWithin = MSO_TRUE
            pf.SpaceWithin = value
        except Exception as e:
            logging.warning(f"段落{index}の行間設定に失敗しました: {e}")
            value = bases[index - 1]
        final.append(value)
    return final


def place_ruby_boxes(slide, shape, blocks, settings, measurer, refresh, spacings):
    """ルビ用のテキストボックスを作って親文字の上（縦書きは右）へ配置する。
    親シェイプの書式には一切書き込まず、ルビ側が親から読み取ってコピーするだけなので、
    縦横・フォント名・色・サイズはそのまま残る。"""
    vertical = _is_vertical(shape)
    orient = _orientation(shape)
    ratio = float(settings["ruby_ratio"]) / 100.0
    offset = float(settings["ruby_offset"])
    full_text = shape.TextFrame.TextRange.Text
    created = []

    for py_start, py_len, ruby in blocks:
        block_text = full_text[py_start:py_start + py_len]
        start = _com_index(full_text, py_start)
        length = _com_length(block_text)
        if length <= 0:
            continue
        try:
            rng = shape.TextFrame.TextRange.Characters(start, length)
            bound_left = float(rng.BoundLeft)
            bound_top = float(rng.BoundTop)
            bound_width = float(rng.BoundWidth)
            bound_height = float(rng.BoundHeight)
            font = _font_of(shape.TextFrame.TextRange.Characters(start, 1))
        except Exception as e:
            logging.warning(f"「{block_text}」の座標取得に失敗したため飛ばします: {e}")
            continue

        # 行ボックスの中でグリフがどこにあるかを求める。
        # 余白は行の「先頭側」（横書き=上、縦書き=右）に入るので、
        # グリフの開始位置は 行ボックスの先頭 + 余白 になる。
        # 行間は段落ごとに違い得るので、このブロックが属する段落の値を使う。
        # （COMの本文では段落の区切りが \r になっている）
        para_index = full_text[:py_start].count("\r")
        if spacings:
            spacing = spacings[min(para_index, len(spacings) - 1)]
        else:
            spacing = 1.0
        glyph_extent = measurer.get(slide, font, vertical)
        lead = max(glyph_extent * (spacing - 1.0), 0.0)

        box = None
        try:
            box = slide.Shapes.AddTextbox(
                ORIENT_VERTICAL_FAREAST if vertical else ORIENT_HORIZONTAL, 0, 0, 20, 20
            )
            tf = box.TextFrame
            tf.MarginLeft = tf.MarginRight = tf.MarginTop = tf.MarginBottom = 0
            tf.WordWrap = MSO_FALSE
            tf.AutoSize = PP_AUTOSIZE_SHAPE_TO_FIT_TEXT
            tr = tf.TextRange
            tr.Text = ruby
            tr.Font.Size = max(round(font["size"] * ratio, 1), 1.0)
            if font["name"]:
                tr.Font.Name = font["name"]
            if font["far_east"]:
                try:
                    tr.Font.NameFarEast = font["far_east"]
                except Exception:
                    pass
            # 親の色が読み取れなかった場合は色を設定しない。
            # 黒で決め打ちすると、濃い背景に白文字のスライドでルビだけ
            # 真っ黒になって読めなくなる。
            if font["rgb"] is not None:
                try:
                    tr.Font.Color.RGB = font["rgb"]
                except Exception:
                    pass
            tr.Font.Bold = font["bold"]
            tr.ParagraphFormat.Alignment = PP_ALIGN_CENTER
            box.Line.Visible = MSO_FALSE
            box.Fill.Visible = MSO_FALSE
            refresh()

            if vertical:
                # 縦書き：行の先頭側＝右。グリフの右隣（＝余白）へ置く。
                # 最終列は BoundWidth が切り詰められることがあるので、
                # 切り詰めの影響を受けない右端（BoundLeft+BoundWidth）を基準にする。
                flow_start = bound_left + bound_width
                box.Left = flow_start - lead + offset
                box.Top = bound_top + bound_height / 2.0 - box.Height / 2.0
            else:
                # 横書き：行の先頭側＝上。グリフの真上（＝余白）へ置く。
                glyph_top = bound_top + lead
                box.Left = bound_left + bound_width / 2.0 - box.Width / 2.0
                box.Top = glyph_top - box.Height - offset

            # ★重要：目印が付かないルビ図形は再実行時に見つけられず、
            # 処理のたびに重複して溜まっていく。付けられなかったら捨てる。
            if not _tag_set(box, TAG_ROLE, ROLE_RUBY):
                logging.warning(f"「{block_text}」のルビに目印を付けられないため取り消します")
                box.Delete()
                continue
            created.append(box)
        except Exception as e:
            logging.warning(f"「{block_text}」のルビ配置に失敗しました: {e}")
            if box is not None:
                try:
                    box.Delete()
                except Exception:
                    pass
    logging.debug(f"ルビ配置: {len(created)}件（縦書き={vertical} orientation={orient}）")
    return created


def process_shape(slide, shape, term_dict, settings, measurer, refresh, warnings):
    """1つのシェイプにルビを振り、グループ化とアニメーション継承まで行う"""
    try:
        rotation = float(shape.Rotation)
    except Exception:
        rotation = 0.0
    if abs(rotation) > 0.01:
        warnings.append(f"回転している図形はルビ対象外です: {shape.Name}")
        return 0

    orientation = _orientation(shape)
    if orientation != ORIENT_HORIZONTAL and orientation not in VERTICAL_ORIENTATIONS:
        warnings.append(f"未対応の文字方向のためルビ対象外です: {shape.Name}")
        return 0

    full_text = shape.TextFrame.TextRange.Text
    blocks = build_ruby_blocks(full_text, term_dict)
    if not blocks:
        return 0

    shape_id = shape.Id
    orig_name = shape.Name

    # ★重要：グループ化するとアニメーションが破棄されるので、必ず先に控える。
    #    さらに、次回の再実行時に復元できるよう親のタグへも保存しておく。
    effects = collect_effects(slide, shape_id)

    # ★重要：目印（ROLE=PARENT）は、シェイプの状態をいじる「前」に付けること。
    # 後ろに置くと、途中でルビを1つも置けなかった場合や例外が出た場合に
    # 「行間だけ広がって自動調整も切れているのに目印が無い」シェイプが残る。
    # そうなると cleanup_slide が復元対象として見つけられず、そのファイルを
    # 処理し直すたびに広げた行間を元の値と誤認して 1.5→2.25→3.375 と
    # 累乗で広がり続けてしまう。
    _tag_set(shape, TAG_ROLE, ROLE_PARENT)
    _tag_set(shape, TAG_ORIG_NAME, orig_name)

    ruby_boxes = []
    try:
        # 自動調整（縮小型）はフォントサイズを勝手に変えてしまうので切る
        try:
            autosize = int(shape.TextFrame2.AutoSize)
        except Exception:
            autosize = MSO_AUTOSIZE_NONE
        if autosize == MSO_AUTOSIZE_TEXT_TO_FIT_SHAPE:
            _tag_set(shape, TAG_ORIG_AUTOSIZE, autosize)
            try:
                shape.TextFrame2.AutoSize = MSO_AUTOSIZE_NONE
            except Exception as e:
                logging.warning(f"自動調整の解除に失敗しました: {e}")

        spacings = adjust_line_spacing(shape, settings["line_spacing"])

        # ★重要：行間を変えると文字位置が動く。必ず再レイアウトさせてから座標を取る。
        refresh()

        ruby_boxes = place_ruby_boxes(
            slide, shape, blocks, settings, measurer, refresh, spacings
        )
    except Exception as e:
        logging.error(f"ルビ配置中にエラーが発生しました（元に戻します）: {e}")
        warnings.append(f"ルビを配置できませんでした: {orig_name}（{e}）")
        for box in ruby_boxes:
            try:
                box.Delete()
            except Exception:
                pass
        ruby_boxes = []

    if not ruby_boxes:
        # 1件も置けなかったら、広げた行間や切った自動調整を元へ戻す
        try:
            restore_parent_shape(slide, shape)
        except Exception as e:
            logging.warning(f"元の状態への復元に失敗しました: {e}")
        return 0

    is_placeholder = shape.Type == MSO_PLACEHOLDER
    if is_placeholder:
        # PowerPointの仕様上、プレースホルダは他の図形とグループ化できない。
        # 親はそのまま残し、同じアニメーションを各ルビへ複製して同時再生させる。
        if effects:
            for box in ruby_boxes:
                apply_effects(
                    slide, box, effects,
                    force_with_previous=True, anchor_shape_id=shape_id,
                )
        return len(ruby_boxes)

    # テキストボックスは 親＋ルビ を1グループにまとめる
    _tag_set(shape, TAG_ORIG_ANIM, json.dumps(effects, ensure_ascii=False))
    names = []
    try:
        unique = f"RUBIGUI_PARENT_{slide.SlideIndex}_{shape_id}"
        shape.Name = unique
        names.append(unique)
        for i, box in enumerate(ruby_boxes, start=1):
            box_name = f"RUBIGUI_RUBY_{slide.SlideIndex}_{shape_id}_{i}"
            box.Name = box_name
            names.append(box_name)
        group = slide.Shapes.Range(names).Group()
        # ★重要：目印はグループ化の直後、名前を付ける前に打つこと。
        # 先に Name を代入して例外が出ると、目印の無いグループだけが残り、
        # 以降の実行で中のルビも親も一切見つけられなくなる。
        _tag_set(group, TAG_ROLE, ROLE_GROUP)
        try:
            group.Name = f"RUBIGUI_GROUP_{slide.SlideIndex}_{shape_id}"
        except Exception:
            pass
        # 図形名はユーザーの見えるところなので元に戻しておく
        try:
            group.GroupItems.Item(1).Name = orig_name
        except Exception:
            pass
        apply_effects(slide, group, effects)
    except Exception as e:
        logging.warning(f"グループ化に失敗しました（ルビは配置済みです）: {e}")
        warnings.append(f"グループ化できませんでした: {orig_name}")
        try:
            shape.Name = orig_name
        except Exception:
            pass
        apply_effects(slide, shape, effects)
    return len(ruby_boxes)


def generate_ruby_pptx(pptx_path, term_dict, settings, ppt_app=None, warnings=None):
    """PowerPointファイルにルビを振って「出力（ルビ付き）」フォルダへ保存する。
    元ファイルには一切書き込まず、先に出力先へコピーしてからコピーだけを編集する。"""
    if warnings is None:
        warnings = []
    src = Path(pptx_path)
    _, output_dir = get_ruby_project_dirs()
    out_path = output_dir / f"{src.stem}（ルビ）.pptx"

    if src.resolve() != out_path.resolve():
        shutil.copy2(src, out_path)
    logging.info(f"ルビ振り開始: {src} → {out_path}")

    if ppt_app is None:
        ppt, own_app = get_powerpoint()
    else:
        ppt, own_app = ppt_app, False
    # ★重要：PowerPointは Visible=False だと不安定で、しかも非表示ウィンドウでは
    #    TextRange.Bound* が取得できない。必ず表示状態で開く。
    ppt.Visible = True

    pres = None
    total_ruby = 0
    try:
        pres = ppt.Presentations.Open(str(out_path), ReadOnly=MSO_FALSE, WithWindow=MSO_TRUE)
        window = pres.Windows(1)
        # 行の高さの実測結果はフォント単位でしか変わらないので、
        # スライドごとではなくファイル全体で使い回す（プローブ生成の往復を減らす）
        measurer = GlyphExtentMeasurer(lambda: None)

        for index in range(1, pres.Slides.Count + 1):
            slide = pres.Slides.Item(index)

            def refresh(idx=index):
                """座標を取る前にスライドを描画させる"""
                try:
                    window.View.GotoSlide(idx)
                except Exception:
                    pass

            refresh()
            cleanup_slide(slide)
            refresh()

            targets = []
            for i in range(1, slide.Shapes.Count + 1):
                shape = slide.Shapes.Item(i)
                if is_target_shape(shape, settings.get("include_title", False)):
                    targets.append(shape)

            measurer.refresh = refresh
            for shape in targets:
                try:
                    total_ruby += process_shape(
                        slide, shape, term_dict, settings, measurer, refresh, warnings
                    )
                except Exception as e:
                    logging.error(f"スライド{index}のシェイプ処理でエラー: {e}")
                    warnings.append(f"スライド{index}: {e}")

        pres.Save()
        logging.info(f"ルビ振り完了: {out_path}（ルビ {total_ruby} 件）")
    finally:
        if pres is not None:
            try:
                pres.Close()
            except Exception as e:
                logging.warning(f"プレゼンテーションのクローズに失敗しました: {e}")
        if own_app:
            try:
                ppt.Quit()
            except Exception as e:
                logging.warning(f"PowerPointの終了に失敗しました: {e}")
    return out_path, total_ruby


# ============================================================
# GUI
# ============================================================
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
        self.root.title(f"ルビ編集ツール(PowerPoint版) v{APP_VERSION}")
        self.root.geometry("800x600")
        self.center_main_window(800, 600)
        self.data = []
        self.override_dict = {}
        self.settings = load_settings()
        self.file_path = None
        self.current_file_path = None
        self.pptx_files = []
        self.current_index = 0
        self.review_ready = False  # 「一括処理(確認あり)」で語句を表示済みか
        self.setup_ui()
        self.load_override_dict()
        self.ruby_dir, self.output_dir = get_ruby_project_dirs()

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
        # 下限を1.0にすると余白がゼロになり、ルビが上の行の文字へ完全に重なる
        self.settings["line_spacing"] = to_float(self.spacing_entry, "line_spacing", 1.1, 4.0)
        self.settings["ruby_offset"] = to_float(self.offset_entry, "ruby_offset", -50, 50)
        self.settings["include_title"] = bool(self.include_title_var.get())
        save_settings(self.settings)
        return self.settings

    def setup_ui(self):
        # ✅ バージョン表記（画面下部に常時表示。サポート対応時に問い合わせてもらいやすくするため）
        version_label = tk.Label(
            self.root,
            text=f"RubiGUI-PPT v{APP_VERSION}",
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
        self.drop_label = tk.Label(
            left_frame, text="ここにPowerPointファイルをドロップ", relief="ridge", height=4
        )
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
        single_file_frame.pack(fill="x", pady=(0, 12))
        tk.Button(single_file_frame, text="TSV保存", command=self.save_only_tsv).pack(fill="x", pady=2)
        tk.Button(
            single_file_frame, text="ルビ付きPowerPoint出力", command=self.run_single_file_output
        ).pack(fill="x", pady=2)

        # ルビ設定
        settings_frame = tk.LabelFrame(right_frame, text="ルビ設定", padx=10, pady=5)
        settings_frame.pack(fill="x", pady=(0, 12))

        tk.Label(settings_frame, text="ルビの大きさ（親文字の％）", anchor="w").pack(fill="x")
        self.ratio_entry = tk.Entry(settings_frame)
        self.ratio_entry.insert(0, str(self.settings["ruby_ratio"]))
        self.ratio_entry.pack(fill="x", pady=(0, 4))

        tk.Label(settings_frame, text="行間の倍率（2行以上のとき）", anchor="w").pack(fill="x")
        self.spacing_entry = tk.Entry(settings_frame)
        self.spacing_entry.insert(0, str(self.settings["line_spacing"]))
        self.spacing_entry.pack(fill="x", pady=(0, 4))

        tk.Label(settings_frame, text="ルビの位置調整（pt）", anchor="w").pack(fill="x")
        self.offset_entry = tk.Entry(settings_frame)
        self.offset_entry.insert(0, str(self.settings["ruby_offset"]))
        self.offset_entry.pack(fill="x", pady=(0, 4))

        self.include_title_var = tk.IntVar(value=1 if self.settings["include_title"] else 0)
        tk.Checkbutton(
            settings_frame,
            text="タイトル枠も対象にする",
            variable=self.include_title_var,
            anchor="w",
            command=self.on_include_title_changed,
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

    def on_include_title_changed(self):
        """タイトル枠の扱いを変えたら、語句一覧も即座に取り直す
        （一覧に出る語句＝実際にルビが振られる語句、という対応を崩さないため）"""
        self.settings["include_title"] = bool(self.include_title_var.get())
        save_settings(self.settings)
        if self.current_file_path and self._confirm_discard_edits():
            self.extract_words(self.current_file_path)

    def _confirm_discard_edits(self):
        """語句一覧を作り直すと手で直した読みが消えるので、消える前に一言確認する"""
        if not self.data:
            return True
        return msgbox.askyesno(
            "語句一覧の再読み込み",
            "語句一覧を読み込み直します。\n"
            "手動で修正した読みは失われますがよろしいですか？"
        )

    # ---------- ファイル受け取り ----------
    def batch_process(self, event):
        logging.debug(f"batch_process() 呼び出し元データ: {event.data}")
        paths = self.root.tk.splitlist(event.data)
        file_paths = [Path(p.strip()) for p in paths]
        pptx_files = [f for f in file_paths if f.suffix.lower() == ".pptx" and not f.name.startswith("~$")]

        if not pptx_files:
            msgbox.showinfo("一括処理", ".pptx ファイルが見つかりません。")
            return

        if len(pptx_files) == 1:
            # ✅ 1ファイルだけ → 抽出＋表示だけに留める
            logging.debug("1ファイルのみ → 語句抽出＋表示のみ処理")
            self.extract_words(pptx_files[0])
            return

        # ✅ 複数ファイル → フル一括処理へ
        self._process_batch_files(pptx_files)

    def process_pptx_batch(self, file_paths):
        logging.debug(f"process_pptx_batch 呼び出し: {[str(f) for f in file_paths]}")
        pptx_files = [f for f in file_paths if f.suffix.lower() == ".pptx" and not f.name.startswith("~$")]
        if not pptx_files:
            msgbox.showinfo("一括処理", ".pptx ファイルが見つかりません。")
            return
        self._process_batch_files(pptx_files)

    def _update_batch_progress(self, text):
        if hasattr(self, "progress_label"):
            self.progress_label.config(text=text)
            self.root.update_idletasks()

    def _process_batch_files(self, pptx_files):
        """複数ファイルの一括処理本体。
        PowerPointは1つだけ起動して使い回し、ファイルごとの起動・終了コストを省く。"""
        settings = self.read_settings_from_ui()
        total = len(pptx_files)
        success_count = 0
        failures = []
        warnings = []

        try:
            ppt, own_app = get_powerpoint()
            ppt.Visible = True
        except Exception as e:
            msgbox.showerror("PowerPoint起動エラー", f"PowerPointの起動に失敗しました。\n{e}")
            return

        try:
            for i, file_path in enumerate(pptx_files, start=1):
                file_path = Path(file_path)
                self._update_batch_progress(f"処理中: {i}/{total} - {file_path.name}")
                try:
                    logging.info(f"処理中: {file_path}")
                    terms = extract_terms(
                        file_path, self.override_dict, settings.get("include_title", False)
                    )
                    logging.debug(f"{file_path.name} 語句抽出数: {len(terms)}")
                    if not terms:
                        logging.warning(f"{file_path} → 語句ゼロ（抽出失敗の可能性）")
                    term_pairs = [(t["word"], t["reading"]) for t in terms]
                    self.save_tsv(term_pairs, file_path)
                    _, count = generate_ruby_pptx(
                        file_path, dict(term_pairs), settings, ppt_app=ppt, warnings=warnings
                    )
                    if count == 0:
                        logging.warning(f"{file_path} → ルビを1件も配置できませんでした")
                    success_count += 1
                except Exception as e:
                    logging.error(f"{file_path} → {e}")
                    failures.append(f"{file_path.name}（{e}）")
        finally:
            if own_app:
                try:
                    ppt.Quit()
                except Exception:
                    pass
            self._update_batch_progress("")
            # ★重要：一括処理中は語句一覧の表示を更新していないので、
            # 処理後の self.data / self.file_path をそのまま残すと
            # 「画面はAのまま、中身は最後に処理したファイルの語句」という
            # ちぐはぐな状態になる。この状態で「ルビ付きPowerPoint出力」や
            # 「TSV保存」を押すと、別ファイルの読みでルビが振られてしまう。
            self._clear_current_file()

        summary = f"{total} 件中 {success_count} 件のファイルを正常に処理しました。"
        if failures:
            summary += "\n\n失敗したファイル:\n" + "\n".join(failures)
        if warnings:
            summary += "\n\n注意:\n" + "\n".join(dict.fromkeys(warnings[:10]))
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

    def _list_pptx(self, folder):
        """フォルダ内の .pptx を列挙する。
        ★重要：glob.glob はフォルダ名の [ ] をワイルドカードとして解釈するため、
        「資料[2024]」のような名前のフォルダで0件になる。必ず Path.glob を使う。"""
        return [f for f in Path(folder).glob("*.pptx") if not f.name.startswith("~$")]

    def start_batch_review(self):
        if not getattr(self, "target_folder", None):
            msgbox.showwarning("警告", "先に「フォルダ選択」で処理対象フォルダを選んでください。")
            return
        self.pptx_files = self._list_pptx(self.target_folder)
        self.current_index = 0
        if not self.pptx_files:
            msgbox.showinfo("一括処理", f"{self.target_folder} に .pptx ファイルが見つかりません。")
            return
        self.process_next_file()

    def process_next_file(self):
        if self.current_index >= len(self.pptx_files):
            msgbox.showinfo("完了", "すべてのファイルを処理しました")
            self._clear_current_file()
            return
        file_path = Path(self.pptx_files[self.current_index])
        self.extract_words(file_path)
        if self.current_file_path is None:
            # 読み込みに失敗（extract_words内でエラー表示済み）。止まらず次へ進める。
            self.current_index += 1
            self.process_next_file()
            return
        self.review_ready = True
        msgbox.showinfo("確認", f"{file_path.name} の語句を確認・修正してください")

    def select_folder(self):
        folder = filedialog.askdirectory(title="処理対象フォルダを選択")
        if folder:
            self.target_folder = folder
            self.pptx_files = self._list_pptx(folder)
            self.current_index = 0
            # フォルダを選び直した時点では、まだどのファイルも確認していない
            self.review_ready = False
            msgbox.showinfo("フォルダ選択", f"選択されたフォルダ:\n{folder}")

    def on_batch_button_click(self):
        logging.debug("on_batch_button_click 呼び出し")
        if getattr(self, "target_folder", None):
            pptx_files = self._list_pptx(self.target_folder)
            if pptx_files:
                self.process_pptx_batch(pptx_files)
            else:
                msgbox.showinfo("情報", "フォルダ内に PowerPoint ファイルが見つかりませんでした。")
        else:
            msgbox.showwarning("警告", "フォルダが選択されていません。")

    def select_file_for_processing(self, event=None):
        file_path = filedialog.askopenfilename(filetypes=[("PowerPointファイル", "*.pptx")])
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
            terms = extract_terms(
                path, self.override_dict, self.settings.get("include_title", False)
            )
        except Exception as e:
            logging.error(f"語句抽出に失敗しました: {e}")
            msgbox.showerror(
                "読み込みエラー",
                f"{Path(path).name} を読み込めませんでした。\n"
                f"PowerPointファイル（.pptx）として開けるか確認してください。\n\n{e}"
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

    def reapply_ruby(self):
        """辞書を読み直して、現在開いているファイルの語句一覧に反映する"""
        if not self.current_file_path:
            msgbox.showwarning("辞書再適用", "先にPowerPointファイルをドロップまたは選択してください。")
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

            # ★重要：語句・読みを空欄にして「保存」すると、空文字の行が残って
            # TSV出力にも空欄行が混ざる。削除は下の「削除」ボタンを使ってもらう。
            if not new_word or not new_reading:
                msgbox.showwarning(
                    "エラー",
                    "語句・読みは空欄のまま保存できません。\n"
                    "この語句を削除したい場合は「削除」ボタンを使用してください。"
                )
                return

            index = self.tree.index(item_id)
            tag = "evenrow" if index % 2 == 0 else "oddrow"
            self.tree.item(item_id, values=(new_word, new_reading), tags=(tag,))

            for i, (w, r) in enumerate(self.data):
                if w == word and r == reading:
                    self.data[i] = (new_word, new_reading)
                    break
            edit_win.destroy()

        def delete_edit():
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
    def save_tsv(self, terms, file_path):
        logging.debug("TSV保存開始")
        base_name = Path(file_path).stem
        tsv_path = self.ruby_dir / f"{base_name}.tsv"

        # ✅ PPTX版はVBAマクロを経由しないため、TSVの文字コードにcp932の制約がない。
        #    Word版で起きていた「cp932非対応文字を含む語句が除外される」問題を避けるため
        #    UTF-8（BOM付き）で出力する。BOM付きなのでExcelでダブルクリックしても
        #    文字化けしない。
        try:
            with open(tsv_path, "w", encoding="utf-8-sig", newline="") as f:
                for word, reading in terms:
                    f.write(f"{word}\t{reading}\n")
        except OSError as e:
            logging.error(f"TSVの保存に失敗しました: {e}")
            msgbox.showerror("TSV保存", f"TSVファイルを保存できませんでした。\n{e}")
            return
        logging.debug(f"TSV保存完了: {tsv_path}（{len(terms)}件）")

    def save_only_tsv(self):
        if not self.file_path or not self.data:
            msgbox.showwarning("エラー", "処理対象がありません")
            return
        terms = [(w, r) for w, r in self.data]
        self.save_tsv(terms, self.file_path)
        msgbox.showinfo("TSV保存", f"{len(terms)}件の語句をTSV形式で保存しました。")

    def run_single_file_output(self):
        if not self.current_file_path:
            msgbox.showwarning("エラー", "先にPowerPointファイルをドロップまたは選択してください。")
            return
        if not self.data:
            msgbox.showwarning("エラー", "ルビを振る語句がありません。")
            return
        settings = self.read_settings_from_ui()
        warnings = []
        try:
            out_path, count = generate_ruby_pptx(
                self.current_file_path, dict(self.data), settings, warnings=warnings
            )
        except Exception as e:
            logging.error(f"ルビ付きPowerPoint出力に失敗しました: {e}")
            msgbox.showerror(
                "出力エラー",
                f"ルビ付きPowerPointの作成中にエラーが発生しました。\n"
                f"PowerPointが起動できるか、ファイルが他で開かれていないか確認してください。\n\n{e}"
            )
            return

        message = f"{Path(out_path).name} を保存しました。\nルビ {count} 件を配置しました。"
        if count == 0:
            message += "\n\n※ルビが1件も配置されていません。語句一覧と対象範囲を確認してください。"
        if warnings:
            message += "\n\n注意:\n" + "\n".join(dict.fromkeys(warnings[:10]))
        msgbox.showinfo("ルビ付きPowerPoint出力", message)

    def advance_to_next(self):
        if not self.pptx_files or self.current_index >= len(self.pptx_files):
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
        current_file = Path(self.pptx_files[self.current_index])
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
        warnings = []
        self.save_tsv(self.data, current_file)
        try:
            generate_ruby_pptx(current_file, dict(self.data), settings, warnings=warnings)
        except Exception as e:
            logging.error(f"{current_file} の処理に失敗しました: {e}")
            msgbox.showerror("保存失敗", f"{current_file.name} の処理に失敗しました。\n{e}")
        self.review_ready = False
        self.current_index += 1
        self.process_next_file()

    # ---------- 辞書 ----------
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
        tree = ttk.Treeview(edit_win, columns=("word", "reading"), show="headings", selectmode="browse")
        tree.heading("word", text="語句")
        tree.heading("reading", text="読み")
        tree.pack(fill="both", expand=True, padx=10, pady=10)
        for word, reading in self.override_dict.items():
            tree.insert("", "end", values=(word, reading))

        # 編集ポップアップ
        def edit_dict_item(event=None):
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
        tree.bind("<Double-1>", edit_dict_item)

        # 削除ボタン
        def delete_entry():
            selected = tree.focus()
            if not selected:
                return
            word, _ = tree.item(selected, "values")
            if not msgbox.askyesno("削除確認", f"「{word}」を辞書から削除しますか？"):
                return
            self.override_dict.pop(word, None)
            tree.delete(selected)
        tk.Button(edit_win, text="選択した語句を削除", command=delete_entry, fg="red").pack(pady=5)

        def save_and_close():
            try:
                with open("override.json", "w", encoding="utf-8") as f:
                    json.dump(self.override_dict, f, indent=2, ensure_ascii=False)
            except OSError as e:
                logging.error(f"override.json の保存に失敗しました: {e}")
                msgbox.showerror("辞書保存", f"override.json を保存できませんでした。\n{e}")
                return
            edit_win.destroy()
        tk.Button(edit_win, text="保存して閉じる", command=save_and_close).pack(pady=(0, 10))
        # ウィジェットを並べ終えてから中央へ寄せる（先に呼ぶと空のサイズで計算される）
        self.center_window_auto(edit_win)


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = RubyEditorApp(root)
    root.mainloop()
