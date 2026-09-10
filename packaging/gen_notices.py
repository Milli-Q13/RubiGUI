"""THIRD-PARTY-NOTICES.txt を生成する。

RubiGUI の exe には複数のオープンソースが同梱されている。いずれも
コピーレフトではないため RubiGUI 自体のソース公開義務は生じないが、
**ライセンス本文と告知を配布物に添える義務**はある。手書きすると
ライブラリを更新したときに更新を忘れるので、生成できるものは生成する。

★すべてを自動収集できるわけではない。
  ・SudachiPy は配布物にライセンスファイルを同梱していない
  ・Python 本体と Tcl/Tk は pip パッケージではない
  これらは packaging/licenses/ に手置きし、収集分と結合する。
  どちらにも無い同梱物があれば、告知漏れのまま配布しないよう停止する。
"""
import importlib.metadata as md
from pathlib import Path

HERE = Path(__file__).resolve().parent
MANUAL_DIR = HERE / "licenses"

# 配布物に同梱されるもの。pip パッケージ名か、手置きファイルの名前（拡張子なし）。
BUNDLED = [
    "SudachiPy",
    "SudachiDict-full",
    "tkinterdnd2",
    "jaconv",
    "pywin32",
    "pyinstaller",
    "Python",
    "TclTk",
]

# 収集対象とみなすファイル名の目印。LEGAL は SudachiDict の告知ファイル。
_MARKERS = ("LICENSE", "NOTICE", "LEGAL", "COPYING")

_HEADER = """\
RubiGUI 第三者ソフトウェアのライセンス告知
============================================================

RubiGUI には次のオープンソースソフトウェアが含まれています。
それぞれのライセンス本文と告知を以下に収録します。

このファイルは packaging/gen_notices.py が自動生成しています。
手で編集しないでください。

"""


def _read_text(path):
    """ライセンスファイルは配布元によって文字コードが異なるので順に試す。

    utf-8 → cp932 の順で正しくデコードできればそれを使う。どちらも失敗した
    場合、以前は latin-1 にフォールバックしていたが、latin-1 は256バイト値
    すべてに対応する全単射のため常に「それらしい」文字列を返してしまい、
    実際には文字コードを取り違えたまま気づかず誤った本文を採用する恐れが
    あった（このファイルは法的義務を果たすためのものなので、誤った本文を
    静かに採用するより、ビルドを止めて人間に気づかせる方が安全）。そこで
    最終手段として errors="replace" で読み、U+FFFD（置換文字）が1つでも
    出た場合は文字コードを特定できなかったとみなして SystemExit で止める。
    """
    for encoding in ("utf-8", "cp932"):
        try:
            return Path(path).read_text(encoding=encoding)
        except (UnicodeDecodeError, LookupError):
            continue

    text = Path(path).read_text(encoding="utf-8", errors="replace")
    if "�" in text:
        raise SystemExit(
            f"文字コードを特定できませんでした: {path}\n"
            "utf-8 / cp932 のいずれでも正しく読めません。"
            "実際の文字コード（EUC-JP・UTF-16 等）を確認し、"
            "必要なら _read_text の対応エンコーディングに追加してください。"
        )
    return text


def _collect_from_package(name):
    """pip パッケージからライセンスファイルを集める。無ければ空リスト。"""
    try:
        dist = md.distribution(name)
        files = md.files(name) or []
    except md.PackageNotFoundError:
        return []

    found = []
    for f in files:
        # パス全体ではなく、ファイル自身の名前（basename）だけを見る。
        # そうしないと "licenses/" のようなディレクトリ名に含まれる語で
        # そのディレクトリ配下の無関係なファイルまで拾ってしまう一方、
        # "LICENSE" 等の語が入ったディレクトリ配下にある jaconv・pyinstaller の
        # 本物のライセンスファイルは basename 判定でも問題なく拾える。
        basename_upper = Path(str(f)).name.upper()
        if any(m in basename_upper for m in _MARKERS):
            path = Path(dist.locate_file(f))
            if path.is_file():
                found.append((str(f), _read_text(path)))
    return found


def _collect_from_manual(name):
    """packaging/licenses/ に手置きされた本文を読む。無ければ空リスト。"""
    path = MANUAL_DIR / f"{name}.txt"
    if path.is_file():
        return [(f"licenses/{path.name}（手置き）", _read_text(path))]
    return []


def _version(name):
    try:
        return md.version(name)
    except md.PackageNotFoundError:
        return None


def generate(out_path):
    """THIRD-PARTY-NOTICES.txt を書き出し、そのパスを返す。"""
    out_path = Path(out_path)
    parts = [_HEADER]
    missing = []

    for name in BUNDLED:
        # package 側と手置き側の両方を集めて結合する（どちらかで打ち切らない）。
        # 手置きファイルは「package から本物のライセンスが取れない」からこそ
        # 用意されている（例: SudachiPy）。以前は `or` で package 側の結果が
        # 1件でもあれば手置きを無視していたため、将来 package が名前だけ似た
        # 無関係なファイルを同梱するようになった場合、正しい手置き本文が
        # 黙って握りつぶされる恐れがあった。
        entries = _collect_from_package(name) + _collect_from_manual(name)
        if not entries:
            missing.append(name)
            continue

        version = _version(name)
        title = f"{name} {version}" if version else name
        parts.append("=" * 60)
        parts.append(title)
        parts.append("=" * 60)
        parts.append("")
        for origin, body in entries:
            parts.append(f"--- {origin} ---")
            parts.append("")
            parts.append(body.rstrip())
            parts.append("")

    if missing:
        raise SystemExit(
            "告知を用意できない同梱物があります: " + "、".join(missing) + "\n"
            "pip パッケージ名が正しいか確認するか、"
            f"{MANUAL_DIR} に <名前>.txt としてライセンス本文を置いてください。"
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(parts), encoding="utf-8")
    return out_path


if __name__ == "__main__":
    written = generate(HERE / "THIRD-PARTY-NOTICES.txt")
    print(f"generated: {written} ({written.stat().st_size} bytes)")
