"""配布物を組み立てる。

使い方（リポジトリのどこからでも可）:
    python packaging/build_exe.py          # 先に exe をビルドしておく
    python packaging/make_dist.py          # 配布物を組んで zip 化する

★方針：入れるものを明示的に列挙する（許可リスト方式）。
  以前は版フォルダをそのまま固めていたが、この方式では rubigui.log が
  配布物に入る。ログには実在の教材名とローカルパス（ユーザー名を含む）が
  記録されている。.gitignore は zip の作成には効かない。
  除外リスト方式にすると新しい種類のファイルが増えたときに漏れるので、
  許可リストにする。

★辞書（359.8MB）は版フォルダに複製していない。--dic で場所を指定する。
"""
import argparse
import hashlib
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path

import gen_notices

HERE = Path(__file__).resolve().parent
REPO = HERE.parent

WORD_DIR = REPO / "RubiGUI_word_v3.1"
PPT_DIR = REPO / "RubiGUI_ppt_v1.3"
DOCS_DIR = REPO / "docs"

BUNDLE_NAME = "RubiGUI_2026-09"
DEFAULT_DIC = REPO / "RubiGUI_word_v3.0" / "system_full.dic"


@dataclass(frozen=True)
class Entry:
    src: Path
    dest: str


# 配布物に入れるものの全量。ここに無いものは入らない。
ALLOWLIST = [
    Entry(WORD_DIR / "RubiGUI_Word_v3.1.exe", "RubiGUI_Word_v3.1.exe"),
    Entry(PPT_DIR / "RubiGUI_PPT_v1.3.exe", "RubiGUI_PPT_v1.3.exe"),
    Entry(WORD_DIR / "RubiGUI_V31.bas", "RubiGUI_V31.bas"),
    Entry(WORD_DIR / "sudachi.json", "sudachi.json"),
    Entry(WORD_DIR / "override.json", "override.json"),
    Entry(WORD_DIR / "ruby_settings.json", "ruby_settings.json"),
    Entry(WORD_DIR / "readme.txt", "readme_Word.txt"),
    Entry(PPT_DIR / "readme.txt", "readme_PPT.txt"),
    Entry(DOCS_DIR / "はじめにお読みください.pdf", "はじめにお読みください.pdf"),
    Entry(HERE / "LICENSE.txt", "LICENSE.txt"),
]

# 混入していたら停止するもの。
FORBIDDEN_SUFFIXES = (".py", ".pyc", ".log", ".spec")
FORBIDDEN_NAMES = ("requirements.txt", "__pycache__")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def check_no_forbidden(folder):
    """配布物に入ってはいけないものが無いか調べ、あれば停止する。"""
    folder = Path(folder)
    hits = []
    for path in folder.rglob("*"):
        if path.name.lower() in FORBIDDEN_NAMES or path.suffix.lower() in FORBIDDEN_SUFFIXES:
            hits.append(str(path.relative_to(folder)))
    if hits:
        raise SystemExit(
            "配布物に入ってはいけないファイルが見つかりました:\n  "
            + "\n  ".join(hits)
            + "\nALLOWLIST を確認してください。"
        )


def build(out_dir, dic):
    """許可リストのファイルだけを集めて配布フォルダを組む。

    コピーを始める前に ALLOWLIST と辞書の存在を確認する（--dic の
    指定ミスはリリース当日に起こりやすい操作ミス）。コピー開始後に
    何かが失敗した場合（辞書コピーの失敗、gen_notices.generate() の
    例外、ディスクフル等）は、途中まで組み上がった out_dir を丸ごと
    削除してから例外を再送出する。中途半端な配布フォルダを残すと、
    オペレーターが目視で「できている」と誤認して手で zip 化してしまい、
    THIRD-PARTY-NOTICES.txt や辞書を欠いた配布物が外部に出かねない。
    """
    out_dir = Path(out_dir)
    dic = Path(dic)
    if out_dir.exists():
        shutil.rmtree(out_dir)

    missing = [e.src for e in ALLOWLIST if not e.src.is_file()]
    if missing:
        raise SystemExit(
            "配布物に入れるファイルが見つかりません:\n  "
            + "\n  ".join(str(m) for m in missing)
            + "\nexe が未ビルドなら packaging/build_exe.py を先に実行してください。"
        )

    if not dic.is_file():
        raise SystemExit(f"辞書が見つかりません: {dic}\n--dic で場所を指定してください。")

    out_dir.mkdir(parents=True)
    try:
        for entry in ALLOWLIST:
            shutil.copy2(entry.src, out_dir / entry.dest)

        print(f"辞書をコピー中（{dic.stat().st_size / 1e6:.1f} MB）...", flush=True)
        shutil.copy2(dic, out_dir / "system_full.dic")

        gen_notices.generate(out_dir / "THIRD-PARTY-NOTICES.txt")

        check_no_forbidden(out_dir)
    except BaseException:
        shutil.rmtree(out_dir, ignore_errors=True)
        raise

    return out_dir


def make_zip(folder):
    """配布フォルダを zip に固める。展開すると folder と同じ名前になる。"""
    folder = Path(folder)
    zip_path = folder.parent / f"{folder.name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    print("zip を作成中...", flush=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(folder.rglob("*")):
            if path.is_file():
                zf.write(path, Path(folder.name) / path.relative_to(folder))
    return zip_path


def main():
    parser = argparse.ArgumentParser(description="RubiGUI の配布物を組み立てる")
    parser.add_argument("--dic", default=str(DEFAULT_DIC), help="system_full.dic の場所")
    parser.add_argument("--out", default=str(REPO / "dist" / BUNDLE_NAME),
                        help="配布フォルダの出力先")
    args = parser.parse_args()

    folder = build(args.out, args.dic)
    zip_path = make_zip(folder)
    digest = sha256(zip_path)

    print("\n=== 組み立て結果 ===")
    print(f"  フォルダ : {folder}")
    print(f"  zip      : {zip_path}  ({zip_path.stat().st_size / 1e6:.1f} MB)")
    print(f"  SHA-256  : {digest}")
    print("\nこの SHA-256 をリリースノートに記載してください。")

    (zip_path.parent / f"{zip_path.name}.sha256.txt").write_text(
        f"{digest}  {zip_path.name}\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
