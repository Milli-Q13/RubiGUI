"""PPT版の出力(.pptx)を開いて「どこにルビが付いたか」を照合する。

PPT版のルビは親文字の上に重ねた別シェイプ。そのシェイプの文字＝読みを
集めて、期待する読みが出ているかで判定する。

★ルビ用シェイプの見分け方に注意
  最初は「名前が RUBIGUI_RUBY_ で始まるシェイプ」で拾おうとしたが、
  それでは取りこぼす。図形名を付けているのはグループ化する経路だけで、
  プレースホルダ（PowerPointの仕様上グループ化できない）に付けたルビは
  既定名（TextBox N）のまま残るため。
  アプリ自身が目印にしているのはシェイプのタグ RUBIGUI_ROLE=RUBY なので、
  こちらで判定する。タグは ppt/tags/tagN.xml にあり、スライドの
  .rels 経由で参照されている。

使い方
    python tests/check_ppt_output.py "出力（ルビ付き）\\ppt_fixture（ルビ）.pptx"
"""
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
NS_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

TAG_ROLE = "RUBIGUI_ROLE"
ROLE_RUBY = "RUBY"

# ルビが付くべき場所（PPT版の対象＝テキストボックスとタイトル以外のプレースホルダ）
SHOULD_HAVE_RUBY = {
    "京都": ("きょうと", "本文プレースホルダ（第二階層）"),
    "大阪": ("おおさか", "縦書きテキストボックス"),
    "名古屋": ("なごや", "本文プレースホルダ"),
    "仙台": ("せんだい", "複数行のテキストボックス"),
    "広島": ("ひろしま", "1行だけのテキストボックス"),
    "神戸": ("こうべ", "行間1.5が設定済みのテキストボックス"),
    "横浜": ("よこはま", "自動調整つきテキストボックス"),
    "奈良": ("なら", "空欄のある本文テキストボックス"),
}

# 対象外の場所（オートシェイプ・表・グループ内・SmartArt・グラフ・
# ワードアート・ノート・回転図形）。ここにルビが付いたら想定外。
SHOULD_NOT_HAVE_RUBY = {
    "福岡": ("ふくおか", "オートシェイプ"),
    "熊本": ("くまもと", "表"),
    "長崎": ("ながさき", "グループ化された図形の中"),
    "鹿児島": ("かごしま", "SmartArt"),
    "沖縄": ("おきなわ", "グラフ"),
    "青森": ("あおもり", "ワードアート"),
    "秋田": ("あきた", "ノート"),
    "岩手": ("いわて", "回転した図形"),
}

# タイトルは「タイトルもルビ対象にする」設定のときだけ付く
TITLE_WORD = ("札幌", "さっぽろ", "タイトルプレースホルダ")


def _tag_targets(z, slide_name):
    """スライドの .rels から「rId → tagN.xml のパス」を作る。"""
    rels_name = slide_name.replace("ppt/slides/", "ppt/slides/_rels/") + ".rels"
    mapping = {}
    try:
        rels = ET.fromstring(z.read(rels_name))
    except KeyError:
        return mapping
    for rel in rels.iter(f"{NS_REL}Relationship"):
        target = rel.get("Target") or ""
        if "tags/" in target:
            mapping[rel.get("Id")] = "ppt/" + target.split("../", 1)[-1]
    return mapping


def _is_ruby_shape(z, sp, tag_targets):
    """シェイプに RUBIGUI_ROLE=RUBY のタグが付いているか。"""
    for tags in sp.iter(f"{NS_P}tags"):
        path = tag_targets.get(tags.get(f"{NS_R}id"))
        if not path:
            continue
        try:
            tag_root = ET.fromstring(z.read(path))
        except KeyError:
            continue
        for tag in tag_root.iter(f"{NS_P}tag"):
            if tag.get("name") == TAG_ROLE and tag.get("val") == ROLE_RUBY:
                return True
    return False


def ruby_texts(path):
    """出力ファイル内の、ルビ用シェイプの文字をすべて集める。"""
    found = []
    with zipfile.ZipFile(path) as z:
        slide_names = sorted(
            (n for n in z.namelist()
             if re.match(r"ppt/slides/slide\d+\.xml$", n)),
            key=lambda n: int(re.search(r"(\d+)", n).group(1)),
        )
        for name in slide_names:
            root = ET.fromstring(z.read(name))
            index = int(re.search(r"(\d+)", name).group(1))
            tag_targets = _tag_targets(z, name)
            for sp in root.iter(f"{NS_P}sp"):
                if not _is_ruby_shape(z, sp, tag_targets):
                    continue
                nv = sp.find(f"{NS_P}nvSpPr")
                c_nv = nv.find(f"{NS_P}cNvPr") if nv is not None else None
                shape_name = (c_nv.get("name") or "") if c_nv is not None else ""
                text = "".join(t.text or "" for t in sp.iter(f"{NS_A}t"))
                found.append((index, shape_name, text))
    return found


def main(path, include_title=False):
    path = Path(path)
    if not path.exists():
        print(f"ファイルがありません: {path}")
        return 1

    found = ruby_texts(path)
    all_ruby = "".join(text for _, _, text in found)

    print(f"ファイル: {path.name}")
    print(f"ルビ用シェイプの数: {len(found)}\n")

    failures = []

    print("--- ルビが付くべき場所 ---")
    for word, (reading, place) in SHOULD_HAVE_RUBY.items():
        hit = reading in all_ruby
        print(f"  {'OK ' if hit else 'NG '} {word}（{place}） 読み={reading}")
        if not hit:
            failures.append(f"{word}（{place}）にルビが付いていない")

    print("\n--- ルビが付かないはずの場所 ---")
    for word, (reading, place) in SHOULD_NOT_HAVE_RUBY.items():
        hit = reading in all_ruby
        print(f"  {'NG ' if hit else 'OK '} {word}（{place}） 読み={reading}")
        if hit:
            failures.append(f"{word}（{place}）に想定外のルビが付いた")

    word, reading, place = TITLE_WORD
    hit = reading in all_ruby
    expected = include_title
    mark = "OK " if hit == expected else "NG "
    print(f"\n--- タイトル（include_title={include_title}） ---")
    print(f"  {mark} {word}（{place}） 読み={reading} → {'付いた' if hit else '付かなかった'}")
    if hit != expected:
        failures.append(
            f"{word}（{place}）: include_title={include_title} なのに"
            f"{'付いた' if hit else '付かなかった'}"
        )

    print("\n--- スライドごとのルビ ---")
    by_slide = {}
    for index, _, text in found:
        by_slide.setdefault(index, []).append(text)
    for index in sorted(by_slide):
        print(f"  スライド{index}: {' '.join(by_slide[index])}")

    if failures:
        print("\n★NG:")
        for f in failures:
            print(f"  - {f}")
        return 1
    print("\nすべて期待どおりです。")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    include_title = len(sys.argv) > 2 and sys.argv[2].lower() in ("1", "true", "yes")
    sys.exit(main(sys.argv[1], include_title))
