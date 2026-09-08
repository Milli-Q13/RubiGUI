"""Word版の出力(.docx)を開いて「どこにルビが付いたか」を地名で照合する。

make_fixture_word.py が作ったフィクスチャを処理した結果に対して使う。
期待どおりなら「OK」、食い違ったら「NG」を出して終了コード1で終わる。

使い方
    python tests/check_word_output.py "出力（ルビ付き）\\word_fixture（ルビ）.docx"
"""
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# 本文(Content)の中にあり、ルビが付くべき場所
SHOULD_HAVE_RUBY = {
    "京都": "通常の段落",
    "大阪": "単純な表のセル",
    "名古屋": "セル結合のある表",
    "札幌": "入れ子の表",
    "仙台": "箇条書き",
    "広島": "段組み",
    "神戸": "同じ語句が複数回出てくる段落",
    "横浜": "全角スペースで字下げした段落",
}

# 本文(Content)に含まれない場所。ここにルビが付いていたら想定外
SHOULD_NOT_HAVE_RUBY = {
    "福岡": "テキストボックス",
    "熊本": "オートシェイプ",
    "長崎": "ワードアート",
    "鹿児島": "SmartArt",
    "沖縄": "グラフ",
    "青森": "ヘッダー",
    "秋田": "フッター",
    "岩手": "脚注",
    "山形": "コメント",
}

# 元からルビが振ってある語句。二重に振られていないかを見る
PRE_EXISTING_RUBY = {"奈良": "元からルビあり（二重ルビになっていないか）"}


def ruby_pairs(root):
    """(親文字, ルビ) の一覧を返す。"""
    pairs = []
    for ruby in root.iter(f"{W}ruby"):
        rt = ruby.find(f"{W}rt")
        rb = ruby.find(f"{W}rubyBase")
        base = "".join(t.text or "" for t in rb.iter(f"{W}t")) if rb is not None else ""
        reading = "".join(t.text or "" for t in rt.iter(f"{W}t")) if rt is not None else ""
        pairs.append((base, reading))
    return pairs


def main(path):
    path = Path(path)
    if not path.exists():
        print(f"ファイルがありません: {path}")
        return 1

    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))

    pairs = ruby_pairs(root)
    # ルビは1文字ずつ w:ruby に分かれることがあるので、親文字を連結した
    # 文字列に対して「その語句が含まれるか」で判定する
    ruby_base_text = "".join(base for base, _ in pairs)

    print(f"ファイル: {path.name}")
    print(f"ルビの総数: {len(pairs)}\n")

    failures = []

    print("--- ルビが付くべき場所 ---")
    for word, place in SHOULD_HAVE_RUBY.items():
        hit = word in ruby_base_text
        print(f"  {'OK ' if hit else 'NG '} {word}（{place}）")
        if not hit:
            failures.append(f"{word}（{place}）にルビが付いていない")

    print("\n--- ルビが付かないはずの場所 ---")
    for word, place in SHOULD_NOT_HAVE_RUBY.items():
        hit = word in ruby_base_text
        print(f"  {'NG ' if hit else 'OK '} {word}（{place}）")
        if hit:
            failures.append(f"{word}（{place}）に想定外のルビが付いた")

    print("\n--- 参考（判定はしない） ---")
    for word, note in PRE_EXISTING_RUBY.items():
        count = sum(1 for base, _ in pairs if word in base)
        readings = [r for b, r in pairs if word in b]
        print(f"  {word}: ルビ {count} 件 {readings}　… {note}")

    print("\n--- 付いたルビの一覧 ---")
    merged = []
    for base, reading in pairs:
        if merged and merged[-1][1] == reading:
            continue
        merged.append((base, reading))
    for base, reading in pairs:
        print(f"  {base} → {reading}")

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
    sys.exit(main(sys.argv[1]))
