"""Word版の動作確認用フィクスチャ(.docx)をCOMで生成する。

なぜスクリプトで作るのか
    手作りの確認用ファイルは「入れ忘れ」が起きる。実際 v2.1 では
    表を入れたファイルで確認しておらず、「表の中を検索すると空回りする」
    不具合を出荷まで見逃した。ここに一通り並べておけば、次の版でも
    同じファイルを再生成して同じ観点で確認できる。

仕掛け
    場所ごとに「そこにしか出てこない地名」を置いてある。処理後の
    出力を check_word_output.py で調べると、どの場所にルビが付いて
    どの場所に付かなかったかが地名で判別できる。

    ルビが付くべき場所   : 京都 大阪 名古屋 札幌 仙台 広島 神戸 横浜 奈良
    ルビが付かないはず   : 福岡 熊本 長崎 鹿児島 沖縄 青森 秋田 岩手 山形
    （後者は本文(Content)に含まれない場所＝テキストボックス・図形・
      ヘッダー・脚注・コメントなど。Python側の抽出にも出てこない）

使い方
    python tests/make_fixture_word.py [出力先フォルダ]
"""
import sys
from pathlib import Path

import win32com.client

# Word の定数（参照設定なしで動かすため直接書く）
WD_FORMAT_DOCX = 16
WD_SEEK_MAIN = 0
WD_SEEK_PRIMARY_HEADER = 1
WD_SEEK_PRIMARY_FOOTER = 4
WD_COLLAPSE_END = 0
WD_COLLAPSE_START = 1
WD_LIST_BULLET = 2
WD_ALIGN_ROW_CENTER = 1
WD_TEXT_ORIENT_VERTICAL_FAR_EAST = 3
MSO_TEXT_ORIENT_HORIZONTAL = 1
WD_FIELD_PAGE = 33
WD_FIELD_TOC = 26


def add_para(doc, text, style=None):
    """本文の末尾に段落を1つ足す。"""
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    rng.InsertParagraphAfter()
    rng = doc.Content
    rng.Collapse(WD_COLLAPSE_END)
    rng.InsertAfter(text)
    if style is not None:
        try:
            doc.Paragraphs(doc.Paragraphs.Count).Style = style
        except Exception as e:
            print(f"  （スタイル {style} の適用に失敗: {e}）")
    return doc.Paragraphs(doc.Paragraphs.Count)


def build(out_path):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = True          # PhoneticGuide 同様、可視の方が安定する
    word.DisplayAlerts = False
    doc = word.Documents.Add()

    try:
        # ---- 1. 通常の段落（ルビが付くべき） ----
        rng = doc.Content
        rng.Text = "確認用の文書です。通常の段落に京都という語句を置きます。"

        # ---- 2. 同じ語句が複数回出る（初回のみ／すべて の切り分け用） ----
        add_para(doc, "神戸は一度目です。神戸は二度目です。神戸は三度目です。")

        # ---- 3. 全角スペースで字下げした段落（空白語句バグの再現用） ----
        add_para(doc, "　　　　横浜のまえに全角スペースを並べています。")

        # ---- 4. 箇条書き ----
        p = add_para(doc, "仙台を箇条書きに入れます。")
        try:
            p.Range.ListFormat.ApplyListTemplate(
                word.ListGalleries(WD_LIST_BULLET).ListTemplates(1))
        except Exception as e:
            print(f"  （箇条書きの適用に失敗: {e}）")

        # ---- 5. 単純な表（今回の不具合の本丸） ----
        add_para(doc, "")
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
        t1 = doc.Tables.Add(rng, 1, 1)
        t1.Cell(1, 1).Range.Text = "大阪を単純な表のセルに入れます。"

        # ---- 6. セル結合のある表 ----
        add_para(doc, "")
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
        t2 = doc.Tables.Add(rng, 2, 2)
        t2.Cell(1, 1).Range.Text = "名古屋"
        t2.Cell(1, 2).Range.Text = "結合前"
        t2.Cell(2, 1).Range.Text = "下段左"
        t2.Cell(2, 2).Range.Text = "下段右"
        try:
            t2.Cell(1, 1).Merge(t2.Cell(1, 2))
        except Exception as e:
            print(f"  （セル結合に失敗: {e}）")

        # ---- 7. 入れ子の表（表のセルの中に表） ----
        add_para(doc, "")
        rng = doc.Content
        rng.Collapse(WD_COLLAPSE_END)
        t3 = doc.Tables.Add(rng, 1, 1)
        try:
            # ★セルの末尾へ Collapse すると「セル終端記号」の上に来てしまい、
            #   Tables.Add が「表の最終行が参照されている」と拒否する。
            #   先頭へ Collapse すること。
            inner_rng = t3.Cell(1, 1).Range
            inner_rng.Collapse(WD_COLLAPSE_START)
            t4 = doc.Tables.Add(inner_rng, 1, 1)
            t4.Cell(1, 1).Range.Text = "札幌を入れ子の表に入れます。"
        except Exception as e:
            print(f"  （入れ子の表の作成に失敗: {e}）")

        # ---- 8. 段組み ----
        add_para(doc, "")
        add_para(doc, "広島を段組みの段落に入れます。" * 3)
        try:
            doc.Paragraphs(doc.Paragraphs.Count).Range.PageSetup.TextColumns.SetCount(2)
        except Exception as e:
            print(f"  （段組みの適用に失敗: {e}）")

        # ---- 9. 既にルビが振ってある語句（二重ルビの確認用） ----
        add_para(doc, "奈良には最初からルビが振ってあります。")
        try:
            target = doc.Content.Find
            target.Text = "奈良"
            target.Forward = True
            target.Wrap = 0
            if target.Execute():
                target.Parent.PhoneticGuide(Text="なら", Alignment=0, Raise=0, FontSize=5)
        except Exception as e:
            print(f"  （既存ルビの設定に失敗: {e}）")

        # ==== ここから下は「ルビが付かないはず」の場所 ====

        # ---- 10. テキストボックス ----
        try:
            box = doc.Shapes.AddTextbox(MSO_TEXT_ORIENT_HORIZONTAL, 300, 60, 180, 50)
            box.TextFrame.TextRange.Text = "福岡はテキストボックスの中です。"
        except Exception as e:
            print(f"  （テキストボックスの作成に失敗: {e}）")

        # ---- 11. オートシェイプ（四角形）内のテキスト ----
        try:
            shp = doc.Shapes.AddShape(1, 300, 130, 180, 50)  # 1 = 四角形
            shp.TextFrame.TextRange.Text = "熊本は図形の中です。"
        except Exception as e:
            print(f"  （オートシェイプの作成に失敗: {e}）")

        # ---- 12. ワードアート ----
        try:
            doc.Shapes.AddTextEffect(
                PresetTextEffect=0, Text="長崎はワードアートです",
                FontName="MS Gothic", FontSize=20,
                FontBold=0, FontItalic=0, Left=300, Top=200)
        except Exception as e:
            print(f"  （ワードアートの作成に失敗: {e}）")

        # ---- 13. SmartArt ----
        try:
            layout = word.Application.SmartArtLayouts(1)
            smart = doc.Shapes.AddSmartArt(layout, 300, 270, 200, 120)
            smart.SmartArt.AllNodes(1).TextFrame2.TextRange.Text = "鹿児島はSmartArtです"
        except Exception as e:
            print(f"  （SmartArtの作成に失敗: {e}）")

        # ---- 14. グラフ ----
        try:
            chart_shape = doc.Shapes.AddChart2(-1, 51, 300, 410, 220, 150)
            chart_shape.Chart.ChartTitle.Text = "沖縄のグラフ"
        except Exception as e:
            print(f"  （グラフの作成に失敗: {e}）")

        # ---- 15. ヘッダー／フッター ----
        try:
            view = doc.Windows(1).ActivePane.View
            section = doc.Sections(1)
            section.Headers(1).Range.Text = "青森はヘッダーです"
            section.Footers(1).Range.Text = "秋田はフッターです"
            view.SeekView = WD_SEEK_MAIN
        except Exception as e:
            print(f"  （ヘッダー／フッターの設定に失敗: {e}）")

        # ---- 16. 脚注 ----
        try:
            rng = doc.Content
            rng.Collapse(WD_COLLAPSE_END)
            doc.Footnotes.Add(Range=rng, Text="岩手は脚注です")
        except Exception as e:
            print(f"  （脚注の作成に失敗: {e}）")

        # ---- 17. コメント ----
        try:
            first_para = doc.Paragraphs(1).Range
            doc.Comments.Add(Range=first_para, Text="山形はコメントです")
        except Exception as e:
            print(f"  （コメントの作成に失敗: {e}）")

        # ---- 18. フィールド（ページ番号・目次） ----
        try:
            rng = doc.Content
            rng.Collapse(WD_COLLAPSE_END)
            rng.InsertParagraphAfter()
            rng = doc.Content
            rng.Collapse(WD_COLLAPSE_END)
            doc.Fields.Add(Range=rng, Type=WD_FIELD_PAGE)
        except Exception as e:
            print(f"  （ページ番号フィールドの作成に失敗: {e}）")

        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        doc.SaveAs2(str(out_path), FileFormat=WD_FORMAT_DOCX)
        print(f"作成しました: {out_path}")
    finally:
        try:
            doc.Close(SaveChanges=0)
        except Exception:
            pass
        try:
            word.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    target_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "fixtures"
    build(target_dir / "word_fixture.docx")
