"""PPT版の動作確認用フィクスチャ(.pptx)をCOMで生成する。

なぜスクリプトで作るのか / 仕掛けは make_fixture_word.py と同じ。
場所ごとに「そこにしか出てこない地名」を置いてある。

    ルビが付くべき場所 : 京都 大阪 名古屋 札幌 仙台 広島 神戸 横浜 奈良
    付かないはず       : 福岡 熊本 長崎 鹿児島 沖縄 青森 秋田 岩手

PPT版の対象は「通常のテキストボックス」と「タイトル以外のプレースホルダ」だけ。
オートシェイプ内テキスト・表・グループ内シェイプ・図・SmartArt・グラフは
意図的に対象外にしてある（is_target_shape / _xml_is_target）。
この2つの判定が食い違うと「一覧に出るのにルビが振られない語句」が
生まれるので、全種類を1つのファイルに置いて突き合わせられるようにしてある。

使い方
    python tests/make_fixture_ppt.py [出力先フォルダ]
"""
import sys
from pathlib import Path

import win32com.client

PP_LAYOUT_BLANK = 12
PP_LAYOUT_TEXT = 2
MSO_TEXT_ORIENT_HORIZONTAL = 1
MSO_TEXT_ORIENT_VERTICAL_FAR_EAST = 3
MSO_SHAPE_RECTANGLE = 1
MSO_TRUE = -1
MSO_FALSE = 0
PP_SAVE_AS_DEFAULT = 24          # ppSaveAsDefault (.pptx)
PP_EFFECT_APPEAR = 1             # msoAnimEffectAppear（0はCustomなので1が正しい）
MSO_ANIM_TRIGGER_ON_CLICK = 1


def add_textbox(slide, left, top, width, height, text,
                orient=MSO_TEXT_ORIENT_HORIZONTAL, name=None):
    box = slide.Shapes.AddTextbox(orient, left, top, width, height)
    box.TextFrame.TextRange.Text = text
    box.TextFrame.TextRange.Font.Size = 18
    if name:
        try:
            box.Name = name
        except Exception:
            pass
    return box


def build(out_path):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True                # 非表示だと Bound* が取れないので必ず可視
    pres = ppt.Presentations.Add()

    try:
        # ================= スライド1: ルビが付くべき場所 =================
        s1 = pres.Slides.Add(1, PP_LAYOUT_TEXT)
        try:
            s1.Shapes.Item(1).TextFrame.TextRange.Text = "札幌はタイトルです"
        except Exception as e:
            print(f"  （タイトルの設定に失敗: {e}）")
        try:
            body = s1.Shapes.Item(2).TextFrame.TextRange
            body.Text = "名古屋は本文プレースホルダです\r第二階層の京都です"
        except Exception as e:
            print(f"  （本文プレースホルダの設定に失敗: {e}）")

        # 縦書きテキストボックス
        try:
            add_textbox(s1, 480, 120, 60, 220, "大阪は縦書きです",
                        orient=MSO_TEXT_ORIENT_VERTICAL_FAR_EAST, name="tb_vertical")
        except Exception as e:
            print(f"  （縦書きテキストボックスの作成に失敗: {e}）")

        # ================= スライド2: 行間まわり =================
        s2 = pres.Slides.Add(2, PP_LAYOUT_BLANK)

        # 複数行（行間が広がる＝ズレ警告の対象になるはず）
        add_textbox(s2, 40, 40, 300, 120,
                    "仙台をふくむ複数行の文章です。\r二行目があるので行間が広がります。",
                    name="tb_multiline")

        # 1行だけ（行間を触らないはず）
        add_textbox(s2, 40, 200, 300, 40, "広島は一行だけです", name="tb_singleline")

        # 既に行間1.5が設定済み（1.5×1.5＝2.25 にならないことの回帰確認）
        try:
            box = add_textbox(s2, 40, 270, 300, 120,
                              "神戸は行間1.5が設定済みです。\n二行目があります。",
                              name="tb_prespaced")
            pf = box.TextFrame.TextRange.Paragraphs().ParagraphFormat
            pf.LineRuleWithin = MSO_TRUE
            pf.SpaceWithin = 1.5
        except Exception as e:
            print(f"  （行間1.5の設定に失敗: {e}）")

        # 自動調整（縮小型）が効いたボックス
        try:
            box = add_textbox(s2, 400, 40, 200, 60,
                              "横浜は自動調整つきの長めの文章です。", name="tb_autosize")
            box.TextFrame2.AutoSize = 2      # msoAutoSizeTextToFitShape
        except Exception as e:
            print(f"  （自動調整の設定に失敗: {e}）")

        # ================= スライド3: 教員の運用（空欄＋答え） =================
        s3 = pres.Slides.Add(3, PP_LAYOUT_BLANK)
        add_textbox(s3, 40, 60, 500, 160,
                    "古代文明のはじまりについて。\r"
                    "大きな川の周辺で（　　　　　　）が発達しました。\r"
                    "これが奈良時代の話につながります。",
                    name="tb_body_with_blank")
        # 空欄の上に重ねる「答え」ボックス。アニメーションで表示する運用。
        try:
            answer = add_textbox(s3, 250, 100, 140, 36, "農業", name="tb_answer")
            effect = s3.TimeLine.MainSequence.AddEffect(
                answer, PP_EFFECT_APPEAR, 0, MSO_ANIM_TRIGGER_ON_CLICK)
        except Exception as e:
            print(f"  （答えボックス／アニメーションの設定に失敗: {e}）")

        # ================= スライド4: ルビが付かないはずの場所 =================
        s4 = pres.Slides.Add(4, PP_LAYOUT_BLANK)

        # オートシェイプ（四角形）内テキスト
        try:
            shp = s4.Shapes.AddShape(MSO_SHAPE_RECTANGLE, 40, 40, 200, 60)
            shp.TextFrame.TextRange.Text = "福岡は図形の中です"
        except Exception as e:
            print(f"  （オートシェイプの作成に失敗: {e}）")

        # 表
        try:
            tbl = s4.Shapes.AddTable(2, 2, 40, 120, 300, 80)
            tbl.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "熊本"
            tbl.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "表の中"
        except Exception as e:
            print(f"  （表の作成に失敗: {e}）")

        # グループ化された図形の中のテキスト
        try:
            a = s4.Shapes.AddTextbox(MSO_TEXT_ORIENT_HORIZONTAL, 40, 230, 140, 40)
            a.TextFrame.TextRange.Text = "長崎はグループ内です"
            b = s4.Shapes.AddTextbox(MSO_TEXT_ORIENT_HORIZONTAL, 200, 230, 140, 40)
            b.TextFrame.TextRange.Text = "グループの相方"
            a.Name, b.Name = "grp_a", "grp_b"
            s4.Shapes.Range(["grp_a", "grp_b"]).Group()
        except Exception as e:
            print(f"  （グループ化に失敗: {e}）")

        # SmartArt
        try:
            layout = ppt.SmartArtLayouts(1)
            smart = s4.Shapes.AddSmartArt(layout, 380, 40, 220, 120)
            smart.SmartArt.AllNodes(1).TextFrame2.TextRange.Text = "鹿児島はSmartArtです"
        except Exception as e:
            print(f"  （SmartArtの作成に失敗: {e}）")

        # グラフ
        try:
            chart_shape = s4.Shapes.AddChart2(-1, 51, 380, 180, 220, 150)
            # ★HasTitle を立ててからでないと ChartTitle は存在しない
            chart_shape.Chart.HasTitle = True
            chart_shape.Chart.ChartTitle.Text = "沖縄のグラフ"
        except Exception as e:
            print(f"  （グラフの作成に失敗: {e}）")

        # ワードアート
        try:
            s4.Shapes.AddTextEffect(
                PresetTextEffect=0, Text="青森はワードアートです",
                FontName="MS Gothic", FontSize=20,
                FontBold=MSO_FALSE, FontItalic=MSO_FALSE, Left=40, Top=300)
        except Exception as e:
            print(f"  （ワードアートの作成に失敗: {e}）")

        # 回転した図形（v1.1から「回転している図形は対象外」の警告が出る仕様）
        try:
            rot = s4.Shapes.AddTextbox(MSO_TEXT_ORIENT_HORIZONTAL, 380, 350, 200, 40)
            rot.TextFrame.TextRange.Text = "岩手は回転しています"
            rot.Rotation = 15
        except Exception as e:
            print(f"  （回転図形の作成に失敗: {e}）")

        # ノート
        try:
            s4.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = "秋田はノートです"
        except Exception as e:
            print(f"  （ノートの設定に失敗: {e}）")

        # PowerPointが最初から持っている空スライドを消す
        try:
            if pres.Slides.Count > 4:
                pres.Slides.Item(pres.Slides.Count).Delete()
        except Exception as e:
            print(f"  （余分なスライドの削除に失敗: {e}）")

        out_path = Path(out_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        pres.SaveAs(str(out_path), PP_SAVE_AS_DEFAULT)
        print(f"作成しました: {out_path}（スライド {pres.Slides.Count} 枚）")
    finally:
        try:
            pres.Close()
        except Exception:
            pass
        try:
            ppt.Quit()
        except Exception:
            pass


if __name__ == "__main__":
    target_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "fixtures"
    build(target_dir / "ppt_fixture.pptx")
