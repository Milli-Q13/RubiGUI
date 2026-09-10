"""版番号とマクロ名が v3.1 / v1.3 に揃っていることの確認。

マクロ名は版ごとに改名する方針になっている。Python側の MACRO_NAME と
.bas 側のマクロ名がずれると「何も起きない（ルビが振られないだけ）」という
分かりにくい失敗になるため、両者が一致していることを機械的に確かめる。
"""
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
BAS = REPO / "RubiGUI_word_v3.1" / "RubiGUI_V31.bas"


def test_word_version(word_app):
    assert word_app.APP_VERSION == "3.1"


def test_ppt_version(ppt_app):
    assert ppt_app.APP_VERSION == "1.3"


def test_word_macro_name(word_app):
    assert word_app.MACRO_NAME == "InsertFuriganaFromTSV_V31"


def test_bas_module_name_matches_version():
    """.bas は CP932。UTF-8 で読むと文字化けするので明示する。"""
    text = BAS.read_text(encoding="cp932")
    assert 'Attribute VB_Name = "RubiGUI_V31"' in text


def test_bas_defines_the_macro_python_calls(word_app):
    text = BAS.read_text(encoding="cp932")
    assert f"Public Sub {word_app.MACRO_NAME}(" in text
