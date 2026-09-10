"""起動時に版番号がログへ残ることの確認。

不具合報告ではログを添付してもらう。版がログから分かれば、報告者が
版番号を書き間違えても調査できる。
"""
import logging
import re
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
WORD_APP = REPO / "RubiGUI_word_v3.1" / "RubiGUI_V3.1.py"
PPT_APP = REPO / "RubiGUI_ppt_v1.3" / "RubiGUI_PPT_V1.3.py"


def test_word_logs_its_version(word_app, caplog):
    with caplog.at_level(logging.INFO):
        word_app.log_startup_banner()
    assert "3.1" in caplog.text
    assert "Word" in caplog.text


def test_ppt_logs_its_version(ppt_app, caplog):
    with caplog.at_level(logging.INFO):
        ppt_app.log_startup_banner()
    assert "1.3" in caplog.text
    assert "PowerPoint" in caplog.text


def _strip_line_comments(text):
    """各行の `#` 以降を落とす。コメントアウトされた呼び出しを実コードと誤認しないため。"""
    return "\n".join(line.split("#", 1)[0] for line in text.splitlines())


@pytest.mark.parametrize("path", [WORD_APP, PPT_APP], ids=["Word版", "PPT版"])
def test_log_startup_banner_is_called_before_sudachi_dictionary_init(path):
    """log_startup_banner() が Sudachi辞書の初期化より前に呼ばれていることをソース上で確認する。

    Sudachi辞書の初期化（dictionary.Dictionary(...)）に失敗すると、その場で
    sys.exit(1) して起動処理が打ち切られる。log_startup_banner() の呼び出しが
    それより後ろにあると、実際に届く起動失敗のログ（zipの中からexeを直接
    実行した場合など）には版番号が一切残らない。これはTask 3の目的そのものを
    潰してしまうため、モジュール読み込み順として辞書初期化より前に
    呼ばれていることを固定する。

    conftest.py はモジュールを importlib で読み込むだけで `__main__` を実行しない
    （GUIを起動させないため）。そのためモジュール読み込み時点の呼び出しを普通に
    呼んで確かめることができず、ソーステキストを直接読んで検査する。
    """
    text = _strip_line_comments(path.read_text(encoding="utf-8"))

    dict_marker = "dictionary.Dictionary("
    assert dict_marker in text, f"{path.name} に Sudachi辞書の初期化が見つからない"

    # "def log_startup_banner():" 自体は呼び出しではないので除外する。
    call_match = re.search(r"(?<!def )log_startup_banner\(\)", text)
    assert call_match is not None, (
        f"{path.name} から log_startup_banner() の呼び出しが消えている"
    )

    call_pos = call_match.start()
    dict_pos = text.index(dict_marker)
    assert call_pos < dict_pos, (
        f"{path.name} で log_startup_banner() が Sudachi辞書の初期化より後になっている"
        "（辞書初期化に失敗した場合の起動失敗ログに版番号が残らない）"
    )
