"""起動時に版番号がログへ残ることの確認。

不具合報告ではログを添付してもらう。版がログから分かれば、報告者が
版番号を書き間違えても調査できる。
"""
import logging
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
def test_log_startup_banner_is_called_before_tk_init_in_main(path):
    """log_startup_banner() が起動時の配線から外れていないことをソース上で確認する。

    conftest.py はモジュールを importlib で読み込むだけで `__main__` を実行しない
    （GUIを起動させないため）。そのため呼び出し側を普通に呼んで確かめることができず、
    `if __name__ == "__main__":` 以降のソーステキストを直接読んで検査する。
    """
    text = path.read_text(encoding="utf-8")
    marker = 'if __name__ == "__main__":'
    assert marker in text, f"{path.name} に __main__ ブロックが見つからない"

    main_block = _strip_line_comments(text[text.index(marker):])
    assert "log_startup_banner()" in main_block, (
        f"{path.name} の __main__ ブロックから log_startup_banner() の呼び出しが消えている"
    )

    call_pos = main_block.index("log_startup_banner()")
    tk_pos = main_block.index("TkinterDnD.Tk()")
    assert call_pos < tk_pos, (
        f"{path.name} で log_startup_banner() が TkinterDnD.Tk() より後になっている"
        "（他の処理より先にログを出す、という要件が壊れている）"
    )
