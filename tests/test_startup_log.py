"""起動時に版番号がログへ残ることの確認。

不具合報告ではログを添付してもらう。版がログから分かれば、報告者が
版番号を書き間違えても調査できる。
"""
import logging


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
