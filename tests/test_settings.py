"""ruby_settings.json の読み書き。

Word版とPPT版は同じフォルダに同居し、同じ ruby_settings.json を共有する。
両版が持つキーは一致しないので、片方の保存がもう片方のキーを消さないことを確かめる。
"""
import json

import pytest

PPT_SETTINGS = {
    "ruby_ratio": 50,
    "line_spacing": 2.0,
    "ruby_offset": 0.0,
    "include_title": True,
    "ruby_mode": "first",
}

WORD_SETTINGS = {
    "ruby_mode": "all",
    "ruby_ratio": 60,
    "ruby_offset": 1.0,
}


@pytest.fixture
def settings_file(tmp_path, word_app, ppt_app, monkeypatch):
    """両アプリの保存先を、テスト用の空フォルダへ向ける。"""
    path = tmp_path / "ruby_settings.json"
    monkeypatch.setattr(word_app, "SETTINGS_PATH", path)
    monkeypatch.setattr(ppt_app, "SETTINGS_PATH", path)
    return path


def test_word_save_keeps_ppt_only_keys(settings_file, word_app, ppt_app):
    """Word版の保存が、PPT版だけが持つキーを消さない。"""
    ppt_app.save_settings(PPT_SETTINGS)
    word_app.save_settings(WORD_SETTINGS)

    saved = json.loads(settings_file.read_text(encoding="utf-8"))
    assert saved["line_spacing"] == 2.0
    assert saved["include_title"] is True


def test_word_save_still_updates_its_own_keys(settings_file, word_app, ppt_app):
    """相手のキーを守るあまり、自分の設定が保存されないのでは意味がない。"""
    ppt_app.save_settings(PPT_SETTINGS)
    word_app.save_settings(WORD_SETTINGS)

    saved = json.loads(settings_file.read_text(encoding="utf-8"))
    assert saved["ruby_mode"] == "all"
    assert saved["ruby_ratio"] == 60
    assert saved["ruby_offset"] == 1.0


def test_ppt_reads_back_its_keys_after_word_saved(settings_file, word_app, ppt_app):
    """利用者から見た症状（設定が戻っている）が起きないことの確認。"""
    ppt_app.save_settings(PPT_SETTINGS)
    word_app.save_settings(WORD_SETTINGS)

    loaded = ppt_app.load_settings()
    assert loaded["line_spacing"] == 2.0
    assert loaded["include_title"] is True


def test_save_survives_a_broken_settings_file(settings_file, word_app):
    """壊れたファイルがあっても、自分の設定は保存できる。"""
    settings_file.write_text("{ これは壊れたJSON", encoding="utf-8")

    word_app.save_settings(WORD_SETTINGS)

    saved = json.loads(settings_file.read_text(encoding="utf-8"))
    assert saved["ruby_mode"] == "all"
