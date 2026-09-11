"""ruby_settings.json の読み書き。

Word版とPPT版は同じフォルダに同居し、同じ ruby_settings.json を共有する。
両版が持つキーは一致しないので、片方の保存がもう片方のキーを消さないことを確かめる。
"""
import json
import types

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


# ============================================================
# 閉じるときの保存
# ============================================================
# 数値設定（ルビの大きさ・高さ、PPT版は行間）は read_settings_from_ui() でしか
# 保存されず、その関数はファイルを処理する経路からしか呼ばれない。入力欄には
# 変更を検知する仕組みが無く、閉じるときにも保存していなかった。
# そのため「設定を変えて閉じただけ」では値が失われ、開き直すと既定値に戻る。
# 選択・チェック系（ルビを振る範囲、タイトル枠）は変更時に即保存するハンドラを
# 持つため残り、数値だけが消えるという分かりにくい壊れ方をしていた。
# この挙動は v3.0 / v1.2 から存在する。

import inspect  # noqa: E402

from pathlib import Path  # noqa: E402

REPO = Path(__file__).resolve().parent.parent
WORD_APP_SRC = REPO / "RubiGUI_word_v3.1" / "RubiGUI_V3.1.py"
PPT_APP_SRC = REPO / "RubiGUI_ppt_v1.3" / "RubiGUI_PPT_V1.3.py"


def _on_close_body(path):
    """`if __name__ == "__main__":` にある on_close() の本体を返す。"""
    text = path.read_text(encoding="utf-8")
    text = "\n".join(line.split("#", 1)[0] for line in text.splitlines())
    start = text.index("def on_close():")
    end = text.index("root.protocol(", start)
    return text[start:end]


@pytest.mark.parametrize("path", [WORD_APP_SRC, PPT_APP_SRC], ids=["Word版", "PPT版"])
def test_closing_the_window_saves_the_numeric_settings(path):
    """ウィンドウを閉じるときに入力欄の値を保存していることをソース上で確認する。

    conftest.py はモジュールを importlib で読み込むだけで `__main__` を実行しない
    （GUIを起動させないため）。閉じる処理は `__main__` の中にあるので、実際に
    呼んで確かめられない。ソーステキストを直接読んで検査する。
    """
    body = _on_close_body(path)

    assert "read_settings_from_ui" in body, (
        f"{path.name} の on_close() が設定を保存していない"
        "（数値設定を変えて閉じても保存されず、開き直すと既定値に戻る）"
    )
    assert body.index("read_settings_from_ui") < body.index("root.destroy()"), (
        f"{path.name} の on_close() で保存が root.destroy() より後になっている"
    )


@pytest.mark.parametrize("app_name", ["word_app", "ppt_app"], ids=["Word版", "PPT版"])
def test_read_settings_from_ui_can_stay_quiet(app_name, request):
    """閉じるときの保存では警告ダイアログを出さないための引数があること。

    入力欄に不正な値が残ったまま閉じられた場合、通常の経路では警告を出して
    既定値に戻す。しかし閉じる最中にダイアログを出すと、利用者は終了を
    妨げられたように感じるうえ、破棄処理を止めかねない。
    """
    module = request.getfixturevalue(app_name)
    sig = inspect.signature(module.RubyEditorApp.read_settings_from_ui)
    assert "quiet" in sig.parameters, (
        "read_settings_from_ui() に quiet 引数が無い"
        "（閉じるときに警告ダイアログを抑えられない）"
    )


class _FakeEntry:
    """tk.Entry の代わり。read_settings_from_ui が使う4つのメソッドだけ持つ。"""

    def __init__(self, text):
        self._text = text

    def get(self):
        return self._text

    def delete(self, start, end):
        self._text = ""

    def insert(self, index, text):
        self._text = text


class _FakeVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


def test_word_numeric_settings_actually_reach_the_file(tmp_path, word_app, monkeypatch):
    """入力欄に打った数値が、ファイルに書き出されることを確かめる。

    これが壊れていたのが今回の不具合。「設定を変えて閉じただけ」では
    保存されず、開き直すと既定値に戻っていた。
    """
    path = tmp_path / "ruby_settings.json"
    monkeypatch.setattr(word_app, "SETTINGS_PATH", path)

    obj = types.SimpleNamespace(
        settings=dict(word_app.DEFAULT_SETTINGS),
        ratio_entry=_FakeEntry("60"),
        offset_entry=_FakeEntry("2"),
        ruby_mode_var=_FakeVar("all"),
    )
    word_app.RubyEditorApp.read_settings_from_ui(obj, quiet=True)

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved["ruby_ratio"] == 60
    assert saved["ruby_offset"] == 2
    assert saved["ruby_mode"] == "all"


def test_ppt_numeric_settings_actually_reach_the_file(tmp_path, ppt_app, monkeypatch):
    """PPT版も同じ。行間はPPT版だけが持つ数値設定。"""
    path = tmp_path / "ruby_settings.json"
    monkeypatch.setattr(ppt_app, "SETTINGS_PATH", path)

    obj = types.SimpleNamespace(
        settings=dict(ppt_app.DEFAULT_SETTINGS),
        ratio_entry=_FakeEntry("60"),
        spacing_entry=_FakeEntry("2.0"),
        offset_entry=_FakeEntry("2"),
        include_title_var=_FakeVar(1),
        ruby_mode_var=_FakeVar("all"),
    )
    ppt_app.RubyEditorApp.read_settings_from_ui(obj, quiet=True)

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved["line_spacing"] == 2.0
    assert saved["ruby_ratio"] == 60
    assert saved["include_title"] is True


def test_quiet_mode_does_not_pop_a_dialog_on_bad_input(tmp_path, word_app, monkeypatch):
    """不正な値のまま閉じられても、警告ダイアログを出さずに既定値へ戻す。"""
    path = tmp_path / "ruby_settings.json"
    monkeypatch.setattr(word_app, "SETTINGS_PATH", path)

    called = []
    monkeypatch.setattr(word_app.msgbox, "showwarning", lambda *a, **k: called.append(a))

    obj = types.SimpleNamespace(
        settings=dict(word_app.DEFAULT_SETTINGS),
        ratio_entry=_FakeEntry("これは数値ではない"),
        offset_entry=_FakeEntry("0"),
        ruby_mode_var=_FakeVar("first"),
    )
    word_app.RubyEditorApp.read_settings_from_ui(obj, quiet=True)

    assert called == [], "quiet=True なのに警告ダイアログが出た"
    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved["ruby_ratio"] == word_app.DEFAULT_SETTINGS["ruby_ratio"]
