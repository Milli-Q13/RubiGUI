"""配布物の組み立て。

配布物に入ってはいけないものが混入しないことを機械的に担保する。
実際に rubigui.log には実在の教材名とローカルパスが記録されていた。
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "packaging"))

import make_dist  # noqa: E402


def test_forbidden_check_passes_on_a_clean_folder(tmp_path):
    (tmp_path / "RubiGUI_Word_v3.1.exe").write_text("dummy")
    (tmp_path / "readme_Word.txt").write_text("dummy")
    make_dist.check_no_forbidden(tmp_path)  # 例外が出なければ合格


@pytest.mark.parametrize("name", [
    "rubigui.log",
    "RubiGUI_V3.1.py",
    "requirements.txt",
])
def test_forbidden_check_stops_on_a_forbidden_file(tmp_path, name):
    (tmp_path / "RubiGUI_Word_v3.1.exe").write_text("dummy")
    (tmp_path / name).write_text("秘密", encoding="utf-8")
    with pytest.raises(SystemExit):
        make_dist.check_no_forbidden(tmp_path)


def test_forbidden_check_stops_on_pycache(tmp_path):
    (tmp_path / "__pycache__").mkdir()
    (tmp_path / "__pycache__" / "x.pyc").write_text("dummy")
    with pytest.raises(SystemExit):
        make_dist.check_no_forbidden(tmp_path)


def test_sha256_matches_a_known_value(tmp_path):
    """空ファイルのSHA-256は既知の値になる。"""
    path = tmp_path / "empty.bin"
    path.write_bytes(b"")
    expected = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
    assert make_dist.sha256(path) == expected


def test_allowlist_has_no_forbidden_pattern():
    """許可リスト自体が禁止パターンに触れていないこと。"""
    for entry in make_dist.ALLOWLIST:
        assert not entry.dest.endswith(".py")
        assert not entry.dest.endswith(".log")
        assert entry.dest != "requirements.txt"
