"""THIRD-PARTY-NOTICES.txt の生成。

同梱物の告知が1つでも欠けたまま配布しないことを機械的に担保する。
特に SudachiDict の LEGAL は、UniDic（著作権表示の再生産が義務）と
NEologd の告知を含むため、抜けると義務違反になる。
"""
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "packaging"))

import gen_notices  # noqa: E402


def test_generates_a_file(tmp_path):
    out = gen_notices.generate(tmp_path / "THIRD-PARTY-NOTICES.txt")
    assert out.exists()
    assert out.stat().st_size > 10000


def test_covers_every_bundled_component(tmp_path):
    text = gen_notices.generate(tmp_path / "n.txt").read_text(encoding="utf-8")
    for name in gen_notices.BUNDLED:
        assert name in text, f"{name} の告知が欠けている"


def test_includes_sudachidict_legal(tmp_path):
    """UniDic と NEologd の告知は SudachiDict の LEGAL にしかない。"""
    text = gen_notices.generate(tmp_path / "n.txt").read_text(encoding="utf-8")
    assert "UniDic" in text
    assert "neologd" in text.lower()


def test_includes_the_apache_license_body(tmp_path):
    text = gen_notices.generate(tmp_path / "n.txt").read_text(encoding="utf-8")
    assert "Apache License" in text


def test_stops_when_a_component_has_no_notice(tmp_path, monkeypatch):
    """告知を用意できない同梱物があれば、黙って続けず停止する。"""
    monkeypatch.setattr(gen_notices, "BUNDLED", ["存在しないパッケージ"])
    with pytest.raises(SystemExit):
        gen_notices.generate(tmp_path / "n.txt")
