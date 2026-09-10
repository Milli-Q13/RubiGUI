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


def test_read_text_raises_when_encoding_is_undeterminable(tmp_path):
    """utf-8 でも cp932 でも読めないファイル（EUC-JP・UTF-16・壊れたバイト列など）を
    latin-1 は256バイト全てに対応するため必ず「それらしく」読めてしまい、例外が
    出ないまま中身が化けた文字列を返してしまう。将来同梱物の更新でこの種の
    ファイルが混入しても、告知本文が誤った内容のまま静かに配布されないよう、
    エンコーディングが特定できない場合は SystemExit で止まらなければならない。
    """
    bad_file = tmp_path / "broken_encoding.txt"
    # utf-8: 0x80 は先頭バイトとして不正。cp932: 0x83 で継続バイト不足として不正。
    # 一方 latin-1 ならどんなバイト列も例外なくデコードできてしまう（真の懸念点）。
    bad_file.write_bytes(bytes([0x80, 0x81, 0x82, 0x83]))
    with pytest.raises(SystemExit):
        gen_notices._read_text(bad_file)


def test_manual_license_is_not_discarded_when_package_yields_a_match(tmp_path, monkeypatch):
    """SudachiPy は配布物にライセンスファイルを同梱していないため、正しい本文は
    packaging/licenses/SudachiPy.txt に手置きしている。将来 SudachiPy の配布物に
    "LICENSE" 等の語を含む無関係なファイルが（誤って、あるいは仕様変更で）追加
    されると、`_collect_from_package(...) or _collect_from_manual(...)` という
    実装では package 側の戻り値がそのまま採用され、手置きの正しい Apache License
    全文が黙って握りつぶされてしまう。package 側が何かを返しても手置きの内容が
    失われないことを確認する。
    """
    original_collect_from_package = gen_notices._collect_from_package

    def fake_collect_from_package(name):
        if name == "SudachiPy":
            # 将来 SudachiPy が同梱するかもしれない、無関係だが名前だけ一致するファイルを模擬する
            return [("dummy/LICENSE_LOOKALIKE.txt", "これは無関係なダミーファイルです")]
        return original_collect_from_package(name)

    monkeypatch.setattr(gen_notices, "_collect_from_package", fake_collect_from_package)

    text = gen_notices.generate(tmp_path / "n.txt").read_text(encoding="utf-8")
    # 手置き SudachiPy.txt 冒頭の署名的な一文（Apache License 本文そのもの）が
    # 出力から消えていないこと
    assert "Works Applications Co., Ltd." in text
