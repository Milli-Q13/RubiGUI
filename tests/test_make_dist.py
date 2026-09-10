"""配布物の組み立て。

配布物に入ってはいけないものが混入しないことを機械的に担保する。
実際に rubigui.log には実在の教材名とローカルパスが記録されていた。
"""
import sys
import zipfile
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
    """許可リスト自体が禁止パターンに触れていないこと。

    禁止パターンを手書きすると、make_dist.FORBIDDEN_SUFFIXES /
    FORBIDDEN_NAMES 側に新しいパターンが増えたときに乖離する
    （実際に .pyc と .spec が手書き版から漏れていた）。モジュール側の
    定義をそのまま参照することで、乖離しようがない形にする。
    """
    for entry in make_dist.ALLOWLIST:
        dest_lower = entry.dest.lower()
        assert not dest_lower.endswith(make_dist.FORBIDDEN_SUFFIXES)
        assert dest_lower not in make_dist.FORBIDDEN_NAMES


@pytest.mark.parametrize("name, is_dir", [
    ("REQUIREMENTS.TXT", False),
    ("__PYCACHE__", True),
])
def test_forbidden_check_is_case_insensitive_on_names(tmp_path, name, is_dir):
    """禁止ファイル名の判定が大文字小文字を区別しないことを保証する。

    Windows はファイル名の大文字小文字を区別しないため、うっかり
    REQUIREMENTS.TXT や __PYCACHE__ のような表記で紛れ込んでも
    見逃してはならない。
    """
    (tmp_path / "RubiGUI_Word_v3.1.exe").write_text("dummy")
    target = tmp_path / name
    if is_dir:
        target.mkdir()
        # 拡張子側の判定（.pyc 等）に頼らず、ディレクトリ名自体の
        # 大文字小文字非依存判定だけで検出できることを確認するため、
        # 中身は禁止拡張子に該当しないファイルにする。
        (target / "x.dat").write_text("dummy")
    else:
        target.write_text("dummy")
    with pytest.raises(SystemExit):
        make_dist.check_no_forbidden(tmp_path)


def _dummy_allowlist(tmp_path):
    """build() のテスト用に、実体を持つ最小の ALLOWLIST を作る。

    実際の exe / PDF はタスク7・8でまだ作られていないため使えない
    （ダミーを作ってはいけない、という制約は本物のビルド成果物の話であり、
    テスト用の一時ファイルはこのテストの中だけで完結する）。
    """
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    exe = src_dir / "dummy.exe"
    exe.write_bytes(b"dummy-exe")
    return [make_dist.Entry(exe, "dummy.exe")]


def test_build_validates_dic_before_copying_anything(tmp_path, monkeypatch):
    """--dic の指定ミスは、1バイトもコピーする前に停止することを保証する。

    辞書パスの指定ミスはリリース当日に起こりやすい操作ミスである。
    コピーを始めた後に気づくのではなく、着手前に止まることで、
    exe 等だけが入った中途半端な配布フォルダを残さない。
    """
    monkeypatch.setattr(make_dist, "ALLOWLIST", _dummy_allowlist(tmp_path))

    wrong_dic = tmp_path / "does_not_exist.dic"
    out_dir = tmp_path / "dist" / "RubiGUI_test"

    with pytest.raises(SystemExit):
        make_dist.build(out_dir, wrong_dic)

    # コピーが1件も始まっていない = out_dir 自体が作られていないはず。
    assert not out_dir.exists()


def test_build_removes_out_dir_when_a_later_step_fails(tmp_path, monkeypatch):
    """組み立て途中で失敗したら、半端な配布フォルダを残さないことを保証する。

    gen_notices.generate() が失敗する（あるいはディスクフルで copy2 が
    失敗する）と、そこまでにコピー済みの exe / json / readme 等だけが
    残ってしまう。オペレーターがそれを見て「できている」と誤認し、
    手で zip 化してしまうと、THIRD-PARTY-NOTICES.txt も辞書も欠けた
    配布物が外部に出てしまう。
    """
    monkeypatch.setattr(make_dist, "ALLOWLIST", _dummy_allowlist(tmp_path))

    dic = tmp_path / "dummy.dic"
    dic.write_bytes(b"dummy-dic")

    def boom(_out_path):
        raise SystemExit("告知ファイルの生成に失敗（テスト用）")

    monkeypatch.setattr(make_dist.gen_notices, "generate", boom)

    out_dir = tmp_path / "dist" / "RubiGUI_test"

    with pytest.raises(SystemExit):
        make_dist.build(out_dir, dic)

    # exe や辞書はコピー済みのはずだが、失敗した以上 out_dir ごと消えている必要がある。
    assert not out_dir.exists()


def test_build_refuses_to_touch_an_out_dir_with_unrelated_contents(tmp_path, monkeypatch):
    """--out の指定ミスで無関係なフォルダの中身を消さないことを保証する。

    build() は先頭で `shutil.rmtree(out_dir)` を無条件に実行していた。
    `--out .` ならリポジトリの作業ツリーが（.git ごと）、`--out D:\\授業資料`
    ならそこが、検証より前に丸ごと消えてしまう。
    ここでは「SystemExit が出た」ことだけでなく「無関係なファイルが実際に
    残っている」ことまで確認する。消してから例外を出しても前者だけの
    テストでは通ってしまうため。
    """
    monkeypatch.setattr(make_dist, "ALLOWLIST", _dummy_allowlist(tmp_path))

    out_dir = tmp_path / "out"
    out_dir.mkdir()
    stranger = out_dir / "無関係なファイル.txt"
    stranger.write_text("大事なデータ", encoding="utf-8")

    with pytest.raises(SystemExit):
        make_dist.build(out_dir, tmp_path / "dummy.dic")

    # 例外が出ただけでなく、無関係なファイルが実際に残っていること。
    assert stranger.exists()
    assert stranger.read_text(encoding="utf-8") == "大事なデータ"


def test_make_zip_uses_forward_slashes_and_utf8_flag_for_non_ascii_names(tmp_path):
    """zip エントリ名がスラッシュ区切りになり、日本語名は UTF-8 フラグ付きで
    格納されることを保証する。

    受け取るのは Windows の学校職員で、「はじめにお読みください.pdf」は
    案内の一番最初に読んでもらう文書である。ここで名前が化けると、
    案内の出だしで躓かせてしまう。サブフォルダ配下のファイルで検証するのは、
    フラットな構成だと OS のパス区切りの違いが表面化せず、
    このテストが何も検証していないのと同じになってしまうため。
    """
    folder = tmp_path / "RubiGUI_test"
    (folder / "docs").mkdir(parents=True)
    (folder / "docs" / "readme.txt").write_text("dummy", encoding="utf-8")
    (folder / "はじめにお読みください.pdf").write_bytes(b"%PDF-dummy")

    zip_path = make_dist.make_zip(folder)

    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        assert "RubiGUI_test/docs/readme.txt" in names
        assert not any("\\" in n for n in names)

        info = zf.getinfo("RubiGUI_test/はじめにお読みください.pdf")
        assert info.flag_bits & 0x800  # UTF-8 フラグ（bit 11）


def test_readmes_folder_structure_section_matches_the_allowlist():
    """両readmeの「■ フォルダ構成」が配布物の一覧（ALLOWLIST）と一致することを保証する。

    現状は手で確認して一致しているだけなので、次に ALLOWLIST へ
    エントリを足したときに readme を更新し忘れてもテストは黙って通って
    しまう。配布物に実際に入るファイル名が、両readmeの本文中のどこかに
    最低1回は書かれていることを機械的に確認する。
    """
    expected_names = {e.dest for e in make_dist.ALLOWLIST} | {
        "system_full.dic",
        "THIRD-PARTY-NOTICES.txt",
    }

    word_readme = (REPO / "RubiGUI_word_v3.1" / "readme.txt").read_text(encoding="utf-8")
    ppt_readme = (REPO / "RubiGUI_ppt_v1.3" / "readme.txt").read_text(encoding="utf-8")

    for readme_text, label in ((word_readme, "Word版readme"), (ppt_readme, "PPT版readme")):
        missing = sorted(name for name in expected_names if name not in readme_text)
        assert not missing, f"{label} の「■ フォルダ構成」に無い（と思われる）ファイル: {missing}"
