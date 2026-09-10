# RubiGUI 外部配布 実装計画

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** RubiGUI Word版とPowerPoint版を1つの配布物にまとめ、職場外へ安全に配れる状態にする。

**Architecture:** 新しい版フォルダ（`RubiGUI_word_v3.1/` / `RubiGUI_ppt_v1.3/`）を作り、同梱に必要な修正をそこだけに入れる。旧版フォルダは配布済みの成果物として一切触らない。配布物は `packaging/make_dist.py` が**許可リストに挙げたファイルだけ**を集めて組み立て、zip 化して SHA-256 を出す。アプリに通信機能は追加しない。

**Tech Stack:** Python 3 / tkinter / SudachiPy / pywin32 / PyInstaller / pytest 9.1.1

**Spec:** `docs/superpowers/specs/2026-09-10-external-distribution-design.md`

## Global Constraints

- ライセンスは **MIT**。著作権者表記は `Copyright (c) 2026 Milli-Q13`（この文字列のまま使う）
- **アプリに通信機能を追加しない。** 更新確認・バージョンチェック・外部への送信を一切書かない
- 配布物に `*.py` / `*.log` / `__pycache__/` / `requirements.txt` を含めない
- 配布版数は Word **v3.1** / PPT **v1.3**。同梱物全体の識別子は **`RubiGUI_2026-09`**
- 同梱辞書は **SudachiDict-full 20250515**（`system_full.dic`）のまま。更新しない
- **`.bas` ファイルは CP932（Shift-JIS）エンコード。** UTF-8 として読み書きしないこと。VBAエディタへのインポートで文字化けする。ASCII文字の置換のみ行う
- コメント・ドキュメント・コミットメッセージは日本語で書く（既存コードに合わせる）
- 旧版フォルダ（`RubiGUI_word_v3.0/` `RubiGUI_ppt_v1.2/` およびそれ以前）は**変更しない**

## ファイル構成

**新規作成**

| パス | 責務 |
|---|---|
| `RubiGUI_word_v3.1/` | Word版 v3.1 のソースと設定（`system_full.dic` は置かない） |
| `RubiGUI_ppt_v1.3/` | PPT版 v1.3 のソースと設定（同上） |
| `tests/conftest.py` | アプリ本体を pytest から読み込む共通処理 |
| `tests/test_version.py` | 版番号とマクロ名の整合の検証 |
| `tests/test_settings.py` | `ruby_settings.json` の読み書きの検証 |
| `tests/test_startup_log.py` | 起動ログに版番号が残ることの検証 |
| `tests/test_notices.py` | 第三者ライセンス告知の生成の検証 |
| `tests/test_make_dist.py` | 配布物の組み立てと混入チェックの検証 |
| `packaging/gen_notices.py` | `THIRD-PARTY-NOTICES.txt` の生成 |
| `packaging/licenses/` | 自動収集できないライセンス本文の手置き場 |
| `packaging/make_dist.py` | 配布物の組み立て・zip化・SHA-256・混入チェック |
| `packaging/LICENSE.txt` | MIT 本文（英文正文＋参考訳） |
| `docs/はじめにお読みください.docx` | 導入手順書の編集元 |
| `docs/はじめにお読みください.pdf` | 導入手順書（配布物に入れる） |

**変更**

| パス | 変更内容 |
|---|---|
| `packaging/build_exe.py` | ビルド対象を v3.1 / v1.3 へ、exe名を変更 |
| `.gitignore` | 配布物の出力先と生成物を除外 |
| `readme.md` | 最新版の表と配布手順の追記 |
| `docs/CHANGELOG.md` | v3.1 / v1.3 の節を追加 |

### 辞書ファイルの扱い

`make_dist.py` が `--dic` で指定されたパスから配布物へコピーする（既定値 `RubiGUI_word_v3.0/system_full.dic`）。

**ただし新しい版フォルダにも辞書の実体が必要である。** アプリはモジュールの読み込み時点で
Sudachi辞書を初期化し、失敗するとエラーダイアログを出して `sys.exit(1)` する。
そのため辞書が無いと `tests/conftest.py` からの import が成立せず、テストが動かない。

ディスクを増やさずにこれを満たすため、**新しい版フォルダの `system_full.dic` は
`RubiGUI_word_v3.0/system_full.dic` へのハードリンク**にする（NTFS。実体は1つ、容量の増加なし）。

```bash
python -c "import os; os.link('RubiGUI_word_v3.0/system_full.dic', 'RubiGUI_word_v3.1/system_full.dic')"
python -c "import os; os.link('RubiGUI_word_v3.0/system_full.dic', 'RubiGUI_ppt_v1.3/system_full.dic')"
```

辞書は `.gitignore` 済みなので、この措置は各自の作業環境だけの話であり、
リポジトリにも配布物にも影響しない。

**動作確認は版フォルダではなく、組み上がった配布フォルダに対して行う。** 配布する物そのものを試すことになるので、この方が確実である。

---

### Task 1: 新版フォルダの作成と pytest の土台

**Files:**
- Create: `RubiGUI_word_v3.1/RubiGUI_V3.1.py`（`RubiGUI_word_v3.0/RubiGUI_V3.0.py` のコピー）
- Create: `RubiGUI_word_v3.1/RubiGUI_V31.bas`（`RubiGUI_V30.bas` のコピー）
- Create: `RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py`（`RubiGUI_ppt_v1.2/RubiGUI_PPT_V1.2.py` のコピー）
- Create: `tests/conftest.py`
- Create: `tests/test_version.py`
- Modify: `packaging/build_exe.py`

**Interfaces:**
- Consumes: なし（最初のタスク）
- Produces: pytest フィクスチャ `word_app` / `ppt_app`（読み込み済みのアプリモジュールを返す）。以降のすべてのテストがこれを使う。モジュール属性 `APP_VERSION`（str）、`SETTINGS_PATH`（Path）、`MACRO_NAME`（str, Word版のみ）、関数 `load_settings()` / `save_settings(dict)` を公開する。

- [ ] **Step 1: 共通の読み込み処理を書く**

`tests/conftest.py` を作る。

```python
"""RubiGUI のアプリ本体を pytest から読み込むための共通処理。

アプリは1ファイル完結のGUIスクリプトで、ファイル名に「.」が入っている
（RubiGUI_V3.1.py）ため、通常の import 文では読み込めない。importlib で
ファイルパスを直接指定して読み込む。

GUI は起動しない。起動処理は `if __name__ == "__main__"` の中にあるため、
モジュールとして読み込むだけなら定義が並ぶだけで済む（実測 0.6 秒）。

なお読み込むとそのフォルダに rubigui.log が作られる。.gitignore 済みで、
配布物にも入らない（make_dist.py の許可リストに無いため）。
"""
import importlib.util
import sys
from pathlib import Path

import pytest

REPO = Path(__file__).resolve().parent.parent

WORD_APP = REPO / "RubiGUI_word_v3.1" / "RubiGUI_V3.1.py"
PPT_APP = REPO / "RubiGUI_ppt_v1.3" / "RubiGUI_PPT_V1.3.py"


def _load(path, name):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="session")
def word_app():
    """Word版のモジュール。"""
    return _load(WORD_APP, "rubigui_word")


@pytest.fixture(scope="session")
def ppt_app():
    """PowerPoint版のモジュール。"""
    return _load(PPT_APP, "rubigui_ppt")
```

- [ ] **Step 2: 失敗するテストを書く**

`tests/test_version.py` を作る。

```python
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
```

- [ ] **Step 3: テストを実行して失敗を確認する**

Run: `python -m pytest tests/test_version.py -v`
Expected: FAIL（`RubiGUI_word_v3.1` が存在しないため収集時にエラー）

- [ ] **Step 4: 新しい版フォルダを作る**

辞書・exe・ログ・`__pycache__` はコピーしない。

```bash
mkdir -p RubiGUI_word_v3.1 RubiGUI_ppt_v1.3

cp RubiGUI_word_v3.0/RubiGUI_V3.0.py  RubiGUI_word_v3.1/RubiGUI_V3.1.py
cp RubiGUI_word_v3.0/RubiGUI_V30.bas  RubiGUI_word_v3.1/RubiGUI_V31.bas
cp RubiGUI_word_v3.0/override.json RubiGUI_word_v3.0/ruby_settings.json RubiGUI_word_v3.0/sudachi.json RubiGUI_word_v3.0/readme.txt RubiGUI_word_v3.1/

cp RubiGUI_ppt_v1.2/RubiGUI_PPT_V1.2.py RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py
cp RubiGUI_ppt_v1.2/override.json RubiGUI_ppt_v1.2/ruby_settings.json RubiGUI_ppt_v1.2/sudachi.json RubiGUI_ppt_v1.2/readme.txt RubiGUI_ppt_v1.3/
```

`requirements.txt` はコピーしない。配布物に入れないものであり、開発時の依存はリポジトリ直下で管理すれば足りる。

- [ ] **Step 5: `.bas` のモジュール名とマクロ名を変える**

置換対象は3箇所すべてASCIIなので、バイト単位の置換で CP932 のまま安全に書き換えられる。

**`-b`（バイナリモード）を必ず付ける。** この環境の GNU sed は `-b` 無しだと
CRLF を LF に変換してしまい、`.bas` の改行コードが壊れる。

```bash
sed -bi 's/V30/V31/g' RubiGUI_word_v3.1/RubiGUI_V31.bas
grep -c "V31" RubiGUI_word_v3.1/RubiGUI_V31.bas
```

Expected: `3`

- [ ] **Step 6: Word版 Python の版番号とマクロ名を変える**

`RubiGUI_word_v3.1/RubiGUI_V3.1.py` の次の箇所を書き換える。

| 現在 | 変更後 |
|---|---|
| `APP_VERSION = "3.0"` | `APP_VERSION = "3.1"` |
| `MACRO_NAME = "InsertFuriganaFromTSV_V30"` | `MACRO_NAME = "InsertFuriganaFromTSV_V31"` |
| コメント `# Word側マクロ（RubiGUI_V30.bas）のマクロ名。` | `RubiGUI_V31.bas` |
| コメント `# 「RubiGUI_V30.bas を入れ直してください」と明示的に案内できる。` | `RubiGUI_V31.bas` |
| メッセージ `"繰り返し発生する場合は、RubiGUI_V30.bas が最新版か確認してください。\n"` | `RubiGUI_V31.bas` |
| メッセージ `f"     {APP_DIR / 'RubiGUI_V30.bas'}\n"` | `'RubiGUI_V31.bas'` |

**改名の履歴を書いたコメントは書き換えず、末尾に追記する。** 過去の記録なので `V30` のまま残す必要がある。

```python
# （v2.0 "InsertFuriganaFromTSV_SaveToNewFile_Stable" → v2.1 "..._V21" → v3.0 "..._V30"
#   → v3.1 "..._V31"）。
```

`# ★v3.0 ではモジュール名も Module1 から RubiGUI_V30 へ変更した。` の行も履歴なので残し、続けて次の1行を足す。

```python
# ★v3.1 でモジュール名を RubiGUI_V31 に改めた（版ごとに追随させる方針）。
```

**一括置換（`sed s/V30/V31/g`）を Python ファイルに使ってはいけない。** 履歴コメントまで書き換わり、過去の記録が失われる。

- [ ] **Step 7: PPT版 Python の版番号を変える**

`RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py` の `APP_VERSION = "1.2"` を `APP_VERSION = "1.3"` にする。PPT版はマクロを使わないため、他に変更はない。

- [ ] **Step 8: テストを実行して通ることを確認する**

Run: `python -m pytest tests/test_version.py -v`
Expected: 5件すべて PASS

- [ ] **Step 9: ビルドスクリプトを新しい版に向ける**

`packaging/build_exe.py` の `TARGETS` を次のように書き換える。exe 名は同一フォルダに2つ並んだときにどちらが Word 用か分かる名前にする。

```python
TARGETS = {
    "word": {
        "script": REPO / "RubiGUI_word_v3.1" / "RubiGUI_V3.1.py",
        "name": "RubiGUI_Word_v3.1",
    },
    "ppt": {
        "script": REPO / "RubiGUI_ppt_v1.3" / "RubiGUI_PPT_V1.3.py",
        "name": "RubiGUI_PPT_v1.3",
    },
}
```

`main()` の末尾にある配布物の案内（`print("\n配布するときは、exe と同じフォルダに次を入れてください:")` から始まる3行）は `make_dist.py` に置き換わるので、次の1行に差し替える。

```python
    print("\n配布物の組み立ては packaging/make_dist.py で行ってください。")
```

- [ ] **Step 10: コミット**

```bash
git add RubiGUI_word_v3.1 RubiGUI_ppt_v1.3 tests/conftest.py tests/test_version.py packaging/build_exe.py
git commit -m "Word v3.1 / PPT v1.3: 版フォルダを作成しマクロ名を追随させた

外部配布にあたって同梱に必要な修正を入れるための版。旧版フォルダは
配布済みの成果物なので触らない。辞書は複製せず、make_dist.py が
配布物へコピーする。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 2: 設定ファイルのマージ書き込み

**Files:**
- Modify: `RubiGUI_word_v3.1/RubiGUI_V3.1.py`（`save_settings`）
- Modify: `RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py`（`save_settings`）
- Test: `tests/test_settings.py`

**Interfaces:**
- Consumes: `word_app` / `ppt_app` フィクスチャ（Task 1）
- Produces: `save_settings(settings: dict) -> None`。既存ファイルの内容を保ったまま `settings` のキーだけを更新して書き戻す。ファイルが壊れていても自分の設定は保存する。

**なぜ必要か:** 両版は同じフォルダに同居し、同じ `ruby_settings.json` を共有する。持っているキーは一致しない（PPT版だけが `line_spacing` と `include_title` を持つ）。現在の `save_settings` は `json.dump(settings)` で自分のキーだけを書き出すため、Word版が保存した時点で PPT版のキーが消える。読み込み側は未知のキーを無視するので気づきにくく、「設定したはずなのに戻っている」という報告になる。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_settings.py` を作る。

```python
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
```

- [ ] **Step 2: テストを実行して失敗を確認する**

Run: `python -m pytest tests/test_settings.py -v`
Expected: `test_word_save_keeps_ppt_only_keys` と `test_ppt_reads_back_its_keys_after_word_saved` が FAIL（`KeyError: 'line_spacing'`）。残り2件は現在の実装でも PASS する。

- [ ] **Step 3: 両版の `save_settings` を書き換える**

`RubiGUI_word_v3.1/RubiGUI_V3.1.py` と `RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py` の `save_settings` を、**まったく同じ内容**に置き換える。

```python
def save_settings(settings):
    """ルビ設定を ruby_settings.json に保存する。

    ★重要：Word版とPPT版は同じフォルダに同居し、同じ ruby_settings.json を
    共有する。両版が持つキーは一致しない（PPT版だけが line_spacing と
    include_title を持つ）ので、自分の設定だけを書き出すと相手版のキーが
    消える。読み込み側は知らないキーを無視するため気づきにくく、
    「設定したはずなのに戻っている」という分かりにくい不具合になる。
    そのため、既存の内容を読んでから自分のキーだけを更新して書き戻す。
    """
    merged = {}
    try:
        if SETTINGS_PATH.exists():
            with open(SETTINGS_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                merged.update(loaded)
    except (json.JSONDecodeError, UnicodeDecodeError, OSError) as e:
        # 読み直せない場合は「相手版のキーは救えないが、自分の設定は保存する」方に倒す。
        # ここで諦めると、ファイルが一度壊れたきり設定を保存できなくなる。
        logging.warning(f"{SETTINGS_PATH.name} を読み直せませんでした（上書きします）: {e}")

    merged.update(settings)

    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8") as f:
            json.dump(merged, f, indent=2, ensure_ascii=False)
    except OSError as e:
        logging.error(f"{SETTINGS_PATH.name} の保存に失敗しました: {e}")
```

- [ ] **Step 4: テストを実行して通ることを確認する**

Run: `python -m pytest tests/test_settings.py -v`
Expected: 4件すべて PASS

- [ ] **Step 5: コミット**

```bash
git add RubiGUI_word_v3.1/RubiGUI_V3.1.py RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py tests/test_settings.py
git commit -m "設定の保存を上書きからマージへ変更（Word/PPT両版）

両版を同じフォルダに同梱すると ruby_settings.json を共有する。
持っているキーが違うため、従来の書き出しでは Word版の保存で
PPT版の line_spacing と include_title が消えていた。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 3: 起動時にログへ版番号を記録

**Files:**
- Modify: `RubiGUI_word_v3.1/RubiGUI_V3.1.py`
- Modify: `RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py`
- Test: `tests/test_startup_log.py`

**Interfaces:**
- Consumes: `word_app` / `ppt_app` フィクスチャ（Task 1）、`APP_VERSION`
- Produces: `log_startup_banner() -> None`。両版が同名で公開する。

**なぜ必要か:** 不具合報告ではログを送ってもらう運用にする。ログに版番号が残っていれば、報告者が版を書き間違えても調査できる。現在はログに版が出ていない。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_startup_log.py` を作る。

```python
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
```

- [ ] **Step 2: テストを実行して失敗を確認する**

Run: `python -m pytest tests/test_startup_log.py -v`
Expected: FAIL（`AttributeError: module 'rubigui_word' has no attribute 'log_startup_banner'`）

- [ ] **Step 3: Word版に関数を足して main から呼ぶ**

`RubiGUI_V3.1.py` の `def load_settings():` の直前に足す。

```python
def log_startup_banner():
    """起動したことと版番号をログに残す。

    不具合報告ではログを送ってもらうので、版がログから分かるようにしておく。
    報告者が版番号を書き間違えても、ログを見れば確実に特定できる。
    """
    logging.info(f"===== RubiGUI Word版 v{APP_VERSION} 起動 =====")
```

`main()` の本体の先頭（`try:` の直後の行）で呼ぶ。

```python
    log_startup_banner()
```

- [ ] **Step 4: PPT版に同じものを足す**

`RubiGUI_PPT_V1.3.py` の `def load_settings():` の直前に足す。文言だけ変える。

```python
def log_startup_banner():
    """起動したことと版番号をログに残す。

    不具合報告ではログを送ってもらうので、版がログから分かるようにしておく。
    報告者が版番号を書き間違えても、ログを見れば確実に特定できる。
    """
    logging.info(f"===== RubiGUI PowerPoint版 v{APP_VERSION} 起動 =====")
```

`main()` の本体の先頭で `log_startup_banner()` を呼ぶ。

- [ ] **Step 5: テストを実行して通ることを確認する**

Run: `python -m pytest tests/ -v`
Expected: Task 1〜3 の 11件すべて PASS

- [ ] **Step 6: コミット**

```bash
git add RubiGUI_word_v3.1/RubiGUI_V3.1.py RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py tests/test_startup_log.py
git commit -m "起動時に版番号をログへ記録

不具合報告ではログを添付してもらう運用にするため、ログだけで
版が特定できるようにした。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 4: ライセンス表記と第三者ライセンス告知の生成

**Files:**
- Create: `packaging/LICENSE.txt`
- Create: `packaging/licenses/SudachiPy.txt`
- Create: `packaging/licenses/Python.txt`
- Create: `packaging/licenses/TclTk.txt`
- Create: `packaging/gen_notices.py`
- Test: `tests/test_notices.py`
- Modify: `.gitignore`

**Interfaces:**
- Consumes: なし
- Produces: `gen_notices.generate(out_path: Path) -> Path`。`THIRD-PARTY-NOTICES.txt` を書き出してそのパスを返す。告知を用意できない同梱物があれば `SystemExit` を送出する。モジュール変数 `BUNDLED`（同梱物名のリスト）。`make_dist.py`（Task 5）がこれを呼ぶ。

**なぜ生成するのか:** 手書きすると、ライブラリを更新したときに更新を忘れる。ただし**すべてを自動収集できるわけではない**。SudachiPy は配布物にライセンスファイルを同梱しておらず、`importlib.metadata.files()` に LICENSE が現れない。Python 本体と Tcl/Tk も pip パッケージではないため収集できない。これらは `packaging/licenses/` に手置きし、スクリプトが両者を結合する。

- [ ] **Step 1: MIT ライセンス本文を置く**

`packaging/LICENSE.txt` を作る。英文が正文であることを明示し、参考訳を併記する。

```
MIT License

Copyright (c) 2026 Milli-Q13

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.


------------------------------------------------------------
参考訳（正文は上記の英文です）
------------------------------------------------------------

MITライセンス

Copyright (c) 2026 Milli-Q13

本ソフトウェアおよび関連する文書ファイル（以下「本ソフトウェア」）の複製を
取得するすべての人に対し、本ソフトウェアを無制限に扱うことを無償で許可します。
これには、使用、複製、改変、結合、公開、頒布、サブライセンス、および販売する
権利、ならびに本ソフトウェアを提供される人に同じことを許可する権利が
含まれます。ただし、以下の条件に従うものとします。

上記の著作権表示および本許諾表示を、本ソフトウェアのすべての複製または
重要な部分に記載しなければなりません。

本ソフトウェアは「現状のまま」提供され、明示または黙示を問わず、商品性、
特定目的への適合性、および権利非侵害についての保証を含め、いかなる保証も
ありません。作者または著作権者は、契約行為、不法行為、またはそれ以外で
あろうと、本ソフトウェアに起因または関連し、あるいは本ソフトウェアの
使用またはその他の扱いによって生じる一切の請求、損害、その他の義務に
ついて何らの責任も負いません。
```

- [ ] **Step 2: 自動収集できないライセンス本文を手置きする**

次の3ファイルを作る。本文は各配布元の公式のものをそのまま貼る。

- `packaging/licenses/SudachiPy.txt` — Apache License 2.0 の全文（https://www.apache.org/licenses/LICENSE-2.0.txt）
- `packaging/licenses/Python.txt` — PSF License Agreement の全文（ローカルの Python インストール先にある `LICENSE.txt` からコピーできる）
- `packaging/licenses/TclTk.txt` — Tcl/Tk のライセンス（BSD系）

各ファイルの先頭に、どの同梱物のものかを示す見出しを入れる。SudachiPy の例：

```
SudachiPy (https://github.com/WorksApplications/SudachiPy)
Copyright (c) 2019 Works Applications Co., Ltd.
Licensed under the Apache License, Version 2.0

------------------------------------------------------------

（以下、Apache License 2.0 の全文）
```

- [ ] **Step 3: 失敗するテストを書く**

`tests/test_notices.py` を作る。

```python
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
```

- [ ] **Step 4: テストを実行して失敗を確認する**

Run: `python -m pytest tests/test_notices.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'gen_notices'`）

- [ ] **Step 5: 生成スクリプトを書く**

`packaging/gen_notices.py` を作る。

```python
"""THIRD-PARTY-NOTICES.txt を生成する。

RubiGUI の exe には複数のオープンソースが同梱されている。いずれも
コピーレフトではないため RubiGUI 自体のソース公開義務は生じないが、
**ライセンス本文と告知を配布物に添える義務**はある。手書きすると
ライブラリを更新したときに更新を忘れるので、生成できるものは生成する。

★すべてを自動収集できるわけではない。
  ・SudachiPy は配布物にライセンスファイルを同梱していない
  ・Python 本体と Tcl/Tk は pip パッケージではない
  これらは packaging/licenses/ に手置きし、収集分と結合する。
  どちらにも無い同梱物があれば、告知漏れのまま配布しないよう停止する。
"""
import importlib.metadata as md
from pathlib import Path

HERE = Path(__file__).resolve().parent
MANUAL_DIR = HERE / "licenses"

# 配布物に同梱されるもの。pip パッケージ名か、手置きファイルの名前（拡張子なし）。
BUNDLED = [
    "SudachiPy",
    "SudachiDict-full",
    "tkinterdnd2",
    "jaconv",
    "pywin32",
    "pyinstaller",
    "Python",
    "TclTk",
]

# 収集対象とみなすファイル名の目印。LEGAL は SudachiDict の告知ファイル。
_MARKERS = ("LICENSE", "NOTICE", "LEGAL", "COPYING")

_HEADER = """\
RubiGUI 第三者ソフトウェアのライセンス告知
============================================================

RubiGUI には次のオープンソースソフトウェアが含まれています。
それぞれのライセンス本文と告知を以下に収録します。

このファイルは packaging/gen_notices.py が自動生成しています。
手で編集しないでください。

"""


def _read_text(path):
    """ライセンスファイルは配布元によって文字コードが異なるので順に試す。"""
    for encoding in ("utf-8", "cp932", "latin-1"):
        try:
            return Path(path).read_text(encoding=encoding)
        except (UnicodeDecodeError, LookupError):
            continue
    return Path(path).read_text(encoding="utf-8", errors="replace")


def _collect_from_package(name):
    """pip パッケージからライセンスファイルを集める。無ければ空リスト。"""
    try:
        dist = md.distribution(name)
        files = md.files(name) or []
    except md.PackageNotFoundError:
        return []

    found = []
    for f in files:
        upper = str(f).upper()
        if any(m in upper for m in _MARKERS):
            path = Path(dist.locate_file(f))
            if path.is_file():
                found.append((str(f), _read_text(path)))
    return found


def _collect_from_manual(name):
    """packaging/licenses/ に手置きされた本文を読む。無ければ空リスト。"""
    path = MANUAL_DIR / f"{name}.txt"
    if path.is_file():
        return [(f"licenses/{path.name}（手置き）", _read_text(path))]
    return []


def _version(name):
    try:
        return md.version(name)
    except md.PackageNotFoundError:
        return None


def generate(out_path):
    """THIRD-PARTY-NOTICES.txt を書き出し、そのパスを返す。"""
    out_path = Path(out_path)
    parts = [_HEADER]
    missing = []

    for name in BUNDLED:
        entries = _collect_from_package(name) or _collect_from_manual(name)
        if not entries:
            missing.append(name)
            continue

        version = _version(name)
        title = f"{name} {version}" if version else name
        parts.append("=" * 60)
        parts.append(title)
        parts.append("=" * 60)
        parts.append("")
        for origin, body in entries:
            parts.append(f"--- {origin} ---")
            parts.append("")
            parts.append(body.rstrip())
            parts.append("")

    if missing:
        raise SystemExit(
            "告知を用意できない同梱物があります: " + "、".join(missing) + "\n"
            "pip パッケージ名が正しいか確認するか、"
            f"{MANUAL_DIR} に <名前>.txt としてライセンス本文を置いてください。"
        )

    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text("\n".join(parts), encoding="utf-8")
    return out_path


if __name__ == "__main__":
    written = generate(HERE / "THIRD-PARTY-NOTICES.txt")
    print(f"generated: {written} ({written.stat().st_size} bytes)")
```

- [ ] **Step 6: テストを実行して通ることを確認する**

Run: `python -m pytest tests/test_notices.py -v`
Expected: 5件すべて PASS

- [ ] **Step 7: 生成物を目視で確認する**

```bash
python packaging/gen_notices.py
```

`packaging/THIRD-PARTY-NOTICES.txt` を**エディタで開いて**次を確認する（コンソールに `cat` すると Windows の文字コードで化けて見えるので、判断材料にしない）。

- SudachiDict の LEGAL に含まれる UniDic の著作権表示（`The UniDic Consortium`）が入っている
- 日本語部分が文字化けしていない
- 8つの同梱物すべてに見出しがある

生成物そのものはコミットしない（`make_dist.py` が毎回作る）。

- [ ] **Step 8: コミット**

```bash
echo "packaging/THIRD-PARTY-NOTICES.txt" >> .gitignore
git add packaging/LICENSE.txt packaging/licenses packaging/gen_notices.py tests/test_notices.py .gitignore
git commit -m "MITライセンスと第三者ライセンス告知の生成を追加

外部配布にあたり、同梱物のライセンス本文と告知を配布物へ添える。
SudachiPy・Python・Tcl/Tk は自動収集できないので手置きし、
どちらにも無い同梱物があれば生成を停止する。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 5: 配布物の組み立て

**Files:**
- Create: `packaging/make_dist.py`
- Test: `tests/test_make_dist.py`
- Modify: `.gitignore`

**Interfaces:**
- Consumes: `gen_notices.generate(out_path)`（Task 4）
- Produces: `make_dist.build(out_dir: Path, dic: Path) -> Path`（配布フォルダを組んでパスを返す）、`make_dist.check_no_forbidden(folder: Path) -> None`（禁止ファイルがあれば `SystemExit`）、`make_dist.sha256(path: Path) -> str`、`make_dist.ALLOWLIST`（`Entry(src, dest)` のリスト）

**なぜ許可リストなのか:** 現在の「フォルダごと zip に固める」方式は危険である。実際に `RubiGUI_word_v3.0/rubigui.log` には実在の授業プリント名とローカルパス（ユーザー名を含む）が記録されていた。`.gitignore` は zip の作成には効かない。除外リスト方式だと新しい種類のファイルが増えたときに漏れるため、**入れるものを明示的に列挙する**方式にする。

- [ ] **Step 1: 失敗するテストを書く**

`tests/test_make_dist.py` を作る。

```python
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
```

- [ ] **Step 2: テストを実行して失敗を確認する**

Run: `python -m pytest tests/test_make_dist.py -v`
Expected: FAIL（`ModuleNotFoundError: No module named 'make_dist'`）

- [ ] **Step 3: 組み立てスクリプトを書く**

`packaging/make_dist.py` を作る。

```python
"""配布物を組み立てる。

使い方（リポジトリのどこからでも可）:
    python packaging/build_exe.py          # 先に exe をビルドしておく
    python packaging/make_dist.py          # 配布物を組んで zip 化する

★方針：入れるものを明示的に列挙する（許可リスト方式）。
  以前は版フォルダをそのまま固めていたが、この方式では rubigui.log が
  配布物に入る。ログには実在の教材名とローカルパス（ユーザー名を含む）が
  記録されている。.gitignore は zip の作成には効かない。
  除外リスト方式にすると新しい種類のファイルが増えたときに漏れるので、
  許可リストにする。

★辞書（359.8MB）は版フォルダに複製していない。--dic で場所を指定する。
"""
import argparse
import hashlib
import shutil
import zipfile
from dataclasses import dataclass
from pathlib import Path

import gen_notices

HERE = Path(__file__).resolve().parent
REPO = HERE.parent

WORD_DIR = REPO / "RubiGUI_word_v3.1"
PPT_DIR = REPO / "RubiGUI_ppt_v1.3"
DOCS_DIR = REPO / "docs"

BUNDLE_NAME = "RubiGUI_2026-09"
DEFAULT_DIC = REPO / "RubiGUI_word_v3.0" / "system_full.dic"


@dataclass(frozen=True)
class Entry:
    src: Path
    dest: str


# 配布物に入れるものの全量。ここに無いものは入らない。
ALLOWLIST = [
    Entry(WORD_DIR / "RubiGUI_Word_v3.1.exe", "RubiGUI_Word_v3.1.exe"),
    Entry(PPT_DIR / "RubiGUI_PPT_v1.3.exe", "RubiGUI_PPT_v1.3.exe"),
    Entry(WORD_DIR / "RubiGUI_V31.bas", "RubiGUI_V31.bas"),
    Entry(WORD_DIR / "sudachi.json", "sudachi.json"),
    Entry(WORD_DIR / "override.json", "override.json"),
    Entry(WORD_DIR / "ruby_settings.json", "ruby_settings.json"),
    Entry(WORD_DIR / "readme.txt", "readme_Word.txt"),
    Entry(PPT_DIR / "readme.txt", "readme_PPT.txt"),
    Entry(DOCS_DIR / "はじめにお読みください.pdf", "はじめにお読みください.pdf"),
    Entry(HERE / "LICENSE.txt", "LICENSE.txt"),
]

# 混入していたら停止するもの。
FORBIDDEN_SUFFIXES = (".py", ".pyc", ".log", ".spec")
FORBIDDEN_NAMES = ("requirements.txt", "__pycache__")


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def check_no_forbidden(folder):
    """配布物に入ってはいけないものが無いか調べ、あれば停止する。"""
    folder = Path(folder)
    hits = []
    for path in folder.rglob("*"):
        if path.name in FORBIDDEN_NAMES or path.suffix.lower() in FORBIDDEN_SUFFIXES:
            hits.append(str(path.relative_to(folder)))
    if hits:
        raise SystemExit(
            "配布物に入ってはいけないファイルが見つかりました:\n  "
            + "\n  ".join(hits)
            + "\nALLOWLIST を確認してください。"
        )


def build(out_dir, dic):
    """許可リストのファイルだけを集めて配布フォルダを組む。"""
    out_dir = Path(out_dir)
    if out_dir.exists():
        shutil.rmtree(out_dir)
    out_dir.mkdir(parents=True)

    missing = [e.src for e in ALLOWLIST if not e.src.is_file()]
    if missing:
        raise SystemExit(
            "配布物に入れるファイルが見つかりません:\n  "
            + "\n  ".join(str(m) for m in missing)
            + "\nexe が未ビルドなら packaging/build_exe.py を先に実行してください。"
        )

    for entry in ALLOWLIST:
        shutil.copy2(entry.src, out_dir / entry.dest)

    dic = Path(dic)
    if not dic.is_file():
        raise SystemExit(f"辞書が見つかりません: {dic}\n--dic で場所を指定してください。")
    print(f"辞書をコピー中（{dic.stat().st_size / 1e6:.1f} MB）...", flush=True)
    shutil.copy2(dic, out_dir / "system_full.dic")

    gen_notices.generate(out_dir / "THIRD-PARTY-NOTICES.txt")

    check_no_forbidden(out_dir)
    return out_dir


def make_zip(folder):
    """配布フォルダを zip に固める。展開すると folder と同じ名前になる。"""
    folder = Path(folder)
    zip_path = folder.parent / f"{folder.name}.zip"
    if zip_path.exists():
        zip_path.unlink()
    print("zip を作成中...", flush=True)
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for path in sorted(folder.rglob("*")):
            if path.is_file():
                zf.write(path, Path(folder.name) / path.relative_to(folder))
    return zip_path


def main():
    parser = argparse.ArgumentParser(description="RubiGUI の配布物を組み立てる")
    parser.add_argument("--dic", default=str(DEFAULT_DIC), help="system_full.dic の場所")
    parser.add_argument("--out", default=str(REPO / "dist" / BUNDLE_NAME),
                        help="配布フォルダの出力先")
    args = parser.parse_args()

    folder = build(args.out, args.dic)
    zip_path = make_zip(folder)
    digest = sha256(zip_path)

    print("\n=== 組み立て結果 ===")
    print(f"  フォルダ : {folder}")
    print(f"  zip      : {zip_path}  ({zip_path.stat().st_size / 1e6:.1f} MB)")
    print(f"  SHA-256  : {digest}")
    print("\nこの SHA-256 をリリースノートに記載してください。")

    (zip_path.parent / f"{zip_path.name}.sha256.txt").write_text(
        f"{digest}  {zip_path.name}\n", encoding="utf-8"
    )


if __name__ == "__main__":
    main()
```

- [ ] **Step 4: テストを実行して通ることを確認する**

Run: `python -m pytest tests/test_make_dist.py -v`
Expected: 7件すべて PASS（parametrize の3件を含む）

- [ ] **Step 5: 出力先を git 管理から外す**

`.gitignore` に追記する。

```
# 配布物の出力先（make_dist.py が生成する）
dist/
```

- [ ] **Step 6: コミット**

```bash
git add packaging/make_dist.py tests/test_make_dist.py .gitignore
git commit -m "配布物の組み立てスクリプトを追加

許可リストに挙げたファイルだけを集めて zip 化し、SHA-256 を出す。
従来のフォルダごと固める方式では rubigui.log が配布物に入り、
実在の教材名とローカルパスが外部に出る状態だった。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 6: readme の外部向け改訂と配布用の既定値

**Files:**
- Modify: `RubiGUI_word_v3.1/readme.txt`
- Modify: `RubiGUI_ppt_v1.3/readme.txt`
- Modify: `RubiGUI_word_v3.1/override.json`
- Modify: `RubiGUI_ppt_v1.3/override.json`
- Modify: `RubiGUI_word_v3.1/RubiGUI_V31.bas`（先頭コメントの版表記のみ）

**Interfaces:**
- Consumes: なし
- Produces: 配布用の readme と既定の `override.json`。`make_dist.py` の `ALLOWLIST` が参照する。

- [ ] **Step 1: 動作確認用の辞書データを配布用の既定値に差し替える**

現在の `override.json` は動作確認用（`ひつまぶし` / `名古屋` / `今日`）。書式の見本になる語に差し替える。**空にはしない**（書き方が分からず問い合わせになるため）。両版とも同じ内容にする。

```json
{
  "河内": "かわち",
  "大和": "やまと",
  "分別": "ふんべつ"
}
```

**配布物に入るのは Word 側の `override.json` だけ**である（同梱すると両版が
同じファイルを共有するため1つで足りる）。PPT 側のものは、版フォルダ単体でも
動かせるようにしておくための開発用の控えである。内容がずれると紛らわしいので
同じ内容にしておく。

readme の `override.json` の節に、この3語を選んだ理由を1行足す。

```
※ 最初から入っている3語は書き方の見本です。いずれも文脈によって読みが
   変わる語で、辞書の推定が外れやすいものを選んでいます。不要であれば
   辞書編集画面から削除してかまいません。
```

- [ ] **Step 2: 配布物の中に残っている旧版の表記を更新する**

版フォルダは v3.0 / v1.2 からのコピーで作られているため、**中身の版表記が古いまま**である。
このまま配ると、v3.1 のフォルダを受け取った人が `RubiGUI_V30.bas` を探すことになる。

`RubiGUI_word_v3.1/readme.txt`:

| 現在 | 変更後 |
|---|---|
| `v3.0` / `RubiGUI_V3.0.py` | `v3.1` / `RubiGUI_V3.1.py` |
| `RubiGUI_V30.bas` | `RubiGUI_V31.bas` |
| モジュール名 `RubiGUI_V30` | `RubiGUI_V31` |
| マクロ名 `InsertFuriganaFromTSV_V30` | `InsertFuriganaFromTSV_V31` |
| 見出し `■ v2.1 から変更された点` | `■ v3.0 から変更された点`（内容も v3.1 の変更に差し替える） |

`RubiGUI_ppt_v1.3/readme.txt`:

| 現在 | 変更後 |
|---|---|
| `v1.2` / `RubiGUI_PPT_V1.2.py` | `v1.3` / `RubiGUI_PPT_V1.3.py` |

**過去の変更履歴を説明している節（「v2.1 から変更された点」「v2.0 から変更された点」の
本文）は、そこが履歴である限り版番号を書き換えない。** 書き換えるのは「今の版が何か」を
述べている箇所と、利用者が実際に触るファイル名である。

あわせて `RubiGUI_word_v3.1/RubiGUI_V31.bas` の先頭コメントに残っている `v3.0` を直す
（2箇所: 5行目の版表記、7行目の `RubiGUI_V3.0.py` への参照）。マクロ名は Task 1 で
更新済みなので触らない。

**`.bas` は CP932 なので、バイト単位の `sed` ではなく Python で読み書きする。**
CP932 の2バイト文字の下位バイトが ASCII と衝突しうるため、バイト置換は安全ではない。

**一括置換してはいけない。** 19行目の `v3.0 は…` は v3.0 の挙動を説明する履歴であり、
書き換えると記録が壊れる。直すのは5行目と7行目だけである。

**読み込み側にも `newline=''` が要る。** `read_text` は CRLF を LF に畳むため、
これを付けないと書き戻したときに改行コードが LF only になって壊れる。

```bash
python -c "
p = 'RubiGUI_word_v3.1/RubiGUI_V31.bas'
with open(p, encoding='cp932', newline='') as f:
    lines = f.readlines()
lines[4] = lines[4].replace('v3.0', 'v3.1')          # 5行目: 版表記
lines[6] = lines[6].replace('V3.0.py', 'V3.1.py')    # 7行目: 呼び出し元の参照
with open(p, 'w', encoding='cp932', newline='') as f:
    f.writelines(lines)
"
```

書き換え後、`v3.0` は**1件**（19行目の履歴）、`v3.1` は**1件**（5行目）になる。
7行目は `RubiGUI_V3.1.py`（大文字 V）なので小文字の集計には入らない。

- [ ] **Step 3: 職場前提・開発者向けの記述を落とす**

`RubiGUI_word_v3.1/readme.txt` から次の節を削除する。

- `● exe にする場合（PyInstaller）`（`■ 起動方法とコンソール画面について` の中）

あわせて readme 全体を通読し、次に当てはまる記述を洗い出して削除または書き換える。

- ソースコードから実行することを前提にした説明（配布物に `.py` は入らない）
- 職場の環境・フォルダ構成・人物を前提にした説明
- `requirements.txt` や pip への言及

`RubiGUI_ppt_v1.3/readme.txt` についても同じ作業を行う。

- [ ] **Step 4: サポート方針の節を両方の readme の冒頭に足す**

版名の直後に置く。文面は両版で同じにする（ログのファイル名だけ読み替える）。

```
------------------------------------------------------------
■ このソフトについて（必ずお読みください）
------------------------------------------------------------

RubiGUI は個人が私的に作成し、無償で配布しているソフトウェアです。
MITライセンスで提供しています（同梱の LICENSE.txt を参照）。

● 無保証です
  このソフトは「現状のまま」提供されます。動作すること、目的に合うことを
  保証しません。万一ファイルが破損しても責任を負いかねます。
  ★処理する前に、必ず元ファイルのバックアップを取ってください。

● ルビを振るとレイアウトが変わることがあります
  ルビを表示する高さを確保するため、行間を広げる処理を行います。
  そのぶん内容が下へ動き、ページ送りが変わることがあります。
  余裕を持った作りの文書でお使いください。

● サポートについて
  不具合の報告や要望は歓迎しますが、対応をお約束するものではありません。
  作者が可能な範囲で対応します。緊急の対応はできません。

● 問い合わせ先
  まず、配布元の取りまとめの方へご連絡ください。
  作者への連絡が必要な場合は、取りまとめの方を通じてお願いします。

  ★このソフトは作者が私的に作成したものです。
    作者の勤務先は配布主体ではなく、問い合わせ先でもありません。
    勤務先へのお問い合わせはご遠慮ください。

● 不具合を報告するとき
  次の3つを添えていただけると、調査がはるかに早くなります。
    1. タイトルバーに出ている版番号（例: ルビ編集ツール v3.1）
    2. このソフトのフォルダにある rubigui.log
       （PowerPoint版は rubigui_ppt.log）
    3. 何をしたら何が起きたか
```

- [ ] **Step 5: 落とし忘れが無いか確認する**

```bash
grep -n "PyInstaller\|requirements\|pip install" RubiGUI_word_v3.1/readme.txt RubiGUI_ppt_v1.3/readme.txt
grep -c "RubiGUI_V30\.bas\|InsertFuriganaFromTSV_V30" RubiGUI_word_v3.1/readme.txt
python -c "from pathlib import Path; t=Path('RubiGUI_word_v3.1/RubiGUI_V31.bas').read_text(encoding='cp932'); print('v3.0:', t.count('v3.0'), 'v3.1:', t.count('v3.1'))"
```

Expected:
- 1つ目（職場前提の記述）: 何も出力されない
- 2つ目: **`2`**。いずれも残って正しいもの。1つは「v3.0→v3.1 で名前が変わった」という
  移行案内の対比表（消すと手順が意味不明になる）、もう1つは過去の変更を説明する履歴の節
- 3つ目: **`v3.0: 1 v3.1: 1`**。`v3.0` の1件は19行目の履歴

- [ ] **Step 6: コミット**

```bash
git add RubiGUI_word_v3.1/readme.txt RubiGUI_ppt_v1.3/readme.txt RubiGUI_word_v3.1/override.json RubiGUI_ppt_v1.3/override.json
git commit -m "readme を外部配布向けに改訂

サポート方針・免責・問い合わせ先を冒頭に明記した。勤務先が配布主体でも
問い合わせ先でもないことを明示している。開発者向けの記述を削除し、
override.json を動作確認用のデータから書式の見本に差し替えた。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 7: 導入手順書（はじめにお読みください.pdf）

**Files:**
- Create: `docs/はじめにお読みください.docx`（編集元）
- Create: `docs/はじめにお読みください.pdf`（配布物に入れる）

**Interfaces:**
- Consumes: なし
- Produces: `docs/はじめにお読みください.pdf`。`make_dist.py` の `ALLOWLIST` が参照する。

**このタスクは手作業である。** スクリーンショットの撮影が必要なため自動化できない。Word で作成し、PDF として出力する（配布相手も Word 使いなので違和感がない）。

- [ ] **Step 1: 章立てのとおりに Word 文書を作る**

つまずく順に並べる。既存 readme と違い、**最初の30分に必要なことだけ**に絞る。

```
RubiGUI 導入の手引き（2026年9月版）
  対象: RubiGUI Word版 v3.1 / PowerPoint版 v1.3

1. はじめに
   ・このソフトは何をするものか（3行）
   ・無保証であること、元ファイルのバックアップを取ること
   ・PowerPoint版の方が導入が簡単なので、まず試すならこちら

2. 準備（共通）
   2-1. zip は必ず展開してから使う
        ★zip の中から直接 exe を実行すると、辞書が見つからず起動に失敗します
   2-2. 置き場所
        ★OneDrive の同期フォルダに置かないでください（例: C:\RubiGUI に置く）
        同期のタイミングで作業用ファイルが消えずに残ることがあります
   2-3. ダウンロードと展開には時間がかかります
        辞書ファイルが 360MB あるため、数分かかることがあります

3. 初回起動時の警告について（スクリーンショット）
   ・「Windows によって PC が保護されました」と出た場合
   ・「詳細情報」→「実行」の順にクリック
   ★このソフトには発行元の署名が付いていないため、この警告が出ます

4. PowerPoint版を使う
   ・RubiGUI_PPT_v1.3.exe をダブルクリック
   ・ファイルをドラッグ＆ドロップ
   ・準備作業はありません

5. Word版を使う（マクロの導入が必要）
   5-1. なぜマクロが必要か（1段落）
   5-2. 導入手順（スクリーンショット必須）
        (1) Word を開き Alt+F11 でVBAエディタを表示
        (2) 古い RubiGUI_V30 や Module1 があれば削除
        (3) Normal を右クリック →「ファイルのインポート」
        (4) RubiGUI_V31.bas を選ぶ
        (5) Word を閉じる（保存を聞かれたら「保存する」）
   5-3. うまくいかないとき
        ・「マクロが見つかりません」→ 手順(3)のインポート先が Normal か確認

6. 困ったときは
   ・タイトルバーの版番号を控える
   ・rubigui.log（PowerPoint版は rubigui_ppt.log）を用意する
   ・取りまとめの方へ連絡する
```

- [ ] **Step 2: スクリーンショットを撮って貼る**

必要な画像は次の5枚。

1. SmartScreen の警告画面（「詳細情報」が見えている状態）
2. SmartScreen の「実行」ボタンが出た状態
3. VBAエディタの全体（`Normal` が見えている状態）
4. `Normal` を右クリックして「ファイルのインポート」を選ぶところ
5. インポート後、`RubiGUI_V31` がモジュール一覧に出ている状態

- [ ] **Step 3: PDF として出力する**

Word で「名前を付けて保存」→ ファイルの種類「PDF」→ `docs/はじめにお読みください.pdf`

- [ ] **Step 4: 内容を読み返す**

配布相手は RubiGUI を見たことがない人である。次を確認する。

- 職場の固有名詞・人名・実在の教材名が入っていないか（**スクリーンショットの中も確認する**。ファイル一覧やタイトルバーに実名が写り込みやすい）
- 「いつもの」「例のフォルダ」のような、文脈を前提にした表現が無いか
- 手順どおりに操作して実際に導入できるか（1回通しで試す）

- [ ] **Step 5: コミット**

```bash
git add "docs/はじめにお読みください.docx" "docs/はじめにお読みください.pdf"
git commit -m "導入手順書を追加

外部配布向けに、最初の30分に必要なことだけを別紙にまとめた。
つまずく順（展開・置き場所・SmartScreen・マクロ導入）に並べている。

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

---

### Task 8: 統合検証

**Files:**
- Modify: `readme.md`
- Modify: `docs/CHANGELOG.md`

**Interfaces:**
- Consumes: Task 1〜7 のすべて
- Produces: 検証済みの配布物 zip と SHA-256

- [ ] **Step 1: すべてのテストを実行する**

Run: `python -m pytest tests/ -v`
Expected: Task 1〜5 のテスト（23件）がすべて PASS

- [ ] **Step 2: exe をビルドする**

```bash
python packaging/build_exe.py
```

Expected: `RubiGUI_Word_v3.1.exe` と `RubiGUI_PPT_v1.3.exe` がそれぞれ約 17.8MB で出力される。**120MB を超えていたら辞書が同梱されている**（`packaging/hooks/hook-sudachipy.py` が効いているか確認する）。

- [ ] **Step 3: 配布物を組み立てる**

```bash
python packaging/make_dist.py
```

Expected: `dist/RubiGUI_2026-09.zip` が生成され、SHA-256 が表示される。サイズは約 396MB（辞書359.8MB + exe 17.8MB×2）。

- [ ] **Step 4: 配布物の中身を確認する**

```bash
python -c "import zipfile; [print(n) for n in sorted(zipfile.ZipFile('dist/RubiGUI_2026-09.zip').namelist())]"
```

Expected: 12ファイル。`.py` / `.log` / `__pycache__` / `requirements.txt` が**1つも無いこと**を目視で確認する。

- [ ] **Step 4b: exe に開発環境のパスが埋め込まれていないか確認する**

PyInstaller がビルド時のパスを実行ファイルに残すことがある。既存の v3.0 / v1.2 の exe を
調べた限りでは埋め込まれていなかったが、これはビルドの呼び出し方に依存する性質なので、
**実際に配る成果物に対して確認する**。

```bash
python -c "
from pathlib import Path
for name in ['RubiGUI_word_v3.1/RubiGUI_Word_v3.1.exe', 'RubiGUI_ppt_v1.3/RubiGUI_PPT_v1.3.exe']:
    b = Path(name).read_bytes()
    hits = [w for w in ['milli','OneDrive','Desktop','RubiGUI_v2']
            if w.encode('ascii') in b or w.encode('utf-16-le') in b]
    print(name, '→', hits or 'クリーン')
"
```

Expected: 両方とも `クリーン`

- [ ] **Step 5: クリーンな環境で展開して起動する**

OneDrive の同期対象外のフォルダ（例 `C:\RubiGUI_test`）に展開する。

- [ ] **`はじめにお読みください.pdf` のファイル名が文字化けせずに表示される**
      zip 内の日本語ファイル名は UTF-8 フラグ（bit 11）付きで格納される。最近の
      Windows エクスプローラーは正しく扱うが、古い Windows や一部の展開ツール
      （旧版の 7-Zip / Lhaplus 等）では文字化けする実績がある。**受け取った人が
      最初に読む文書のファイル名**なので、ここが化けると出だしで躓く。
      配布先に近い環境で確認できるとなお良い（取りまとめ役の方に依頼してもよい）
- [ ] `RubiGUI_PPT_v1.3.exe` が起動し、タイトルバーに `v1.3` と出る
- [ ] `RubiGUI_Word_v3.1.exe` が起動し、タイトルバーに `v3.1` と出る
- [ ] 辞書編集画面を開き、`河内` `大和` `分別` の3件が読み込まれている（0件なら設定ファイルの読み先がずれている）

- [ ] **Step 6: 設定衝突の回帰テストを実機で行う**

これは Task 2 の単体テストが担保している内容を、実際の配布物で確かめるもの。

- [ ] PPT版を起動し、行間を `2.0`、「タイトル枠も対象にする」をONにして閉じる
- [ ] Word版を起動し、ルビの大きさを変えて保存し、閉じる
- [ ] PPT版を再度起動し、**行間が 2.0 のまま、タイトル枠もONのまま**であることを確認

- [ ] **Step 7: Wordマクロを導入して1ファイル処理する**

導入手順書（Task 7）のとおりに操作する。**手順書の検証も兼ねている。**

- [ ] `RubiGUI_V31.bas` を Normal へインポートできる
- [ ] 表を含む文書を1ファイル処理し、ルビが振られる
- [ ] `rubigui.log` の先頭に `RubiGUI Word版 v3.1 起動` が記録されている

- [ ] **Step 8: フィクスチャで出力を照合する**

```bash
python tests/make_fixture_word.py
python tests/check_word_output.py
python tests/make_fixture_ppt.py
python tests/check_ppt_output.py
```

Expected: v3.0 / v1.2 と同じ結果（Word版17項目、PPT版は `include_title` の ON/OFF 両方）。**今回の修正は設定の保存とログだけなので、出力が変わってはいけない。**

- [ ] **Step 9: リポジトリの readme と CHANGELOG を更新する**

`readme.md` の「最新版」の表を v3.1 / v1.3 に更新し、配布手順の節を足す。

```markdown
## ■ 配布物の作り方

    python packaging/build_exe.py     # exe をビルド
    python packaging/make_dist.py     # 配布物を組んで zip 化

`dist/RubiGUI_2026-09.zip` と SHA-256 が出力される。
配布物には許可リスト（`packaging/make_dist.py` の `ALLOWLIST`）に
挙げたファイルだけが入る。**版フォルダをそのまま固めてはいけない**
（rubigui.log に実在の教材名とローカルパスが記録されているため）。
```

`docs/CHANGELOG.md` に v3.1 / v1.3 の節を足す。設定保存のマージ化・起動ログ・マクロ改名・外部配布向けの整備を記載する。

- [ ] **Step 10: コミット**

```bash
git add readme.md docs/CHANGELOG.md
git commit -m "Word v3.1 / PPT v1.3: 外部配布向けの整備を完了

Co-Authored-By: Claude Opus 5 <noreply@anthropic.com>"
```

- [ ] **Step 11: 配布前の最終確認（人手）**

**この確認が終わるまで公開しない。**

- [ ] 取りまとめ役の方に「御校で GitHub は開けるか」を確認した（開けない場合は Google Drive を正本に切り替える）
- [ ] zip を実際にダウンロードし直し、SHA-256 が記録値と一致した
- [ ] 導入手順書に職場の固有名詞・実在の教材名が入っていない（スクリーンショットの中も含む）
- [ ] `THIRD-PARTY-NOTICES.txt` に UniDic の著作権表示が入っている
- [ ] 不具合報告用の Google フォームを作成し、取りまとめ役の方へ URL を伝えた
      （項目: 版番号 / OS / Office の版 / ログの添付 / 何をしたら何が起きたか）

---

### Task 9: 公開

**Files:** なし（リポジトリ外の作業）

**Interfaces:**
- Consumes: Task 8 で検証済みの `dist/RubiGUI_2026-09.zip` と SHA-256
- Produces: 配布可能な URL

**このタスクはすべて人手で行う。** 外部への公開は取り消しが難しいので、
Task 8 Step 11 のチェックがすべて済んでいることを確認してから着手する。

- [ ] **Step 1: GitHub Release を下書きで作る**

リポジトリ `Milli-Q13/RubiGUI` で新しいリリースを作る。

- タグ: `v2026-09`
- タイトル: `RubiGUI 2026年9月版（Word v3.1 / PowerPoint v1.3）`
- `dist/RubiGUI_2026-09.zip` を添付する（約396MB。Releases は1ファイル2GBまで置ける）
- **下書き（Draft）のまま保存する。** この時点では公開しない

リリースノートに次を記載する。

```
Word版 v3.1 / PowerPoint版 v1.3

■ 同梱物
  RubiGUI_Word_v3.1.exe / RubiGUI_PPT_v1.3.exe / 辞書 / 導入の手引き

■ はじめての方へ
  展開したフォルダの「はじめにお読みください.pdf」からお読みください。
  ★zip の中から直接 exe を実行しないでください（辞書が見つからず起動に失敗します）

■ 変更点
  ・Word版とPowerPoint版を1つにまとめ、辞書と読みの登録（override.json）を共有できるようにした
  ・両版の設定が互いを打ち消さないようにした
  ・ログに版番号を記録するようにした
  ・Wordマクロの名前が RubiGUI_V31 に変わりました（入れ替えが必要です）

■ SHA-256（RubiGUI_2026-09.zip）
  <make_dist.py が出力した値をここに貼る>

■ ライセンス
  MIT License（同梱の LICENSE.txt）
  同梱している第三者ソフトウェアの権利表記は THIRD-PARTY-NOTICES.txt を参照してください。
```

- [ ] **Step 2: 同じ zip を Google Drive にも置く**

GitHub が開けない環境向けのミラー。**Task 8 で検証したものと同一のファイルを置く**
（作り直さない。作り直すと SHA-256 が変わる）。共有リンクは「リンクを知っている全員が閲覧可」にする。

- [ ] **Step 3: 取りまとめ役の方に連絡する**

伝える内容は次の4つ。

1. ダウンロード先（GitHub / Google Drive の両方）
2. SHA-256 の値と、照合の必要は無い旨（心配な方向けの情報であること）
3. 不具合報告フォームの URL
4. **まず「はじめにお読みください.pdf」を読んでもらいたいこと**

- [ ] **Step 4: 公開する**

取りまとめ役の方から「アクセスできた」と確認が取れてから、GitHub Release の
下書きを公開に切り替える。

- [ ] **Step 5: Obsidian の進捗ログに記録する**

`D:\obsidian\プロジェクト\RubiGUI\RubiGUI.md` の「進捗ログ」に、
配布日・配布先・配布した版・SHA-256 を1行で残す。
次に不具合報告が来たとき、誰がどの版を持っているかの手がかりになる。
