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
