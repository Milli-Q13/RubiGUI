"""PyInstaller 標準の sudachipy フックを差し替えるためのフック。

★なぜ必要か
  pyinstaller-hooks-contrib の hook-sudachipy.py は、インストールされている
  sudachidict_small / sudachidict_core / sudachidict_full を見つけると、
  その辞書ファイルを無条件で exe へ同梱する。
  実測では sudachidict_full（359.8MB）と sudachidict_core（217.1MB）が
  取り込まれ、exe が 215MB に膨らんでいた。

  RubiGUI は辞書パッケージを使わない。exe と同じフォルダに置いた
  sudachi.json / system_full.dic を実行時に読む方式なので、
  同梱すると「使われない辞書」を配布することになる。

  --exclude-module では止められない（フックが collect_data_files を
  直接呼んでいるため）。--additional-hooks-dir はこのフォルダを
  標準フックより先に探すので、ここへ同名のフックを置いて上書きする。

★同梱するもの
  sudachipy/resources/ の char.def / rewrite.def / unk.def / sudachi.json。
  これらは形態素解析の実行時に必要（合計15KB程度）。
"""
from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files("sudachipy")

# 0.6.8 以降、バイナリ拡張から参照されるため明示が必要
hiddenimports = [
    "sudachipy.config",
    "sudachipy.errors",
]

# sudachidict_* は意図的に同梱しない
