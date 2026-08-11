"""RubiGUI（Word版 / PowerPoint版）の exe をまとめてビルドする。

使い方（リポジトリのどこからでも可）:
    python packaging/build_exe.py            # 両方ビルド
    python packaging/build_exe.py word       # Word版だけ
    python packaging/build_exe.py ppt        # PowerPoint版だけ

出力先は各バージョンのフォルダ（sudachi.json や system_full.dic と同じ場所）。
exe はそのフォルダの設定ファイルを読むので、フォルダごと配布する。

★ビルドの要点
  ・--noconsole    … 黒いコンソール画面を出さない
  ・--onefile      … 配布物を exe 1つにまとめる
  ・--collect-all tkinterdnd2
                   … ドラッグ＆ドロップに必要な tkdnd のバイナリを同梱する。
                     これが無いと起動時に TclError で落ちる。
  ・--additional-hooks-dir packaging/hooks
                   … sudachidict_* が同梱されて exe が 215MB に膨らむのを防ぐ。
                     詳細は packaging/hooks/hook-sudachipy.py を参照。
  ・--hidden-import win32timezone
                   … pywin32 が実行時に動的 import することがあるため。
"""
import argparse
import shutil
import subprocess
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
HOOKS = REPO / "packaging" / "hooks"
WORK = REPO / "build"

TARGETS = {
    "word": {
        "script": REPO / "RubiGUI_word_v2.1" / "RubiGUI_V2.1.py",
        "name": "RubiGUI_V2.1",
    },
    "ppt": {
        "script": REPO / "RubiGUI_ppt_v1.1" / "RubiGUI_PPT_V1.1.py",
        "name": "RubiGUI_PPT_V1.1",
    },
}


def build(key):
    target = TARGETS[key]
    script = target["script"]
    if not script.exists():
        raise SystemExit(f"スクリプトが見つかりません: {script}")
    dist = script.parent

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm", "--clean", "--onefile", "--noconsole",
        "--name", target["name"],
        "--collect-all", "tkinterdnd2",
        "--additional-hooks-dir", str(HOOKS),
        "--hidden-import", "win32timezone",
        "--exclude-module", "sudachidict_core",
        "--exclude-module", "sudachidict_full",
        "--exclude-module", "sudachidict_small",
        "--workpath", str(WORK / key),
        "--specpath", str(WORK),
        "--distpath", str(dist),
        str(script),
    ]
    print(f"\n=== {key} をビルド中 ===\n{' '.join(cmd)}\n", flush=True)
    started = time.time()
    proc = subprocess.run(cmd, cwd=REPO)
    if proc.returncode != 0:
        raise SystemExit(f"{key} のビルドに失敗しました（終了コード {proc.returncode}）")

    exe = dist / f"{target['name']}.exe"
    size = exe.stat().st_size / 1e6
    print(f"完了: {exe}  {size:.1f} MB  ({time.time() - started:.0f}秒)", flush=True)
    if size > 120:
        print("  ※想定より大きいです。sudachidict_* が同梱されていないか確認してください。",
              flush=True)
    return exe


def main():
    parser = argparse.ArgumentParser(description="RubiGUI の exe をビルドする")
    parser.add_argument("targets", nargs="*", choices=list(TARGETS) + [],
                        help="省略すると両方ビルドします")
    args = parser.parse_args()
    keys = args.targets or list(TARGETS)

    built = [build(k) for k in keys]

    print("\n=== ビルド結果 ===")
    for exe in built:
        print(f"  {exe.stat().st_size / 1e6:>7.1f} MB  {exe}")
    print("\n配布するときは、exe と同じフォルダに次を入れてください:")
    print("  sudachi.json / system_full.dic / override.json / ruby_settings.json / readme.txt")
    print("  （Word版は Module1.bas も）")

    if WORK.exists():
        shutil.rmtree(WORK, ignore_errors=True)


if __name__ == "__main__":
    main()
