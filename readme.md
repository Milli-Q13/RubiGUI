# RubiGUI Repository（自分用メモ）

このリポジトリは、RubiGUI（Word版／PPT版）のソースコードと設定ファイルを保管するための自分用メモ置き場。  
完成した exe は GitHub には置かず、直接配布する。

---

## ■ フォルダ構成

```
RubiGUI/
 ├─ RubiGUI_word_v2.0/      # Word版（旧・保管用）
 ├─ RubiGUI_word_v2.1/      # Word版 最新（ソースコード＋設定ファイル＋Module1.bas）
 ├─ RubiGUI_ppt_v1.0/       # PPT版（旧・保管用）
 ├─ RubiGUI_ppt_v1.1/       # PPT版 最新（ソースコード＋設定ファイル）
 ├─ docs/
 │   └─ CHANGELOG.md        # バージョンごとの変更履歴
 └─ readme.md               # このファイル
```

※ exe と dic は GitHub に置かない（.gitignore で除外）

### ● 最新版

| 版 | バージョン | 本体 | 備考 |
|---|---|---|---|
| Word | v2.1 | `RubiGUI_word_v2.1/RubiGUI_V2.1.py` | **Module1.bas の入れ替えが必須**（マクロ名が `InsertFuriganaFromTSV_V21` に変更） |
| PPT  | v1.1 | `RubiGUI_ppt_v1.1/RubiGUI_PPT_V1.1.py` | 既定のルビ範囲が「初回のみ」に変更（v1.0 は全出現） |

変更内容の詳細は [docs/CHANGELOG.md](docs/CHANGELOG.md) を参照。

---

## ■ exe のビルド

```
python packaging/build_exe.py          # 両方
python packaging/build_exe.py word     # Word版だけ
python packaging/build_exe.py ppt      # PPT版だけ
```

exe は各バージョンのフォルダに出力される（`sudachi.json` などと同じ場所）。
1ファイル約 18MB。フォルダごと配布する。

### ● 注意：`--collect-all sudachipy` は使わない

PyInstaller 標準の sudachipy フックは、インストール済みの
`sudachidict_full`（359MB）と `sudachidict_core`（217MB）を無条件に同梱し、
exe が **215MB** に膨らむ。RubiGUI はこれらを使わず、exe と同じフォルダの
`system_full.dic` を実行時に読む。

`packaging/hooks/hook-sudachipy.py` で標準フックを上書きしてこれを防いでいる。
`build_exe.py` はこのフックを使うので、必ずスクリプト経由でビルドすること。

---

## ■ GitHub 運用メモ（自分用）

### ● 基本操作
```
git add .
git commit -m "update"
git push
```

### ● 新バージョンを追加するとき
1. 新しいフォルダを作る（例：RubiGUI_Word_v2.1/）
2. ソースコードと設定ファイルを入れる
3. exe は入れない
4. git add → commit → push

### ● 巨大ファイルを入れない
- *.exe と *.dic は .gitignore に入れておく  
- exe は GitHub Releases にも置かない（直接配布）

### ● filter-repo を使ったときの注意
filter-repo 実行後に origin が消えることがある  
→ その場合は再登録する

```
git remote add origin https://github.com/Milli-Q13/RubiGUI.git
git push origin main --force
```

---

## ■ RubiGUI（Word版 / PPT版）について

各バージョンの使い方は、それぞれのフォルダ内の readme.txt に記載。  
GitHub ではコードと設定ファイルのみ管理する。
