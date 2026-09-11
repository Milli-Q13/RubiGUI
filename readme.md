# RubiGUI Repository（自分用メモ）

このリポジトリは、RubiGUI（Word版／PPT版）のソースコードと設定ファイルを保管するための自分用メモ置き場。  
exe と辞書は git 管理外。配布物は `packaging/make_dist.py` で組み立てる。

---

## ■ フォルダ構成

```
RubiGUI/
 ├─ RubiGUI_word_v2.0/      # Word版（旧・保管用）
 ├─ RubiGUI_word_v2.1/      # Word版（旧・保管用）
 ├─ RubiGUI_word_v3.0/      # Word版（旧・職場配布中）
 ├─ RubiGUI_word_v3.1/      # Word版 最新（ソース＋設定＋RubiGUI_V31.bas）
 ├─ RubiGUI_ppt_v1.0/       # PPT版（旧・保管用）
 ├─ RubiGUI_ppt_v1.1/       # PPT版（旧・保管用）
 ├─ RubiGUI_ppt_v1.2/       # PPT版（旧・職場配布中）
 ├─ RubiGUI_ppt_v1.3/       # PPT版 最新（ソース＋設定）
 ├─ packaging/
 │   ├─ build_exe.py        # exe のビルド
 │   ├─ make_dist.py        # 配布物の組み立て（許可リスト方式）
 │   ├─ gen_notices.py      # 第三者ライセンス告知の生成
 │   ├─ LICENSE.txt         # MIT（配布物に同梱する正本）
 │   └─ licenses/           # 自動収集できないライセンス本文の手置き場
 ├─ tests/                  # pytest（41件）＋ フィクスチャ生成・照合スクリプト
 ├─ docs/
 │   ├─ CHANGELOG.md        # バージョンごとの変更履歴
 │   └─ はじめにお読みください.docx / .pdf   # 配布物に同梱する導入手順書
 ├─ LICENSE                 # MIT（リポジトリ用）
 └─ readme.md               # このファイル
```

※ exe と dic は GitHub に置かない（.gitignore で除外）

### ● 最新版

| 版 | バージョン | 本体 | 備考 |
|---|---|---|---|
| Word | v3.1 | `RubiGUI_word_v3.1/RubiGUI_V3.1.py` | **`RubiGUI_V31.bas` の入れ替えが必須**（マクロ名が `InsertFuriganaFromTSV_V31` に変更） |
| PPT  | v1.3 | `RubiGUI_ppt_v1.3/RubiGUI_PPT_V1.3.py` | マクロ不要 |

v3.1 / v1.3 は**外部配布向けの版**。Word版とPPT版を1つの配布物にまとめ、辞書と
`override.json` を共有する。そのため両版の設定保存がマージ方式になっている。

変更内容の詳細は [docs/CHANGELOG.md](docs/CHANGELOG.md) を参照。

---

## ■ exe のビルド

```
python packaging/build_exe.py          # 両方
python packaging/build_exe.py word     # Word版だけ
python packaging/build_exe.py ppt      # PPT版だけ
```

exe は各バージョンのフォルダに出力される。1ファイル約 17.8MB。

PyInstaller の作業フォルダはリポジトリの外（システムの一時領域）に置いている。
**このリポジトリは OneDrive 配下にあり、作業フォルダを中に置くと同期と衝突して
`PermissionError` でビルドが落ちる。**

### ● 注意：`--collect-all sudachipy` は使わない

PyInstaller 標準の sudachipy フックは、インストール済みの
`sudachidict_full`（359MB）と `sudachidict_core`（217MB）を無条件に同梱し、
exe が **215MB** に膨らむ。RubiGUI はこれらを使わず、exe と同じフォルダの
`system_full.dic` を実行時に読む。

`packaging/hooks/hook-sudachipy.py` で標準フックを上書きしてこれを防いでいる。
`build_exe.py` はこのフックを使うので、必ずスクリプト経由でビルドすること。

---

## ■ 配布物の作り方

```
python packaging/build_exe.py                               # exe をビルド
python packaging/make_dist.py --out "C:/RubiGUI_dist/RubiGUI_2026-09"
```

zip と SHA-256 が出力される。**この SHA-256 をリリースノートに載せる。**

### ● 組み立て先はリポジトリの外にする

既定の出力先 `dist/` はリポジトリ内＝OneDrive 内なので、359.8MB を書いては消す
処理が同期と衝突し、`PermissionError` で中断する。**`--out` で OneDrive の外を
指定すること。**

### ● 出力をパイプしない

`| tail` などに繋ぐと終了コードが最後のコマンドのものになり、失敗しても成功に
見える。実際にこれで中断を見落とした。

### ● 配布物には許可リストにあるものだけが入る

`packaging/make_dist.py` の `ALLOWLIST` に挙げたファイルだけが配布物に入る。
**版フォルダをそのまま zip に固めてはいけない。** `rubigui.log` には実在の教材名と
ローカルパス（ユーザー名を含む）が記録されており、`.gitignore` は zip の作成には
効かない。

`THIRD-PARTY-NOTICES.txt` は毎回生成される。同梱物の告知を1つでも用意できない
場合はその場で停止する（法的義務を欠いたまま配布しないため）。

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
