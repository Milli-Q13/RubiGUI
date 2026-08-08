# RubiGUI Repository（自分用メモ）

このリポジトリは、RubiGUI（Word版／PPT版）のソースコードと設定ファイルを保管するための自分用メモ置き場。  
完成した exe は GitHub には置かず、直接配布する。

---

## ■ フォルダ構成

```
RubiGUI/
 ├─ RubiGUI_Word_v2.0/      # Word版（ソースコード＋設定ファイル）
 ├─ RubiGUI_PPT_v1.0/       # PPT版（ソースコード＋設定ファイル）
 ├─ docs/                   # 使い方メモなど
 └─ README.md               # このファイル
```

※ exe と dic は GitHub に置かない（.gitignore で除外）

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
