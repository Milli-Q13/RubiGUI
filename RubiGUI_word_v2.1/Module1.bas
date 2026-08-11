Attribute VB_Name = "Module1"
Option Explicit

'============================================================
' RubiGUI（Word版）v2.1  ルビ付与マクロ
'
' RubiGUI_V2.1.py から次の形で呼び出される：
'   word.Run("InsertFuriganaFromTSV_V21", 元ファイル, TSVファイル, 出力ファイル)
'
' ★v2.0 からの主な変更点
'  1) 「すべての出現箇所」にルビを振れるようにした。
'     v2.0 は Range を作り直すたびに .Find の条件が初期化されてしまい、
'     2回目以降の検索が成立せず、最初の1箇所にしかルビが振られなかった。
'     v2.1 は検索のたびに条件を設定し直したうえで、処理を2段階に分ける。
'       第1段階：文書を変更せずに「どこへ何のルビを振るか」を決める。
'                確定した文字位置は occupied に記録し、重なる候補は捨てる。
'       第2段階：決めた位置へ、文書の末尾から前へ向かってルビを振る。
'     PhoneticGuide は対象文字列をルビ付きの構造へ置き換えるため、実行の
'     たびに以降の文字位置がずれる。先に全部の位置を決めて後ろから詰めれば、
'     未処理側（前方）の位置は動かない。詳しくは InsertAllGuides を参照。
'  2) 一時ファイル（_temp_〇〇.docx）を元ファイルの近くに作るのをやめた。
'     v2.0 は fso.GetParentFolderName(doc.Path) を使っており、これは
'     「文書のあるフォルダ」ではなく「そのさらに親フォルダ」を返すため、
'     想定外の場所に _temp_ ファイルが残っていた。
'     v2.1 は出力フォルダの中に作業用のコピー（~rubigui_〇〇.docx）を作り、
'     それだけを編集して、成功したときにだけ正式な出力名へ差し替える。
'     元ファイルには一切書き込まない。
'  3) パスを推測せず、すべて引数で受け取る。
'     v2.0 は WScript.Shell の SpecialFolders("Desktop") でデスクトップを
'     推測していたが、Python側はレジストリの User Shell Folders から
'     引いており、OneDrive でデスクトップがリダイレクトされている環境では
'     両者が食い違うことがあった。
'  4) TSV を UTF-8 で読む（ADODB.Stream）。
'     VBA の Open ... For Input は ANSI 固定のため、v2.0 では機種依存文字を
'     含む語句を Python 側で除外せざるを得なかった。
'  5) エラー時に MsgBox を出さず Err.Raise で Python へ返す。
'     MsgBox はクリックされるまで戻らないので、一括処理の途中で止まっていた。
'============================================================

' ADODB.Stream 用の定数（参照設定なしで使えるよう明示的に定義）
Private Const ADO_TYPE_TEXT As Long = 2
Private Const ADO_READ_LINE As Long = -2
Private Const ADO_LF As Long = 10

' 既定値（TSVの設定行が無い・壊れている場合に使う）
Private Const DEFAULT_MODE As String = "first"
Private Const DEFAULT_RATIO As Double = 50#
Private Const DEFAULT_OFFSET As Double = 0#

' 同一語句の検索ループが万一抜けられなくなったときの安全弁
Private Const MAX_HITS_PER_TERM As Long = 20000


Public Sub InsertFuriganaFromTSV_V21(ByVal srcPath As String, _
                                     ByVal tsvPath As String, _
                                     ByVal outPath As String)

    Dim fso As Object
    Dim docNew As Document
    Dim terms As Collection
    Dim rubyMode As String
    Dim rubyRatio As Double
    Dim rubyOffset As Double
    Dim savedScreenUpdating As Boolean
    Dim applied As Long
    Dim skipped As Long
    Dim errMessage As String
    Dim workPath As String
    ' ルビ付与が終わって作業用コピーの保存まで済んだか。
    ' ★これが True になった後は、失敗しても作業用コピーを消してはいけない。
    '   完成したルビ付き文書がそこにしか無い状態があり得るため。
    Dim rubyDone As Boolean

    workPath = ""
    rubyDone = False
    rubyMode = DEFAULT_MODE
    rubyRatio = DEFAULT_RATIO
    rubyOffset = DEFAULT_OFFSET
    savedScreenUpdating = Application.ScreenUpdating

    On Error GoTo ErrHandler

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FileExists(srcPath) Then
        Err.Raise vbObjectError + 901, "RubiGUI", _
            "元のWordファイルが見つかりません: " & srcPath
    End If
    If Not fso.FileExists(tsvPath) Then
        Err.Raise vbObjectError + 902, "RubiGUI", _
            "TSVファイルが見つかりません: " & tsvPath & vbCrLf & _
            "先に「TSV保存」を実行してください。"
    End If

    ' --- TSV（設定行＋語句）を読み込む ---
    Set terms = LoadTsv(tsvPath, rubyMode, rubyRatio, rubyOffset)

    ' --- 出力先を用意し、作業用のコピーを %TEMP% に作る ---
    ' ★正式な出力名ではなく作業用の名前へコピーする。正式名へ直接コピーすると、
    '   この後の処理が失敗したときに「ルビの振られていないコピー」が
    '   完成品の顔で出力フォルダに残ってしまう。
    '   置き場所は %TEMP%（同期対象外）。理由は PrepareOutput のコメント参照。
    PrepareOutput fso, srcPath, outPath, workPath

    ' --- 作業用コピーだけを開いて編集する ---
    ' ★重要：Visible:=False で開いてはいけない。
    '   ウィンドウが表示されていない文書に対して PhoneticGuide を呼ぶと、
    '   Word が wwlib.dll でアクセス違反（0xC0000005）を起こして落ちる。
    '   このとき VBA のエラーにもならないので ErrHandler も動かず、
    '   Python 側には「リモート プロシージャ コールに失敗しました
    '   （0x800706BE）」だけが返る。実測：語句0件なら完走、1件でも
    '   ルビを振ろうとすると必ずクラッシュした。
    '   Application 自体は Python 側が非表示にしているので、
    '   文書にウィンドウを持たせても画面には出てこない。
    Application.ScreenUpdating = False
    Set docNew = Documents.Open(FileName:=workPath, _
                                ConfirmConversions:=False, _
                                ReadOnly:=False, _
                                AddToRecentFiles:=False, _
                                Visible:=True)
    ' ルビ付与はアクティブな文書を前提にした処理なので、明示的に前面へ出す
    On Error Resume Next
    docNew.Activate
    On Error GoTo ErrHandler

    ' フィールドコードを検索対象から外す（元の文書に入っているフィールドの
    ' 中身にヒットして、そこへふりがなを差し込んでしまうのを防ぐ）
    On Error Resume Next
    docNew.Windows(1).View.ShowFieldCodes = False
    On Error GoTo ErrHandler

    InsertAllGuides docNew, terms, rubyMode, rubyRatio, rubyOffset, applied, skipped

    docNew.Save
    docNew.Close SaveChanges:=wdDoNotSaveChanges
    Set docNew = Nothing
    rubyDone = True

    ' --- ここまで来て初めて出力先へ書き出す ---
    ' ★ここから先で失敗しても作業用コピーは消さない（rubyDone が True）。
    '   完成したルビ付き文書が作業用コピーにしか無い状態になるため、
    '   消すと成果物が失われる。
    ' ★「削除してから移動」ではなく上書きコピー1回にする。削除と移動の間に
    '   失敗すると、出力が消えたままになる隙間ができるため。
    fso.CopyFile workPath, outPath, True
    fso.DeleteFile workPath, True
    workPath = ""
    rubyDone = False

    Application.ScreenUpdating = savedScreenUpdating

    ' 結果を Python が読めるようファイルに残す（GUIの完了メッセージで使う）
    WriteResult tsvPath, rubyMode, applied, skipped, ""
    Exit Sub

ErrHandler:
    errMessage = Err.Description
    On Error Resume Next
    If Not docNew Is Nothing Then
        docNew.Close SaveChanges:=wdDoNotSaveChanges
    End If
    If Len(workPath) > 0 Then
        If rubyDone Then
            ' ルビ付与は完了している。正式名への変更だけが失敗した状態なので、
            ' 作業用コピーは消さずに残し、手で改名できるよう場所を伝える。
            errMessage = errMessage & vbCrLf & vbCrLf & _
                "※ルビ付与自体は完了しています。次のファイルを" & vbCrLf & _
                outPath & vbCrLf & _
                "へ手動で改名してご利用ください：" & vbCrLf & workPath
        Else
            ' 中途半端な作業用コピーを片付ける（前回成功した出力は残す）
            If fso.FileExists(workPath) Then fso.DeleteFile workPath, True
        End If
    End If
    Application.ScreenUpdating = savedScreenUpdating
    WriteResult tsvPath, rubyMode, applied, skipped, errMessage
    On Error GoTo 0
    Err.Raise vbObjectError + 900, "RubiGUI", errMessage
End Sub


'------------------------------------------------------------
' TSV を UTF-8 で読み込み、設定行と語句を取り出す。
' 設定行は「#キー<TAB>値」の形式で、語句より前に置かれている。
'------------------------------------------------------------
Private Function LoadTsv(ByVal tsvPath As String, _
                         ByRef rubyMode As String, _
                         ByRef rubyRatio As Double, _
                         ByRef rubyOffset As Double) As Collection

    Dim stream As Object
    Dim result As Collection
    Dim lineData As String
    Dim parts() As String
    Dim key As String
    ' ★設定行として扱うのは、終端マーカー「#DATA」より前の「#」行だけ。
    '   全ての「#」行を設定行とみなすと、「#タグ」のような語句を辞書に
    '   登録したときに、その語句が黙って捨てられてしまう。
    Dim inHeader As Boolean

    inHeader = True
    Set result = New Collection

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = ADO_TYPE_TEXT
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile tsvPath
    ' Python は "\n" 区切りで書き出す（既定の adCRLF のままだと1行に繋がる）
    stream.LineSeparator = ADO_LF

    Do Until stream.EOS
        lineData = stream.ReadText(ADO_READ_LINE)

        ' 改行コードが CRLF だった場合の CR と、先頭行の BOM を取り除く
        If Len(lineData) > 0 Then
            If Right$(lineData, 1) = vbCr Then
                lineData = Left$(lineData, Len(lineData) - 1)
            End If
        End If
        If Len(lineData) > 0 Then
            If Left$(lineData, 1) = ChrW$(&HFEFF) Then
                lineData = Mid$(lineData, 2)
            End If
        End If

        If Len(lineData) > 0 Then
            parts = Split(lineData, vbTab)
            If inHeader And Left$(lineData, 1) = "#" Then
                key = LCase$(Trim$(Mid$(parts(0), 2)))
                If key = "data" Then
                    ' 設定の終わり。これ以降は「#」で始まっていても語句として読む
                    inHeader = False
                ElseIf UBound(parts) >= 1 Then
                    Select Case key
                        Case "mode"
                            rubyMode = LCase$(Trim$(parts(1)))
                        Case "ratio"
                            rubyRatio = Val(Trim$(parts(1)))
                        Case "offset"
                            rubyOffset = Val(Trim$(parts(1)))
                    End Select
                End If
            Else
                inHeader = False
                If UBound(parts) >= 1 Then
                    If Len(parts(0)) > 0 And Len(parts(1)) > 0 Then
                        result.Add Array(parts(0), parts(1))
                    End If
                End If
            End If
        End If
    Loop
    stream.Close

    ' 設定が壊れていても止まらないよう、範囲外は既定値へ戻す
    If rubyMode <> "all" Then rubyMode = "first"
    If rubyRatio < 10# Or rubyRatio > 100# Then rubyRatio = DEFAULT_RATIO
    ' PhoneticGuide の Raise は「親文字からの距離」なので負値は指定できない
    If rubyOffset < 0# Or rubyOffset > 50# Then rubyOffset = DEFAULT_OFFSET

    Set LoadTsv = result
End Function


'------------------------------------------------------------
' 出力先フォルダを作り、元ファイルを「作業用のコピー」として複製する。
' 正式な出力名への差し替えは、ルビ付与が最後まで成功してから呼び出し元が行う。
'
' ★workPath は ByRef で、コピーを始める「前」に呼び出し元へ返す。
'   戻り値で返すと、CopyFile が途中で失敗して壊れた作業用コピーが
'   できたときに、呼び出し元の ErrHandler がその存在を知らず
'   片付けられないまま残ってしまう。
'------------------------------------------------------------
Private Sub PrepareOutput(ByVal fso As Object, _
                          ByVal srcPath As String, _
                          ByVal outPath As String, _
                          ByRef workPath As String)
    Dim outFolder As String
    Dim tempFolder As String
    Dim i As Long
    Dim fullName As String
    Dim openName As String
    Dim isSaved As Boolean
    Dim blockedName As String

    outFolder = fso.GetParentFolderName(outPath)
    If Len(outFolder) > 0 Then
        If Not fso.FolderExists(outFolder) Then
            fso.CreateFolder outFolder
        End If
    End If

    ' ★作業用コピーは %TEMP%（クラウド同期の対象外）に作る。
    '   出力フォルダ（デスクトップ配下＝OneDriveの同期対象になりがち）に
    '   作ると、コピー→改名の直後に同期エンジンが元の名前のファイルを
    '   復活させることがあり、成功しているのに ~rubigui_〇〇 が残る。
    '   実測：一括処理で「出力は正常・作業用コピーだけが元ファイルと
    '   同一内容で残る」状態を確認した。出力フォルダへの書き込みは
    '   最後の1回だけにして、この競合そのものを避ける。
    ' ★名前は実行ごとに一意にする（日時を挟む）。固定名にすると、前回
    '   「ルビは振れたが書き出しだけ失敗した」ときに残した完成品を、
    '   次回の実行が上書きしたり後片付けで削除したりしてしまう。
    '   代わりに自己修復（次回の上書き）が効かなくなるので、
    '   古い残骸は下で掃除する。
    tempFolder = Environ$("TEMP")
    If Len(tempFolder) = 0 Then tempFolder = Environ$("TMP")
    If Len(tempFolder) = 0 Or Not fso.FolderExists(tempFolder) Then
        ' %TEMP% が取れない環境では、やむを得ず出力フォルダを使う
        tempFolder = outFolder
    End If

    CleanupOldWorkFiles fso, tempFolder, fso.GetFileName(outPath)
    ' 旧版が出力フォルダに作った残骸もここで掃除する
    CleanupOldWorkFiles fso, outFolder, fso.GetFileName(outPath)

    workPath = tempFolder & "\~rubigui_" & Format$(Now, "yyyymmddhhnnss") & _
               "_" & fso.GetFileName(outPath)

    ' 前回の出力や作業用コピーが Word で開かれたままだと差し替えられないので
    ' 先に閉じる。
    ' ★fullName / openName / isSaved は毎回リセットしてから読む。
    '   On Error Resume Next の下では取得に失敗しても前の周回の値が残るため、
    '   リセットしないとまったく無関係な文書を閉じてしまうことがある。
    ' ★未保存の編集がある文書は閉じない。無確認で破棄するとユーザーの
    '   作業が失われるので、代わりにエラーにして知らせる。
    '   保存済みかどうかを読めなかった場合も「閉じない」側に倒す。
    blockedName = ""
    On Error Resume Next
    For i = Documents.Count To 1 Step -1
        fullName = ""
        fullName = Documents(i).FullName
        openName = LCase$(fullName)
        If Len(openName) > 0 Then
            If openName = LCase$(outPath) Or openName = LCase$(workPath) Then
                isSaved = False
                isSaved = Documents(i).Saved
                If isSaved Then
                    Documents(i).Close SaveChanges:=wdDoNotSaveChanges
                Else
                    blockedName = fullName
                End If
            End If
        End If
    Next i
    On Error GoTo 0

    If Len(blockedName) > 0 Then
        Err.Raise vbObjectError + 903, "RubiGUI", _
            "出力先のファイルが、保存されていない状態でWordに開かれています。" & vbCrLf & _
            "保存または閉じてから、もう一度実行してください：" & vbCrLf & blockedName
    End If

    fso.CopyFile srcPath, workPath, True
End Sub


'------------------------------------------------------------
' 指定フォルダに残っている古い作業用コピーを掃除する。
' 通常は %TEMP% を渡すが、旧版が出力フォルダに残した分も掃除するため
' 呼び出し元は両方のフォルダに対して呼ぶ。
'
' ★作業用コピーの名前を実行ごとに一意にしたことで、「次回の実行が
'   上書きする」という自己修復が効かなくなった。Wordの強制終了などで
'   後片付けが走らなかった分が溜まり続けるのを防ぐため、同じ出力名に
'   対応する古いもの（24時間より前）だけをここで消す。
' ★24時間の猶予を置くのは、「改名だけ失敗して意図的に残した完成品」を
'   ユーザーが取り出す前に消してしまわないようにするため。この完成品は
'   保存した時刻が更新日時になるので、24時間ルールで確実に守られる。
'   なお、コピー直後（保存前）の作業用コピーの更新日時は CopyFile が
'   コピー元の日時を引き継ぐため「24時間より古い」と判定され得るが、
'   掃除はコピーより前に走るので自分の分を消すことはない。処理中の
'   ファイルを他から消されないのは、Wordが掛けるロックによる。
'------------------------------------------------------------
Private Sub CleanupOldWorkFiles(ByVal fso As Object, _
                                ByVal outFolder As String, _
                                ByVal outName As String)
    Dim f As Object
    Dim n As String

    If Len(outFolder) = 0 Or Len(outName) = 0 Then Exit Sub

    On Error Resume Next
    For Each f In fso.GetFolder(outFolder).Files
        n = ""
        n = f.Name
        ' 「>=」なのは、日時の入っていない旧形式（~rubigui_〇〇.docx）も
        ' 掃除の対象にするため。ちょうど一致するのはその形だけなので誤爆しない。
        If Len(n) >= Len(outName) + 9 Then
            If LCase$(Left$(n, 9)) = "~rubigui_" Then
                If LCase$(Right$(n, Len(outName) + 1)) = "_" & LCase$(outName) Then
                    If f.DateLastModified < Now - 1 Then
                        f.Delete True
                    End If
                End If
            End If
        End If
    Next f
    On Error GoTo 0
End Sub


'------------------------------------------------------------
' 語句ごとにルビを振る。処理は2段階に分ける。
'
'  第1段階：文書を一切変更せずに「どこへ何のルビを振るか」を決める。
'  第2段階：決めた位置へ、文書の末尾から前に向かって実際にルビを振る。
'
' ★2段階に分ける理由
'   PhoneticGuide は対象文字列をルビ付きの構造へ置き換えるため、
'   実行するたびにそれ以降の文字位置がずれる。位置決めと付与を混ぜると
'   「次にどこを探すべきか」が分からなくなり、v2.0 では取りこぼしや
'   二重付与が起きていた。先に全部の位置を決めてしまい、末尾から前へ
'   処理すれば、まだ処理していない前方の位置は一切ずれない。
'
' ★語句の重なり（「日本」と「日本人」）について
'   語句は Python 側で「表層形の長い順」に並べて渡される。第1段階では
'   確定済みの範囲を occupied で記録しておき、重なる候補は採用しない。
'   これにより長い語句が優先され、その内側へ短い語句のルビが
'   二重に載ることがない。Word のふりがなの内部表現に依存しないので、
'   Wordのバージョン差にも左右されない。
'------------------------------------------------------------
Private Sub InsertAllGuides(ByVal docNew As Document, _
                            ByVal terms As Collection, _
                            ByVal rubyMode As String, _
                            ByVal rubyRatio As Double, _
                            ByVal rubyOffset As Double, _
                            ByRef applied As Long, _
                            ByRef skipped As Long)

    Dim docLen As Long
    Dim occupied() As Boolean
    Dim startOf() As Long
    Dim matches As Collection
    Dim idx As Long
    Dim entry As Variant
    Dim targetWord As String
    Dim furigana As String
    Dim rng As Range
    Dim searchStart As Long
    Dim hitStart As Long
    Dim hitEnd As Long
    Dim guardCount As Long
    Dim pos As Long
    Dim conflict As Boolean
    Dim info As Variant
    Dim processed As Long

    docLen = docNew.Content.End
    If docLen <= 0 Then Exit Sub

    ' 文字位置ごとの「もう別の語句のルビが確定している」印と、
    ' 「その位置から始まる採用済み候補の番号」
    ReDim occupied(0 To docLen)
    ReDim startOf(0 To docLen)
    Set matches = New Collection

    ' ======== 第1段階：位置決め（文書は変更しない） ========
    For idx = 1 To terms.Count
        entry = terms(idx)
        targetWord = CStr(entry(0))
        furigana = CStr(entry(1))

        If Len(targetWord) > 0 And Len(furigana) > 0 Then
            searchStart = 0
            guardCount = 0
            Do
                guardCount = guardCount + 1
                If guardCount > MAX_HITS_PER_TERM Then Exit Do
                If searchStart >= docLen Then Exit Do
                ' ★1つの語句が何千回も出てくる文書では、ここを回り続けている間
                '   Wordが「応答なし」に見えてしまう。ヒット単位でも息継ぎする。
                If (guardCount Mod 200) = 0 Then DoEvents

                Set rng = docNew.Range(searchStart, docLen)
                ' ★見つけるたびに条件を設定し直すのが要点。
                '   Range を作り直すと .Find の条件は保持されない。
                '   v2.0 はこれを忘れており、2回目以降の検索が成立せず
                '   最初の1箇所にしかルビが振られなかった。
                If Not FindWord(rng, targetWord) Then Exit Do

                hitStart = rng.Start
                hitEnd = rng.End
                If hitEnd <= hitStart Then Exit Do
                searchStart = hitEnd

                If IsInsideField(rng) Then
                    ' 元からある差し込みフィールドなどの中は触らない
                    skipped = skipped + 1
                Else
                    conflict = False
                    For pos = hitStart To hitEnd - 1
                        If pos >= 0 And pos <= docLen Then
                            If occupied(pos) Then
                                conflict = True
                                Exit For
                            End If
                        End If
                    Next pos

                    If conflict Then
                        ' もっと長い語句のルビが既に確定している範囲
                        skipped = skipped + 1
                    Else
                        For pos = hitStart To hitEnd - 1
                            If pos >= 0 And pos <= docLen Then occupied(pos) = True
                        Next pos
                        matches.Add Array(hitStart, hitEnd, furigana)
                        startOf(hitStart) = matches.Count
                        ' 「初回のみ」は1件採用できた時点で次の語句へ。
                        ' ★重なりで見送った分は数に入れない。最初のヒットが
                        '   他の語句のルビと重なっていても、次の出現箇所を探す。
                        If rubyMode <> "all" Then Exit Do
                    End If
                End If
            Loop
        End If

        If (idx Mod 50) = 0 Then DoEvents
    Next idx

    ' ======== 第2段階：末尾から前へ実際にルビを振る ========
    ' 重なりを排除済みなので、1つの開始位置に候補は高々1つしかない。
    ' 位置の大きい方から処理すれば、前方の位置はずれない。
    processed = 0
    For pos = docLen To 0 Step -1
        If startOf(pos) > 0 Then
            info = matches(startOf(pos))
            Set rng = docNew.Range(CLng(info(0)), CLng(info(1)))
            If ApplyGuide(rng, CStr(info(2)), rubyRatio, rubyOffset) Then
                applied = applied + 1
            Else
                skipped = skipped + 1
            End If
            processed = processed + 1
            If (processed Mod 50) = 0 Then DoEvents
        End If
    Next pos
End Sub


'------------------------------------------------------------
' 検索条件を毎回設定し直してから、範囲の先頭から前方へ1件検索する。
' 見つかると rng は一致した範囲へ置き換わる。
'------------------------------------------------------------
Private Function FindWord(ByRef rng As Range, _
                          ByVal targetWord As String) As Boolean
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = targetWord
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindStop
        ' ★日本語は分かち書きしないため、MatchWholeWord を True にすると
        '   文中の語句を取りこぼす。v2.0 の取りこぼしの一因でもあった。
        .MatchWholeWord = False
        .MatchCase = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Format = False
        ' 日本語版Word固有の検索オプションは、ユーザーごとに設定が残るため
        ' 必ず明示的に切る（有効なままだと別の語句にヒットする）。
        ' MatchByte を True にしないと「ＡＩ」を探して「AI」に当たるなど、
        ' 半角と全角を取り違える。東アジア言語版以外では存在しないので
        ' On Error Resume Next の中で設定する。
        On Error Resume Next
        .MatchFuzzy = False
        .IgnoreSpace = False
        .IgnorePunct = False
        .MatchByte = True
        On Error GoTo 0
    End With
    FindWord = rng.Find.Execute
End Function


'------------------------------------------------------------
' 元の文書に入っているフィールド（差し込み印刷・目次・数式など）の
' 中かどうか。フィールドコードの途中にふりがなを差し込むと壊れるため飛ばす。
'
' 本ツールが振ったルビとの重なりは、この判定ではなく InsertAllGuides の
' occupied（確定済みの文字位置）で防いでいる。Wordのふりがなの内部表現は
' バージョンによって EQ フィールドだったり w:ruby だったりするため、
' フィールドの有無で判定すると取りこぼす。
'------------------------------------------------------------
Private Function IsInsideField(ByRef rng As Range) As Boolean
    Dim inside As Boolean
    inside = False
    On Error Resume Next
    If rng.Fields.Count > 0 Then inside = True
    If rng.Information(wdInFieldCode) Then inside = True
    On Error GoTo 0
    IsInsideField = inside
End Function


'------------------------------------------------------------
' 実際にルビを振る。ルビの大きさは親文字サイズに対する％で決める。
'------------------------------------------------------------
Private Function ApplyGuide(ByRef rng As Range, _
                            ByVal furigana As String, _
                            ByVal rubyRatio As Double, _
                            ByVal rubyOffset As Double) As Boolean
    Dim baseSize As Double
    Dim guideSize As Double
    Dim baseFont As String
    Dim ok As Boolean

    baseSize = 0#
    baseFont = ""
    On Error Resume Next
    baseSize = rng.Font.Size
    ' ★ルビは必ず仮名なので、日本語用フォント(NameFarEast)を優先する。
    '   欧文フォント(Name)を先に採ると、本文が「欧文=Calibri／日本語=游ゴシック」
    '   のような一般的な設定のときに、ルビだけ本文と違う書体になってしまう。
    baseFont = rng.Font.NameFarEast
    If Len(baseFont) = 0 Then baseFont = rng.Font.Name
    On Error GoTo 0

    ' 範囲内で書式が混在していると Word は 9999999 を返す
    If baseSize <= 0# Or baseSize > 1000# Then baseSize = 10.5
    guideSize = Int(baseSize * rubyRatio / 100# * 10# + 0.5) / 10#
    If guideSize < 1# Then guideSize = 1#

    ok = False
    On Error Resume Next
    If Len(baseFont) = 0 Then
        rng.PhoneticGuide Text:=furigana, _
            Alignment:=wdPhoneticGuideAlignmentCenter, _
            Raise:=rubyOffset, FontSize:=guideSize
    Else
        rng.PhoneticGuide Text:=furigana, _
            Alignment:=wdPhoneticGuideAlignmentCenter, _
            Raise:=rubyOffset, FontSize:=guideSize, FontName:=baseFont
    End If
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    ApplyGuide = ok
End Function


'------------------------------------------------------------
' 処理結果を TSV と同じ場所へ書き出す（拡張子 .result、ASCIIのみ）。
' Python はこれを読んで「ルビ〇件」を表示する。
'------------------------------------------------------------
Private Sub WriteResult(ByVal tsvPath As String, _
                        ByVal rubyMode As String, _
                        ByVal applied As Long, _
                        ByVal skipped As Long, _
                        ByVal errMessage As String)
    Dim resultPath As String
    Dim fileNum As Integer

    On Error Resume Next
    resultPath = tsvPath & ".result"
    fileNum = FreeFile
    Open resultPath For Output As #fileNum
    Print #fileNum, "mode" & vbTab & rubyMode
    Print #fileNum, "applied" & vbTab & CStr(applied)
    Print #fileNum, "skipped" & vbTab & CStr(skipped)
    If Len(errMessage) > 0 Then
        Print #fileNum, "error" & vbTab & "1"
    End If
    Close #fileNum
    On Error GoTo 0
End Sub
