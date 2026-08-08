Attribute VB_Name = "Module1"
Option Explicit
Sub InsertFuriganaFromTSV_SaveToNewFile_Stable()
Dim docOriginal As Document, docNew As Document
Dim docName As String, nameOnly As String, extOnly As String
Dim basePath As String, tsvPath As String, savePath As String
Dim fso As Object, FileNum As Integer
Dim LineData As String, WordParts() As String
Dim TargetWord As String, Furigana As String
Dim rng As Range
' 元文書を取得
Set docOriginal = ActiveDocument
docName = docOriginal.Name
nameOnly = Left(docName, InStrRev(docName, ".") - 1)
extOnly = Mid(docName, InStrRev(docName, "."))

' 親フォルダの取得
Set fso = CreateObject("Scripting.FileSystemObject")
basePath = fso.GetParentFolderName(docOriginal.Path)

' 元文書をテンポラリ保存
Dim tempPath As String
tempPath = basePath & "\_temp_" & nameOnly & extOnly
docOriginal.SaveAs2 tempPath
' テンポラリ文書を開いて編集
Set docNew = Documents.Open(tempPath)
If docNew Is Nothing Then
    MsgBox "テンポラリファイルが開けませんでした", vbCritical
    Exit Sub
End If

' TSVファイルのパス構築
Dim desktopPath As String
desktopPath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
tsvPath = desktopPath & "\ルビ振り\ルビデータ\" & nameOnly & ".tsv"
If Dir(tsvPath) = "" Then
    MsgBox "TSVファイルが見つかりません：" & vbCrLf & tsvPath, vbCritical
    Exit Sub
End If

Dim outputFolder As String
outputFolder = desktopPath & "\ルビ振り\出力（ルビ付き）"

If Not fso.FolderExists(outputFolder) Then
    fso.CreateFolder outputFolder
End If

savePath = outputFolder & "\" & nameOnly & "（ルビ）" & extOnly

' TSVを読み込んでルビ処理開始
FileNum = FreeFile
Open tsvPath For Input As FileNum
Do Until EOF(FileNum)
    Line Input #FileNum, LineData
    WordParts = Split(LineData, vbTab)

    If UBound(WordParts) = 1 Then
        TargetWord = WordParts(0)
        Furigana = WordParts(1)

        Set rng = docNew.Range(0, 0)
        With rng.Find
            .Text = TargetWord
            .Forward = True
            .Wrap = wdFindStop
            .MatchWholeWord = True
            .MatchCase = True
        End With

        Do While rng.Find.Execute
            rng.PhoneticGuide Text:=Furigana, Alignment:=wdPhoneticGuideAlignmentCenter, _
                Raise:=0, FontSize:=6, FontName:=rng.Font.Name

            ' ?? 範囲を明示的に進めてループ回避！
            Set rng = docNew.Range(rng.End, docNew.Content.End)
            DoEvents
        Loop
    End If
Loop
Close FileNum

' 仕上げの保存処理
docNew.SaveAs2 FileName:=savePath, FileFormat:=wdFormatXMLDocument
If Dir(tempPath) = "" Then
    MsgBox "テンポラリファイルが保存されませんでした", vbCritical
    Exit Sub
End If
Kill tempPath

End Sub

