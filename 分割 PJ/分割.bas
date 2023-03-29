Attribute VB_Name = "Module1"
Sub listUp()

Application.ScreenUpdating = False

Dim lb As Workbook, wh As Worksheet, ti As Worksheet, cp As Worksheet

'全体の流れ----------------
'タイトルシートにある番号でAY列をフィルタ、値書式貼り付け
'＜初回のみ＞　任意の文字列入力、フォルダ選択
'新規ブック立ち上げ、保存

Set lb = ThisWorkbook
Set wh = Sheets("全体リスト")
Set ti = Sheets("タイトル")
Set cp = Sheets("分割")

Dim i, tanto, j
Dim tLastRow, wLastRow
Dim fld As String, fileN As String, pathW As String

yn = MsgBox("分割処理を開始しますか？", vbOKCancel)
If yn = 2 Then
    MsgBox "キャンセルされました"
    Exit Sub
End If
wh.Range("A1").AutoFilter
tLastRow = ti.Cells(9999, 1).End(xlUp).Row
wLastRow = wh.Cells(9999, 1).End(xlUp).Row

For i = 2 To tLastRow
'タイトルシートからフィルタをかける番号（tanto）を取得
    tanto = ti.Cells(i, 1)
    seg = ti.Cells(i, 2)
    seg = "(" & seg & ")"
    With wh
    '担当番号でフィルタ
        .Range("A1").AutoFilter 51, tanto
        Set Rng = .Range("A1:AY" & wLastRow)
        '見えてる部分だけコピー
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
        Rng.Select
        Selection.Copy
        For j = 0 To 100
         DoEvents
        Next j
        
        '書式は事前に用意、値のみ貼り付け
        cp.Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        '初回のみ処理
        If i = 2 Then
            MsgBox "分割ファイルの保存先を選択してください"
            If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
                fld = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
                MsgBox "ファイル名を入力してください"
                fileN = InputBox("ファイル名は「ここで入力した言葉 （勤務先）」となります。", "ファイル名入力")
                
            End If
        End If
        
        pathW = fld & "\" & fileN & seg & ".xlsx"
        
        Application.CutCopyMode = False
        
        '分割シートをコピーして保存、閉じる
        cp.Copy
        ActiveWorkbook.SaveAs Filename:= _
            pathW, FileFormat:= _
            xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWindow.Close
            
    
    End With

    cp.Range("A2:H" & cp.Cells(9999, 1).End(xlUp).Row).ClearContents

Next i
Application.ScreenUpdating = True

MsgBox "処理が終了しました"

End Sub


