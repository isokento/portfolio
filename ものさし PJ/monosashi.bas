Attribute VB_Name = "Module1"
Option Explicit
Sub allDel()

Application.ScreenUpdating = False

Call a_1Del
Call a_2Del
Call a_3Del

Application.ScreenUpdating = True

MsgBox "すべて削除 を完了しました"

End Sub

Sub allDo()

Application.ScreenUpdating = False

Call a_1
Call a_2
Call a_3

Application.ScreenUpdating = True

MsgBox "すべて配置 を完了しました"

End Sub

Sub a_1()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range


Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")
Set rr = Sheets("ものさし")
'問題の位置を取得
    '定規スタート37
    '字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 1
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226
r = 11
imgWidth = 9

With Qsh

    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    If QRng.Cells(11, cm1) <> "" Then
        Call a_1Del
        Application.ScreenUpdating = False
    End If
End With

With QRng

    For j = 9 To 13 Step 2
        
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
'▼▼▼▼▼▼▼▼▼▼▼▼ランダム▼▼▼▼▼▼▼▼▼▼▼▼▼
        
        r = r + 2
        min1 = rr.Cells(r, "EA").Value
        max1 = rr.Cells(r, "EB").Value
        min2 = rr.Cells(r + 1, "EA").Value
        max2 = rr.Cells(r + 1, "EB").Value
        
        cm1Cel.Value = Int((max1 - min1 + 1) * Rnd + min1)
        If j = 9 Then
            mm1Cel.Value = Int(9 * Rnd + 1)
        Else
            mm1Cel.Value = Int(10 * Rnd + 0)
        End If
        If j > 9 Then
            If (cm1Cel.Value * 10 + mm1Cel.Value) - (cm2Cel.Offset(-2, 0).Value * 10 + mm2Cel.Offset(-2, 0).Value) < 5 Then mm1Cel.Value = mm1Cel.Value + 5
        End If
        cm2Cel.Value = Int((max2 - min2 + 1) * Rnd + min2)
        mm2Cel.Value = Int(10 * Rnd + 0)
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        If (cm2Cel.Value * 10 + mm2Cel.Value) - (cm1Cel.Value * 10 + mm1Cel.Value) < 5 Then mm2Cel.Value = mm2Cel.Value + 5
        
        
'▲▲▲▲▲▲▲▲▲▲▲▲ランダム▲▲▲▲▲▲▲▲▲▲▲▲▲
        
        
        'If cm1cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm1cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm1Cel.Value * 10 + mm1Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt1).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        For n = 1 To 100
            DoEvents
        Next n
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        
        ActiveSheet.Pictures.Paste(Link:=True).Select
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
    Next j

    For j = 9 To 13 Step 2
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
        'If cm2Cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm2Cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm2Cel.Value * 10 + mm2Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt2).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        For n = 1 To 100
            DoEvents
        Next n
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
nextj:
    Next j


End With

Application.ScreenUpdating = True


End Sub

Sub a_1Del()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range


Dim shp As Shape
Dim rng As Range
Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")

'問題の位置を取得
'定規スタート37
'字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 1
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226

With Qsh
    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set delRng = .Range(.Cells(QRow - 1, QClm + 27), .Cells(QRow - 1, QClm + 1000))
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    
End With


'▼▼▼▼▼▼▼▼▼▼▼▼図形削除▼▼▼▼▼▼▼▼▼▼▼▼▼

With QRng

    For Each shp In ActiveSheet.Shapes
      ' 図形の配置されているセル範囲をオブジェクト変数にセット
      Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
      ' 図形の配置されているセル範囲と
      ' 選択されているセル範囲が重なっているときに図形を削除
      If Not (Intersect(rng, delRng) Is Nothing) Then
        shp.Delete
      End If
    Next
    
'▲▲▲▲▲▲▲▲▲▲▲▲図形削除▲▲▲▲▲▲▲▲▲▲▲▲▲

'▼▼▼▼▼▼▼▼▼▼▼▼数字削除▼▼▼▼▼▼▼▼▼▼▼▼▼
    For j = 9 To 13 Step 2
    
        .Cells(j, cm1).Value = ""
        .Cells(j, mm1).Value = ""
        .Cells(j, cm2).Value = ""
        .Cells(j, mm2).Value = ""
    
    Next

'▲▲▲▲▲▲▲▲▲▲▲▲数字削除▲▲▲▲▲▲▲▲▲▲▲▲▲

End With

Application.ScreenUpdating = True


End Sub

Sub a_2()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range

Dim exSum As Long

Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")
Set rr = Sheets("ものさし")
'問題の位置を取得
    '定規スタート37
    '字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 2
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226
r = 11
imgWidth = 9

With Qsh

    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    If QRng.Cells(11, cm1) <> "" Then
        Call a_2Del
        Application.ScreenUpdating = False
    End If
End With

With QRng

    For j = 9 To 13 Step 2
        
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
'▼▼▼▼▼▼▼▼▼▼▼▼ランダム▼▼▼▼▼▼▼▼▼▼▼▼▼
        
        r = r + 2
        min1 = rr.Cells(r, "EA").Value
        max1 = rr.Cells(r, "EB").Value
        min2 = rr.Cells(r + 1, "EA").Value
        max2 = rr.Cells(r + 1, "EB").Value
        exSum = cm1Cel.Offset(-15, 0).Value * 10 + mm1Cel.Offset(-15, 0).Value
        
        cm1Cel.Value = Int((max1 - min1 + 1) * Rnd + min1)
        If j = 9 Then
            mm1Cel.Value = Int(9 * Rnd + 1)
        Else
            mm1Cel.Value = Int(10 * Rnd + 0)
        End If
                
        Do While cm1Cel.Value * 10 + mm1Cel.Value = exSum
            cm1Cel.Value = Int((max1 - min1 + 1) * Rnd + min1)
            If j = 9 Then
                mm1Cel.Value = Int(9 * Rnd + 1)
            Else
                mm1Cel.Value = Int(10 * Rnd + 0)
            End If
        Loop
        
        If j > 9 Then
            If (cm1Cel.Value * 10 + mm1Cel.Value) - (cm2Cel.Offset(-2, 0).Value * 10 + mm2Cel.Offset(-2, 0).Value) < 5 Then mm1Cel.Value = mm1Cel.Value + 5
        End If
        
        exSum = cm2Cel.Offset(-15, 0).Value * 10 + mm2Cel.Offset(-15, 0).Value
        
        cm2Cel.Value = Int((max2 - min2 + 1) * Rnd + min2)
        mm2Cel.Value = Int(10 * Rnd + 0)
        
        Do While cm2Cel.Value * 10 + mm2Cel.Value = exSum
            cm2Cel.Value = Int((max2 - min2 + 1) * Rnd + min2)
            mm2Cel.Value = Int(10 * Rnd + 0)
        Loop
        
        If (cm2Cel.Value * 10 + mm2Cel.Value) - (cm1Cel.Value * 10 + mm1Cel.Value) < 5 Then mm2Cel.Value = mm2Cel.Value + 5
        
        
'▲▲▲▲▲▲▲▲▲▲▲▲ランダム▲▲▲▲▲▲▲▲▲▲▲▲▲
        
        
        'If cm1cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm1cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm1Cel.Value * 10 + mm1Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt1).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        For n = 1 To 100
            DoEvents
        Next n
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        
        ActiveSheet.Pictures.Paste(Link:=True).Select
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
    Next j
    For n = 1 To 100
        DoEvents
    Next n
        
    For j = 9 To 13 Step 2
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
        'If cm2Cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm2Cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm2Cel.Value * 10 + mm2Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt2).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        For n = 1 To 100
            DoEvents
        Next n
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
nextj:
    Next j


End With

Application.ScreenUpdating = True


End Sub

Sub a_2Del()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range


Dim shp As Shape
Dim rng As Range
Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")

'問題の位置を取得
'定規スタート37
'字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 2
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226

With Qsh
    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set delRng = .Range(.Cells(QRow - 1, QClm + 27), .Cells(QRow - 1, QClm + 1000))
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    
End With


'▼▼▼▼▼▼▼▼▼▼▼▼図形削除▼▼▼▼▼▼▼▼▼▼▼▼▼

With QRng

    For Each shp In ActiveSheet.Shapes
      ' 図形の配置されているセル範囲をオブジェクト変数にセット
      Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
      ' 図形の配置されているセル範囲と
      ' 選択されているセル範囲が重なっているときに図形を削除
      If Not (Intersect(rng, delRng) Is Nothing) Then
        shp.Delete
      End If
    Next
    
'▲▲▲▲▲▲▲▲▲▲▲▲図形削除▲▲▲▲▲▲▲▲▲▲▲▲▲

'▼▼▼▼▼▼▼▼▼▼▼▼数字削除▼▼▼▼▼▼▼▼▼▼▼▼▼
    For j = 9 To 13 Step 2
    
        .Cells(j, cm1).Value = ""
        .Cells(j, mm1).Value = ""
        .Cells(j, cm2).Value = ""
        .Cells(j, mm2).Value = ""
    
    Next

'▲▲▲▲▲▲▲▲▲▲▲▲数字削除▲▲▲▲▲▲▲▲▲▲▲▲▲

End With

Application.ScreenUpdating = True


End Sub

Sub a_3()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range

Dim exSum As Long, exSum2 As Long

Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")
Set rr = Sheets("ものさし")
'問題の位置を取得
    '定規スタート37
    '字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 3
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226
r = 11
imgWidth = 9

With Qsh

    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    If QRng.Cells(11, cm1) <> "" Then
        Call a_3Del
        Application.ScreenUpdating = False
    End If
End With

With QRng

    For j = 9 To 13 Step 2
        
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
'▼▼▼▼▼▼▼▼▼▼▼▼ランダム▼▼▼▼▼▼▼▼▼▼▼▼▼
        
        r = r + 2
        min1 = rr.Cells(r, "EA").Value
        max1 = rr.Cells(r, "EB").Value
        min2 = rr.Cells(r + 1, "EA").Value
        max2 = rr.Cells(r + 1, "EB").Value
        exSum = cm1Cel.Offset(-15, 0).Value * 10 + mm1Cel.Offset(-15, 0).Value
        exSum2 = cm1Cel.Offset(-30, 0).Value * 10 + mm1Cel.Offset(-30, 0).Value
        
        cm1Cel.Value = Int((max1 - min1 + 1) * Rnd + min1)
        If j = 9 Then
            mm1Cel.Value = Int(9 * Rnd + 1)
        Else
            mm1Cel.Value = Int(10 * Rnd + 0)
        End If
                
        Do While cm1Cel.Value * 10 + mm1Cel.Value = exSum Or cm1Cel.Value * 10 + mm1Cel.Value = exSum2
            cm1Cel.Value = Int((max1 - min1 + 1) * Rnd + min1)
            If j = 9 Then
                mm1Cel.Value = Int(9 * Rnd + 1)
            Else
                mm1Cel.Value = Int(10 * Rnd + 0)
            End If
        Loop
        
        If j > 9 Then
            If (cm1Cel.Value * 10 + mm1Cel.Value) - (cm2Cel.Offset(-2, 0).Value * 10 + mm2Cel.Offset(-2, 0).Value) < 5 Then mm1Cel.Value = mm1Cel.Value + 5
        End If
        
        exSum = cm2Cel.Offset(-15, 0).Value * 10 + mm2Cel.Offset(-15, 0).Value
        exSum2 = cm1Cel.Offset(-30, 0).Value * 10 + mm1Cel.Offset(-30, 0).Value
        
        cm2Cel.Value = Int((max2 - min2 + 1) * Rnd + min2)
        mm2Cel.Value = Int(10 * Rnd + 0)
        
        Do While cm1Cel.Value * 10 + mm1Cel.Value = exSum Or cm1Cel.Value * 10 + mm1Cel.Value = exSum2
            cm2Cel.Value = Int((max2 - min2 + 1) * Rnd + min2)
            mm2Cel.Value = Int(10 * Rnd + 0)
        Loop
        
        If (cm2Cel.Value * 10 + mm2Cel.Value) - (cm1Cel.Value * 10 + mm1Cel.Value) < 5 Then mm2Cel.Value = mm2Cel.Value + 5
        
        
'▲▲▲▲▲▲▲▲▲▲▲▲ランダム▲▲▲▲▲▲▲▲▲▲▲▲▲
        
        
        'If cm1cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm1cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm1Cel.Value * 10 + mm1Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt1).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        
        ActiveSheet.Pictures.Paste(Link:=True).Select
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
    Next j

    For j = 9 To 13 Step 2
        
        Set cm1Cel = .Cells(j, cm1)
        Set mm1Cel = .Cells(j, mm1)
        Set cm2Cel = .Cells(j, cm2)
        Set mm2Cel = .Cells(j, mm2)
        
        'If cm2Cel.Value = "" Then
        '    GoTo nextj
        'ElseIf mm2Cel.Value = "" Then
        '    GoTo nextj
        'End If
        
        mmSum = cm2Cel.Value * 10 + mm2Cel.Value
        Set imgCel = .Cells(1, pt1 + 1 + mmSum * 2 - imgWidth)
        pt = .Cells(j, pt2).Value
        
        IMGsh.Activate
        IMGsh.Range(pt).Select
        Selection.Copy
        For n = 1 To 100
            DoEvents
        Next n
        Qsh.Activate
        imgCel.Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        For n = 1 To 100
            DoEvents
        Next n
        imgCel.Offset(0, 290).Select
        For n = 1 To 100
            DoEvents
        Next n
        ActiveSheet.Pictures.Paste(Link:=True).Select
        Application.CutCopyMode = False
        .Cells(1, 1).Select
        
nextj:
    Next j


End With

Application.ScreenUpdating = True


End Sub

Sub a_3Del()

Application.ScreenUpdating = False

Dim i, j, n, m, r, rr, min1, max1, min2, max2
Dim QNum, QRow As Long, QClm As Long, QRng As Range
Dim pt1 As Long, cm1 As Long, mm1 As Long, pt2 As Long, cm2 As Long, mm2 As Long, mmSum As Long, pt As String
Dim cm1Cel As Range, mm1Cel As Range, cm2Cel As Range, mm2Cel As Range
Dim imgWidth As Long, imgCel As Range
Dim Qsh As Worksheet, IMGsh As Worksheet

Dim delRng As Range


Dim shp As Shape
Dim rng As Range
Set Qsh = Sheets("製本")
Set IMGsh = Sheets("矢印")

'問題の位置を取得
'定規スタート37
'字、cm、mm：１：36, 63, 107　２： 161, 189, 233、283

QNum = 3
QClm = 9
pt1 = 28
cm1 = 56
mm1 = 100
pt2 = 154
cm2 = 182
mm2 = 226

With Qsh
    QRow = WorksheetFunction.Match(QNum, .Range(.Cells(1, QClm), .Cells(9999, QClm)), 0)
    Set delRng = .Range(.Cells(QRow - 1, QClm + 27), .Cells(QRow - 1, QClm + 1000))
    Set QRng = .Range(.Cells(QRow - 1, QClm), .Cells(QRow + 11, QClm + 274))
    
End With


'▼▼▼▼▼▼▼▼▼▼▼▼図形削除▼▼▼▼▼▼▼▼▼▼▼▼▼

With QRng

    For Each shp In ActiveSheet.Shapes
      ' 図形の配置されているセル範囲をオブジェクト変数にセット
      Set rng = Range(shp.TopLeftCell, shp.BottomRightCell)
      ' 図形の配置されているセル範囲と
      ' 選択されているセル範囲が重なっているときに図形を削除
      If Not (Intersect(rng, delRng) Is Nothing) Then
        shp.Delete
      End If
    Next
    
'▲▲▲▲▲▲▲▲▲▲▲▲図形削除▲▲▲▲▲▲▲▲▲▲▲▲▲

'▼▼▼▼▼▼▼▼▼▼▼▼数字削除▼▼▼▼▼▼▼▼▼▼▼▼▼
    For j = 9 To 13 Step 2
    
        .Cells(j, cm1).Value = ""
        .Cells(j, mm1).Value = ""
        .Cells(j, cm2).Value = ""
        .Cells(j, mm2).Value = ""
    
    Next

'▲▲▲▲▲▲▲▲▲▲▲▲数字削除▲▲▲▲▲▲▲▲▲▲▲▲▲

End With

Application.ScreenUpdating = True


End Sub

