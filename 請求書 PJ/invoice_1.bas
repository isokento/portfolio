Attribute VB_Name = "Module1"
Option Explicit

Sub Invoice()
Application.ScreenUpdating = False

Dim yn

yn = MsgBox("処理を開始しますか？", vbOKCancel)
If yn = 2 Then Exit Sub

Dim Inv
Dim listS, HaikanS, KonyuS, UnitS, HozenS
Dim llR, i, j, k, judC, a, b, c, srcRng, r, tgtSh
Dim pat

Set Inv = ActiveWorkbook
With Inv
    Set listS = .Sheets("KOナンバー毎の集計金額")
    Set HaikanS = .Sheets("配管")
    Set KonyuS = .Sheets("購入")
    Set UnitS = .Sheets("ユニット")
    Set HozenS = .Sheets("保全")
    
End With

With listS
    llR = .Cells(9999, 1).End(xlUp).Row
    judC = .Cells(1, 999).End(xlToLeft).Column + 1
    
    .Cells(1, judC) = "判定"
    For i = 2 To llR
        .Cells(i, judC).Value = Replace(Replace(Replace(Mid(.Cells(i, 1).Value, InStr(.Cells(i, 1), "(") + 1, 9), ")", ""), "費", ""), "工事", "")
        
    Next
    
    
    Dim sh()
    sh = Array("配管", "購入", "ユニット", "保全")
    Set srcRng = .Range(.Cells(1, judC), .Cells(llR, judC))
End With

Set tgtSh = Sheets(sh(0))
Dim p0, p1, p2, p3, p4, p5, ad

p0 = 15
p1 = 43
p2 = 46
p3 = 52
p4 = 96
p5 = 99

With tgtSh
    For i = 0 To UBound(sh)
        r = WorksheetFunction.Match("現　　場　　件　　名　　　", .Range("A:A"), 0) + 1
        Set tgtSh = Sheets(sh(i))
        a = WorksheetFunction.Match(sh(i), srcRng, 0)
        b = WorksheetFunction.CountIf(srcRng, sh(i))
        c = a + b - 1
        '29以下：1　32以下：2　77以下：3　80以下：4
        '1、３はそのまま。2は2ページ目の下部に合計欄。4は行追加して合計欄を作成
        If b <= 29 Then
            pat = 1
            .Range(.Cells(p4, "E"), .Cells(p5, "G")).ClearContents
        ElseIf b <= 32 Then
            pat = 2
            .Range(.Cells(p1, "E"), .Cells(p2, "G")).ClearContents
        ElseIf b <= 77 Then
            pat = 3
            .Range(.Cells(p1, "E"), .Cells(p2, "G")).ClearContents
        Else:
            pat = 4
            .Range(.Cells(p1, "E"), .Cells(p2, "G")).ClearContents
            ad = b - p5
            For k = 0 To ad
                .Rows(p4 - 2).Add
            Next k
        End If

        For j = a To c
            If r = p2 + 1 Then r = p3
            .Cells(r, "A") = listS.Cells(j, "F").Value
            .Cells(r, "E") = listS.Cells(j, "E").Value
            .Cells(r, "G") = listS.Cells(j, "G").Value
            r = r + 1
        Next j

'        If pat = 0 Then k = 44 Else GoTo 97
'        .Cells(k, "E") = "小 計"
'        .Cells(k, "G").Formula = "=SUM(G15:G43)"
'        .Cells(k + 1, "E") = "消　費　税"
'        .Cells(k + 1, "G") = listS.Cells(c, "D").Value
'        .Cells(k + 2, "E") = "税　込　合　計"
'        .Cells(k + 2, "G") = listS.Cells(c, "D").Value
'        End If
'        .Cells(12, "B") = listS.Cells(c, "C").Value
        .Cells(12, "B").Value = .Cells(WorksheetFunction.Match("税　込　合　計", .Range("E:E"), 0), 7)
        
    Next i
End With

Application.ScreenUpdating = True

MsgBox "終了しました"

End Sub
