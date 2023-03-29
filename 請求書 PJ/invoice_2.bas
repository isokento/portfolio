Attribute VB_Name = "Module2"
Option Explicit

Sub allClear()

Dim sh(), i, delRng, delRng2, tgtSh
sh = Array("配管", "購入", "ユニット", "保全")
Set tgtSh = Sheets(sh(0))

Dim s
Set s = Sheets("KOナンバー毎の集計金額")

If s.Cells(2, "H") = "" Then MsgBox ("消すものがありません")

With tgtSh
    For i = 0 To UBound(sh)
        Set tgtSh = Sheets(sh(i))
        Set delRng = .Range("A15:H46")
        Set delRng2 = .Range("A52:H99")
        .Range(.Cells(12, "B"), .Cells(12, "C")).ClearContents
        delRng.ClearContents
        delRng2.ClearContents
        
        .Cells(44, "E").Value = "小　　　　　　計"
        .Cells(45, "E").Value = "消　　費　　税"
        .Cells(46, "E").Value = "税　込　合　計"
        .Cells(44, "G").Formula = "=SUM(G15:G43)"
        .Cells(45, "G").Formula = "=ROUND(G44*KOナンバー毎の集計金額!$L$8,0)"
        .Cells(46, "G").Formula = "=G44+G45"
        .Cells(97, "E").Value = "小　　　　　　計"
        .Cells(98, "E").Value = "消　　費　　税"
        .Cells(99, "E").Value = "税　込　合　計"
        .Cells(97, "G").Formula = "=SUM(G15:G96)"
        .Cells(98, "G").Formula = "=ROUND(G97*KOナンバー毎の集計金額!$L$8,0)"
        .Cells(99, "G").Formula = "=G97+G98"
        
        
    Next
    
End With

s.Range("H1:H999").ClearContents



End Sub

