Attribute VB_Name = "Module2"
Option Explicit

Sub allClear()

Dim sh(), i, delRng, delRng2, tgtSh
sh = Array("�z��", "�w��", "���j�b�g", "�ۑS")
Set tgtSh = Sheets(sh(0))

Dim s
Set s = Sheets("KO�i���o�[���̏W�v���z")

If s.Cells(2, "H") = "" Then MsgBox ("�������̂�����܂���")

With tgtSh
    For i = 0 To UBound(sh)
        Set tgtSh = Sheets(sh(i))
        Set delRng = .Range("A15:H46")
        Set delRng2 = .Range("A52:H99")
        .Range(.Cells(12, "B"), .Cells(12, "C")).ClearContents
        delRng.ClearContents
        delRng2.ClearContents
        
        .Cells(44, "E").Value = "���@�@�@�@�@�@�v"
        .Cells(45, "E").Value = "���@�@��@�@��"
        .Cells(46, "E").Value = "�Ł@���@���@�v"
        .Cells(44, "G").Formula = "=SUM(G15:G43)"
        .Cells(45, "G").Formula = "=ROUND(G44*KO�i���o�[���̏W�v���z!$L$8,0)"
        .Cells(46, "G").Formula = "=G44+G45"
        .Cells(97, "E").Value = "���@�@�@�@�@�@�v"
        .Cells(98, "E").Value = "���@�@��@�@��"
        .Cells(99, "E").Value = "�Ł@���@���@�v"
        .Cells(97, "G").Formula = "=SUM(G15:G96)"
        .Cells(98, "G").Formula = "=ROUND(G97*KO�i���o�[���̏W�v���z!$L$8,0)"
        .Cells(99, "G").Formula = "=G97+G98"
        
        
    Next
    
End With

s.Range("H1:H999").ClearContents



End Sub

