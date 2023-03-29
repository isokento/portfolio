Attribute VB_Name = "Module1"
Option Explicit

Sub Invoice()
Application.ScreenUpdating = False

Dim yn

yn = MsgBox("�������J�n���܂����H", vbOKCancel)
If yn = 2 Then Exit Sub

Dim Inv
Dim listS, HaikanS, KonyuS, UnitS, HozenS
Dim llR, i, j, k, judC, a, b, c, srcRng, r, tgtSh
Dim pat

Set Inv = ActiveWorkbook
With Inv
    Set listS = .Sheets("KO�i���o�[���̏W�v���z")
    Set HaikanS = .Sheets("�z��")
    Set KonyuS = .Sheets("�w��")
    Set UnitS = .Sheets("���j�b�g")
    Set HozenS = .Sheets("�ۑS")
    
End With

With listS
    llR = .Cells(9999, 1).End(xlUp).Row
    judC = .Cells(1, 999).End(xlToLeft).Column + 1
    
    .Cells(1, judC) = "����"
    For i = 2 To llR
        .Cells(i, judC).Value = Replace(Replace(Replace(Mid(.Cells(i, 1).Value, InStr(.Cells(i, 1), "(") + 1, 9), ")", ""), "��", ""), "�H��", "")
        
    Next
    
    
    Dim sh()
    sh = Array("�z��", "�w��", "���j�b�g", "�ۑS")
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
        r = WorksheetFunction.Match("���@�@��@�@���@�@���@�@�@", .Range("A:A"), 0) + 1
        Set tgtSh = Sheets(sh(i))
        a = WorksheetFunction.Match(sh(i), srcRng, 0)
        b = WorksheetFunction.CountIf(srcRng, sh(i))
        c = a + b - 1
        '29�ȉ��F1�@32�ȉ��F2�@77�ȉ��F3�@80�ȉ��F4
        '1�A�R�͂��̂܂܁B2��2�y�[�W�ڂ̉����ɍ��v���B4�͍s�ǉ����č��v�����쐬
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
'        .Cells(k, "E") = "�� �v"
'        .Cells(k, "G").Formula = "=SUM(G15:G43)"
'        .Cells(k + 1, "E") = "���@��@��"
'        .Cells(k + 1, "G") = listS.Cells(c, "D").Value
'        .Cells(k + 2, "E") = "�Ł@���@���@�v"
'        .Cells(k + 2, "G") = listS.Cells(c, "D").Value
'        End If
'        .Cells(12, "B") = listS.Cells(c, "C").Value
        .Cells(12, "B").Value = .Cells(WorksheetFunction.Match("�Ł@���@���@�v", .Range("E:E"), 0), 7)
        
    Next i
End With

Application.ScreenUpdating = True

MsgBox "�I�����܂���"

End Sub
