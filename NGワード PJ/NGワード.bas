Attribute VB_Name = "Module1"
Sub NGword()

Dim can

can = MsgBox("�J�n���܂�", vbOKCancel)
If can = 2 Then Exit Sub

Application.ScreenUpdating = False

Dim wSh, cSh
Dim w, c, wlastRow, clastRow
Dim tgt, src, cnt


Set wSh = Sheets("NG�L�[���[�h")
Set cSh = Sheets("NG�`�F�b�N")

wlastRow = wSh.Cells(99999, 1).End(xlUp).Row
clastRow = cSh.Cells(99999, 4).End(xlUp).Row

For c = 2 To clastRow
    tgt = cSh.Cells(c, 4).Value
    tgt = Replace(tgt, " ", "")
    For w = 2 To wlastRow
        src = wSh.Cells(w, 1)
        src = Replace(src, " ", "")
        
        On Error Resume Next
        cnt = WorksheetFunction.Search(src, tgt)
        If cnt > 0 Then
'            cSh.Cells(c, "K").Value = wSh.Cells(w, 1).Value
            cSh.Cells(c, "K") = cSh.Cells(c, "K") + 1
            cnt = 0
'            GoTo nextC
        End If
        On Error GoTo 0
    Next w

'nextC:
Next c

Application.ScreenUpdating = True

MsgBox "�I�����܂���"

End Sub

Sub �N���A()
Dim can

can = MsgBox("���Z�b�g���܂����H", vbOKCancel)

If can = 2 Then Exit Sub

Sheets("NG�`�F�b�N").Range("K2:K5000").ClearContents

End Sub

Sub updating()
Application.ScreenUpdating = True
End Sub
