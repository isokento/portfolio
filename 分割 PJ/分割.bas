Attribute VB_Name = "Module1"
Sub listUp()

Application.ScreenUpdating = False

Dim lb As Workbook, wh As Worksheet, ti As Worksheet, cp As Worksheet

'�S�̗̂���----------------
'�^�C�g���V�[�g�ɂ���ԍ���AY����t�B���^�A�l�����\��t��
'������̂݁��@�C�ӂ̕�������́A�t�H���_�I��
'�V�K�u�b�N�����グ�A�ۑ�

Set lb = ThisWorkbook
Set wh = Sheets("�S�̃��X�g")
Set ti = Sheets("�^�C�g��")
Set cp = Sheets("����")

Dim i, tanto, j
Dim tLastRow, wLastRow
Dim fld As String, fileN As String, pathW As String

yn = MsgBox("�����������J�n���܂����H", vbOKCancel)
If yn = 2 Then
    MsgBox "�L�����Z������܂���"
    Exit Sub
End If
wh.Range("A1").AutoFilter
tLastRow = ti.Cells(9999, 1).End(xlUp).Row
wLastRow = wh.Cells(9999, 1).End(xlUp).Row

For i = 2 To tLastRow
'�^�C�g���V�[�g����t�B���^��������ԍ��itanto�j���擾
    tanto = ti.Cells(i, 1)
    seg = ti.Cells(i, 2)
    seg = "(" & seg & ")"
    With wh
    '�S���ԍ��Ńt�B���^
        .Range("A1").AutoFilter 51, tanto
        Set Rng = .Range("A1:AY" & wLastRow)
        '�����Ă镔�������R�s�[
        Set Rng = Rng.SpecialCells(xlCellTypeVisible)
        Rng.Select
        Selection.Copy
        For j = 0 To 100
         DoEvents
        Next j
        
        '�����͎��O�ɗp�ӁA�l�̂ݓ\��t��
        cp.Cells(1, 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        
        '����̂ݏ���
        If i = 2 Then
            MsgBox "�����t�@�C���̕ۑ����I�����Ă�������"
            If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then
                fld = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1)
                MsgBox "�t�@�C��������͂��Ă�������"
                fileN = InputBox("�t�@�C�����́u�����œ��͂������t �i�Ζ���j�v�ƂȂ�܂��B", "�t�@�C��������")
                
            End If
        End If
        
        pathW = fld & "\" & fileN & seg & ".xlsx"
        
        Application.CutCopyMode = False
        
        '�����V�[�g���R�s�[���ĕۑ��A����
        cp.Copy
        ActiveWorkbook.SaveAs Filename:= _
            pathW, FileFormat:= _
            xlOpenXMLWorkbook, CreateBackup:=False
        ActiveWindow.Close
            
    
    End With

    cp.Range("A2:H" & cp.Cells(9999, 1).End(xlUp).Row).ClearContents

Next i
Application.ScreenUpdating = True

MsgBox "�������I�����܂���"

End Sub


