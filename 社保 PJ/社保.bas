Attribute VB_Name = "Module1"
Sub fileOpen()

Application.ScreenUpdating = False

'�N�x�̏I���i�͂��߁j���ƂɒP�N���Ə�������O��

Dim thisBk
Dim fldPath As String, fullPath As String, sh, r
Dim fileC, i, YYYY, M, YM
Dim buf, cnt, nendo, tgtNendo As Long, openCnt

Const flp As String = "*.*"
    
Set thisBk = ActiveWorkbook

''�t�H���_�I��
With Application.FileDialog(msoFileDialogFolderPicker)
    .Show
    On Error Resume Next
    fldPath = .SelectedItems(1)
    If fldPath = "" Then Exit Sub
    On Error GoTo 0
End With


'Dir�֐����p�O�@�Q�Ɛݒ�̖��Ŗv------------------------------------------

'Set fso = New FileSystemObject
'Set fileC = fso.GetFolder(fldPath).Files

'�f�t�@�C�������W�J����
'For Each i In fileC
'    If InStr(i.Name, ".xls") = 0 Then GoTo nexti
'    Workbooks.Open i
'    Call copyToCalBk(i.Name, thisBk)
'nexti:
'Next

'For Each f In fso.myPath.Files
'    fldPath = myPath.SelectedItems(1)
'Next f
    
'---------------------------------------------------------------------------
    
'Dir�֐��ŊJ��------

tgtNendo = StrConv(InputBox("�W�v����N�x����͂��Ă��������i�����̂݁j", "�N�x�i�����̂݁j"), vbNarrow)

buf = Dir(fldPath & "\" & flp)
fullPath = fldPath & "\" & buf
openCnt = 0
f = 0
Do While buf <> ""
    Debug.Print buf
    YYYY = Left(buf, InStr(buf, ".") - 1)
    M = Right("0" & Replace(Mid(buf, InStr(buf, ".") + 1, 2), ".", ""), 2)
    YM = YYYY & M
    
    '�N�x����
    If Val(M) <= 6 Then
        nendo = YYYY - 1
    Else
        nendo = YYYY
    End If
    If nendo <> tgtNendo Then GoTo continue
    
    If f = 0 And r <> 6 Then
        For Each sh In Worksheets
            If sh.Name = nendo Then
                r = MsgBox("���łɏW�v�ς݂̔N�x�ł��B�㏑�����܂����H", vbYesNo)
                If r = 7 Then GoTo EOS
                GoTo break
            Else
                f = 1
            End If
        Next
    End If
break:

    
                
    Debug.Print fullPath
    fullPath = fldPath & "\" & buf
    If InStr(buf, ".xl") = 0 Then GoTo continue
    Workbooks.Open fullPath
    nendo = copyToCalBk(buf, thisBk)
    Application.DisplayAlerts = False
    Workbooks(buf).Close
    Application.DisplayAlerts = True
    openCnt = openCnt + 1
continue:
    buf = Dir()
Loop
    
'i = MsgBox("�N�x���Ƃ̃V�[�g�փR�s�[���܂����H", vbYesNo)


If openCnt = 0 Then
    MsgBox "�Y���N�x�̃f�[�^����������܂���ł���"
Else
    thisBk.Activate
    Call makeHit(tgtNendo)
End If

Application.ScreenUpdating = True
    
EOS:
    Sheets("�Y����").Activate
    MsgBox "�������I�����܂���"
End Sub

Function copyToCalBk(ByVal kyuyoBkName As String, ByVal macroBk As Object) As String

Dim i, j, lastRowR, cnt, lastClm, lastRowK
Dim YYYY, M, YM, nendo
Dim rareSh, r, sCnt

Set rareSh = macroBk.Sheets("���f�[�^")

With ActiveWorkbook.ActiveSheet
    Debug.Print ActiveWorkbook.Name
    Debug.Print ActiveSheet.Name
    lastRowK = .Cells(9999, 1).End(xlUp).Row
    For i = 2 To lastRowK
        .Cells(i, "HI") = _
        .Cells(i, "Z") + .Cells(i, "AA") + .Cells(i, "AB") + .Cells(i, "AC") + .Cells(i, "AD") + .Cells(i, "AT") + .Cells(i, "AE") + .Cells(i, "AF") + .Cells(i, "AH") + _
        .Cells(i, "AI") + .Cells(i, "AJ") + .Cells(i, "AK") + .Cells(i, "AL") + .Cells(i, "AM") + .Cells(i, "AN") + .Cells(i, "AO") + .Cells(i, "AP") + .Cells(i, "BH") + _
        .Cells(i, "BI") + .Cells(i, "BJ") + .Cells(i, "BK") + .Cells(i, "BL") + .Cells(i, "BM") + .Cells(i, "BN") + .Cells(i, "BO") + .Cells(i, "BP") + .Cells(i, "BT") + _
        .Cells(i, "BW") + .Cells(i, "BX") + .Cells(i, "CD") + .Cells(i, "CE") + .Cells(i, "CG") + .Cells(i, "CI") + .Cells(i, "CW") + .Cells(i, "DB") + .Cells(i, "EL") + _
        .Cells(i, "EM") + .Cells(i, "FJ") + .Cells(i, "GU") + .Cells(i, "GY") + .Cells(i, "GZ") + .Cells(i, "HA") + .Cells(i, "HB") + .Cells(i, "HC") + .Cells(i, "HD") + _
        .Cells(i, "HF") + .Cells(i, "HG") + .Cells(i, "HH")
    Next i


End With

YYYY = Left(kyuyoBkName, InStr(kyuyoBkName, ".") - 1)
M = Right("0" & Replace(Mid(kyuyoBkName, InStr(kyuyoBkName, ".") + 1, 2), ".", ""), 2)
YM = YYYY & M
Debug.Print M

'�N�x����
If Val(M) <= 6 Then
    nendo = YYYY - 1
Else
    nendo = YYYY
End If



With macroBk.Sheets("���f�[�^")
    
    lastClm = .Cells(1, 9999).End(xlToLeft).Column
    On Error Resume Next
    cnt = Application.WorksheetFunction.Match(Val(nendo), .Range("2:2"), 0)
End With

'���łɓ��N�x�̃f�[�^�������Ă���ꍇ
'�N�x�V�[�g�Ɉړ��A�s�F�ԍ��A��FYM�ŏZ������A���z�̂ݎQ��
'�]�ƈ��ԍ��Ȃ��ꍇ�A���ʏ�����GO�i��ԉ��̍s�ɒ��O�ƃR�s�[���ԍ��������z�R�s�[�j

'���N�x�f�[�^���Ȃ��ꍇ
'���f�[�^�ɂ͐��07-12-06�܂ō���Ă��܂�
'�N�x�V�[�g�̍쐬�A�N�����̖���

'���ʏ���
'���f�[�^
'�Q�s�ڈȍ~�AB-0��͏����̂݃R�s�[�AP-X�͑S�R�s


'�Ȃ��ꍇ����t�H�[�}�b�g�쐬
If cnt < 1 Then
    With macroBk.Sheets("���f�[�^")
    .Cells(1, lastClm + 1).Value = YYYY & "07"
    .Cells(2, lastClm + 1).Value = nendo
    Debug.Print YM
    Debug.Print nendo
    For j = 1 To 11
        lastClm = .Cells(1, 9999).End(xlToLeft).Column
        
        If Right(.Cells(1, lastClm).Value, 2) = 12 Then
            .Cells(1, lastClm + 1).Value = .Cells(1, lastClm) + 89
        Else
            .Cells(1, lastClm + 1).Value = .Cells(1, lastClm) + 1
        End If
            .Cells(2, lastClm + 1).Value = nendo
    Next j
    End With
    
    '�V�[�g�����
    macroBk.Activate
    macroBk.Sheets("�N�x���^").Copy after:=Sheets("�N�x���^")
    Do While r < 100
        r = r + 1
    Loop

'    Sheets("�N�x���^(2)").Activate
    With ActiveSheet
        .Name = nendo
        .Cells(5, "B").Value = nendo & "�N�x"
        For j = 7 To 18
            If j <= 12 Then
                .Cells(6, j - 3).Value = nendo & "�N" & j & "��"
                .Cells(1, j - 3).Value = nendo & Right("0" & j, 2)
            Else
                .Cells(6, j - 3).Value = nendo + 1 & "�N" & j - 12 & "��"
                .Cells(1, j - 3).Value = nendo + 1 & Right("0" & j, 2)
            End If
        
        Next
    End With
End If

'�����܂Ńt�H�[�}�b�g�쐬

On Error Resume Next
tgtclm = WorksheetFunction.Match(Val(YM), rareSh.Range("1:1"), 0)
If tgtclm < 1 Then
    MsgBox "���f�[�^�V�[�g�ɊY�����錎�����݂��܂���"
    Exit Function
End If


'�������狋�^�f�[�^���琶�f�[�^�R�s�[

For i = 2 To lastRowK
    With Workbooks(kyuyoBkName).ActiveSheet
        On Error Resume Next
        cnt = WorksheetFunction.Match(.Cells(i, 1), rareSh.Range("B:B"), 0)
'���łɎЈ������邩
        If cnt < 1 Then
            
            '�l�R�s�[
            lastRowR = rareSh.Cells(9999, 2).End(xlUp).Row
            rareSh.Cells(lastRowR + 1, 2) = .Cells(i, 1)
            rareSh.Cells(lastRowR + 1, 3) = .Cells(i, 2)
            rareSh.Cells(lastRowR + 1, tgtclm) = .Cells(i, "HI")
        Else
            rareSh.Cells(cnt, tgtclm) = .Cells(i, "HI")
            cnt = 0
        End If
        
        
            '�s�R�s�[
'            rareSh.Range(rareSh.Cells(lastRowR, "B"), rareSh.Cells(lastRowR, "O")).Select
'            Do While r < 100
'                r = r + 1
'            Loop
'            Selection.Copy
'            rareSh.Range(rareSh.Cells(lastRowR + 1, "B"), rareSh.Cells(lastRowR + 1, "O")).PasteSpecial (xlPasteFormats)
'            Application.CutCopyMode = False
'            rareSh.Range(rareSh.Cells(lastRowR, "P"), rareSh.Cells(lastRowR, "X")).Copy
'            Do While r < 100
'                r = r + 1
'            Loop
'            rareSh.Range(rareSh.Cells(lastRowR + 1, "P"), rareSh.Cells(lastRowR + 1, "X")).Paste
'            Application.CutCopyMode = False
    
    
        On Error GoTo 0
        
    End With
Next

copyToCalBk = nendo

On Error GoTo 0

End Function

Sub makeHit(ByVal nendo As String)

Dim lastRowR, i, nendoSh, lastRowN, r, insClm, lastRowS

Set rareSh = ActiveWorkbook.Sheets("���f�[�^")

If nendo = "" Then
    nendo = InputBox("�R�s�[����N�x����͂��Ă�������")
Else
    
    With rareSh
    '�ԍ������@�Ɓ@���񂩂����
        Set nendoSh = ActiveWorkbook.Sheets(nendo)
        lastRowR = .Cells(9999, 2).End(xlUp).Row
        nendoclm = WorksheetFunction.Match(Val(nendo), .Range("2:2"), 0)
        .Range(.Cells(4, "B"), .Cells(lastRowR, 3)).Copy nendoSh.Cells(7, "B")
        .Range(.Cells(4, nendoclm), .Cells(lastRowR, nendoclm + 11)).Copy nendoSh.Cells(7, "D")
    
    End With
        
    With nendoSh
        .Activate
        lastRowN = lastRowR + 3
         .Range(.Cells(7, 2), .Cells(lastRowN, "O")).Borders.LineStyle = xlContinuous
         .Range(.Cells(7, "P"), .Cells(7, "X")).Copy
         Do While r < 100
            r = r + 1
        Loop
        ActiveSheet.Paste Destination:=.Range(.Cells(8, "P"), .Cells(lastRowN, "X"))
        
    '    With .Range(.Cells(6, "P"), .Cells(lastRowN, "P")).Borders(xlEdgeLeft)
    '        .LineStyle = xlbold
    '        .Weight = xlMedium
    '    End With
    End With
End If

Application.CutCopyMode = False


'�Y���҂̒��o-------------------

Dim srcSh As Object
Set srcSh = Sheets("�Y����")
Dim prevClm

With srcSh
    On Error Resume Next
    prevClm = WorksheetFunction.Match(nendo & "�N�x", .Range("4:4"), 0)
    If prevClm > 0 Then
        insClm = prevClm
    Else
        insClm = WorksheetFunction.Match("���݂̏���", .Range("4:4"), 0)
        .Columns(insClm).Insert
        .Cells(4, insClm).Value = nendo & "�N�x"
    End If
    
    For i = 7 To lastRowN
        lastRowS = .Cells(9999, 2).End(xlUp).Row
        r = 0
        If nendoSh.Cells(i, "X") = "��" Then
            r = WorksheetFunction.Match(nendoSh.Cells(i, 2), .Range("B:B"), 0)
            If r > 0 Then
                .Cells(r, insClm).Value = "��"
            Else
                .Rows(lastRowS + 1).Insert
                .Rows(lastRowS).Copy .Rows(lastRowS + 1)
                .Range(.Cells(lastRowS + 1, "B"), .Cells(lastRowS + 1, insClm)).ClearContents
                .Cells(lastRowS + 1, "B") = nendoSh.Cells(i, "B")
                .Cells(lastRowS + 1, "C") = nendoSh.Cells(i, "C")
                .Cells(lastRowS + 1, insClm).Value = "��"
            End If
        End If
    Next
End With

On Error GoTo 0

End Sub
