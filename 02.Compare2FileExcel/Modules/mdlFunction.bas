Attribute VB_Name = "mdlFunction"
'***********************************************************
' �@�\      : �J�����̃A�h���X����z��p�X�t�@�C�����擾����
' ������    : ARG1 - ���t�@�C���J����
'           : ARG2 - �V�t�@�C���J����
' �Ԃ�l    : ��z��p�X�t�@�C��
' ����      : BAODTQ
'***********************************************************
Public Function GetArrFile(strColumnF1 As String, strColumnF2) As Variant
' �n���h���G���[
On Error GoTo ErrHandle
    
    Dim row As Integer

    row = Cells(Rows.Count, strColumnF1).End(xlUp).row
    If row >= 6 Then
        GetArrFile = Range(strColumnF1 & 6, strColumnF2 & row).Value
    Else
        GetArrFile = Empty
    End If
    
ErrHandle:
    If Err.Number <> 0 Then
        GetArrFile = Empty
    End If
End Function

'***************************************************
' �@�\      : ���s����O�̃o���f�[�V����
' ������    : ARG1 - ���[�N�u�b�N
'           : ARG2 - ���̓t�@�C���̔z��
' �Ԃ�l    : True/False
' ����      : BAODTQ
'***************************************************
Public Function Validate(wbName As String, arrFileInput As Variant) As Boolean
' �n���h���G���[
On Error GoTo ErrHandle
    
    '�ϐ����`����
    Dim File1_Path As String
    Dim File2_Path As String
    Dim countFile As Integer
    
    Validate = False
    
    For countFile = 1 To UBound(arrFileInput)
        File1_Path = CStr(arrFileInput(countFile, 1))
        File2_Path = CStr(arrFileInput(countFile, 2))
        
        If File1_Path = File2_Path And Not StrIsEmpty(File1_Path) Then
            If Not CheckFileExist(File1_Path) Then
                MsgBox RepMessage(MSG_003, " "), vbCritical, G_PROJECT_NAME
            Else
                MsgBox RepMessage(MSG_006), vbCritical, G_PROJECT_NAME
            End If
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_OLD_FILE).Interior.Color = vbRed
                .Cells(5 + countFile, COLUMN_NEW_FILE).Interior.Color = vbRed
            End With
            Validate = True
            Exit Function
        End If
        
        Application.ScreenUpdating = False
        
        Call OutPutError(File1_Path, COLUMN_OLD_FILE, wbName, Validate, countFile)
        
        Call OutPutError(File2_Path, COLUMN_NEW_FILE, wbName, Validate, countFile)
        
        Call Check2File(File1_Path, File2_Path, COLUMN_OLD_FILE, COLUMN_NEW_FILE, wbName, Validate, countFile)
        
        Application.ScreenUpdating = True
    Next countFile
    
ErrHandle:
    If Err.Number <> 0 Then
        Validate = True
    End If
End Function

'***************************************************
' �@�\      : ���́E�o�̓t�@�C���ɂ̓G���[���������邱�Ƃ��m�F����
' ������    : ARG1 - �t�@�C���p�X
'           : ARG2 - �J����
'           : ARG3 - ���[�N�u�b�N
'           : ARG4 - �u�[���l
'           : ARG5 - �t�@�C���̃C���f�b�N�X
' �Ԃ�l    : �Ȃ�
' ����      : BAODTQ
'***************************************************
Public Sub OutPutError(File_Path As String, COLUMN_FILE As String, wbName As String, Validate As Boolean, countFile As Integer)
' �n���h���G���[
On Error GoTo ErrHandle

        '�t�@�C������ł��邱�Ƃ��m�F����
        If StrIsEmpty(File_Path) Then
            MsgBox RepMessage(MSG_001), vbCritical, G_PROJECT_NAME
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_FILE).Interior.Color = vbRed
            End With
            Validate = True
        ElseIf Not CheckFileExist(File_Path) Then
            '�t�@�C�������݂��Ă��邱�Ƃ��m�F����
            If COLUMN_FILE = COLUMN_NEW_FILE Then
                MsgBox RepMessage(MSG_003, "�V"), vbCritical, G_PROJECT_NAME
            Else
                MsgBox RepMessage(MSG_003, "��"), vbCritical, G_PROJECT_NAME
            End If
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_FILE).Interior.Color = vbRed
            End With
            Validate = True
        ElseIf Not checkFileType(File_Path) Then
            '�t�@�C���^�C�v���m�F����
            MsgBox RepMessage(MSG_005), vbCritical, G_PROJECT_NAME
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_FILE).Interior.Color = vbRed
            End With
            Validate = True
        End If
ErrHandle:
    If Err.Number <> 0 Then
        Validate = True
    End If
End Sub

'***************************************************
' �@�\      : ���́E�o�̓t�@�C���ɂ̓G���[���������邱�Ƃ��m�F����
' ������    : ARG1 - �t�@�C���p�X
'           : ARG2 - �J����
'           : ARG3 - ���[�N�u�b�N
'           : ARG4 - �u�[���l
'           : ARG5 - �t�@�C���̃C���f�b�N�X
' �Ԃ�l    : �Ȃ�
' ����      : AnhTT
'***************************************************
Public Sub Check2File(File1_Path As String, File2_Path As String, COLUMN_OLD_FILE As String, COLUMN_NEW_FILE As String, wbName As String, Validate As Boolean, countFile As Integer)
' �n���h���G���[
On Error GoTo ErrHandle
        
    Dim F1_Workbook As Workbook, F2_Workbook As Workbook, sh As Integer
    Dim ShNameF1 As String, ShNameF2 As String
    
    Set F1_Workbook = Workbooks.Open(File1_Path)
    Set F2_Workbook = Workbooks.Open(File2_Path)
    
    If F1_Workbook.Sheets.Count <> F2_Workbook.Sheets.Count Then
        MsgBox RepMessage(MSG_004, " "), vbCritical, G_PROJECT_NAME
        With Workbooks(wbName).Sheets(1)
            .Cells(5 + countFile, COLUMN_OLD_FILE).Interior.Color = vbRed
            .Cells(5 + countFile, COLUMN_NEW_FILE).Interior.Color = vbRed
        End With
        Validate = True
    End If
    For sh = 1 To F2_Workbook.Sheets.Count
        ShNameF1 = F1_Workbook.Sheets(sh).Name
        ShNameF2 = F2_Workbook.Sheets(sh).Name
        
        If ShNameF1 <> ShNameF2 Then
            MsgBox RepMessage(MSG_008, ShNameF2, File2_Path, File1_Path), vbCritical, G_PROJECT_NAME
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_OLD_FILE).Interior.Color = vbRed
                .Cells(5 + countFile, COLUMN_NEW_FILE).Interior.Color = vbRed
            End With
            Validate = True
        End If
    Next sh
    
    F2_Workbook.Close savechanges:=False
    F1_Workbook.Close savechanges:=False
    Set F2_Workbook = Nothing
    Set F1_Workbook = Nothing
    
ErrHandle:
    If Err.Number <> 0 Then
        Validate = True
    End If
End Sub

