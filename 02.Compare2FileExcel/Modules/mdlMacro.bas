Attribute VB_Name = "mdlMacro"
Option Explicit

'***************************************************
' �@�\      : �f�[�^�̃G�N�X�|�[�g����
' �Ԃ�l    : �Ȃ�
' ����      : UNI
'***************************************************
Sub Comparison_Click()
' �n���h���G���[
On Error GoTo ErrHandle
    
    '�ϐ����`����
    Dim filePathArr As Variant
    Dim F1_Workbook As Workbook, F2_Workbook As Workbook
    
    Dim sh As Integer, ShName As String, lColIdx As Long, sIdx As Long, ssh As String, strWorkbookMain As String
    
    Dim iRow As Double, iCol As Double, iRow_Max As Double, iCol_Max As Double
    
    Dim File1_Path As String, File2_Path As String, F1_Data As String, F2_Data As String
    
    Dim pctdone As Single, countFile As Integer, row As Integer
    
    Debug.Print "Start=" & VBA.DateTime.Now
    
    strWorkbookMain = ActiveWorkbook.Name
    
    '�z��Ƀt�@�C���p�X���v�b�V�����܂��B
    filePathArr = GetArrFile(COLUMN_OLD_FILE, COLUMN_NEW_FILE)
    If IsEmpty(filePathArr) Then
        MsgBox RepMessage(MSG_001), vbCritical, G_PROJECT_NAME
        Debug.Print "End=" & VBA.DateTime.Now
        Exit Sub
    End If
    
    If Validate(strWorkbookMain, filePathArr) Then
        Debug.Print "End=" & VBA.DateTime.Now
        Exit Sub
    Else
        '�V�[�g�̌��ʂ��폜���܂��B
        Call ClearSheetResult
        
        '�v���O���X�o�[��\������
        ufProgress.LabelProgress.Width = 0
        ufProgress.Show
        
        sIdx = 1
        
        For countFile = 1 To UBound(filePathArr)
            File1_Path = CStr(filePathArr(countFile, 1))
            File2_Path = CStr(filePathArr(countFile, 2))
    
            '�v���O���X�o�[�����I�ɍX�V����
            pctdone = countFile / UBound(filePathArr)
            With ufProgress
                .LabelCaption.Caption = UBound(filePathArr) & "�s����" & countFile & "�s�ڂŎ��s��"
                .LabelProgress.Width = pctdone * (.FrameProgress.Width)
            End With
            DoEvents
            
            With Workbooks(strWorkbookMain).Sheets(1)
                .Cells(5 + countFile, COLUMN_OLD_FILE).Interior.Color = vbWhite
                .Cells(5 + countFile, COLUMN_NEW_FILE).Interior.Color = vbWhite
            End With

            'F1�̃t�@�C����F2�t�@�C�����J��
            Application.ScreenUpdating = False
            Set F1_Workbook = Workbooks.Open(File1_Path)
            Set F2_Workbook = Workbooks.Open(File2_Path)
            
            If F1_Workbook.Sheets.Count = F2_Workbook.Sheets.Count Then
            
            '���t�@�C���ƐV�t�@�C���̃V�[�g�������[�v���ĊJ��
            For sh = 1 To F2_Workbook.Sheets.Count
         
                ShName = F2_Workbook.Sheets(sh).Name
             
                iRow_Max = F2_Workbook.Sheets(ShName).Range("A:A").SpecialCells(xlLastCell).row
                iCol_Max = F2_Workbook.Sheets(ShName).Range("A:A").SpecialCells(xlLastCell).Column
                For iRow = 1 To iRow_Max
                    For iCol = 1 To iCol_Max
                        F1_Data = F1_Workbook.Sheets(ShName).Cells(iRow, iCol)
                        F2_Data = F2_Workbook.Sheets(ShName).Cells(iRow, iCol)
                        
                        '�G�N�Z���̊e�V�[�g����f�[�^���r���č������o�͂���B
                        If F1_Data <> F2_Data Then
                            '�J����No
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 2) = sIdx
                            '���s���t�̃J����
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 3) = Date
                            '�t�@�C�����i�V�j�̃J����
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 4) = File2_Path
                            '�V�[�g���̂̃J����
                            If ssh <> F2_Workbook.Sheets(sh).Name Then
                                ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 5) = F2_Workbook.Sheets(sh).Name
                                ssh = F2_Workbook.Sheets(sh).Name
                            Else
                                ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 5) = F2_Workbook.Sheets(sh).Name
                            End If
                            '���t�@�C���̃J����
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 6) = F1_Data
                            '�V�t�@�C���̃J����
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 7) = F2_Data
                            '�t���O
                            sIdx = sIdx + 1
                        End If
                    Next iCol
                Next iRow
            Next sh
            
            End If
                        
            F2_Workbook.Close savechanges:=False
            F1_Workbook.Close savechanges:=False
            Set F2_Workbook = Nothing
            Set F1_Workbook = Nothing
            
            ' �v���O���X�o�[�̕\�����~
            If countFile = UBound(filePathArr) Then Unload ufProgress
        Next countFile
        
        '�v���Z�X���������܂���
        ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Activate
        Debug.Print "End=" & VBA.DateTime.Now
        
        With Workbooks(strWorkbookMain).Sheets(1)
            row = .Cells(Rows.Count, GetColumnLetter(Range(COLUMN_OLD_FILE & 6).Column - 1)).End(xlUp).row
            .Range(GetColumnLetter(Range(COLUMN_OLD_FILE & 6).Column - 1) & 6, COLUMN_NEW_FILE & row).Interior.Color = vbWhite
        End With
        
        Application.ScreenUpdating = True
    End If
ErrHandle:
    If Err.Description Like "*'Open' ���\�b�h�͎��s���܂���: 'Workbooks' �I�u�W�F�N�g*" Then
        MsgBox RepMessage(MSG_007), vbCritical, G_PROJECT_NAME
        Unload ufProgress
        Application.ScreenUpdating = True
        Exit Sub
    End If
End Sub

'***************************************************
' �@�\      : �V�[�g�̌��ʂ��폜���܂�
' �Ԃ�l    : �Ȃ�
' ����      : AnhTT
'***************************************************
Public Sub ClearSheetResult()
    With ThisWorkbook.Sheets(G_SHEET_NAME_RESULT)
        .Rows(4 & ":" & .Rows.Count).ClearContents
    End With
End Sub
