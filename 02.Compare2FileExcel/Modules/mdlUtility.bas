Attribute VB_Name = "mdlUtility"
Option Explicit

'***************************************************
' �@�\      : �񖼂��擾����
' �Ԃ�l    : �����l
' ������    : ARG1 - ��A�h���X
' ����      : THAOVTT
'***************************************************
Function GetColumnLetter(ColNum As Long) As String
    ' �ϐ����`����
    Dim vArr
    
    ' �ϐ��ɒl��ݒ肷��
    vArr = Split(Cells(1, ColNum).Address(True, False), "$")
    
    GetColumnLetter = vArr(0)
End Function

'***************************************************
' �@�\      : �t�@�C�������݂��邱�Ƃ��m�F����
' �Ԃ�l    : True/False
' ������    : ARG1 - �t�@�C���p�X
' ����      : THAOVTT
'***************************************************
Public Function CheckFileExist(strPath As String) As Boolean
    ' �n���h���G���[
    On Error GoTo ErrHandle

    If (Len(strPath) > 0) And (Len(Dir(strPath)) > 0) Then
        CheckFileExist = True
    Else
        CheckFileExist = False
    End If
ErrHandle:
    ' �P�[�X�G���[
    If Err.Number <> 0 Then
        CheckFileExist = False
    End If
End Function

'***************************************************
' �@�\      : �t�H���_�����݂��邱�Ƃ��m�F����
' �Ԃ�l    : True/False
' ������    : ARG1 - �t�H���_�p�X
' ����      : THAOVTT
'***************************************************
Public Function CheckFolderIsExits(strPath As String) As Boolean
    Dim fso As Object
    ' FileSystemObject�̃C���X�^���X����N������
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(strPath) Then
        CheckFolderIsExits = True
    Else
        CheckFolderIsExits = False
    End If
End Function

'***************************************************
' �@�\      : �����̋�`�F�b�N
' �Ԃ�l    : True/False
' ������    : ARG1 - ������
' ����      : ANHTT
'***************************************************
Public Function StrIsEmpty(strName As String) As Boolean
    StrIsEmpty = False
    If Trim(strName & vbNullString) = vbNullString Then
        StrIsEmpty = True
    End If
End Function

'***************************************************
' �@�\      : �t�@�C���^�C�v�̃`�F�b�N
' ������    : ARG1 - ������
' �Ԃ�l    : True/False
' ����      : BaoDTQ
'***************************************************
Public Function checkFileType(strPath As String) As Boolean
On Error Resume Next
    
    Dim fso As Object
    Dim strFileName As String
    
    ' FileSystemObject�̃C���X�^���X����N������
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strFileName = fso.GetFileName(strPath)
    
    If strFileName Like "*xls*" Then
        checkFileType = True
    Else
        checkFileType = False
    End If
    
End Function
