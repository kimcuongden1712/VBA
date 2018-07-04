Attribute VB_Name = "mdlUtility"
Option Explicit

'***************************************************
' 機能      : 列名を取得する
' 返り値    : 文字値
' 引き数    : ARG1 - 列アドレス
' 著者      : THAOVTT
'***************************************************
Function GetColumnLetter(ColNum As Long) As String
    ' 変数を定義する
    Dim vArr
    
    ' 変数に値を設定する
    vArr = Split(Cells(1, ColNum).Address(True, False), "$")
    
    GetColumnLetter = vArr(0)
End Function

'***************************************************
' 機能      : ファイルが存在することを確認する
' 返り値    : True/False
' 引き数    : ARG1 - ファイルパス
' 著者      : THAOVTT
'***************************************************
Public Function CheckFileExist(strPath As String) As Boolean
    ' ハンドルエラー
    On Error GoTo ErrHandle

    If (Len(strPath) > 0) And (Len(Dir(strPath)) > 0) Then
        CheckFileExist = True
    Else
        CheckFileExist = False
    End If
ErrHandle:
    ' ケースエラー
    If Err.Number <> 0 Then
        CheckFileExist = False
    End If
End Function

'***************************************************
' 機能      : フォルダが存在することを確認する
' 返り値    : True/False
' 引き数    : ARG1 - フォルダパス
' 著者      : THAOVTT
'***************************************************
Public Function CheckFolderIsExits(strPath As String) As Boolean
    Dim fso As Object
    ' FileSystemObjectのインスタンスを一つ起動する
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If fso.FolderExists(strPath) Then
        CheckFolderIsExits = True
    Else
        CheckFolderIsExits = False
    End If
End Function

'***************************************************
' 機能      : 文字の空チェック
' 返り値    : True/False
' 引き数    : ARG1 - 文字列
' 著者      : ANHTT
'***************************************************
Public Function StrIsEmpty(strName As String) As Boolean
    StrIsEmpty = False
    If Trim(strName & vbNullString) = vbNullString Then
        StrIsEmpty = True
    End If
End Function

'***************************************************
' 機能      : ファイルタイプのチェック
' 引き数    : ARG1 - 文字列
' 返り値    : True/False
' 著者      : BaoDTQ
'***************************************************
Public Function checkFileType(strPath As String) As Boolean
On Error Resume Next
    
    Dim fso As Object
    Dim strFileName As String
    
    ' FileSystemObjectのインスタンスを一つ起動する
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    strFileName = fso.GetFileName(strPath)
    
    If strFileName Like "*xls*" Then
        checkFileType = True
    Else
        checkFileType = False
    End If
    
End Function
