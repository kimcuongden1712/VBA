Attribute VB_Name = "mdlMessage"
Option Explicit

' 定数を定義する
Public Const MSG_001   As String = "入力ファイルが存在していません。"
Public Const MSG_002   As String = "旧入力ファイルと新入力ファイルの数量が異なっているので、再度ご確認ください。"
Public Const MSG_003   As String = "(1)入力ファイルのパスが間違っているので、再度ご確認ください。"
Public Const MSG_004   As String = "２つのワークブックを比較する際にシート数を確認ください。"
Public Const MSG_005   As String = "ファイルタイプは正しくありません。"
Public Const MSG_006   As String = "ファイルが重複しています。"
Public Const MSG_007   As String = "ファイルが開けませんでした。"
Public Const MSG_008   As String = "「（２）」ファイルの（１）シート名は「（３）」ファイルのシート名に異なっています。"

'***************************************************
' 機能      : 文字を置き換える
' 返り値    : 文字値
' 引き数    : ARG1 -置き換える文字1
'           : ARG2 -置き換える文字2
'           : ARG3 -置き換える文字3
'           : ARG4 -置き換える文字4
' 著者      : THAOVTT
'***************************************************
Public Function RepMessage(strMessage As String _
                    , Optional strReplace1 As String _
                    , Optional strReplace2 As String _
                    , Optional strReplace3 As String _
                    , Optional strReplace4 As String) As String
    
    ' 置き換える文字1を入力した場合
    If strReplace1 <> vbNullString Then
        strMessage = Replace(strMessage, "(1)", strReplace1)
    End If
    
    ' 置き換える文字2を入力した場合
    If strReplace2 <> vbNullString Then
        strMessage = Replace(strMessage, "(2)", strReplace2)
    End If

    ' 置き換える文字3を入力した場合
    If strReplace3 <> vbNullString Then
        strMessage = Replace(strMessage, "(3)", strReplace3)
    End If

    ' 置き換える文字4を入力した場合
    If strReplace4 <> vbNullString Then
        strMessage = Replace(strMessage, "(4)", strReplace4)
    End If
    
    RepMessage = strMessage
End Function





