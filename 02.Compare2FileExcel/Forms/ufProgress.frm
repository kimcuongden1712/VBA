VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "処理しています。。。"
   ClientHeight    =   1620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5760
   OleObjectBlob   =   "ufProgress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***************************************************
' @(f)
' 機能      : ユーザーフォームを初期するイベント
' 返り値    :
' 引き数    :
' 機能説明  :
' 備考      :
' 著者      : THAOVTT
'***************************************************
Private Sub UserForm_Initialize()
#If IsMac = False Then
    ' Windowsマシンで作業している場合は、タイトルバーを非表示にします。 それ以外の場合は、通常どおりに表示してください
    Me.Height = Me.Height - 10
    HideTitleBar.HideTitleBar Me
#End If
End Sub
