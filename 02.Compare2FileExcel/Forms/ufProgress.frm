VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufProgress 
   Caption         =   "�������Ă��܂��B�B�B"
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
' �@�\      : ���[�U�[�t�H�[������������C�x���g
' �Ԃ�l    :
' ������    :
' �@�\����  :
' ���l      :
' ����      : THAOVTT
'***************************************************
Private Sub UserForm_Initialize()
#If IsMac = False Then
    ' Windows�}�V���ō�Ƃ��Ă���ꍇ�́A�^�C�g���o�[���\���ɂ��܂��B ����ȊO�̏ꍇ�́A�ʏ�ǂ���ɕ\�����Ă�������
    Me.Height = Me.Height - 10
    HideTitleBar.HideTitleBar Me
#End If
End Sub
