Attribute VB_Name = "mdlMessage"
Option Explicit

' �萔���`����
Public Const MSG_001   As String = "���̓t�@�C�������݂��Ă��܂���B"
Public Const MSG_002   As String = "�����̓t�@�C���ƐV���̓t�@�C���̐��ʂ��قȂ��Ă���̂ŁA�ēx���m�F���������B"
Public Const MSG_003   As String = "(1)���̓t�@�C���̃p�X���Ԉ���Ă���̂ŁA�ēx���m�F���������B"
Public Const MSG_004   As String = "�Q�̃��[�N�u�b�N���r����ۂɃV�[�g�����m�F���������B"
Public Const MSG_005   As String = "�t�@�C���^�C�v�͐���������܂���B"
Public Const MSG_006   As String = "�t�@�C�����d�����Ă��܂��B"
Public Const MSG_007   As String = "�t�@�C�����J���܂���ł����B"
Public Const MSG_008   As String = "�u�i�Q�j�v�t�@�C���́i�P�j�V�[�g���́u�i�R�j�v�t�@�C���̃V�[�g���ɈقȂ��Ă��܂��B"

'***************************************************
' �@�\      : ������u��������
' �Ԃ�l    : �����l
' ������    : ARG1 -�u�������镶��1
'           : ARG2 -�u�������镶��2
'           : ARG3 -�u�������镶��3
'           : ARG4 -�u�������镶��4
' ����      : THAOVTT
'***************************************************
Public Function RepMessage(strMessage As String _
                    , Optional strReplace1 As String _
                    , Optional strReplace2 As String _
                    , Optional strReplace3 As String _
                    , Optional strReplace4 As String) As String
    
    ' �u�������镶��1����͂����ꍇ
    If strReplace1 <> vbNullString Then
        strMessage = Replace(strMessage, "(1)", strReplace1)
    End If
    
    ' �u�������镶��2����͂����ꍇ
    If strReplace2 <> vbNullString Then
        strMessage = Replace(strMessage, "(2)", strReplace2)
    End If

    ' �u�������镶��3����͂����ꍇ
    If strReplace3 <> vbNullString Then
        strMessage = Replace(strMessage, "(3)", strReplace3)
    End If

    ' �u�������镶��4����͂����ꍇ
    If strReplace4 <> vbNullString Then
        strMessage = Replace(strMessage, "(4)", strReplace4)
    End If
    
    RepMessage = strMessage
End Function





