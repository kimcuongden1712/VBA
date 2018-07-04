Attribute VB_Name = "mdlFunction"
'***********************************************************
' 機能      : カラムのアドレスから配列パスファイルを取得する
' 引き数    : ARG1 - 旧ファイルカラム
'           : ARG2 - 新ファイルカラム
' 返り値    : ら配列パスファイル
' 著者      : BAODTQ
'***********************************************************
Public Function GetArrFile(strColumnF1 As String, strColumnF2) As Variant
' ハンドルエラー
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
' 機能      : 実行する前のバリデーション
' 引き数    : ARG1 - ワークブック
'           : ARG2 - 入力ファイルの配列
' 返り値    : True/False
' 著者      : BAODTQ
'***************************************************
Public Function Validate(wbName As String, arrFileInput As Variant) As Boolean
' ハンドルエラー
On Error GoTo ErrHandle
    
    '変数を定義する
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
' 機能      : 入力・出力ファイルにはエラーが発生することを確認する
' 引き数    : ARG1 - ファイルパス
'           : ARG2 - カラム
'           : ARG3 - ワークブック
'           : ARG4 - ブール値
'           : ARG5 - ファイルのインデックス
' 返り値    : なし
' 著者      : BAODTQ
'***************************************************
Public Sub OutPutError(File_Path As String, COLUMN_FILE As String, wbName As String, Validate As Boolean, countFile As Integer)
' ハンドルエラー
On Error GoTo ErrHandle

        'ファイルが空であることを確認する
        If StrIsEmpty(File_Path) Then
            MsgBox RepMessage(MSG_001), vbCritical, G_PROJECT_NAME
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_FILE).Interior.Color = vbRed
            End With
            Validate = True
        ElseIf Not CheckFileExist(File_Path) Then
            'ファイルが存在していることを確認する
            If COLUMN_FILE = COLUMN_NEW_FILE Then
                MsgBox RepMessage(MSG_003, "新"), vbCritical, G_PROJECT_NAME
            Else
                MsgBox RepMessage(MSG_003, "旧"), vbCritical, G_PROJECT_NAME
            End If
            With Workbooks(wbName).Sheets(1)
                .Cells(5 + countFile, COLUMN_FILE).Interior.Color = vbRed
            End With
            Validate = True
        ElseIf Not checkFileType(File_Path) Then
            'ファイルタイプを確認する
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
' 機能      : 入力・出力ファイルにはエラーが発生することを確認する
' 引き数    : ARG1 - ファイルパス
'           : ARG2 - カラム
'           : ARG3 - ワークブック
'           : ARG4 - ブール値
'           : ARG5 - ファイルのインデックス
' 返り値    : なし
' 著者      : AnhTT
'***************************************************
Public Sub Check2File(File1_Path As String, File2_Path As String, COLUMN_OLD_FILE As String, COLUMN_NEW_FILE As String, wbName As String, Validate As Boolean, countFile As Integer)
' ハンドルエラー
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

