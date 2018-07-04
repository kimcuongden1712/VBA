Attribute VB_Name = "mdlMacro"
Option Explicit

'***************************************************
' 機能      : データのエクスポート処理
' 返り値    : なし
' 著者      : UNI
'***************************************************
Sub Comparison_Click()
' ハンドルエラー
On Error GoTo ErrHandle
    
    '変数を定義する
    Dim filePathArr As Variant
    Dim F1_Workbook As Workbook, F2_Workbook As Workbook
    
    Dim sh As Integer, ShName As String, lColIdx As Long, sIdx As Long, ssh As String, strWorkbookMain As String
    
    Dim iRow As Double, iCol As Double, iRow_Max As Double, iCol_Max As Double
    
    Dim File1_Path As String, File2_Path As String, F1_Data As String, F2_Data As String
    
    Dim pctdone As Single, countFile As Integer, row As Integer
    
    Debug.Print "Start=" & VBA.DateTime.Now
    
    strWorkbookMain = ActiveWorkbook.Name
    
    '配列にファイルパスをプッシュします。
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
        'シートの結果を削除します。
        Call ClearSheetResult
        
        'プログレスバーを表示する
        ufProgress.LabelProgress.Width = 0
        ufProgress.Show
        
        sIdx = 1
        
        For countFile = 1 To UBound(filePathArr)
            File1_Path = CStr(filePathArr(countFile, 1))
            File2_Path = CStr(filePathArr(countFile, 2))
    
            'プログレスバーを定期的に更新する
            pctdone = countFile / UBound(filePathArr)
            With ufProgress
                .LabelCaption.Caption = UBound(filePathArr) & "行中の" & countFile & "行目で実行中"
                .LabelProgress.Width = pctdone * (.FrameProgress.Width)
            End With
            DoEvents
            
            With Workbooks(strWorkbookMain).Sheets(1)
                .Cells(5 + countFile, COLUMN_OLD_FILE).Interior.Color = vbWhite
                .Cells(5 + countFile, COLUMN_NEW_FILE).Interior.Color = vbWhite
            End With

            'F1のファイルとF2ファイルを開く
            Application.ScreenUpdating = False
            Set F1_Workbook = Workbooks.Open(File1_Path)
            Set F2_Workbook = Workbooks.Open(File2_Path)
            
            If F1_Workbook.Sheets.Count = F2_Workbook.Sheets.Count Then
            
            '旧ファイルと新ファイルのシート毎をループして開く
            For sh = 1 To F2_Workbook.Sheets.Count
         
                ShName = F2_Workbook.Sheets(sh).Name
             
                iRow_Max = F2_Workbook.Sheets(ShName).Range("A:A").SpecialCells(xlLastCell).row
                iCol_Max = F2_Workbook.Sheets(ShName).Range("A:A").SpecialCells(xlLastCell).Column
                For iRow = 1 To iRow_Max
                    For iCol = 1 To iCol_Max
                        F1_Data = F1_Workbook.Sheets(ShName).Cells(iRow, iCol)
                        F2_Data = F2_Workbook.Sheets(ShName).Cells(iRow, iCol)
                        
                        'エクセルの各シートからデータを比較して差分を出力する。
                        If F1_Data <> F2_Data Then
                            'カラムNo
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 2) = sIdx
                            '実行日付のカラム
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 3) = Date
                            'ファイル名（新）のカラム
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 4) = File2_Path
                            'シート名称のカラム
                            If ssh <> F2_Workbook.Sheets(sh).Name Then
                                ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 5) = F2_Workbook.Sheets(sh).Name
                                ssh = F2_Workbook.Sheets(sh).Name
                            Else
                                ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 5) = F2_Workbook.Sheets(sh).Name
                            End If
                            '旧ファイルのカラム
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 6) = F1_Data
                            '新ファイルのカラム
                            ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Cells(3 + sIdx, 7) = F2_Data
                            'フラグ
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
            
            ' プログレスバーの表示を停止
            If countFile = UBound(filePathArr) Then Unload ufProgress
        Next countFile
        
        'プロセスが完了しました
        ThisWorkbook.Sheets(G_SHEET_NAME_RESULT).Activate
        Debug.Print "End=" & VBA.DateTime.Now
        
        With Workbooks(strWorkbookMain).Sheets(1)
            row = .Cells(Rows.Count, GetColumnLetter(Range(COLUMN_OLD_FILE & 6).Column - 1)).End(xlUp).row
            .Range(GetColumnLetter(Range(COLUMN_OLD_FILE & 6).Column - 1) & 6, COLUMN_NEW_FILE & row).Interior.Color = vbWhite
        End With
        
        Application.ScreenUpdating = True
    End If
ErrHandle:
    If Err.Description Like "*'Open' メソッドは失敗しました: 'Workbooks' オブジェクト*" Then
        MsgBox RepMessage(MSG_007), vbCritical, G_PROJECT_NAME
        Unload ufProgress
        Application.ScreenUpdating = True
        Exit Sub
    End If
End Sub

'***************************************************
' 機能      : シートの結果を削除します
' 返り値    : なし
' 著者      : AnhTT
'***************************************************
Public Sub ClearSheetResult()
    With ThisWorkbook.Sheets(G_SHEET_NAME_RESULT)
        .Rows(4 & ":" & .Rows.Count).ClearContents
    End With
End Sub
