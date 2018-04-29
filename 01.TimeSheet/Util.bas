Attribute VB_Name = "Util"
Option Explicit

'***************************************************
' @ (f)
' Function       : GetDateFromString
' Returns        : String Array
' Argument       : fullDate
' Description    : Get Date From String
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Function GetDateFromString(fullDate As String) As String()
    Dim arrDate() As String
    If IsEmpty(fullDate) = False Then
        arrDate = Split(fullDate, " ")
    End If
    GetDateFromString = arrDate
End Function

'***************************************************
' @ (f)
' Function       : CompareDate
' Returns        : Boolean
' Argument       : curDate - Current Date
'                : tmpDate - Tem Date
' Description    : Compare Date
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Function CompareDate(curDate As String, tmpDate As String) As Boolean
    CompareDate = False
    If curDate = tmpDate Then
        CompareDate = True
    End If
End Function

'***************************************************
' @ (f)
' Function       : CheckValidateInput
' Returns        : Boolean
' Argument       : fromDate - Start Date
'                : endDate - End Date
' Description    : Check valdate Input
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Function CheckValidateInput(fromDate As String, endDate As String) As Boolean
    ' Check date
    If IsEmpty(fromDate) = True Or IsEmpty(endDate) = True Then
        MsgBox "期間が入力または選択してください。"
        CheckValidateInput = False
        Exit Function
    ElseIf Not IsDate(fromDate) Or Not IsDate(endDate) Then
        MsgBox "期間が不正です。修正してください。"
        CheckValidateInput = False
        Exit Function
    ElseIf (endDate < fromDate) Then
        MsgBox "開始日より大きい終了日を入力して下さい。"
        CheckValidateInput = False
        Exit Function
    End If
End Function

'***************************************************
' @ (f)
' Function       : GetTimeStartWorks
' Returns        : String
' Argument       : startTime
' Description    : Get Time Start Works
' Author         : AnhTT
' Create         : 2018-08-07
' Update         : 2018-04-19
' Remarks        :
'***************************************************

Public Function GetTimeStartWorks(startTime As String) As String
    
    Dim dt As Date
    Dim h, m, s As Integer
    
    dt = CDate(startTime)
    
    If dt <= CDate("8:00:00 AM") Then
        'time <= 8:00:00 is set 8:00
        GetTimeStartWorks = "8:00"
        Exit Function
    
    Else
        h = hour(dt)
        m = Minute(dt)
        s = Second(dt)
        
        '   time  > 8:00:00
        '   Case 00:01 > 29:59 set + 30
        '   Case 30:00 set + 30
        '   Case 30:01 > 59:59 set + 60
        
            If (h >= 8 And m <= 29 And s >= 1) Then
                m = 30
            ElseIf (h = 8 And m = 30 And s = 0) Then
                m = 30
            ElseIf (h >= 8 And m >= 30 And s >= 1) Then
                m = 0
                h = h + 1
            End If
            
            GetTimeStartWorks = h & ":" & m
        Exit Function
    End If
    
End Function

'***************************************************
' @ (f)
' Function       : GetTimeEndWorks
' Returns        : String Array
' Argument       : endTime
' Description    : Get Time End Works
' Author         : AnhTT
' Create         : 2018-08-07
' Update         : 2018-04-19
' Remarks        :
'***************************************************

Public Function GetTimeEndWorks(endTime As String) As String()

    Dim dt As Date
    Dim timeOT, timeOut, dateTemp As String
    Dim h, m, s As Integer
    Dim arr(1) As String
    ReDim a(1)
    
    
    dt = CDate(endTime)
    dateTemp = Format(dt, "hh:mm:ss")
    
    'Get Hour, Minute, Second
    h = hour(dt)
    m = Minute(dt)
    s = Second(dt)
    
    'Time <17:30:00 set Time OT = 1.00
    If CDate(dateTemp) < CDate("17:00:00") Then
        If (h >= 1 And m <= 29 And s >= 1) Then
            m = 0
        ElseIf (h = 1 And m = 30 And s = 0) Then
            m = 30
        ElseIf (h >= 1 And m >= 30 And s >= 1) Then
            m = 30
            h = h
        End If
        
        timeOut = h & ":" & m
        timeOT = "1.00"
        
    ElseIf CDate(dateTemp) >= CDate("17:00:00") And CDate(dateTemp) <= CDate("17:30:00") Then
        m = 30
        timeOut = h & ":" & m
        timeOT = "1.00"
        
    ElseIf CDate(dateTemp) > CDate("17:30:00") Then
        'Time >17:30:00
        '   Case 00:01 > 29:59 set + 30
        '   Case 30:00 set + 30
        '   Case 30:01 > 59:59 set + 60
        '   And Set Time OT = 1.50
        If (h >= 17 And m <= 29 And s >= 1) Then
            m = 30
        ElseIf (h = 17 And m = 30 And s = 0) Then
            m = 30
        ElseIf (h >= 17 And m >= 30 And s >= 1) Then
            m = 0
            h = h + 1
        End If
        
        timeOut = h & ":" & m
        timeOT = "1.5"
    End If
    
        arr(0) = timeOut
        arr(1) = timeOT
        
    GetTimeEndWorks = arr
End Function

'***************************************************
' @ (f)
' Function       : ConvertStringToDate
' Returns        :
' Argument       : dateTime - Date
'                : flagTime -
' Description    : Convert String To Date
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Function ConvertStringToDate(dateTime As String, flagTime As String)
    ConvertStringToDate = TimeValue(dateTime)
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
