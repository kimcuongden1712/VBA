Attribute VB_Name = "Timesheet"
'***********************************************************************************************
'*                               Macro Calculator timesheet                                    *
'*Author        : AnhTT                                                                        *
'*Email         : trantheanh.se@gmail.com                                                      *
'*Create        :  2018-08-07                                                                  *
'*Update        :  2018-08-07                                                                  *
'*TODO          :                                                                              *
'*1. Caculator Working Time                                                                    *
'*2. Sum Working Date and OT neu co                                                            *
'*3.                                                                                           *
'*4.                                                                                           *
'*5.                                                                                           *
'***********************************************************************************************

'***************************************************
'@ (s)
' Function       : TimeSheetMacro
' Returns        : None
' Argument       : None
' Description    : Caculator TimeSheet Macro
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Sub TimeSheetMacro()
    Dim rng As Range
    
    Dim curDate As String, tmpDate As String, startDate As String, endDate As String, fullDate As String, oldDate As String, hour As String, flagDate As String
    Dim iRow, jCol, x, flagIncerment As Integer
    
    Dim dateTime() As String
    Dim b() As String
    
    Dim userName As String
    userName = Cells(1, 10).Value
    
    If CheckUserName(userName) = False Then
    
        Exit Sub
    End If
    
    
    'Get ranger
    Set rgn = Range(GetRange(userName))
    
    If rgn.Rows.Count < 0 Then Exit Sub
    
    'Draw format
    Call DrawFormat
    
    flagIncerment = 1
    
    'Set Value
    For x = 1 To rgn.Rows.Count
    '1. Get value range
        fullDate = rgn.Cells(x, 1).Value
        
        If x > 1 Then
            oldDate = rgn.Cells(x - 1, 1).Value
        End If
        
        If x = rgn.Rows.Count Then
            oldDate = rgn.Cells(x, 1).Value
            dateTime = GetDateFromString(oldDate)
            b = GetTimeEndWorks(oldDate)
            'Cells(2 + flagIncerment - 1, 14).Value = dateTime(1)
            Cells(2 + flagIncerment - 1, 14).Value = b(0)
            Cells(2 + flagIncerment - 1, 15).Value = IIf(flagDate = "AM", "S", "C")
            Cells(2 + flagIncerment - 1, 16).Value = b(1)
        End If
        
    '2. Get StartDate va EndDate
        dateTime = GetDateFromString(fullDate)
        hour = dateTime(1)
        flagDate = dateTime(2)
        curDate = dateTime(0)

        If CompareDate(curDate, tmpDate) = False Then
            'Print start Works
            Dim a As String
            a = GetTimeStartWorks(dateTime(1))
            
            Cells(2 + flagIncerment, 11).Value = curDate
            Cells(2 + flagIncerment, 12).Value = a
            Cells(2 + flagIncerment, 13).Value = IIf(flagDate = "AM", "S", "C")
            
            'Print end Works
            If oldDate <> "" Then
                dateTime = GetDateFromString(oldDate)
                
                b = GetTimeEndWorks(oldDate)
                flagDate = dateTime(2)
                'Cells(2 + flagIncerment - 1, 14).Value = dateTime(1)
                Cells(2 + flagIncerment - 1, 14).Value = b(0)
                Cells(2 + flagIncerment - 1, 15).Value = IIf(flagDate = "AM", "S", "C")
                Cells(2 + flagIncerment - 1, 16).Value = b(1)
            End If
            
            'Set parameter
            tmpDate = curDate
            flagIncerment = flagIncerment + 1
        End If
    Next x
End Sub

'***************************************************
'@ (s)
' Function       : DrawFormat
' Returns        : None
' Argument       : None
' Description    : Draw Format
' Author         : AnhTT
' Create         : 2018-04-19
' Update         :
' Remarks        :
'***************************************************

Public Sub DrawFormat()
    Range("K:Q").Clear
    Range("K:K").NumberFormat = "dd/mm/yyyy"
    Range("L:L").NumberFormat = "hh:mm"
    Range("N:N").NumberFormat = "hh:mm:ss"
    Range("K1:K1").Value = "Working Day"
    Range("L1:L1").Value = "Start Work"
    Range("N1:N1").Value = "End Work"
End Sub

'***************************************************
'@ (s)
' Function       : CheckUserName
' Returns        : Boolean
' Argument       : userName
' Description    : Check User Name
' Author         : AnhTT
' Create         : 2018-04-19
' Update         :
' Remarks        :
'***************************************************

Public Function CheckUserName(userName As String) As Boolean
    Dim FindRow As Range
    
    If StrIsEmpty(userName) = True Then
        Call DrawFormat
        Cells(2, 11).Value = "Input User Name"
        CheckUserName = False
        Exit Function
    End If
    'Range("B:B") 'Range name
    Set FindRow = Range("B:B").Find(What:=userName, LookIn:=xlValues)
    If FindRow Is Nothing Then
        Call DrawFormat
        Cells(2, 11).Value = "User Name had not been exits"
        CheckUserName = False
        Exit Function
    End If
    
    CheckUserName = True
End Function

'***************************************************
' @ (f)
' Function       : GetRange
' Returns        :
' Argument       : name - username
' Description    : Get Range ActiveSheet by username
' Author         : AnhTT
' Create         : 2018-08-07
' Update         :
' Remarks        :
'***************************************************
Public Function GetRange(name As String) As String
    Dim nameRng As Range
    Dim timeRng As Range
    Dim newRng As Range
    Dim col As Range
    Dim lngResponse As Long
    Set nameRng = Range("B:B") 'Range name
    Set timeRng = Range("D:D") 'Range time

    Dim n As String, addressBeginStr As String, addressEndStr As String, tempBef As String, tempAft As String

    For Each col In nameRng.Rows
        n = nameRng.Cells(col.Row, 1).Value
        If n = vbNullString Then Exit For
        If col.Row > 1 Then
            tempBef = nameRng.Cells(col.Row - 1, 1).Value
            tempAft = nameRng.Cells(col.Row + 1, 1).Value
        End If
        If (StrComp(n, name, vbTextCompare) = 0) And (StrComp(n, tempBef, vbTextCompare) <> 0) Then
            addressBeginStr = vbNullString & timeRng.Cells(col.Row, 1).Address
        End If
        If (StrComp(n, name, vbTextCompare) = 0) And (StrComp(n, tempAft, vbTextCompare) <> 0) Then
            addressEndStr = vbNullString & timeRng.Cells(col.Row, 1).Address
        End If
    Next col

    If Not IsEmpty(addressBeginStr) And Not IsEmpty(addressBeginStr) Then
        GetRange = addressBeginStr & ":" & addressEndStr
    Else
         Exit Function
    End If
End Function
