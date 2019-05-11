Option Explicit


Function DateFromString(str As String, format As String) As Date
    Dim elements() As String
    Dim dte As Date
    Dim vDay As Variant
    Dim vMonth As Variant
    Dim vYear As Variant
    
    Dim nDay As Long
    Dim nMonth As Long
    Dim nYear As Long
    
    'if isDate(
    
    elements = Split(str, "/")
    If UBound(elements) = 2 Then
        vDay = elements(0)
        vMonth = elements(1)
        vYear = elements(2)
        
        If IsNumeric(vYear) And IsNumeric(vMonth) And IsNumeric(vDay) Then
          nYear = CInt(vYear)
          nMonth = CInt(vMonth)
          nDay = CInt(vDay)
          If IsValidYear(nYear) And IsValidMonth(nMonth) And IsValidDay(nDay, nMonth, nYear) Then
            dte = DateSerial(nYear, nMonth, nDay)
          End If
        End If
    Else
        dte = 0
    End If
    
    DateFromString = dte
End Function

Private Function IsValidYear(year As Long)
  If year > 1970 And year < 3000 Then
    IsValidYear = True
  Else
    IsValidYear = False
  End If
End Function

Private Function IsValidMonth(month As Long)
    If month >= 1 And month <= 12 Then
        IsValidMonth = True
    Else
        IsValidMonth = False
    End If
End Function

Private Function IsValidDay(day As Long, month As Long, year As Long)
    Dim bValid As Boolean
    Dim bLeapYear As Boolean

    bValid = False
    Select Case month
        Case 1, 3, 5, 7, 8, 10, 12
            If day >= 1 And day <= 31 Then
                bValid = True
            End If
        Case 4, 6, 9, 11
            If day >= 1 And day <= 30 Then
                bValid = True
            End If
        Case 2
        
            'if year is not divisible by 4 then not leap year
            If year Mod 4 <> 0 Then
                bLeapYear = False
            'if year is not divisible by 100 this is a leap year
            ElseIf (year Mod 100 <> 0) Then
                bLeapYear = True
            'if year is not divisible by 400 then is not a leap year
            ElseIf (year Mod 400 <> 0) Then
                bLeapYear = False
            'else is a leap year
            Else
                bLeapYear = True
            End If
                    
            If bLeapYear = False And day >= 1 And day <= 28 Then
                bValid = True
            End If
            
            If bLeapYear And day >= 1 And day <= 29 Then
                bValid = True
            End If

        Case Else
            'do nothing
    End Select

    IsValidDay = bValid
End Function

