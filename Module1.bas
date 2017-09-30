Attribute VB_Name = "Module1"
Option Explicit

Declare Function apiGetPrivateProfileString Lib "kernel32" _
            Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
            As String, ByVal lpKeyName As Any, ByVal lpDefault As _
            String, ByVal lpReturnedString As String, ByVal nSize As _
            Long, ByVal lpFileName As String) As Long

Declare Function apiWritePrivateProfileString Lib _
            "kernel32" Alias "WritePrivateProfileStringA" _
            (ByVal lpApplicationName As String, ByVal lpKeyName _
            As Any, ByVal lpString As Any, ByVal lpFileName As _
            String) As Long
' Write INI Profile String
Public Function SetIni(ByVal Section As String, ByVal Keyword As String, ByVal DefVal As String) As String
  Dim ResultString As String * 128
  Dim temp As Integer

  temp = apiWritePrivateProfileString(Section, Keyword, DefVal, App.Path & "\" & App.EXEName & ".ini")
  SetIni = Left$(ResultString, temp)
  
    
End Function

' Retrieve INI Profile String
Public Function GetIni(ByVal Section As String, ByVal Keyword As String, ByVal DefVal As String) As String
  Dim ResultString As String * 256  '-alice-140803
  Dim temp As Integer

  temp = apiGetPrivateProfileString(Section, Keyword, DefVal, ResultString, Len(ResultString), App.Path & "\" & App.EXEName & ".ini")
  GetIni = Left$(ResultString, temp)
    
    
End Function

Public Function GetBinary(ByVal iInput As Integer) As String

  ' Returns the 8-bit binary representation
  ' of an integer iInput where 0 <= iInput <= 255
  
  Dim s As String, i As Integer
  
  If iInput < 0 Or iInput > 255 Then
     GetBinary = ""
     Exit Function
  End If
  
  s = ""
  For i = 1 To 8
     s = CStr(iInput Mod 2) & s
     iInput = iInput \ 2
  Next
  
  GetBinary = s

End Function


Function GetInvBinary(ByVal iInput As Integer) As String

' Returns the 8-bit inverted binary representation
' of an integer iInput where 0 <= iInput <= 255

  Dim s As String
  Dim i As Integer
  Dim j As Integer
  
  If iInput < 0 Or iInput > 255 Then
     GetInvBinary = ""
     Exit Function
  End If
  
  s = ""
  For i = 1 To 8
    j = iInput Mod 2
    If j = 1 Then
      j = 0
    Else:
      j = 1
    End If
    s = CStr(j) & s
    iInput = iInput \ 2
    
  Next
  
  GetInvBinary = s

End Function

Public Function GetHEX(ByVal iInput As Integer) As String

  ' Returns the 8-bit HEX representation
  ' of an integer iInput where 0 <= iInput <= 255
  
  Dim s As String, i As Integer
  Dim ih1 As Integer
  Dim ih2 As Integer
  Dim sh1 As String
  Dim sh2 As String
  Dim h As String
  Dim iBuf As Integer
    
  If iInput < 0 Or iInput > 255 Then
     GetHEX = ""
     Exit Function
  End If
  
  s = ""
  For i = 1 To 8
     s = CStr(iInput Mod 2) & s
     iInput = iInput \ 2
  Next
  
  ih1 = 0
  For i = 1 To 4
    iBuf = CInt(Mid$(s, i, 1))
    If iBuf = 1 Then
      Select Case i
        Case 1:
          ih1 = ih1 + 8
        Case 2:
          ih1 = ih1 + 4
        Case 3:
          ih1 = ih1 + 2
        Case 4:
          ih1 = ih1 + 1
      End Select
    End If
  Next
  If ih1 > 9 Then
    Select Case ih1
      Case 10:
        sh1 = "A"
      Case 11:
        sh1 = "B"
      Case 12:
        sh1 = "C"
      Case 13:
        sh1 = "D"
      Case 14:
        sh1 = "E"
      Case 15:
        sh1 = "F"
    End Select
  Else:
    sh1 = str(ih1)
  End If
  
  ih2 = 0
  For i = 5 To 8
    iBuf = CInt(Mid$(s, i, 1))
    If iBuf = 1 Then
      Select Case i
        Case 5:
          ih2 = ih2 + 8
        Case 6:
          ih2 = ih2 + 4
        Case 7:
          ih2 = ih2 + 2
        Case 8:
          ih2 = ih2 + 1
      End Select
    End If
  Next
  If ih2 > 9 Then
    Select Case ih2
      Case 10:
        sh2 = "A"
      Case 11:
        sh2 = "B"
      Case 12:
        sh2 = "C"
      Case 13:
        sh2 = "D"
      Case 14:
        sh2 = "E"
      Case 15:
        sh2 = "F"
    End Select
  Else:
    sh2 = str(ih2)
  End If
  
  h = Trim(sh1) & Trim(sh2)
  
  GetHEX = Trim(h)
End Function




Public Function MediumDate(str)
  Dim aDay
  Dim aMonth
  Dim aYear
  
  aDay = Day(str)
  aMonth = MonthName(Month(str), True)
  aYear = Year(str)
  
  MediumDate = aDay & "-" & aMonth & "-" & aYear
  
End Function
