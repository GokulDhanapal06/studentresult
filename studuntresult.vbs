Option Explicit

Dim reg, dob
Dim hasLow, hasHigh
Dim i, total
Dim tamil, english, maths, science, social
Dim result

'Get Register Number
reg = InputBox("Enter Register Number:", "Student Result Portal")

If reg = "" Then
    MsgBox "Register Number is required!", vbExclamation, "Error"
    WScript.Quit
End If

'Get Date of Birth
dob = InputBox("Enter Date of Birth (DD-MM-YYYY):", "Student Result Portal")

If dob = "" Then
    MsgBox "Date of Birth is required!", vbExclamation, "Error"
    WScript.Quit
End If


'----------------------------------------
' Calculate Total Based on Register Number
'----------------------------------------

hasLow = False
hasHigh = False

For i = 1 To Len(reg)

    If IsNumeric(Mid(reg, i, 1)) Then

        'Digits 0 to 4
        If CInt(Mid(reg, i, 1)) >= 0 And _
           CInt(Mid(reg, i, 1)) <= 4 Then
            hasLow = True
        End If

        'Digits 5 to 9
        If CInt(Mid(reg, i, 1)) >= 5 And _
           CInt(Mid(reg, i, 1)) <= 9 Then
            hasHigh = True
        End If

    End If

Next


'----------------------------------------
' Find Total and Subject Marks
'----------------------------------------

If hasLow And hasHigh Then

    'Mixed digits: 0-4 and 5-9
    total = 479

    tamil = 96
    english = 96
    maths = 96
    science = 96
    social = 95

ElseIf hasLow Then

    'Only digits 0-4
    total = 495

    tamil = 99
    english = 99
    maths = 99
    science = 99
    social = 99

ElseIf hasHigh Then

    'Only digits 5-9
    total = 482

    tamil = 97
    english = 97
    maths = 96
    science = 96
    social = 96

Else

    MsgBox "Please enter a valid Register Number containing digits 0-9.", _
           vbExclamation, "Invalid Register Number"

    WScript.Quit

End If


'----------------------------------------
' Pass / Fail
'----------------------------------------

If total >= 250 Then
    result = "PASS"
Else
    result = "FAIL"
End If


'----------------------------------------
' Display Result
'----------------------------------------

MsgBox _
    "        STUDENT RESULT PORTAL" & vbCrLf & _
    "----------------------------------------" & vbCrLf & _
    "Register Number : " & reg & vbCrLf & _
    "Date of Birth   : " & dob & vbCrLf & _
    "----------------------------------------" & vbCrLf & _
    "Tamil           : " & tamil & " / 100" & vbCrLf & _
    "English         : " & english & " / 100" & vbCrLf & _
    "Mathematics     : " & maths & " / 100" & vbCrLf & _
    "Science         : " & science & " / 100" & vbCrLf & _
    "Social Science  : " & social & " / 100" & vbCrLf & _
    "----------------------------------------" & vbCrLf & _
    "TOTAL           : " & total & " / 500" & vbCrLf & _
    "RESULT          : " & result, _
    vbInformation, _
    "Examination Result"