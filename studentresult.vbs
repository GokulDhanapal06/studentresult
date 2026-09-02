' Student Result Portal - VBScript
' Technology: VBScript + WSH

Option Explicit

Dim regNo, dob, name
Dim marks(5), subjects(5)
Dim i, total, average, result, grade, status
Dim isValid, marksInput

subjects(0) = "Tamil"
subjects(1) = "English"
subjects(2) = "Mathematics"
subjects(3) = "Science"
subjects(4) = "Social Science"

' 1. Register Number Input & Validation
Do
    regNo = InputBox("Enter Register Number (Ex: 123456):" & vbCrLf & vbCrLf & "Must be 6 digits", "Student Result Portal - Step 1/3")
    If regNo = "" Then WScript.Quit ' User cancelled
    
    isValid = True
    If Len(regNo) <> 6 Then
        isValid = False
    Else
        For i = 1 To Len(regNo)
            If Not IsNumeric(Mid(regNo, i, 1)) Then
                isValid = False
                Exit For
            End If
        Next
    End If

    If Not isValid Then
        MsgBox "Invalid Register Number!" & vbCrLf & "Please enter exactly 6 digits.", vbExclamation, "Validation Error"
    End If
Loop While Not isValid

' 2. Date of Birth Input
Do
    dob = InputBox("Enter Date of Birth (DD-MM-YYYY):" & vbCrLf & vbCrLf & "Ex: 15-08-2008", "Student Result Portal - Step 2/3")
    If dob = "" Then WScript.Quit
    
    If IsDate(dob) Then
        Exit Do
    Else
        MsgBox "Invalid Date of Birth!" & vbCrLf & "Please enter in DD-MM-YYYY format.", vbExclamation, "Validation Error"
    End If
Loop While True

' 3. Student Name (for display)
name = InputBox("Enter Student Name:", "Student Result Portal - Step 3/3")
If name = "" Then name = "Student"

' 4. Subject-wise Marks Input
total = 0
For i = 0 To 4
    Do
        marksInput = InputBox("Enter marks for " & subjects(i) & " (0-100):", "Marks Entry - " & subjects(i))
        If marksInput = "" Then WScript.Quit

        If IsNumeric(marksInput) Then
            marks(i) = CInt(marksInput)
            If marks(i) >= 0 And marks(i) <= 100 Then
                Exit Do
            End If
        End If
        MsgBox "Invalid Marks! Enter a number between 0 and 100.", vbExclamation, "Validation Error"
    Loop While True
    total = total + marks(i)
Next

average = total / 5

' 5. Pass/Fail & Grade Calculation
status = "PASS"
For i = 0 To 4
    If marks(i) < 35 Then
        status = "FAIL"
        Exit For
    End If
Next

If status = "PASS" Then
    If average >= 90 Then grade = "A+"
    If average >= 75 And average < 90 Then grade = "A"
    If average >= 60 And average < 75 Then grade = "B"
    If average >= 50 And average < 60 Then grade = "C"
    If average >= 35 And average < 50 Then grade = "D"
Else
    grade = "E"
    result = "Fail - Need to reappear"
End If

If status = "PASS" Then result = "Pass"

' 6. Final Result Display
Dim finalMsg
finalMsg = "---------------------------------" & vbCrLf
finalMsg = finalMsg & "   STUDENT RESULT PORTAL" & vbCrLf
finalMsg = finalMsg & "---------------------------------" & vbCrLf & vbCrLf
finalMsg = finalMsg & "Name           : " & name & vbCrLf
finalMsg = finalMsg & "Register No    : " & regNo & vbCrLf
finalMsg = finalMsg & "Date of Birth  : " & dob & vbCrLf & vbCrLf
finalMsg = finalMsg & "---------------------------------" & vbCrLf
For i = 0 To 4
    finalMsg = finalMsg & subjects(i) & " : " & marks(i) & "/100" & vbCrLf
Next
finalMsg = finalMsg & "---------------------------------" & vbCrLf & vbCrLf
finalMsg = finalMsg & "Total Marks    : " & total & "/500" & vbCrLf
finalMsg = finalMsg & "Average        : " & Round(average, 2) & "%" & vbCrLf
finalMsg = finalMsg & "Grade          : " & grade & vbCrLf
finalMsg = finalMsg & "Result         : " & status & vbCrLf & vbCrLf

If status = "PASS" Then
    finalMsg = finalMsg & "Congratulations! You have Passed."
Else
    finalMsg = finalMsg & "Sorry! You have Failed."
End If

MsgBox finalMsg, vbInformation, "Examination Result - " & regNo
    vbInformation, _
    "Examination Result"
