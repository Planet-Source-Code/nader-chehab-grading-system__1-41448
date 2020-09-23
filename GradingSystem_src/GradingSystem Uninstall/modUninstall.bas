Attribute VB_Name = "modUninstall"
Sub Main()

    On Error GoTo UninstallError
    
    If Len(Dir(App.Path & "\GradingSystem.mdb")) = 0 Then Err.Raise 1001
    If Len(Dir(App.Path & "\GradingSystem.exe")) = 0 Then Err.Raise 1002
    
    Dim answer As Integer
    answer = MsgBox("Are you sure you want to Uninstall the Grading System ?", vbYesNoCancel Or vbQuestion, "Grading System - Uninstall")
    If answer = vbYes Then
        Kill App.Path & "\GradingSystem.mdb"
        Kill App.Path & "\GradingSystem.exe"
        Call DeleteSectionFromINI("GradingSystem")
        MsgBox "Deleted GradingSystem.mdb" & vbCrLf & _
               "Deleted GradingSystem.exe" & vbCrLf & _
               "Removed GradingSystem from Registry." & vbCrLf & vbCrLf & _
               "Grading System Unistalled successfully.", vbInformation, "Grading System"

    End If

Exit Sub

UninstallError:
Select Case Err.Number
    Case 1001
    MsgBox "The Database (GradingSystem.mdb) was not found."
    Case 1002
    MsgBox "The Executable (GradingSystem.exe) is missing."
    Case Else
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "Grading System"
End Select

End Sub

