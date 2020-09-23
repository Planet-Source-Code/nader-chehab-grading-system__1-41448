Attribute VB_Name = "modMain"
Option Explicit

'API declarations to allow form dragging
Public Declare Function ReleaseCapture Lib "User32" () As Long
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_SYSCOMMAND = &H112

''API declarations to allow "Always on top"
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'ADO Objects
Public objDBConnection As New ADODB.Connection
Public objDBRecordset As New ADODB.Recordset

'Excel Objects
Public xlsApp As Excel.Application
Public xlsWorkbook As Excel.Workbook
Public xlsWorksheet As Excel.Worksheet

'If project is Always on Top or not
Public AlwaysOnTop As Boolean

'General Counters
Public i, j As Integer

'Space for SQL statements
Public strSQL As String

'Username of the currently logged user

'Username of the currently logged user
Public strUsername As String

'Currently Selected Course (its CourseID)
Public strSelectedCourse As String

Sub Main()
AlwaysOnTop = False
    If Len(Dir(App.Path & "\GradingSystem.mdb")) > 0 Then
        Call ConnectToDatabase
        frmLogin.Show
    Else
        frmFirst.Show
    End If
    
End Sub

'Database connection
Public Sub ConnectToDatabase()
With objDBConnection
    .ConnectionString = _
    "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & "\GradingSystem.mdb;" & _
    "Persist Security Info=False;" & _
    "Jet OLEDB:Database Password = " & Transform(ReadFromFile())
    .Open
End With
End Sub

'XOR Encryption function
Public Function Transform(text As String) As String
Dim i, a As Byte
For i = 1 To Len(text)
    a = i Mod Len(text): If a = 0 Then a = Len(text)
    Transform = Transform & Chr(Asc(Mid(text, i, 1)) Xor Asc(Mid("9354123761874238", a, 1)))
Next i
End Function

'Writes a string to file
Public Sub WriteToFile(ByVal strData As String)
Open "GradingSystem.dat" For Output As #1
Write #1, strData
Close #1
End Sub

'Reads a string from file
Public Function ReadFromFile() As String
Open "GradingSystem.dat" For Input As #1
Input #1, ReadFromFile
Close #1
End Function
'The Found() function finds a 'value' in a the database.
'if it is found, it returns True else it returns False.
Public Function found(Value As String, field As String, table As String) As Boolean
Dim objDBRecordset As New ADODB.Recordset
strSQL = "SELECT " & field & " FROM " & table
Set objDBRecordset = objDBConnection.Execute(strSQL)

Do While Not objDBRecordset.EOF
    If objDBRecordset(field) = Value Then
        found = True
        Exit Function
    Else
        found = False
    End If
    objDBRecordset.MoveNext
Loop
End Function

'Update Teacher's List
Public Sub UpdateTeachersList()
strSQL = "SELECT TeacherID, FirstName, LastName FROM Teachers"
Set objDBRecordset = objDBConnection.Execute(strSQL)
'do not put the Administrator in the teacher's list.
objDBRecordset.MoveFirst
objDBRecordset.MoveNext
'add all teachers
With frmTeachersList.grdTeachers
    .Clear
    .Cols = 4
    .Rows = 2
    .ColWidth(0) = 500
    .ColWidth(1) = 1200
    .ColWidth(2) = 1660
    .ColWidth(3) = 1660
    .TextMatrix(0, 1) = "Teacher ID"
    .TextMatrix(0, 2) = "First Name"
    .TextMatrix(0, 3) = "Last Name"
    SendKeys "{RIGHT}"
    While Not objDBRecordset.EOF
    .TextMatrix(.Rows - 1, 1) = objDBRecordset("TeacherID")
    .TextMatrix(.Rows - 1, 2) = objDBRecordset("FirstName")
    .TextMatrix(.Rows - 1, 3) = objDBRecordset("LastName")
    objDBRecordset.MoveNext
    .Rows = .Rows + 1
    Wend
.Rows = .Rows - 1
End With
End Sub

'Delete the selected Teacher
Public Sub DeleteTeacher()
With frmTeachersList.grdTeachers
strSQL = "DELETE FROM TeacherCourse WHERE TeacherID = '" & .TextMatrix(.RowSel, 1) & "'"
objDBConnection.Execute strSQL
strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & .TextMatrix(.RowSel, 1) & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)

'delete teacher's tables
Dim Username As String
Username = objDBRecordset("Username")
strSQL = "SELECT CourseID FROM Courses"
Set objDBRecordset = objDBConnection.Execute(strSQL)
On Error Resume Next
While Not objDBRecordset.EOF
strSQL = "DROP TABLE [" & Username & objDBRecordset("CourseID") & "]"
objDBConnection.Execute (strSQL)
If Err.Number <> -2147217865 And Err.Number <> 0 Then MsgBox "error!"
objDBRecordset.MoveNext
Wend
strSQL = "DELETE FROM Teachers WHERE TeacherID = '" & .TextMatrix(.RowSel, 1) & "'"
objDBConnection.Execute strSQL
End With
Call UpdateTeachersList
End Sub

'Returns the TeacherID of the currently logged in teacher
Public Function getTeacherID() As String
strSQL = "SELECT TeacherID FROM Teachers WHERE Username = '" & strUsername & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
getTeacherID = objDBRecordset("TeacherID")
End Function

'returns Title of a course
Public Function getTitle(courseID As String) As String
strSQL = "SELECT CourseTitle FROM Courses WHERE CourseID = '" & courseID & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
getTitle = objDBRecordset("CourseTitle")
End Function

'Update Course's List
Public Sub UpdateCoursesList()
strSQL = "SELECT CourseID, CourseTitle, NumberOfHours FROM Courses"
Set objDBRecordset = objDBConnection.Execute(strSQL)
'add all Courses
With frmCoursesList.grdCourses
    .Clear
    .Cols = 4
    .Rows = 2
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 2100
    .ColWidth(3) = 1000
    .TextMatrix(0, 1) = "Course ID"
    .TextMatrix(0, 2) = "Course Title"
    .TextMatrix(0, 3) = "Number Of Hours"
    SendKeys "{RIGHT}"
    While Not objDBRecordset.EOF
        .TextMatrix(.Rows - 1, 1) = objDBRecordset("CourseID")
        .TextMatrix(.Rows - 1, 2) = objDBRecordset("CourseTitle")
        .TextMatrix(.Rows - 1, 3) = objDBRecordset("NumberOfHours")
        objDBRecordset.MoveNext
        .Rows = .Rows + 1
    Wend
.Rows = .Rows - 1
End With
End Sub

'Delete the selected Course
Public Sub DeleteCourse()
With frmCoursesList.grdCourses
    strSQL = "DELETE FROM TeacherCourse WHERE CourseID = '" & .TextMatrix(.RowSel, 1) & "'"
    objDBConnection.Execute strSQL
    strSQL = "SELECT Username FROM Teachers"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    On Error Resume Next
    While Not objDBRecordset.EOF
        strSQL = "DROP TABLE [" & objDBRecordset("Username") & .TextMatrix(.RowSel, 1) & "]"
        objDBConnection.Execute (strSQL)
        If Err.Number <> -2147217865 And Err.Number <> 0 Then MsgBox "error!"
        objDBRecordset.MoveNext
    Wend
    strSQL = "DELETE FROM Courses WHERE CourseID = '" & .TextMatrix(.RowSel, 1) & "'"
    objDBConnection.Execute strSQL
End With
Call UpdateCoursesList
End Sub

'Update Student's List
Public Sub UpdateStudentsList()
strSQL = "SELECT StudentID, FirstName, LastName FROM Students"
Set objDBRecordset = objDBConnection.Execute(strSQL)
'add all Students
With frmStudentsList.grdStudents
    .Clear
    .Cols = 4
    .Rows = 2
    .ColWidth(0) = 500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .TextMatrix(0, 1) = "Student ID"
    .TextMatrix(0, 2) = "Frist Name"
    .TextMatrix(0, 3) = "Last Name"
    SendKeys "{RIGHT}"
    While Not objDBRecordset.EOF
        .TextMatrix(.Rows - 1, 1) = objDBRecordset("StudentID")
        .TextMatrix(.Rows - 1, 2) = objDBRecordset("FirstName")
        .TextMatrix(.Rows - 1, 3) = objDBRecordset("LastName")
        objDBRecordset.MoveNext
        .Rows = .Rows + 1
    Wend
.Rows = .Rows - 1
End With
End Sub

'Delete the selected Student
Public Sub DeleteStudent()
With frmStudentsList.grdStudents
strSQL = "SELECT Username, CourseID FROM TeacherCourse, Teachers WHERE Teachers.TeacherID = TeacherCourse.TeacherID"
Set objDBRecordset = objDBConnection.Execute(strSQL)
While Not objDBRecordset.EOF
If found(.TextMatrix(.RowSel, 1), "StudentID", "[" & objDBRecordset("Username") & objDBRecordset("CourseID") & "]") Then
strSQL = "DELETE FROM [" & objDBRecordset("Username") & objDBRecordset("CourseID") & "] WHERE StudentID = '" & .TextMatrix(.RowSel, 1) & "'"
objDBConnection.Execute strSQL
End If
objDBRecordset.MoveNext
Wend
strSQL = "DELETE FROM Students WHERE StudentID = '" & .TextMatrix(.RowSel, 1) & "'"
objDBConnection.Execute strSQL
End With
Call UpdateStudentsList
End Sub
