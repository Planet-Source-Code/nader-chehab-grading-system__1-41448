VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRegistrations 
   BorderStyle     =   0  'None
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9975
   Icon            =   "frmRegistrations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register a Student"
      Height          =   735
      Left            =   8040
      Picture         =   "frmRegistrations.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Main Menu"
      Height          =   735
      Left            =   240
      Picture         =   "frmRegistrations.frx":0D28
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid grdStudents 
      Height          =   3615
      Left            =   240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   6376
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdCourses 
      Height          =   3615
      Left            =   3720
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6376
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Registered Courses:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3720
      TabIndex        =   5
      Top             =   600
      Width           =   1980
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Select a Student:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1785
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu cmdUnregisterCourse 
         Caption         =   "&Unregister this Course"
      End
   End
   Begin VB.Menu mnuList2 
      Caption         =   "List2"
      Visible         =   0   'False
      Begin VB.Menu cmdUpdate 
         Caption         =   "&Register"
      End
      Begin VB.Menu cmdUnregisterStudent 
         Caption         =   "&Unregister from all Courses"
      End
   End
End
Attribute VB_Name = "frmRegistrations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmAdminMenu.Show
Unload Me
End Sub

Private Sub cmdRegister_Click()
frmRegister.Show
Unload Me
frmRegister.cboStudentID.SetFocus
SendKeys "{Home}+{End}"
End Sub

'Unregister the selected course of the current student
Private Sub cmdUnregisterCourse_Click()
Dim ans As Integer
With grdStudents
ans = MsgBox("Are you sure you want to unregister '" & .TextMatrix(.RowSel, 2) & " " & .TextMatrix(.RowSel, 3) & "' from '" & grdCourses.TextMatrix(grdCourses.RowSel, 3) & "' ?", vbYesNo Or vbQuestion, "Grading System - Question")
End With
If ans = vbNo Then Exit Sub
With grdCourses
strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & .TextMatrix(.RowSel, 1) & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
strSQL = "DELETE FROM [" & objDBRecordset("Username") & .TextMatrix(.RowSel, 2) & "] WHERE StudentID = '" & grdStudents.TextMatrix(grdStudents.RowSel, 1) & "'"
objDBConnection.Execute (strSQL)
Call UpdateRegCourses
Call getCourses
End With
End Sub

'Unregister the selected student from all his courses
Private Sub cmdUnregisterStudent_Click()
Dim ans As Integer
ans = MsgBox("Are you sure you want to unregister '" & grdStudents.TextMatrix(grdStudents.RowSel, 2) & " " & grdStudents.TextMatrix(grdStudents.RowSel, 3) & "' from all courses ?", vbYesNo Or vbQuestion, "Grading System - Question")
If ans = vbNo Then Exit Sub
With grdCourses
'if student registered, unregister him
If .TextMatrix(1, 1) <> "" Then
For i = 1 To grdCourses.Rows - 1
strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & .TextMatrix(i, 1) & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
strSQL = "DELETE FROM [" & objDBRecordset("Username") & .TextMatrix(i, 2) & "] WHERE StudentID = '" & grdStudents.TextMatrix(grdStudents.RowSel, 1) & "'"
objDBConnection.Execute (strSQL)
Next i
Call UpdateRegCourses
Call getCourses
End If
End With
End Sub

Private Sub cmdUpdate_Click()
frmRegister.Show
frmRegister.cboStudentID.text = grdStudents.TextMatrix(grdStudents.RowSel, 1)
frmRegister.cboCourseID.SetFocus
SendKeys "{Home}+{End}"
Unload Me
End Sub

Private Sub Form_Load()
Me.Height = 5920
'Update Students List
strSQL = "SELECT StudentID, FirstName, LastName FROM Students"
Set objDBRecordset = objDBConnection.Execute(strSQL)
'add all Students
With frmRegistrations.grdStudents
    .Clear
    .Cols = 4
    .Rows = 2
    .ColWidth(0) = 0
    .ColWidth(1) = 1000
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
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
Call UpdateRegCourses
End Sub

Private Sub grdStudents_Click()
Call getCourses
End Sub

'Update Courses List
Private Sub UpdateRegCourses()
With frmRegistrations.grdCourses
    .Clear
    .Cols = 5
    .Rows = 2
    .ColWidth(0) = 0
    .ColWidth(1) = 1000
    .ColWidth(2) = 1000
    .ColWidth(3) = 2400
    .ColWidth(4) = 1400
    .TextMatrix(0, 1) = "Teacher ID"
    .TextMatrix(0, 2) = "Course ID"
    .TextMatrix(0, 3) = "Course Title"
    .TextMatrix(0, 4) = "Number Of Hours"
    SendKeys "{RIGHT}"
End With
End Sub

'show the courses of the selected student
Private Sub getCourses()
Dim coursesFound As Boolean
strSQL = "SELECT Username, Courses.CourseID, Teachers.TeacherID, CourseTitle, NumberOfHours FROM Teachers, TeacherCourse, Courses " & _
         "WHERE Teachers.TeacherID = TeacherCourse.TeacherID " & _
         "AND Courses.CourseID = TeacherCourse.CourseID"
Set objDBRecordset = objDBConnection.Execute(strSQL)
Call UpdateRegCourses
While Not objDBRecordset.EOF
If found(grdStudents.TextMatrix(grdStudents.RowSel, 1), "StudentID", "[" & objDBRecordset("Username") & objDBRecordset("CourseID") & "]") Then
coursesFound = True
grdCourses.TextMatrix(grdCourses.Rows - 1, 1) = objDBRecordset("TeacherID")
grdCourses.TextMatrix(grdCourses.Rows - 1, 2) = objDBRecordset("CourseID")
grdCourses.TextMatrix(grdCourses.Rows - 1, 3) = objDBRecordset("CourseTitle")
grdCourses.TextMatrix(grdCourses.Rows - 1, 4) = objDBRecordset("NumberOfHours")
grdCourses.Rows = grdCourses.Rows + 1
End If
objDBRecordset.MoveNext
Wend
If coursesFound Then grdCourses.Rows = grdCourses.Rows - 1
End Sub

Private Sub grdStudents_RowColChange()
Call getCourses
End Sub

Private Sub grdCourses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If grdCourses.TextMatrix(1, 1) <> "" Then
If Button = 2 And grdCourses.Rows > 1 Then PopupMenu mnuList
End If
End Sub

Private Sub grdStudents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If grdStudents.Rows > 1 Then
If Button = 2 And grdCourses.Rows > 1 Then PopupMenu mnuList2
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

