VERSION 5.00
Begin VB.Form frmTeacherUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teacher Update"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   ControlBox      =   0   'False
   Icon            =   "frmTeacherUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRemoveTeacher 
      Cancel          =   -1  'True
      Caption         =   "&Remove Teacher"
      Height          =   855
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teacher Information"
      Height          =   2295
      Left            =   600
      TabIndex        =   10
      Top             =   480
      Width           =   4095
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify Password"
         Height          =   285
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   11
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblPassword 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*shadowed*"
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   765
         Width           =   765
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Teacher ID:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   405
         Width           =   855
      End
      Begin VB.Label lblTeacherID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1845
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1485
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdAddCourse 
      Caption         =   "&Add Course"
      Height          =   285
      Left            =   4800
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   615
      Left            =   1755
      MaskColor       =   &H00DEEBEF&
      Picture         =   "frmTeacherUpdate.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   2535
   End
   Begin VB.CommandButton cmdRemoveCourse 
      Caption         =   "&Remove Course"
      Height          =   525
      Left            =   4800
      MaskColor       =   &H00DEEBEF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4260
      Width           =   1095
   End
   Begin VB.ComboBox cboAllCourses 
      Height          =   315
      Left            =   555
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4905
      Width           =   4095
   End
   Begin VB.ListBox lstCoursesTaught 
      Height          =   1425
      Left            =   555
      TabIndex        =   3
      Top             =   3360
      Width           =   4095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Courses taught:"
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   9
      Top             =   3120
      Width           =   1110
   End
End
Attribute VB_Name = "frmTeacherUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdModify_Click()
frmTeacherChangePassword.Show
Me.Enabled = False
End Sub

Private Sub cmdRemoveTeacher_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to remove teacher '" & txtFirstName & "' ?", vbYesNoCancel, "GradingSystem - Removing")
If answer = vbYes Then
Call DeleteTeacher
frmTeachersList.Enabled = True
Unload Me
End If
End Sub

Private Sub Form_Load()
With frmTeachersList.grdTeachers
strSQL = "SELECT TeacherID, FirstName, LastName, Username, Password FROM Teachers " & _
         "WHERE TeacherID = " & "'" & .TextMatrix(.RowSel, 1) & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
lblTeacherID.Caption = objDBRecordset("TeacherID")
txtFirstName.text = objDBRecordset("FirstName")
txtLastName.text = objDBRecordset("LastName")
txtUserName.text = objDBRecordset("Username")
End With

lstCoursesTaught.Clear
cboAllCourses.Clear
'Fill CoursesTaught list
strSQL = "SELECT Courses.CourseID, CourseTitle FROM TeacherCourse, Courses " & _
         "WHERE Courses.CourseID = TeacherCourse.CourseID " & _
         "AND TeacherID = '" & lblTeacherID & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Not objDBRecordset.EOF Then
objDBRecordset.MoveFirst
While Not objDBRecordset.EOF
lstCoursesTaught.AddItem objDBRecordset("CourseTitle") & " - " & objDBRecordset("CourseID")
objDBRecordset.MoveNext
Wend
End If
'Fill AllCourses combo
strSQL = "SELECT CourseID, CourseTitle FROM Courses"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Not objDBRecordset.EOF Then
objDBRecordset.MoveFirst
While Not objDBRecordset.EOF
cboAllCourses.AddItem objDBRecordset("CourseTitle") & " - " & objDBRecordset("CourseID")
objDBRecordset.MoveNext
Wend
cboAllCourses.text = cboAllCourses.list(0)
End If
End Sub

Private Sub cmdAddCourse_Click()
If Not ExistsIn(cboAllCourses.text, lstCoursesTaught) Then

'if the teacher-course table does not exist, create it
strSQL = "SELECT * FROM [" & txtUserName & Right$(cboAllCourses, 10) & "]"
On Error Resume Next
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Err.Number = -2147217865 Or Err.Number = -2147467259 Then

strSQL = "CREATE TABLE [" & txtUserName & Right$(cboAllCourses, 10) & "]"
objDBConnection.Execute (strSQL)

'Add the StudentID column
strSQL = "ALTER TABLE [" & txtUserName & Right$(cboAllCourses, 10) & "] " & _
         "ADD StudentID VARCHAR(7) PRIMARY KEY"
objDBConnection.Execute (strSQL)

'Add the Total Grade column
strSQL = "ALTER TABLE [" & txtUserName & Right$(cboAllCourses, 10) & "] " & _
         "ADD Total VARCHAR(5)"
objDBConnection.Execute (strSQL)

'Add the Absences column
strSQL = "ALTER TABLE [" & txtUserName & Right$(cboAllCourses, 10) & "] " & _
         "ADD Absences VARCHAR(2)"
objDBConnection.Execute (strSQL)

'Add the starting date and ending date columns
strSQL = "ALTER TABLE [" & txtUserName & Right$(cboAllCourses, 10) & "] " & _
         "ADD StartingDate DATE, EndingDate DATE"
objDBConnection.Execute (strSQL)

strSQL = "INSERT INTO TeacherCourse(TeacherID, CourseID) VALUES('" & lblTeacherID & "', '" & Right$(cboAllCourses.text, 10) & "')"
Set objDBRecordset = objDBConnection.Execute(strSQL)
lstCoursesTaught.AddItem cboAllCourses.text
End If

Else
MsgBox "Course already added!", vbInformation, "Grading System - Error"
End If
End Sub

'If teacher's info is changed, update it in database
Private Sub cmdOK_Click()
strSQL = "SELECT TeacherID, FirstName, LastName, Username FROM Teachers"
Set objDBRecordset = objDBConnection.Execute(strSQL)
While Not objDBRecordset.EOF And objDBRecordset("TeacherID") <> lblTeacherID
objDBRecordset.MoveNext
Wend
If txtFirstName <> objDBRecordset("FirstName") Or txtLastName <> objDBRecordset("LastName") Then
    strSQL = "UPDATE Teachers SET FirstName = '" & txtFirstName & "' WHERE TeacherID = '" & lblTeacherID & "'"
    objDBConnection.Execute (strSQL)
    strSQL = "UPDATE Teachers SET LastName = '" & txtLastName & "' WHERE TeacherID = '" & lblTeacherID & "'"
    objDBConnection.Execute (strSQL)
End If
frmTeachersList.Enabled = True
Call UpdateTeachersList
Unload Me
frmTeachersList.SetFocus
End Sub

Private Sub cmdRemoveCourse_Click()
If lstCoursesTaught.ListCount > 0 Then
    If lstCoursesTaught.SelCount = 1 Then
    
    strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & lblTeacherID & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    
    strSQL = "DELETE FROM TeacherCourse WHERE CourseID = '" & Right$(lstCoursesTaught.list(lstCoursesTaught.ListIndex), 10) & "' AND TeacherID = '" & lblTeacherID & "'"
    objDBConnection.Execute (strSQL)
    
    On Error Resume Next
    strSQL = "DROP TABLE [" & objDBRecordset("Username") & Right$(lstCoursesTaught.list(lstCoursesTaught.ListIndex), 10) & "]"
    objDBConnection.Execute (strSQL)
    
    lstCoursesTaught.RemoveItem lstCoursesTaught.ListIndex
    
    Else
    MsgBox "You must select a course to remove!", vbCritical, "Grading System - Error"
    End If
Else
    MsgBox "This teacher does not have any course yet!", vbCritical, "Grading System - Error"
    End If
End Sub

'Finds out if a record already exists in a ListBox.
Private Function ExistsIn(course As String, list As ListBox) As Boolean
For i = 0 To list.ListCount
If course = list.list(i) Then
ExistsIn = True
Exit Function
Else
ExistsIn = False
End If
Next i
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
