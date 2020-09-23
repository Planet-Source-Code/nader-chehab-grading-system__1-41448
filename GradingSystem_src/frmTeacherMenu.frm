VERSION 5.00
Begin VB.Form frmTeacherMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Teacher's Menu"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   ControlBox      =   0   'False
   Icon            =   "frmTeacherMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin GradingSystem.HoverCommand cmdExit 
      Height          =   615
      Left            =   5400
      TabIndex        =   5
      Top             =   5400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "&Exit"
      Style           =   3
      Picture         =   "frmTeacherMenu.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdLogout 
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Caption         =   "&Logout"
      Style           =   3
      Picture         =   "frmTeacherMenu.frx":09A3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdEvaluationChart 
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2760
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      Caption         =   "Courses Properties"
      Style           =   16
      Picture         =   "frmTeacherMenu.frx":0ACA
      HovPicture      =   "frmTeacherMenu.frx":0BFE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdStudentGrade 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   3360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      Caption         =   "Student's Grades and Absences"
      Style           =   16
      Picture         =   "frmTeacherMenu.frx":0D32
      HovPicture      =   "frmTeacherMenu.frx":0E66
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdPersonnalInfo 
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   3960
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   873
      Caption         =   "Personnal Information"
      Style           =   16
      Picture         =   "frmTeacherMenu.frx":0F9A
      HovPicture      =   "frmTeacherMenu.frx":10CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4680
      TabIndex        =   0
      Top             =   5640
      Width           =   75
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Grading System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label lblWelcome 
      Alignment       =   2  'Center
      Caption         =   "Welcome, "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   7
      Top             =   1920
      Width           =   5895
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "T e a c h e r ' s  M e n u"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   3615
   End
End
Attribute VB_Name = "frmTeacherMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Fill the combo-box in the CourseSelection form and show it
Private Sub cmdEvaluationChart_Click()
strSQL = "SELECT Courses.CourseID, CourseTitle FROM Courses, TeacherCourse " & _
         "WHERE Courses.CourseID = TeacherCourse.CourseID " & _
         "AND TeacherCourse.TeacherID = '" & getTeacherID & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If objDBRecordset.EOF Then
MsgBox "You do not have any course to teach! Please contact the Administrator.", vbInformation, "Grading System - Information"
Exit Sub
End If
frmCourseSelection.Show
objDBRecordset.MoveFirst
While Not objDBRecordset.EOF
frmCourseSelection.cboCourses.AddItem objDBRecordset("CourseTitle") & " - " & objDBRecordset("CourseID")
objDBRecordset.MoveNext
Wend
frmCourseSelection.cboCourses.text = frmCourseSelection.cboCourses.list(0)
Me.Enabled = False
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLogout_Click()
strUsername = ""
frmLogin.Show
Unload Me
End Sub

'Show personal information
Private Sub cmdPersonnalInfo_Click()
With frmTeacherModifyInfo
.Show
.lblTeacherID = getTeacherID
strSQL = "SELECT FirstName, LastName, Username FROM Teachers " & _
         "WHERE TeacherID = " & "'" & .lblTeacherID & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
.txtFirstName.text = objDBRecordset("FirstName")
.txtLastName.text = objDBRecordset("LastName")
.txtUserName.text = objDBRecordset("Username")
Me.Enabled = False
End With
End Sub

'Fill the combo-box in the CourseSelection2 form and show it
Private Sub cmdStudentGrade_Click()
strSQL = "SELECT Courses.CourseID, CourseTitle FROM Courses, TeacherCourse " & _
         "WHERE Courses.CourseID = TeacherCourse.CourseID " & _
         "AND TeacherCourse.TeacherID = '" & getTeacherID & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If objDBRecordset.EOF Then
MsgBox "You do not have any course to teach! Please contact the Administrator.", vbInformation, "Grading System - Information"
Exit Sub
End If
frmCourseSelection2.Show
objDBRecordset.MoveFirst
While Not objDBRecordset.EOF
frmCourseSelection2.cboCourses.AddItem objDBRecordset("CourseTitle") & " - " & objDBRecordset("CourseID")
objDBRecordset.MoveNext
Wend
frmCourseSelection2.cboCourses.text = frmCourseSelection2.cboCourses.list(0)
Me.Enabled = False
End Sub

Private Sub Form_Load()
lblWelcome.Caption = lblWelcome.Caption & strUsername & "!"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
