VERSION 5.00
Begin VB.Form frmCourseUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Update"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   ControlBox      =   0   'False
   Icon            =   "frmCourseUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6510
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumberOfHours 
      Height          =   285
      Left            =   2400
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox txtCourseTitle 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   360
      Picture         =   "frmCourseUpdate.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2340
      Width           =   1695
   End
   Begin VB.CommandButton cmdUpdateCourse 
      Caption         =   "&Update"
      Height          =   615
      Left            =   4440
      Picture         =   "frmCourseUpdate.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2340
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Course ID:"
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   7
      Top             =   765
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number Of Hours:"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   6
      Top             =   1485
      Width           =   1275
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Course Title:"
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   5
      Top             =   1125
      Width           =   885
   End
   Begin VB.Label lblCourseID 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   2400
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
End
Attribute VB_Name = "frmCourseUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmCoursesList.Enabled = True
Unload Me
End Sub

Private Sub cmdUpdateCourse_Click()
On Error GoTo CourseUpdateError
    If txtCourseTitle = "" Then Err.Raise 1001
    If txtNumberOfHours = "" Then Err.Raise 1002

    strSQL = "UPDATE Courses SET CourseTitle = '" & txtCourseTitle & "', " & _
             "NumberOfHours = '" & txtNumberOfHours & "' " & _
             "WHERE CourseID = '" & lblCourseID & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    frmCoursesList.Enabled = True
    Unload Me
    Call UpdateCoursesList
    frmCoursesList.SetFocus
Exit Sub

CourseUpdateError:
Select Case Err.Number
    Case 1001
        MsgBox "Please specify a course title.", vbInformation, "Grading System - Information"
        txtCourseTitle.SetFocus
    Case 1002
        MsgBox "Please specify the number of hours of this course.", vbInformation, "Grading System - Information"
        txtNumberOfHours.SetFocus
    Case Else
       MsgBox "Error " & CStr(Err.Number) & ": Err.Description, vbExclamation, App.Title"
End Select
End Sub

Private Sub Form_Load()
With frmCoursesList.grdCourses
lblCourseID = .TextMatrix(.RowSel, 1)
txtCourseTitle = .TextMatrix(.RowSel, 2)
txtNumberOfHours = .TextMatrix(.RowSel, 3)
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub txtNumberOfHours_KeyPress(KeyAscii As Integer)
If Not IsNumeric(CStr(Chr$(KeyAscii))) Then KeyAscii = 0
End Sub
