VERSION 5.00
Begin VB.Form frmCourseSelection2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Selection"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5250
   ControlBox      =   0   'False
   Icon            =   "frmCourseSelection2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboCourses 
      Height          =   315
      Left            =   523
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   730
      Left            =   3868
      Picture         =   "frmCourseSelection2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1560
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   730
      Left            =   283
      Picture         =   "frmCourseSelection2.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1100
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Please select a Grading Table:"
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
      Left            =   525
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "frmCourseSelection2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
frmTeacherMenu.Show
frmTeacherMenu.Enabled = True
frmTeacherMenu.SetFocus
Unload Me
End Sub

'Checks if students are registered to the selected course
'and if the course's properties have been specifies.
Private Sub cmdValidate_Click()
strSelectedCourse = Right$(cboCourses, 10)

strSQL = "SELECT COUNT(*) FROM [" & strUsername & strSelectedCourse & "]"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If objDBRecordset.Fields.Item(0) = 1 Then
MsgBox "You need to register students to this course first! Please contact the Administrator.", vbInformation, "Grading System - Information"
Exit Sub
End If

strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "]"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If objDBRecordset.Fields.Count = 5 Then
MsgBox "You need to specifiy the properties of this course first! Select 'Courses Properties' from the Main Menu.", vbInformation, "Grading System - Information"
Exit Sub
End If

frmGradingTable.Show
Unload Me
frmTeacherMenu.Enabled = True
Unload frmTeacherMenu
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

