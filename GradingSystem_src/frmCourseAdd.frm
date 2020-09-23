VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCourseAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding a Course"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   ControlBox      =   0   'False
   Icon            =   "frmCourseAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   6180
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Course Infomation"
      Height          =   1695
      Left            =   623
      TabIndex        =   5
      Top             =   360
      Width           =   4935
      Begin VB.TextBox txtCourseTitle 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   2535
      End
      Begin MSMask.MaskEdBox mskCourseID 
         Height          =   255
         Left            =   1680
         TabIndex        =   0
         Top             =   360
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         Mask            =   ">AAA-AAA-AA"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskNumberOfHours 
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   2
         Mask            =   "##"
         PromptChar      =   "_"
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Course ID:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   405
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Number Of Hours:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Course Title:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   765
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdAddCourse 
      Caption         =   "&Add Course"
      Height          =   615
      Left            =   4200
      Picture         =   "frmCourseAdd.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   240
      Picture         =   "frmCourseAdd.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
   End
End
Attribute VB_Name = "frmCourseAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddCourse_Click()
On Error GoTo CourseAddError

    If found(mskCourseID, "CourseID", "Courses") Then Err.Raise 1001
    If Right(mskCourseID.text, 1) = "_" Then Err.Raise 1002
    If txtCourseTitle = "" Then Err.Raise 1003
    If Left(mskNumberOfHours, 1) = "_" Then Err.Raise 1004
    
    strSQL = "INSERT INTO Courses VALUES('" & mskCourseID & "', '" & txtCourseTitle & "', '" & mskNumberOfHours & "')"
    objDBConnection.Execute (strSQL)
    Call UpdateCoursesList
    frmCoursesList.Enabled = True
    Unload Me
    frmCoursesList.SetFocus
Exit Sub

CourseAddError:
Select Case Err.Number
    Case 1001
        MsgBox "CourseID already exists!", vbInformation, "Grading System - Information"
        Dim temp As String
        temp = mskCourseID.Mask
        mskCourseID.Mask = ""
        mskCourseID.text = ""
        mskCourseID.Mask = temp
        mskCourseID.SetFocus
    Case 1002
        MsgBox "Course ID must be 8 characters", vbInformation, "Grading System - Information"
        mskCourseID.SetFocus
    Case 1003
        MsgBox "Please specify a course title", vbInformation, "Grading System - Information"
        txtCourseTitle.SetFocus
    Case 1004
        MsgBox "Please specify the number of hours of this course", vbInformation, "Grading System - Information"
        mskNumberOfHours.SetFocus
    Case Else
       MsgBox "Error " & CStr(Err.Number) & ": Err.Description, vbExclamation, App.Title"
End Select
End Sub

Private Sub cmdCancel_Click()
frmCoursesList.Enabled = True
frmCoursesList.SetFocus
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
