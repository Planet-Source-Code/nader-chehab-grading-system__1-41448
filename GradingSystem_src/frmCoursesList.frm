VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCoursesList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Records"
   ClientHeight    =   5250
   ClientLeft      =   150
   ClientTop       =   525
   ClientWidth     =   5985
   ControlBox      =   0   'False
   FillStyle       =   3  'Vertical Line
   ForeColor       =   &H00000000&
   Icon            =   "frmCoursesList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Main Menu"
      Height          =   734
      Left            =   360
      Picture         =   "frmCoursesList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddCourse 
      Caption         =   "&Add a Course"
      Height          =   734
      Left            =   4440
      Picture         =   "frmCoursesList.frx":0D42
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdCourses 
      Height          =   3015
      Left            =   360
      TabIndex        =   2
      Top             =   1200
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5318
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Courses Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   4
      Left            =   1830
      TabIndex        =   4
      Top             =   360
      Width           =   2190
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Double-click on a record to view it or update it:"
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
      Left            =   405
      TabIndex        =   3
      Top             =   840
      Width           =   4770
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu cmdAddCourse2 
         Caption         =   "&Add a Course"
      End
      Begin VB.Menu cmdUpdateCourse 
         Caption         =   "&Update Course"
      End
      Begin VB.Menu cmdDeleteCourse 
         Caption         =   "&Delete Course"
      End
   End
End
Attribute VB_Name = "frmCoursesList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddCourse_Click()
Me.Enabled = False
frmCourseAdd.Show
End Sub

Private Sub cmdAddCourse2_Click()
Me.Enabled = False
frmCourseAdd.Show
End Sub

Private Sub cmdCancel_Click()
frmAdminMenu.Show
Unload Me
End Sub

Private Sub cmdUpdateCourse_Click()
Me.Enabled = False
frmCourseUpdate.Show
End Sub

Private Sub cmdDeleteCourse_Click()
Dim answer As Integer
answer = MsgBox("Warning: All students and teachers registrations to this course will be removed too." & vbCrLf & "Are you sure you want to remove the selected Course ?", vbYesNoCancel Or vbQuestion, "GradingSystem - Removing")
If answer = vbYes Then Call DeleteCourse
End Sub

Private Sub Form_Load()
Call UpdateCoursesList
End Sub

Private Sub grdCourses_DblClick()
If grdCourses.Rows > 1 Then
Me.Enabled = False
frmCourseUpdate.Show
End If
End Sub

Private Sub grdCourses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And grdCourses.Rows > 1 Then PopupMenu mnuList
End Sub

'Drag the Form from anywhere
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

