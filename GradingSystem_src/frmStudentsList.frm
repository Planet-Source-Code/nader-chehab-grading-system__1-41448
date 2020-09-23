VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmStudentsList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Records"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmStudentsList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Main Menu"
      Height          =   734
      Left            =   360
      Picture         =   "frmStudentsList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddStudent 
      Caption         =   "&Add a Student"
      Height          =   734
      Left            =   4313
      Picture         =   "frmStudentsList.frx":0D42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdStudents 
      Height          =   3015
      Left            =   315
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5318
      _Version        =   393216
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
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
      Left            =   360
      TabIndex        =   4
      Top             =   720
      Width           =   4770
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Students Records"
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
      Left            =   1785
      TabIndex        =   3
      Top             =   240
      Width           =   2235
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu cmdAddStudent2 
         Caption         =   "&Add a Student"
      End
      Begin VB.Menu cmdUpdateStudent 
         Caption         =   "&Update Student"
      End
      Begin VB.Menu cmdDeleteStudent 
         Caption         =   "&Delete Student"
      End
   End
End
Attribute VB_Name = "frmStudentsList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddStudent_Click()
Me.Enabled = False
frmStudentAdd.Show
End Sub

Private Sub cmdAddStudent2_Click()
Me.Enabled = False
frmStudentAdd.Show
End Sub

Private Sub cmdCancel_Click()
frmAdminMenu.Show
Unload Me
End Sub

Private Sub cmdUpdateStudent_Click()
Me.Enabled = False
frmStudentUpdate.Show
End Sub

Private Sub cmdDeleteStudent_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to remove the selected Student ?", vbYesNoCancel Or vbQuestion, "GradingSystem - Removing")
If answer = vbYes Then Call DeleteStudent
End Sub

Private Sub Form_Load()
Call UpdateStudentsList
End Sub

Private Sub grdStudents_DblClick()
If grdStudents.Rows > 1 Then
Me.Enabled = False
frmStudentUpdate.Show
End If
End Sub

Private Sub grdStudents_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And grdStudents.Rows > 1 Then PopupMenu mnuList
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

