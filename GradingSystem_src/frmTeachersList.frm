VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTeachersList 
   BorderStyle     =   0  'None
   Caption         =   "Teachers Records"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5985
   ControlBox      =   0   'False
   Icon            =   "frmTeachersList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5143.636
   ScaleMode       =   0  'User
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Main Menu"
      Height          =   734
      Left            =   360
      Picture         =   "frmTeachersList.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddTeacher 
      Caption         =   "&Add a Teacher"
      Height          =   734
      Left            =   4313
      Picture         =   "frmTeachersList.frx":0D42
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdTeachers 
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
      Caption         =   "Teachers Records"
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
      Width           =   2340
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu cmdAddTeacher2 
         Caption         =   "&Add a Teacher"
      End
      Begin VB.Menu cmdDeleteTeacher 
         Caption         =   "&Delete Teacher"
      End
      Begin VB.Menu cmdUpdateTeacher 
         Caption         =   "&Update Teacher"
      End
   End
End
Attribute VB_Name = "frmTeachersList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddTeacher_Click()
Me.Enabled = False
frmTeacherAdd.Show
End Sub

Private Sub cmdAddTeacher2_Click()
Me.Enabled = False
frmTeacherAdd.Show
End Sub

Private Sub cmdCancel_Click()
frmAdminMenu.Show
Unload Me
End Sub

Private Sub cmdUpdateTeacher_Click()
Me.Enabled = False
frmTeacherUpdate.Show
End Sub

Private Sub cmdDeleteTeacher_Click()
Dim answer As Integer
answer = MsgBox("Are you sure you want to remove the selected Teacher ? " & _
        vbCrLf & "(his course registrations will also be deleted.)", vbYesNoCancel Or vbQuestion, "GradingSystem - Removing")
If answer = vbYes Then Call DeleteTeacher
End Sub

Private Sub Form_Load()
Call UpdateTeachersList
End Sub

Private Sub grdTeachers_DblClick()
If grdTeachers.Rows > 1 Then
Me.Enabled = False
frmTeacherUpdate.Show
End If
End Sub

Private Sub grdTeachers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And grdTeachers.Rows > 1 Then PopupMenu mnuList
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
