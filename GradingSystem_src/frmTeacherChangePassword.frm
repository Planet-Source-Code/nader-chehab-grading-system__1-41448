VERSION 5.00
Begin VB.Form frmTeacherChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changing Teacher's Passowrd"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmTeacherChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   240
      Picture         =   "frmTeacherChangePassword.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Accept"
      Height          =   615
      Left            =   3240
      Picture         =   "frmTeacherChangePassword.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtRetype 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Retype Password:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1245
      Width           =   1290
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "New Password:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   765
      Width           =   1110
   End
End
Attribute VB_Name = "frmTeacherChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmTeacherUpdate.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
If txtNew <> "" Then
If txtRetype = txtNew Then
strSQL = "UPDATE Teachers " & _
         "SET [Password] = '" & Transform(txtNew) & "' " & _
         "WHERE Username = '" & frmTeacherUpdate.txtUserName & "'"
objDBConnection.Execute strSQL
MsgBox "Password changed. Be sure to remember it!", vbInformation, "Grading System - Information"
frmTeacherUpdate.Enabled = True
Unload Me
Else
MsgBox "Retype password correctly.", vbCritical, "Grading System - Security"
txtRetype = ""
txtNew = ""
txtNew.SetFocus
End If
Else
MsgBox "Type a new password", vbCritical, "Grading System - Security"
txtNew.SetFocus
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
