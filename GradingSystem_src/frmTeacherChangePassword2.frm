VERSION 5.00
Begin VB.Form frmTeacherChangePassword2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changing Password"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmTeacherChangePassword2.frx":0000
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
      Picture         =   "frmTeacherChangePassword2.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Accept"
      Height          =   615
      Left            =   3240
      Picture         =   "frmTeacherChangePassword2.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtOld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1793
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1793
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtRetype 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1793
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Old Password:"
      Height          =   195
      Index           =   0
      Left            =   353
      TabIndex        =   7
      Top             =   525
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "New Password:"
      Height          =   195
      Index           =   1
      Left            =   353
      TabIndex        =   6
      Top             =   1005
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Retype Password:"
      Height          =   195
      Index           =   2
      Left            =   353
      TabIndex        =   5
      Top             =   1485
      Width           =   1290
   End
End
Attribute VB_Name = "frmTeacherChangePassword2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmTeacherModifyInfo.Enabled = True
Unload Me
End Sub

Private Sub cmdOK_Click()
strSQL = "SELECT Password FROM Teachers WHERE Username = '" & frmTeacherModifyInfo.txtUserName & "'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Transform(txtOld) = objDBRecordset("Password") Then
If txtNew <> "" Then
If txtRetype = txtNew Then
strSQL = "UPDATE Teachers " & _
         "SET [Password] = '" & Transform(txtNew) & "' " & _
         "WHERE Username = '" & frmTeacherModifyInfo.txtUserName & "'"
objDBConnection.Execute strSQL
MsgBox "Password changed. Be sure to remember it!", vbInformation, "Grading System - Information"
frmTeacherModifyInfo.Enabled = True
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
Else
MsgBox "Old password invalid.", vbCritical, "Grading System - Security"
txtNew = ""
txtRetype = ""
txtOld.SetFocus
SendKeys "{Home}+{End}"
End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
