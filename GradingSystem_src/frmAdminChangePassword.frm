VERSION 5.00
Begin VB.Form frmAdminChangePassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Changing Administrator's Password"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4770
   ControlBox      =   0   'False
   Icon            =   "frmAdminChangePassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   4770
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOld 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtNew 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox txtRetype 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Accept"
      Height          =   615
      Left            =   3240
      Picture         =   "frmAdminChangePassword.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   240
      Picture         =   "frmAdminChangePassword.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Old Password:"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   645
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "New Password:"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1125
      Width           =   1110
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Retype Password:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   1605
      Width           =   1290
   End
End
Attribute VB_Name = "frmAdminChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmAdminMenu.Enabled = True
frmAdminMenu.SetFocus
Unload Me
End Sub

Private Sub cmdOK_Click()
  
strSQL = "SELECT Password FROM Teachers WHERE Username = 'Administrator'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Transform(txtOld) = objDBRecordset("Password") Then
If Len(txtNew) >= 6 Then
If txtRetype = txtNew Then
strSQL = "UPDATE Teachers " & _
         "SET [Password] = '" & Transform(txtNew) & "' " & _
         "WHERE Username = 'Administrator'"
objDBConnection.Execute strSQL

objDBConnection.Close
objDBConnection.Mode = adModeShareExclusive
Call ConnectToDatabase
Call WriteToFile(Transform(txtNew))

'change database password to admin's password
strSQL = "ALTER Database Password " & txtNew & " " & txtOld
GradingSystem.objDBConnection.Execute (strSQL)

MsgBox "Password changed! Do not forget it.", vbInformation, "Grading System - Information"
frmAdminMenu.Enabled = True
frmAdminMenu.SetFocus
Unload Me
objDBConnection.Close
objDBConnection.Mode = adModeShareDenyNone
Call ConnectToDatabase
Else
MsgBox "Retype password correctly.", vbCritical, "Grading System - Security"
txtRetype = ""
txtNew = ""
txtNew.SetFocus
End If
Else
MsgBox "The password must be at least 6 characters.", vbCritical, "Grading System - Security"
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
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

