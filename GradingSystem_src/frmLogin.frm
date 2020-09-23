VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grading System"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin GradingSystem.HoverCommand cmdExit 
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "&Exit"
      Style           =   3
      Picture         =   "frmLogin.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   1875
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1470
      Width           =   2415
   End
   Begin VB.ComboBox cboUserName 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin GradingSystem.HoverCommand cmdLogin 
      Height          =   615
      Left            =   3480
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "&Login"
      Style           =   3
      Picture         =   "frmLogin.frx":09A3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   900
      TabIndex        =   6
      Top             =   330
      Width           =   705
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   240
      Picture         =   "frmLogin.frx":0D51
      Top             =   240
      Width           =   585
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Password:"
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
      Index           =   1
      Left            =   645
      TabIndex        =   5
      Top             =   1530
      Width           =   1020
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "&Username:"
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
      Left            =   645
      TabIndex        =   4
      Top             =   960
      Width           =   1050
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bytRetries As Byte

Private Sub Form_Load()
        
bytRetries = 0

'populate the Username combo-box
strSQL = "SELECT Username, Password, LastName FROM Teachers"
Set objDBRecordset = objDBConnection.Execute(strSQL)
objDBRecordset.MoveFirst
While Not objDBRecordset.EOF
cboUserName.AddItem objDBRecordset("Username")
objDBRecordset.MoveNext
Wend
cboUserName.text = "Administrator"
End Sub

Private Sub cmdLogin_Click()
Call ValidatePassword
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Call ValidatePassword
End If
End Sub

Private Sub ValidatePassword()
'go to the selected username
objDBRecordset.MoveFirst
While objDBRecordset("Username") <> cboUserName.text And Not objDBRecordset.EOF
objDBRecordset.MoveNext
Wend
'if the passwords match, login
If Transform(txtPassword.text) = objDBRecordset("Password") Then
    strUsername = cboUserName.text
    Unload Me
    If strUsername = "Administrator" Then
    frmAdminMenu.Show
    frmAdminMenu.Hide
    ExplodeForm frmAdminMenu, 1000
    frmAdminMenu.Show
    Else
    frmTeacherMenu.Show
    frmTeacherMenu.Hide
    ExplodeForm frmTeacherMenu, 1000
    frmTeacherMenu.Show
    End If
Else
    'after 3 retires, quit
    If bytRetries = 3 Then
    MsgBox "Guessing is not allowed!!", vbCritical, "Grading System - Security"
    End
    Else
    bytRetries = bytRetries + 1
    MsgBox "Invalid password. Access denied. Try again.", vbCritical, "Grading System - Security"
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    End If
End If
End Sub

'Form always on top. Taken from AllAPI.net
Private Sub Form_Activate()
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

