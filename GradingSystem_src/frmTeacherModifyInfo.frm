VERSION 5.00
Begin VB.Form frmTeacherModifyInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modify Personal Information"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "frmTeacherModifyInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Personal Information"
      Height          =   2415
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   5775
      Begin VB.CommandButton cmdModify 
         Caption         =   "&Modify Password"
         Height          =   405
         Left            =   4080
         TabIndex        =   2
         Top             =   1140
         Width           =   1575
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblPassword 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "*shadowed*"
         ForeColor       =   &H80000011&
         Height          =   285
         Left            =   1440
         TabIndex        =   13
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblTeacherID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   1605
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   10
         Top             =   1965
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Teacher ID:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   885
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   7
         Top             =   1245
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   615
      Left            =   4890
      Picture         =   "frmTeacherModifyInfo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   450
      Picture         =   "frmTeacherModifyInfo.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmTeacherModifyInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
frmTeacherMenu.Enabled = True
Unload Me
frmTeacherMenu.SetFocus
End Sub

Private Sub cmdModify_Click()
frmTeacherChangePassword2.Show
Me.Enabled = False
End Sub

Private Sub cmdValidate_Click()
strSQL = "SELECT TeacherID, FirstName, LastName, Username FROM Teachers"
Set objDBRecordset = objDBConnection.Execute(strSQL)
While Not objDBRecordset.EOF And objDBRecordset("TeacherID") <> lblTeacherID
objDBRecordset.MoveNext
Wend
If txtFirstName <> objDBRecordset("FirstName") Or txtLastName <> objDBRecordset("LastName") Then
    strSQL = "UPDATE Teachers SET FirstName = '" & txtFirstName & "' WHERE TeacherID = '" & lblTeacherID & "'"
    objDBConnection.Execute (strSQL)
    strSQL = "UPDATE Teachers SET LastName = '" & txtLastName & "' WHERE TeacherID = '" & lblTeacherID & "'"
    objDBConnection.Execute (strSQL)
End If
Unload Me
frmTeacherMenu.Enabled = True
frmTeacherMenu.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
