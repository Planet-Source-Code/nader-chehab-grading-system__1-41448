VERSION 5.00
Begin VB.Form frmTeacherAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding a Teacher"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ControlBox      =   0   'False
   Icon            =   "frmTeacherAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   6945
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Teacher Information"
      Height          =   3015
      Left            =   1080
      TabIndex        =   7
      Top             =   360
      Width           =   4695
      Begin VB.TextBox txtRetype 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2280
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1920
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label lblTeacherID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Retype Pass:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   13
         Top             =   2325
         Width           =   945
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   885
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   1245
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Teacher ID:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   1605
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   8
         Top             =   1965
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   360
      Picture         =   "frmTeacherAdd.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddTeacher 
      Caption         =   "&Add Teacher"
      Height          =   615
      Left            =   4920
      Picture         =   "frmTeacherAdd.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   1695
   End
End
Attribute VB_Name = "frmTeacherAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
strSQL = "SELECT TeacherID FROM Teachers"

'objDBRecordset.MoveLast won't work with: Set objDBRecordset = objDBConnection.Execute(strSQL)
'so we use this instead:
objDBRecordset.Close
objDBRecordset.Open strSQL, objDBConnection, 1, 3
objDBRecordset.MoveLast

'We create an ID by adding 1 to the last ID in the field
lblTeacherID = Format(Val(objDBRecordset("TeacherID")) + 1, "0000")
End Sub

Private Sub cmdAddTeacher_Click()

'Data Validation
If txtFirstName = "" Then
MsgBox "Enter a First Name", vbCritical, "Grading System - Error"
txtFirstName.SetFocus
Exit Sub
End If

If txtLastName = "" Then
MsgBox "Enter a Last Name", vbCritical, "Grading System - Error"
txtLastName.SetFocus
Exit Sub
End If

If txtUserName = "" Then
MsgBox "Enter a Username", vbCritical, "Grading System - Error"
txtUserName.SetFocus
Exit Sub
End If

If found(txtUserName, "Username", "Teachers") Then
MsgBox "Username already exists!", vbCritical, "Grading System - Error"
txtUserName.SetFocus
SendKeys "{Home}+{End}"
Exit Sub
End If

If txtPassword = "" Then
MsgBox "Enter a Password", vbCritical, "Grading System - Error"
txtPassword.SetFocus
Exit Sub
End If

If txtRetype <> txtPassword Then
MsgBox "Retype the Password correctly.", vbCritical, "Grading System - Error"
txtPassword = ""
txtRetype = ""
txtPassword.SetFocus
Exit Sub
End If

'All info is valid, we add the Teacher:
strSQL = "INSERT INTO Teachers VALUES('" & lblTeacherID & "', '" & txtFirstName & "', '" & txtLastName & "', '" & txtUserName & "', '" & Transform(txtPassword) & "')"
Set objDBRecordset = objDBConnection.Execute(strSQL)

Dim answer As Integer
answer = MsgBox("Teacher Added. Do you want to specify now which courses " & txtFirstName & " " & txtLastName & " will teach ? ", vbYesNoCancel Or vbQuestion, "Grading System - Adding a Teacher")
frmTeachersList.Enabled = True
Unload Me
Call UpdateTeachersList
If answer = vbYes Then
frmTeachersList.grdTeachers.SetFocus
For i = 1 To frmTeachersList.grdTeachers.Rows - 1
SendKeys "{Down}", 500
Next i
frmTeachersList.Enabled = False
frmTeacherUpdate.Show
frmTeacherUpdate.cboAllCourses.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmTeachersList.Enabled = True
frmTeachersList.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

