VERSION 5.00
Begin VB.Form frmStudentAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adding a Student"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   ControlBox      =   0   'False
   Icon            =   "frmStudentAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      Height          =   1935
      Left            =   720
      TabIndex        =   5
      Top             =   360
      Width           =   4815
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtStudentID 
         Height          =   285
         Left            =   1680
         MaxLength       =   7
         TabIndex        =   0
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   885
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1245
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student ID:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   6
         Top             =   525
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   360
      Picture         =   "frmStudentAdd.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdAddStudent 
      Caption         =   "&Add Student"
      Height          =   615
      Left            =   4275
      Picture         =   "frmStudentAdd.frx":0993
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2640
      Width           =   1695
   End
End
Attribute VB_Name = "frmStudentAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAddStudent_Click()
On Error GoTo StudentAddError
   If found(txtStudentID, "StudentID", "Students") Then Err.Raise 1001
   If Len(txtStudentID) < 7 Then Err.Raise 1002
   If txtFirstName = "" Then Err.Raise 1003
   If txtLastName = "" Then Err.Raise 1004
    
    strSQL = "INSERT INTO Students VALUES('" & txtStudentID & "', '" & txtFirstName & "', '" & txtLastName & "')"
    objDBConnection.Execute (strSQL)
    Call UpdateStudentsList
    frmStudentsList.Enabled = True
    Unload Me
    frmStudentsList.SetFocus
Exit Sub

StudentAddError:
Select Case Err.Number
    Case 1001
        MsgBox "Student ID already exists!", vbInformation, "Grading System - Information"
        txtStudentID.SetFocus
        SendKeys "{Home}+{End}"
    Case 1002
        MsgBox "Student ID must be 7 characters.", vbInformation, "Grading System - Information"
        txtStudentID.SetFocus
        SendKeys "{Home}+{End}"
    Case 1003
        MsgBox "Please specify student's First Name.", vbInformation, "Grading System - Information"
        txtFirstName.SetFocus
    Case 1004
        MsgBox "Please specify student's Last Name.", vbInformation, "Grading System - Information"
        txtLastName.SetFocus
    Case Else
       MsgBox "Error " & CStr(Err.Number) & ": Err.Description, vbExclamation, App.Title"
End Select

End Sub

Private Sub cmdCancel_Click()
frmStudentsList.Enabled = True
Unload Me
frmStudentsList.SetFocus
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub txtStudentID_KeyPress(KeyAscii As Integer)
If Not IsNumeric(CStr(Chr$(KeyAscii))) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then KeyAscii = 0
End Sub
