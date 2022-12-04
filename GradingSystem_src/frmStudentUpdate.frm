VERSION 5.00
Begin VB.Form frmStudentUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Update"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   ControlBox      =   0   'False
   Icon            =   "frmStudentUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Student Information"
      Height          =   1815
      Left            =   675
      TabIndex        =   4
      Top             =   360
      Width           =   5055
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student ID:"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   8
         Top             =   525
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   7
         Top             =   1245
         Width           =   810
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   6
         Top             =   885
         Width           =   795
      End
      Begin VB.Label lblStudentID 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
   End
   Begin VB.CommandButton cmdUpdateStudent 
      Caption         =   "&Update"
      Height          =   615
      Left            =   4230
      Picture         =   "frmStudentUpdate.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   390
      Picture         =   "frmStudentUpdate.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmStudentUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmStudentsList.Enabled = True
Unload Me
End Sub

Private Sub cmdUpdateStudent_Click()
On Error GoTo StudentAddError
   If txtFirstName = "" Then Err.Raise 1001
   If txtLastName = "" Then Err.Raise 1002
    
    strSQL = "UPDATE Students SET FirstName = '" & txtFirstName & "', " & _
             "LastName = '" & txtLastName & "' " & _
             "WHERE StudentID = '" & lblStudentID & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    frmStudentsList.Enabled = True
    Unload Me
    Call UpdateStudentsList
    frmStudentsList.SetFocus
Exit Sub

StudentAddError:
Select Case Err.Number
    Case 1001
        MsgBox "Please specify student's First Name.", vbInformation, "Grading System - Information"
        txtFirstName.SetFocus
    Case 1001
        MsgBox "Please specify student's Last Name.", vbInformation, "Grading System - Information"
        txtLastName.SetFocus
    Case Else
       MsgBox "Error " & CStr(Err.Number) & ": Err.Description, vbExclamation, App.Title"
End Select
End Sub

Private Sub Form_Load()
With frmStudentsList.grdStudents
lblStudentID = .TextMatrix(.RowSel, 1)
txtFirstName = .TextMatrix(.RowSel, 2)
txtLastName = .TextMatrix(.RowSel, 3)
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

