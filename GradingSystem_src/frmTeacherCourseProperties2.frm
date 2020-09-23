VERSION 5.00
Begin VB.Form frmTeacherCourseProperties2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Properties - Specify your tests"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "frmTeacherCourseProperties2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Notice"
      Height          =   1215
      Left            =   720
      TabIndex        =   12
      Top             =   1560
      Width           =   6495
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         Picture         =   "frmTeacherCourseProperties2.frx":08CA
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblNote 
         Caption         =   $"frmTeacherCourseProperties2.frx":0E6F
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame4 
      Height          =   3615
      Left            =   1320
      TabIndex        =   11
      Top             =   3000
      Width           =   5415
      Begin VB.TextBox txtTest 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   2415
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   375
         Left            =   3720
         TabIndex        =   1
         Top             =   435
         Width           =   1095
      End
      Begin VB.ListBox lstTests 
         Height          =   2400
         Left            =   1200
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Next"
      Height          =   615
      Left            =   6600
      Picture         =   "frmTeacherCourseProperties2.frx":0EF6
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6960
      Width           =   1100
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   720
      TabIndex        =   6
      Top             =   240
      Width           =   6495
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Course:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lblTeacher0 
         AutoSize        =   -1  'True
         Caption         =   "Teacher:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   285
         Width           =   1065
      End
      Begin VB.Label lblCourse 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label lblTeacher 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   615
      Left            =   120
      Picture         =   "frmTeacherCourseProperties2.frx":1016
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1100
   End
End
Attribute VB_Name = "frmTeacherCourseProperties2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
Call AddTest
End Sub

Private Sub cmdExit_Click()
lstTests.Clear
frmTeacherMenu.Show
Unload Me
End Sub

Private Sub cmdRemove_Click()
If lstTests.ListIndex > -1 Then
lstTests.RemoveItem lstTests.ListIndex
End If
End Sub

Private Sub cmdValidate_Click()
'data validation
If lstTests.ListCount = 0 Then
MsgBox "Please add at least one test for this course.", vbInformation, "Grading System - Information"
Exit Sub
End If

'Add the Tests Columns
On Error Resume Next
For i = 0 To lstTests.ListCount - 1
strSQL = "ALTER TABLE [" & strUsername & strSelectedCourse & "] " & _
         "ADD [" & lstTests.list(i) & "] VARCHAR(5)"
objDBConnection.Execute (strSQL)
If Err.Number = -2147467259 Then
MsgBox Err.Description, vbInformation, "Grading System - Information"
Exit Sub
End If
Next i

strSQL = "INSERT INTO [" & strUsername & strSelectedCourse & "]" & _
         "(StudentID) VALUES('0000000')"
objDBConnection.Execute (strSQL)

frmTeacherCourseProperties3.Show
Unload Me
End Sub

Private Sub Form_Load()
lblTeacher = strUsername
lblCourse = getTitle(strSelectedCourse) & " - " & strSelectedCourse
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub txtTest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If txtTest <> "" Then
KeyAscii = 0
Call AddTest
End If
End If
End Sub

Private Function AddTest()
If lstTests.ListCount > 9 Then
    MsgBox "The maximum number of tests allowed is 10 tests.", vbInformation, "Grading System - Information"
    txtTest = ""
    txtTest.SetFocus
    Exit Function
End If

For i = 0 To lstTests.ListCount - 1
    If txtTest.text = lstTests.list(i) Then
    MsgBox "This test already exists!", vbInformation, "Grading System - Information"
    txtTest.SetFocus
    SendKeys "{Home}+{End}"
    Exit Function
    End If
Next i

lstTests.AddItem txtTest.text
txtTest = ""
txtTest.SetFocus
End Function
