VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Registration"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Course Dates"
      Height          =   1815
      Left            =   5400
      TabIndex        =   30
      Top             =   360
      Width           =   5055
      Begin VB.ComboBox cboSYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cboSMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cboEMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox cboSDay 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboEDay 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboEYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblStartingDate0 
         AutoSize        =   -1  'True
         Caption         =   "Starting Date:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   525
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ending Date:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1005
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Teacher Information"
      Height          =   1815
      Left            =   5640
      TabIndex        =   26
      Top             =   2400
      Width           =   4815
      Begin VB.ComboBox cboTeacherID 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblTeacherName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1800
         TabIndex        =   29
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "TeacherID:"
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
         Index           =   6
         Left            =   120
         TabIndex        =   28
         Top             =   517
         Width           =   1110
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Teacher Name:"
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
         Index           =   8
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Student Information"
      Height          =   1815
      Left            =   240
      TabIndex        =   20
      Top             =   360
      Width           =   5055
      Begin VB.ComboBox cboStudentID 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Text            =   "Student ID"
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "First Name:"
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
         Left            =   120
         TabIndex        =   25
         Top             =   810
         Width           =   1140
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Last Name:"
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
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   1125
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Student ID:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   23
         Top             =   420
         Width           =   1170
      End
      Begin VB.Label lblFirstName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   780
         Width           =   3015
      End
      Begin VB.Label lblLastName 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Course Information"
      Height          =   1815
      Left            =   240
      TabIndex        =   14
      Top             =   2400
      Width           =   5295
      Begin VB.ComboBox cboCourseID 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Text            =   "Course ID"
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label lblNumberOfHours 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label lblCourseTitle 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   810
         Width           =   2895
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Course Number:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   390
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Number Of Hours:"
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
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1200
         Width           =   1755
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Course Title:"
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
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   795
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdCourses 
      Height          =   1935
      Left            =   616
      TabIndex        =   13
      Top             =   5400
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3413
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Appearance      =   0  'Flat
      Caption         =   "Add Course to List"
      Height          =   615
      Left            =   4388
      Picture         =   "frmRegister.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Back"
      Height          =   735
      Left            =   428
      Picture         =   "frmRegister.frx":09B3
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "&Register"
      Height          =   735
      Left            =   8588
      MaskColor       =   &H00DEEBEF&
      Picture         =   "frmRegister.frx":0E9D
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Registered Courses:"
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
      Left            =   600
      TabIndex        =   12
      Top             =   6480
      Width           =   1980
   End
   Begin VB.Menu mnuList 
      Caption         =   "List"
      Visible         =   0   'False
      Begin VB.Menu cmdRemove 
         Caption         =   "&Remove Selected"
      End
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

'Array that will contain the list of months
Dim strMon(12) As String

'Course dates
Dim startingDate, endingDate As Date

'Counter
Dim i As Integer

Private Sub cboTeacherID_Click()
strSQL = "SELECT FirstName, LastName FROM Teachers, TeacherCourse WHERE CourseID = '" & cboCourseID & "' AND TeacherCourse.TeacherID = '" & cboTeacherID & "' AND Teachers.TeacherID = TeacherCourse.TeacherID"
Set objDBRecordset = objDBConnection.Execute(strSQL)
lblTeacherName = objDBRecordset("FirstName") & " " & objDBRecordset("LastName")
End Sub

Private Sub cmdAdd_Click()

'Data Validation
startingDate = CDate(cboSMonth & "-" & cboSDay & "-" & cboSYear)
endingDate = CDate(cboEMonth & "-" & cboEDay & "-" & cboEYear)

If endingDate < startingDate Then
MsgBox "The Course's Ending Date must be after its Starting Date!", vbCritical, "Error"
Exit Sub
End If

If lblCourseTitle = "" Then
MsgBox "Please select a valid coaurse to add!", vbInformation, "Grading System - Information"
Exit Sub
End If
        
If cboTeacherID.text = "" Then
MsgBox "No teacher teaches '" & lblCourseTitle & "'" & vbCrLf & "Please assign a teacher to this course.", vbInformation, "Grading System - Information"
Exit Sub
End If

If found(cboCourseID, "CourseID", "Courses") Then
    With grdCourses
        'if course not added to grid, add it
        For i = 1 To .Rows - 1
            If cboCourseID.text = .TextMatrix(i, 1) And cboTeacherID.text = .TextMatrix(i, 6) Then
            MsgBox "Course already added!", vbInformation, "Grading System - Information"
            Exit Sub
            End If
        Next i
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = cboCourseID.text
        .TextMatrix(.Rows - 1, 2) = lblCourseTitle.Caption
        .TextMatrix(.Rows - 1, 3) = lblNumberOfHours.Caption
        .TextMatrix(.Rows - 1, 4) = startingDate
        .TextMatrix(.Rows - 1, 5) = endingDate
        .TextMatrix(.Rows - 1, 6) = cboTeacherID.text
        .TextMatrix(.Rows - 1, 7) = lblTeacherName
    End With
    If lblFirstName <> "" Then
    cmdRegister.Enabled = True
    cboStudentID.Enabled = False
    Else
    cmdRegister.Enabled = False
    End If
    cboCourseID.SetFocus
    SendKeys "{Home}+{End}"
End If
End Sub

Private Sub cmdRemove_Click()
With grdCourses
If .Rows > 2 Then
.RemoveItem .RowSel
Else
.Rows = 1
cmdRegister.Enabled = False
cboStudentID.Enabled = True
End If
End With
End Sub

Private Sub Form_Load()
    
cmdRegister.Enabled = False
   
    'Populate studentID combo
    strSQL = "SELECT StudentID FROM Students"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    While Not objDBRecordset.EOF
        cboStudentID.AddItem objDBRecordset("StudentID")
        objDBRecordset.MoveNext
    Wend
    
    'Populate CourseID combo
    strSQL = "SELECT CourseID FROM Courses"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    While Not objDBRecordset.EOF
        cboCourseID.AddItem objDBRecordset("CourseID")
        objDBRecordset.MoveNext
    Wend
    
    'Initialize flexgrid
    With grdCourses
    .Clear
    .Cols = 8
    .Rows = 1
    .ColWidth(0) = 0
    .ColWidth(1) = 1000
    .ColWidth(2) = 2100
    .ColWidth(3) = 1400
    .ColWidth(4) = 1050
    .ColWidth(5) = 1000
    .ColWidth(6) = 950
    .ColWidth(7) = 1400
    
    .TextMatrix(0, 1) = "Course ID"
    .TextMatrix(0, 2) = "Course Title"
    .TextMatrix(0, 3) = "Number Of Hours"
    .TextMatrix(0, 4) = "Starting Date"
    .TextMatrix(0, 5) = "Ending Date"
    .TextMatrix(0, 6) = "Teacher ID"
    .TextMatrix(0, 7) = "Teacher Name"
    SendKeys "{RIGHT}"
    End With
    
    
    'The array strMonth will contain the list of months
    strMon(1) = "January"
    strMon(2) = "February"
    strMon(3) = "March"
    strMon(4) = "April"
    strMon(5) = "May"
    strMon(6) = "June"
    strMon(7) = "July"
    strMon(8) = "August"
    strMon(9) = "September"
    strMon(10) = "October"
    strMon(11) = "November"
    strMon(12) = "December"
    
    'Populate date combos
    For i = 1 To 12
    cboSMonth.AddItem strMon(i)
    cboEMonth.AddItem strMon(i)
    Next i
    
    For i = 1 To 31
    cboSDay.AddItem i
    cboEDay.AddItem i
    Next i
    
    For i = 1990 To 2010
    cboSYear.AddItem i
    cboEYear.AddItem i
    Next i
    
    'Initialize date combos to some values
    cboSMonth.text = "January"
    cboSDay.text = "1"
    cboSYear.text = "2003"
    cboEMonth.text = "January"
    cboEDay.text = "1"
    cboEYear.text = "2004"
    
End Sub

Private Sub cboStudentID_Click()

If found(cboStudentID.text, "StudentID", "Students") Then
'if the chosen value is found in the database,
'show the student's FirstName and LastName.
    strSQL = "SELECT FirstName, LastName FROM Students WHERE StudentID = '" & cboStudentID.text & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    lblFirstName.Caption = objDBRecordset("FirstName")
    lblLastName.Caption = objDBRecordset("LastName")

    'If there is no course chosen yet,
    'lock the register button,
    'else, unlock it.
    If grdCourses.Rows = 1 Then
    cmdRegister.Enabled = False
    Else: cmdRegister.Enabled = True
    End If
End If
End Sub

Private Sub cboStudentID_Change()
cmdRegister.Enabled = False
lblFirstName = ""
lblLastName = ""

If Len(cboStudentID.text) = 7 Then
'if 8 characters are typed
    If found(cboStudentID.text, "StudentID", "Students") Then
    'If the entered StudentID is found in the database,
    'show FirstName and LastName
        strSQL = "SELECT FirstName, LastName FROM Students WHERE StudentID = '" & cboStudentID.text & "'"
        Set objDBRecordset = objDBConnection.Execute(strSQL)
        lblFirstName.Caption = objDBRecordset("FirstName")
        lblLastName.Caption = objDBRecordset("LastName")
        
        'If there is no course chosen yet,
        'lock the register button,
        'else, unlock it.
        If grdCourses.Rows = 1 Then
        cmdRegister.Enabled = False
        Else: cmdRegister.Enabled = True
        End If
    End If
End If

End Sub

Private Sub cboCourseID_Click()
If found(cboCourseID.text, "CourseID", "Courses") Then
    strSQL = "SELECT CourseTitle, NumberOfHours FROM Courses WHERE CourseID = " & "'" & cboCourseID.text & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    lblCourseTitle.Caption = objDBRecordset("CourseTitle")
    lblNumberOfHours.Caption = objDBRecordset("NumberOfHours")
    
    'Populate TeacherID combo
    strSQL = "SELECT TeacherID FROM TeacherCourse WHERE CourseID = '" & cboCourseID & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    cboTeacherID.Clear
    lblTeacherName.Caption = ""
    While Not objDBRecordset.EOF
    cboTeacherID.AddItem objDBRecordset("TeacherID")
    objDBRecordset.MoveNext
    Wend
    cboTeacherID.SetFocus
    SendKeys "{DOWN}", 200
    cboCourseID.SetFocus
End If
End Sub

Private Sub cboCourseID_Change()
'if the user types a CourseID manually...
'same procedure as for the StudentID
lblCourseTitle = ""
lblNumberOfHours = ""
cboTeacherID.Clear
lblTeacherName.Caption = ""

If Len(cboCourseID.text) = 10 Then
    
    If found(cboCourseID.text, "CourseID", "Courses") Then
        strSQL = "SELECT CourseTitle, NumberOfHours FROM Courses WHERE CourseID = " & "'" & cboCourseID & "'"
        Set objDBRecordset = objDBConnection.Execute(strSQL)
        lblCourseTitle.Caption = objDBRecordset("CourseTitle")
        lblNumberOfHours.Caption = objDBRecordset("NumberOfHours")
        
        'Populate TeacherID combo
        strSQL = "SELECT TeacherID FROM TeacherCourse WHERE CourseID = '" & cboCourseID & "'"
        Set objDBRecordset = objDBConnection.Execute(strSQL)
        While Not objDBRecordset.EOF
        cboTeacherID.AddItem objDBRecordset("TeacherID")
        objDBRecordset.MoveNext
        Wend
        cboTeacherID.text = cboTeacherID.list(0)
        cboTeacherID.SetFocus
        SendKeys "{DOWN}", 200
        cboCourseID.SetFocus
    End If
End If

End Sub

'Register the student
Private Sub cmdRegister_Click()
Dim strCourseList As String

With grdCourses
    For i = 1 To .Rows - 1
    strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & .TextMatrix(i, 6) & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    If found(cboStudentID, "StudentID", "[" & objDBRecordset("Username") & .TextMatrix(i, 1) & "]") Then
    MsgBox lblFirstName & " " & lblLastName & " is already registered to " & .TextMatrix(i, 2) & vbCrLf & "Please remove this course from the list.", vbInformation, "Grading System - Error"
    Exit Sub
    End If
    Next i
    
    For i = 1 To .Rows - 1
    strSQL = "SELECT Username FROM Teachers WHERE TeacherID = '" & .TextMatrix(i, 6) & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    strSQL = "INSERT INTO [" & objDBRecordset("Username") & .TextMatrix(i, 1) & "] " & _
    "(StudentID, StartingDate, EndingDate) VALUES('" & cboStudentID & "', #" & startingDate & "#, #" & endingDate & "#)"
    objDBConnection.Execute (strSQL)
    Next i
    
    For i = 1 To .Rows - 1
    strCourseList = strCourseList & "-  " & .TextMatrix(i, 2) & "  (with " & .TextMatrix(i, 7) & ")" & vbCrLf
    Next i
End With
MsgBox lblFirstName & " " & lblLastName & " has been registered to the following courses: " & vbCrLf & vbCrLf & strCourseList, vbInformation, "Information"
frmRegistrations.Show
Unload Me
End Sub

Private Sub cmdCancel_Click()
frmRegistrations.Show
Unload Me
End Sub

Private Sub grdCourses_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And grdCourses.Rows > 1 Then PopupMenu mnuList
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

