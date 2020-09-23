VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmGradingTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grading Table"
   ClientHeight    =   6615
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12315
   ControlBox      =   0   'False
   Icon            =   "frmGradingTable.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   12315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Show Excel Report"
      Height          =   855
      Left            =   10560
      Picture         =   "frmGradingTable.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&Main Menu"
      Height          =   855
      Left            =   240
      Picture         =   "frmGradingTable.frx":09E8
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Calculate Totals"
      Height          =   855
      Left            =   8760
      Picture         =   "frmGradingTable.frx":0E60
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   5160
      TabIndex        =   5
      Top             =   120
      Width           =   6975
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
         TabIndex        =   9
         Top             =   240
         Width           =   5535
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
         Width           =   5535
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
         TabIndex        =   7
         Top             =   285
         Width           =   1065
      End
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
         TabIndex        =   6
         Top             =   720
         Width           =   945
      End
   End
   Begin VB.TextBox txtGrade 
      Alignment       =   1  'Right Justify
      Height          =   220
      Left            =   0
      MaxLength       =   5
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSFlexGridLib.MSFlexGrid grdGrades 
      Height          =   3975
      Left            =   240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1560
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   7011
      _Version        =   393216
      FocusRect       =   2
      HighLight       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "S t u d e n t   G r a d e s"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   360
      Width           =   3825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "All grades are over 100."
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
      Left            =   720
      TabIndex        =   10
      Top             =   840
      Width           =   2955
   End
End
Attribute VB_Name = "frmGradingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim updateTab As Boolean ' Update the database ? (true, false)

Private Sub cmdOK_Click()
frmTeacherMenu.Show
Unload Me
End Sub

'Show report in Excel
Private Sub cmdReport_Click()

'Initialize Excel Objects
Set xlsApp = CreateObject("Excel.Application")
Set xlsWorkbook = xlsApp.Workbooks.Add
Set xlsWorksheet = xlsWorkbook.Sheets("Sheet2")
xlsWorksheet.Delete
Set xlsWorksheet = xlsWorkbook.Sheets("Sheet3")
xlsWorksheet.Delete
Set xlsWorksheet = xlsWorkbook.Sheets("Sheet1")

'Fill in the grades
With grdGrades
    For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
            xlsWorksheet.Cells(i + 7, j + 1) = .TextMatrix(i, j)
        Next j
    Next i
End With


'Format the data
With xlsWorksheet.Range("A7", "N7")
.VerticalAlignment = xlVAlignCenter
.Font.Bold = True
.EntireColumn.AutoFit
End With


'Fill in report information
With xlsWorksheet
    .Name = "Report"
    .Cells(4, 1) = "Teacher: "
    .Cells(4, 3) = lblTeacher
    .Cells(5, 1) = "Course: "
    .Cells(5, 3) = lblCourse
'    .Cells.Font.Size = 12
End With

frmToolbar.Show
Unload Me
frmToolbar.SetFocus
End Sub

Private Sub Form_Load()
lblTeacher = strUsername
lblCourse = getTitle(strSelectedCourse) & " - " & strSelectedCourse
Call UpdateGradeList
End Sub

Private Sub grdGrades_Click()
txtGrade.Visible = False
End Sub

'when user double click on a cell, allow input
Private Sub grdGrades_dblClick()
'place the textbox on the selected cell
With grdGrades
    If .ColSel <> .Cols - 2 Then
        txtGrade.Left = .Left + .ColWidth(0) + .ColWidth(1) * (.ColSel - .LeftCol) + 80
        txtGrade.Top = .Top + .RowHeight(0) + .RowHeight(1) * (.RowSel - .TopRow) + 80
        txtGrade.Visible = True
        txtGrade.text = Val(.TextMatrix(.RowSel, .ColSel))
        txtGrade.SetFocus
        SendKeys "{Home}+{End}"
        updateTab = True
        
        'delete totals
        If .ColSel <> .Cols - 1 Then
        For i = 1 To .Rows - 1
        .TextMatrix(i, .Cols - 2) = ""
        Next i
        
        strSQL = "UPDATE [" & strUsername & strSelectedCourse & "] " & _
                 "SET [" & .TextMatrix(0, .Cols - 2) & "] = Null"
        objDBConnection.Execute (strSQL)
        End If
    End If
End With
End Sub

'when the user finishes his input, update the database
Private Sub grdGrades_LeaveCell()
If updateTab = True Then
    'check if it's a valid grade
    If IsNumeric(txtGrade) And Val(txtGrade) >= 0 And Val(txtGrade) <= 100 Then
        Call UpdateSelectedGrade
    Else
        MsgBox "The entry must be an integer between 0 and 100.", vbInformation, "Grading System - Information"
        grdGrades.SetFocus
        txtGrade.Visible = False
    End If
End If
updateTab = False
End Sub

'add the new entry to the database
Private Sub UpdateSelectedGrade()
With grdGrades
strSQL = "UPDATE [" & strUsername & strSelectedCourse & "] " & _
         "SET [" & .TextMatrix(0, .ColSel) & "] = '" & txtGrade & "'" & _
         " WHERE StudentID = '" & Mid$(.TextMatrix(.RowSel, 0), Len(.TextMatrix(.RowSel, 0)) - 7, 7) & "'"
objDBConnection.Execute (strSQL)
End With
Call UpdateGradeList
End Sub

'Refresh the grades list
Private Sub UpdateGradeList()
'Get the grades table from the database
strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "]"
Set objDBRecordset = objDBConnection.Execute(strSQL)

With grdGrades
    .Rows = 2
    .Cols = objDBRecordset.Fields.Count - 2
    .ColWidth(0) = 2600
    .RowHeight(0) = 320
    For i = 1 To .Cols - 1
    .ColWidth(i) = 1500
    Next i
    
    'fill in the student names
    .TextMatrix(0, 0) = "Student"
    strSQL = "SELECT FirstName, LastName, Students.StudentID FROM Students, [" & strUsername & strSelectedCourse & "] StudentCourse WHERE Students.StudentID = StudentCourse.StudentID"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    While Not objDBRecordset.EOF
    .TextMatrix(.Rows - 1, 0) = objDBRecordset("FirstName") & " " & objDBRecordset("LastName") & " (" & objDBRecordset("StudentID") & ")"
    objDBRecordset.MoveNext
    .Rows = .Rows + 1
    Wend
    .Rows = .Rows - 1
    strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "]"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    
    'fill in test names
    .TextMatrix(0, .Cols - 2) = objDBRecordset.Fields(1).Name
    .TextMatrix(0, .Cols - 1) = objDBRecordset.Fields(2).Name
    For i = 1 To .Cols - 3
       .TextMatrix(0, i) = objDBRecordset.Fields(i + 4).Name
    Next i
    
    'fill in grades
    For i = 1 To .Rows - 1
    strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "] WHERE StudentID = '" & Mid$(.TextMatrix(i, 0), Len(.TextMatrix(i, 0)) - 7, 7) & "'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    For j = 1 To .Cols - 3
    If Not IsNull(objDBRecordset.Fields.Item(j + 4)) Then
    .TextMatrix(i, j) = objDBRecordset.Fields.Item(j + 4)
    End If
    Next j
    If Not IsNull(objDBRecordset.Fields.Item(2)) Then
    .TextMatrix(i, .Cols - 1) = objDBRecordset.Fields.Item(2)
    End If
    If Not IsNull(objDBRecordset.Fields.Item(1)) Then .TextMatrix(i, .Cols - 2) = objDBRecordset.Fields.Item(1)
    Next i
End With
End Sub

Private Sub cmdCalculate_Click()
Dim Total As Single ' Each Student's Total

With grdGrades
    'check if all the grades have been entered
    For i = 1 To .Rows - 1
        For j = 1 To .Cols - 3
            If .TextMatrix(i, j) = "" Then
            MsgBox "Please enter a grade for student " & .TextMatrix(i, 0) & " under the " & .TextMatrix(0, j) & " column.", vbInformation, "Grading System - Information"
            Exit Sub
            End If
        Next j
    Next i

    'Calculate the totals
    strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "] WHERE StudentID = '0000000'"
    Set objDBRecordset = objDBConnection.Execute(strSQL)
    
    For i = 1 To .Rows - 1
        'For every student, calculate Total
        Total = 0
        For j = 1 To .Cols - 3
           Total = Total + _
             (Val(.TextMatrix(i, j)) * objDBRecordset(.TextMatrix(0, j)) / 100)
        Next j
        .TextMatrix(i, .Cols - 2) = Total
        strSQL = "UPDATE [" & strUsername & strSelectedCourse & "] " & _
                 "SET Total = '" & Total & "' WHERE StudentID = '" & Mid$(.TextMatrix(i, 0), Len(.TextMatrix(i, 0)) - 7, 7) & "'"
        objDBConnection.Execute (strSQL)
    Next i
  
End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

