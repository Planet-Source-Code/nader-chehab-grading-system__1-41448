VERSION 5.00
Begin VB.Form frmFirst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "First Run"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7470
   Icon            =   "frmFirst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRetype 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3600
      Width           =   2415
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3240
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Notice"
      Height          =   1095
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   6855
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         Picture         =   "frmFirst.frx":08CA
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblNote 
         Caption         =   $"frmFirst.frx":0E6F
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
   End
   Begin GradingSystem.HoverCommand cmdCreate 
      Height          =   735
      Left            =   2408
      TabIndex        =   2
      Top             =   4800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1296
      Caption         =   "&Create Database"
      Style           =   3
      Picture         =   "frmFirst.frx":0F09
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
   Begin VB.Label lblNote 
      Caption         =   "Retype Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   9
      Top             =   3735
      Width           =   1935
   End
   Begin VB.Label lblNote 
      Caption         =   "Administrator's Password:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   480
      TabIndex        =   8
      Top             =   3255
      Width           =   2535
   End
   Begin VB.Label lblNote 
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3188
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Grading System"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   1508
      TabIndex        =   3
      Top             =   240
      Width           =   4455
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCreate_Click()

On Error GoTo CreateDatabase_Error
    
    If Len(txtPassword) < 6 Then Err.Raise 1001
    If txtPassword <> txtRetype Then Err.Raise 1002
    
    Dim objcat As ADOX.Catalog
    Dim strConnection As String
    
    strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                    "Data Source=" & App.Path & "\GradingSystem.mdb;" & _
                     "Jet OLEDB:Database Password = " & txtPassword.text
    
    Set objcat = New ADOX.Catalog
    objcat.Create strConnection
    objcat.ActiveConnection = strConnection
    
    WriteToFile Transform(txtPassword)
    
    Call ConnectToDatabase
         
    'Create Students table
    strSQL = "CREATE TABLE Students " & _
             "(StudentID VARCHAR(7) PRIMARY KEY, " & _
             "FirstName VARCHAR(35), LastName VARCHAR(35))"
    objDBConnection.Execute (strSQL)
    
    'Create Courses table
    strSQL = "CREATE TABLE Courses " & _
             "(CourseID VARCHAR(10) PRIMARY KEY, " & _
             "CourseTitle VARCHAR(50), NumberOfHours VARCHAR(2))"
    objDBConnection.Execute (strSQL)
        
    'Create Teachers table
    strSQL = "CREATE TABLE Teachers " & _
             "(TeacherID VARCHAR(4) PRIMARY KEY, " & _
             "FirstName VARCHAR(35), LastName VARCHAR(35), " & _
             "Username VARCHAR(15), [Password] VARCHAR(12))"
    objDBConnection.Execute (strSQL)
    
    'Create TeacherCourse table
    strSQL = "CREATE TABLE TeacherCourse " & _
             "(TeacherID VARCHAR(4), " & _
             "CourseID VARCHAR(10), " & _
             "CONSTRAINT TeacherCourse_PK PRIMARY KEY(TeacherID, CourseID), " & _
             "CONSTRAINT TeacherID_FK FOREIGN KEY (TeacherID) REFERENCES Teachers(TeacherID), " & _
             "CONSTRAINT CourseID_FK FOREIGN KEY (CourseID) REFERENCES Courses(CourseID))"
    objDBConnection.Execute (strSQL)
    
   
    'Create Administrator
    strSQL = "INSERT INTO Teachers VALUES('0001', 'Administrator', 'Administrator', 'Administrator', '" & Transform(txtPassword) & "')"
    objDBConnection.Execute (strSQL)
    frmLogin.Show
    Unload Me
    
Exit Sub

CreateDatabase_Error:
Select Case Err.Number
    Case 1001
    MsgBox "The password must be at least 6 characters", vbInformation, "Grading System"
    txtRetype = ""
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    
    Case 1002
    MsgBox "Retype password correctly.", vbInformation, "Grading System"
    txtRetype = ""
    txtPassword.SetFocus
    SendKeys "{Home}+{End}"
    Exit Sub
    
    Case Else
    MsgBox "Error " & CStr(Err.Number) & ": " & Err.Description, vbExclamation, App.Title
End Select

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
'If KeyAscii < vbKeyA + 32 And KeyAscii < vbKeyA Or KeyAscii > vbKeyZ + 32 And KeyAscii > vbKeyZ Then KeyAscii = 0
End Sub
