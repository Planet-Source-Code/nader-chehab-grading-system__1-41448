VERSION 5.00
Begin VB.Form frmAdminMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administrator's Menu"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7005
   ControlBox      =   0   'False
   Icon            =   "frmAdminMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   358
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   467
   StartUpPosition =   2  'CenterScreen
   Begin GradingSystem.HoverCommand cmdExit 
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      Caption         =   "&Exit"
      Style           =   3
      Picture         =   "frmAdminMenu.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdCourses 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Caption         =   "Courses List"
      Style           =   16
      Picture         =   "frmAdminMenu.frx":09A3
      HovPicture      =   "frmAdminMenu.frx":0AD7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdStudents 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Caption         =   "Students List"
      Style           =   16
      Picture         =   "frmAdminMenu.frx":0C0B
      HovPicture      =   "frmAdminMenu.frx":0D3F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdTeachers 
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   2520
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Caption         =   "Teachers List (+ assign courses )"
      Style           =   16
      Picture         =   "frmAdminMenu.frx":0E73
      HovPicture      =   "frmAdminMenu.frx":0FA7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdRegistrations 
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   3000
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Caption         =   "Registrations List"
      Style           =   16
      Picture         =   "frmAdminMenu.frx":10DB
      HovPicture      =   "frmAdminMenu.frx":120F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdModifyPassword 
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3600
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   873
      Caption         =   "Admin Password"
      Style           =   16
      Picture         =   "frmAdminMenu.frx":1343
      HovPicture      =   "frmAdminMenu.frx":1477
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin GradingSystem.HoverCommand cmdLogout 
      Height          =   615
      Left            =   3720
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1085
      Caption         =   "&Logout"
      Style           =   3
      Picture         =   "frmAdminMenu.frx":15AB
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   4680
      TabIndex        =   0
      Top             =   4680
      Width           =   75
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "A d m i n i s t r a t o r ' s  M e n u"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   5
      Left            =   2295
      TabIndex        =   9
      Top             =   960
      Width           =   2295
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
      Left            =   1275
      TabIndex        =   8
      Top             =   360
      Width           =   4455
   End
End
Attribute VB_Name = "frmAdminMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCourses_Click()
frmCoursesList.Show
Unload Me
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdLogout_Click()
strUsername = ""
frmLogin.Show
Unload Me
End Sub

Private Sub cmdModifyPassword_Click()
Me.Enabled = False
frmAdminChangePassword.Show
End Sub

Private Sub cmdRegistrations_Click()
frmRegistrations.Show
Unload Me
End Sub

Private Sub cmdStudents_Click()
frmStudentsList.Show
Unload Me
End Sub

Private Sub cmdTeachers_Click()
frmTeachersList.Show
Unload Me
End Sub

'Drag the Form from anywhere
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
'Or use:  SendMessage hwnd, WM_SYSCOMMAND, &HF012&, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub



