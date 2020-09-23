VERSION 5.00
Begin VB.Form frmToolbar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Grading System - Excel Toolbar"
   ClientHeight    =   990
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3060
   ControlBox      =   0   'False
   Icon            =   "frmToolbar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   990
   ScaleWidth      =   3060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Report"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close Report"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'taken from AllAPI.net
SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub cmdClose_Click()
'exit without saving (could be coded better)
On Error Resume Next
xlsWorkbook.SaveAs App.Path & "\temp"
xlsApp.Quit
Set xlsWorkbook = Nothing
Set xlsWorksheet = Nothing
Set xlsApp = Nothing
Kill App.Path & "\temp.xls"
frmGradingTable.Show
Unload Me
End Sub

Private Sub cmdPrint_Click()
Set xlsWorksheet = xlsWorkbook.ActiveSheet
xlsWorksheet.PrintPreview
End Sub

Private Sub Form_Load()
xlsApp.Visible = True
Me.Top = 9000
Me.Left = 11000
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
