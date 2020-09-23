VERSION 5.00
Begin VB.Form frmTeacherCourseProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course's Properties - Step 1"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Notice"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   120
         Picture         =   "frmNumberOfExams.frx":0000
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblNote 
         Caption         =   "This is the first time you view this course.  Please specify the course's properties."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmTeacherCourseProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
