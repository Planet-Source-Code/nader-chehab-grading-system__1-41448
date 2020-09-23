VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTeacherCourseProperties3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Properties - Evalutaion Chart"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8910
   ControlBox      =   0   'False
   Icon            =   "frmTeacherCourseProperties3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   8910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Absences"
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
      Left            =   480
      TabIndex        =   73
      Top             =   7440
      Width           =   7935
      Begin VB.CommandButton cmdPlusAb 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6840
         TabIndex        =   75
         Top             =   270
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusAb 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6480
         TabIndex        =   74
         Top             =   270
         Width           =   255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Maximum allowed Absences for this course:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   1200
         TabIndex        =   77
         Top             =   300
         Width           =   4455
      End
      Begin VB.Label lblMaxAb 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
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
         Left            =   5760
         TabIndex        =   76
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "&Validate"
      Height          =   735
      Left            =   7200
      Picture         =   "frmTeacherCourseProperties3.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   735
      Left            =   480
      Picture         =   "frmTeacherCourseProperties3.frx":097A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Evalutaion Chart"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   480
      TabIndex        =   7
      Top             =   2040
      Width           =   7935
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   0
         Left            =   2880
         TabIndex        =   61
         Top             =   600
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   7080
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   6720
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   7080
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   6720
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   7080
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   6720
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   7080
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   6720
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   7080
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6720
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   7080
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   6720
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7080
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6720
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1680
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   7080
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6720
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6720
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdMinusTest 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6720
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton cmdPlusTest 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   7080
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   1
         Left            =   2880
         TabIndex        =   62
         Top             =   960
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   2
         Left            =   2880
         TabIndex        =   63
         Top             =   1320
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   3
         Left            =   2880
         TabIndex        =   64
         Top             =   1680
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   4
         Left            =   2880
         TabIndex        =   65
         Top             =   2040
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   5
         Left            =   2880
         TabIndex        =   66
         Top             =   2400
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   6
         Left            =   2880
         TabIndex        =   67
         Top             =   2760
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   7
         Left            =   2880
         TabIndex        =   68
         Top             =   3120
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   8
         Left            =   2880
         TabIndex        =   69
         Top             =   3480
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin ComctlLib.ProgressBar barTest 
         Height          =   300
         Index           =   9
         Left            =   2880
         TabIndex        =   70
         Top             =   3840
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   529
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Course Total:"
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
         Left            =   4920
         TabIndex        =   72
         Top             =   4477
         Width           =   1665
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
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
         Index           =   7
         Left            =   120
         TabIndex        =   71
         Top             =   3120
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   2040
         TabIndex        =   59
         Top             =   3840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   2520
         TabIndex        =   58
         Top             =   3840
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   120
         TabIndex        =   55
         Top             =   3855
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   2040
         TabIndex        =   54
         Top             =   3480
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   2520
         TabIndex        =   53
         Top             =   3480
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   50
         Top             =   3495
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   2040
         TabIndex        =   49
         Top             =   3120
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   2520
         TabIndex        =   48
         Top             =   3120
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   45
         Top             =   2775
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   2040
         TabIndex        =   44
         Top             =   2760
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   2520
         TabIndex        =   43
         Top             =   2760
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   40
         Top             =   2415
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2040
         TabIndex        =   39
         Top             =   2400
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2520
         TabIndex        =   38
         Top             =   2400
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   35
         Top             =   2055
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2040
         TabIndex        =   34
         Top             =   2040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2520
         TabIndex        =   33
         Top             =   2040
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   30
         Top             =   1695
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2040
         TabIndex        =   29
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2520
         TabIndex        =   28
         Top             =   1680
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   615
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2040
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2520
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         TabIndex        =   22
         Top             =   4410
         Width           =   615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Course Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   11
         Left            =   5040
         TabIndex        =   21
         Top             =   8175
         Width           =   1410
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6600
         TabIndex        =   20
         Top             =   8130
         Width           =   615
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2520
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2040
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   17
         Top             =   975
         Visible         =   0   'False
         Width           =   90
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2520
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label lblTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2040
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   14
         Top             =   1335
         Visible         =   0   'False
         Width           =   90
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   7935
      Begin VB.Label lblCourse0 
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
         Left            =   600
         TabIndex        =   6
         Top             =   765
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
         Left            =   480
         TabIndex        =   5
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
         Left            =   1680
         TabIndex        =   4
         Top             =   720
         Width           =   5535
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
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   5535
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please choose a percentage for every test:"
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
      Left            =   600
      TabIndex        =   60
      Top             =   1560
      Width           =   5310
   End
End
Attribute VB_Name = "frmTeacherCourseProperties3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
frmTeacherMenu.Show
Unload Me
End Sub

Private Sub cmdMinusAb_Click()
If Val(lblMaxAb) > 0 Then lblMaxAb = Val(lblMaxAb) - 1
End Sub

Private Sub cmdPlusAb_Click()
If Val(lblMaxAb) < 15 Then lblMaxAb = Val(lblMaxAb) + 1
End Sub


Private Sub cmdValidate_Click()
If Val(lblTotal.Caption) < 100 Then
MsgBox "Course total must reach 100.", vbInformation, "Grading System - Information"
Exit Sub
End If
For i = 0 To objDBRecordset.Fields.Count - 6
strSQL = "UPDATE [" & strUsername & strSelectedCourse & "] " & _
         "SET [" & lblLabels(i) & "] = '" & lblTest(i) & "' " & _
         "WHERE StudentID = '0000000'"
objDBConnection.Execute (strSQL)
Next i
'Update the Abcenses column
strSQL = "UPDATE [" & strUsername & strSelectedCourse & "] " & _
         "SET Absences = '" & lblMaxAb & "' WHERE StudentID = '0000000'"
objDBConnection.Execute (strSQL)

MsgBox "Evaluation Chart has been updated succefully.", vbInformation, "Grading System - Information"
frmTeacherMenu.Show
Unload Me
End Sub

Private Sub lblTest_Change(Index As Integer)
barTest(Index).value = Val(lblTest(Index).Caption)
End Sub

Private Sub cmdMinusTest_Click(Index As Integer)
If Val(lblTest(Index).Caption) > 0 And Val(lblTotal.Caption) > 0 Then
lblTest(Index).Caption = Val(lblTest(Index).Caption) - 5
lblTotal = Val(lblTotal.Caption) - 5
End If
End Sub

Private Sub cmdPlusTest_Click(Index As Integer)
If Val(lblTest(Index).Caption) < 100 And Val(lblTotal.Caption) < 100 Then
lblTest(Index).Caption = lblTest(Index).Caption + 5
lblTotal = Val(lblTotal.Caption) + 5
End If
End Sub

Private Sub Form_Load()
Dim strTests As String
lblTotal = 0
lblTeacher = strUsername
lblCourse = getTitle(strSelectedCourse) & " - " & strSelectedCourse
strSQL = "SELECT * FROM [" & strUsername & strSelectedCourse & "] WHERE StudentID = '0000000'"
Set objDBRecordset = objDBConnection.Execute(strSQL)
If Not IsNull(objDBRecordset.Fields.Item(2)) Then lblMaxAb = objDBRecordset.Fields.Item(2)
For j = 0 To objDBRecordset.Fields.Count - 6
lblLabels(j).Visible = True
lblLabels(j).Caption = objDBRecordset.Fields(j + 5).Name
lblTest(j).Visible = True
barTest(j).Visible = True
lblPercent(j).Visible = True
cmdPlusTest(j).Visible = True
cmdMinusTest(j).Visible = True
If Not IsNull(objDBRecordset.Fields.Item(j + 5)) Then
lblTest(j) = objDBRecordset.Fields.Item(j + 5)
lblTotal = Val(lblTotal) + lblTest(j)
End If
Next j
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Activate()
If AlwaysOnTop Then If AlwaysOnTop Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
