VERSION 5.00
Begin VB.Form frmCourses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Course Records"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCourses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
grdCourses.ColWidth(0) = 500
grdCourses.ColWidth(1) = 1200
grdCourses.ColWidth(2) = 2390
grdCourses.ColWidth(3) = 660
End Sub
