Attribute VB_Name = "modAnimation"
'Expoding Animation
'From http://cuinl.tripod.com

#If Win16 Then
Type RECT
Left As Integer
Top As Integer
Right As Integer
Bottom As Integer
End Type
#Else
Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
#End If

#If Win16 Then
Declare Sub GetWindowRect Lib "User" (ByVal hWnd As Integer, lpRect As RECT)
Declare Function GetDC Lib "User" (ByVal hWnd As Integer) As Integer
Declare Function ReleaseDC Lib "User" (ByVal hWnd As Integer, ByVal hdc As _
Integer) As Integer
Declare Sub SetBkColor Lib "GDI" (ByVal hdc As Integer, ByVal crColor As Long)
Declare Sub Rectangle Lib "GDI" (ByVal hdc As Integer, ByVal X1 As Integer, _
ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
Declare Function CreateSolidBrush Lib "GDI" (ByVal crColor As Long) As Integer
Declare Function SelectObject Lib "GDI" (ByVal hdc As Integer, ByVal hObject _
As Integer) As Integer
Declare Sub DeleteObject Lib "GDI" (ByVal hObject As Integer)
#Else
Declare Function GetWindowRect Lib "User32" (ByVal hWnd As Long, _
lpRect As RECT) As Long
Declare Function GetDC Lib "User32" (ByVal hWnd As Long) As Long
Declare Function ReleaseDC Lib "User32" (ByVal hWnd As Long, ByVal _
hdc As Long) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal _
crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, _
ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "User32" (ByVal hdc As Long, ByVal hObject _
As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
#End If

Sub ExplodeForm(f As Form, Movement As Integer)
Dim myRect As RECT
Dim formWidth%, formHeight%, i%, X%, Y%, cx%, cy%
Dim TheScreen As Long
Dim Brush As Long
GetWindowRect f.hWnd, myRect
formWidth = (myRect.Right - myRect.Left)
formHeight = myRect.Bottom - myRect.Top
TheScreen = GetDC(0)
Brush = CreateSolidBrush(f.BackColor)
For i = 1 To Movement
cx = formWidth * (i / Movement)
cy = formHeight * (i / Movement)
X = myRect.Left + (formWidth - cx) / 2
Y = myRect.Top + (formHeight - cy) / 2
Rectangle TheScreen, X, Y, X + cx, Y + cy
Next i
X = ReleaseDC(0, TheScreen)
DeleteObject (Brush)
End Sub


