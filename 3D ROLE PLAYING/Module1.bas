Attribute VB_Name = "Module1"
'I DO NOT TAKE CREDIT FOR THIS PORTION.
'I BORROWED SOMEONE'S TRANSPARENT BITBLT CODE TO MAKE THE
'HANDS TRANSPARENT.
'THIS IS THE ONLY PART THAT ISNT MINE.

Declare Function TransparentBlt Lib "msimg32" _
                (ByVal hdcDest As Long, ByVal nXOriginDest As Long, _
                  ByVal nYOriginDest As Long, ByVal nWidthDest As Long, _
                  ByVal nHeightDest As Long, ByVal hdcSrc As Long, _
                  ByVal nXOriginSrc As Long, ByVal nYOriginSrc As Long, _
                 ByVal nWidthSrc As Long, ByVal nHeightSrc As Long, _
                 ByVal crTransparent As Long) As Long

Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Type Bitmap '14 bytes
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
    Public Const LR_CREATEDIBSECTION = &H2000
    Public Const LR_LOADFROMFILE = &H10




'___________________________________________________________________________


Public mY As Long


Public weaponY As Long 'stores the vertical coords of the hands
