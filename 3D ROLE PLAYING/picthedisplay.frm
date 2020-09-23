VERSION 5.00
Begin VB.Form picthedisplay 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Use Arrow Keys to Move.    Use Mouse to Look Around!  ESC TO QUIT"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9975
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   9975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   2760
      Picture         =   "picthedisplay.frx":0000
      ScaleHeight     =   795
      ScaleWidth      =   915
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer makeMOVE 
      Interval        =   1
      Left            =   1920
      Top             =   1560
   End
End
Attribute VB_Name = "picthedisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim halfhEight As Long
Dim Widthplus As Long

Dim sT As Long
Dim sL As Long
Dim sW As Long
Dim sH As Long

Dim mCOMMAND As String
Dim tmrBOOST As Integer

'ROTATATION
'Dim rANGLE As Long
'______________________


'TRANSPARENT BITBLT DIMS
Dim hdcSrc As Long
Dim hdcWidth As Integer
Dim hdcHeight As Integer
Dim TransparentColor As Long
'__________________________________

Sub GetDimensions() 'GET THE DIMENSIONS OF THE WEAPON PICTURE
Dim dbitmap As Bitmap
Dim hbitmap As Long
Dim filename As String

filename = App.Path & "\weapon1.bmp" 'LOCATION OF WEAPON PIC
hbitmap = LoadImage(ByVal 0&, filename, 0, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)
' Get bitmap information
GetObject hbitmap, Len(dbitmap), dbitmap

hdcHeight = dbitmap.bmHeight
hdcWidth = dbitmap.bmWidth
DeleteObject hbitmap
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
'MOVEMENTS

Case vbKeyUp
    mCOMMAND = "foward"

Case vbKeyDown
    mCOMMAND = "backward"
    
Case vbKeyRight 'STRAFF
    mCOMMAND = "right"

Case vbKeyLeft 'STRAFF
    mCOMMAND = "left"
    
Case vbKeyEscape
    End

End Select

'makeMOVE.Enabled = True

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'makeMOVE.Enabled = False
mCOMMAND = "" 'AFTER USER RELEASES KEY, RESET COMMANDS
End Sub

Private Sub Form_Load()
'TRANSPARENT BITBLT DIMS
Dim TransparentColor As Long
    hdcSrc = CreateCompatibleDC(Picture1.hdc)
    SelectObject hdcSrc, LoadPicture(App.Path & "\weapon1.bmp") 'LOCATION OF WEAPON PIC
    GetDimensions
'_________________________________________________________________


halfhEight = picthedisplay.Height / 2 'THIS IS A REPEAT FROM FORM1
Widthplus = picthedisplay.Width + 120 'EXPLAINED IN FORM1 [LOAD]
mY = halfhEight

'RECORDS STARTING LOCATION ON MAP

sT = Form1.Shape1.Top
sL = Form1.Shape1.Left
sW = Form1.Shape1.Width
sH = Form1.Shape1.Height
'___________________________________

weaponY = 450 'STARTING Y POS OF THE "HANDS"

picthedisplay.PaintPicture Form1.backdrop.Picture, 0, 0, Widthplus, halfhEight + 360 'DRAWS BG
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

mY = Y

   ' picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 260 + ((mY - halfhEight) * -1), Widthplus, 580 + VerticalDisplacement, Form1.Shape1.Left, Form1.Shape1.Top, Form1.Shape1.Width, Form1.Shape1.Height
   ' picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 840 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, 700, Form1.Shape1.Left, Form1.Shape1.Top, Form1.Shape1.Width, Form1.Shape1.Height
   ' picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 960 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, halfhEight, Form1.Shape1.Left, Form1.Shape1.Top, Form1.Shape1.Width, Form1.Shape1.Height


End Sub


Private Sub DrawFrame(VerticalDisplacement As Long) 'DRAWS IN 3D
On Error Resume Next
  With picthedisplay
    picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 260 + ((mY - halfhEight) * -1), Widthplus, 580 + VerticalDisplacement, sL, sT, sW, sH
    picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 840 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, 700, sL, sT, sW, sH
    picthedisplay.PaintPicture Form1.Picture1.Picture, 0, halfhEight + 960 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, halfhEight + (mY - halfhEight), sL, sT, sW, sH

  End With 'picThedisplay
  
'    With picthedisplay
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 260, Widthplus, 580 + VerticalDisplacement, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 840 + VerticalDisplacement, Widthplus, 700, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 960 + VerticalDisplacement, Widthplus, halfhEight, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'  End With 'picThedisplay

End Sub


Private Sub makeMOVE_Timer() 'EXECUTES THE COMMANDS GIVEN BY THE KEYBOARD


Select Case mCOMMAND

Case "foward"
    sT = sT - 60

Case "backward"
    sT = sT + 60

Case "right"
    sL = sL + 60

Case "left"
    sL = sL - 60


End Select


picthedisplay.Cls 'CLEAR DRAWING AREA
''REDRAW''
picthedisplay.PaintPicture Form1.backdrop.Picture, 0, (halfhEight + 260 + ((mY - halfhEight) * -1)) - halfhEight + 360, Widthplus, halfhEight + 360
''REDRAW''
DrawFrame 1


'''Animates the hands
If weaponY < 435 Then
    weaponY = weaponY + 5
Else
    weaponY = weaponY - 5
End If
'___________________________

'MAKES THE HANDS TRANSPARENT
TransparentColor = GetPixel(hdcSrc, 0, 0)
'TransparentBlt picDest.hdc, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, hdcSrc, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, TransparentColor
TransparentBlt picthedisplay.hdc, 100, weaponY, hdcWidth, hdcHeight, hdcSrc, 0, 0, hdcWidth, hdcHeight, TransparentColor
'_____________________________

End Sub


