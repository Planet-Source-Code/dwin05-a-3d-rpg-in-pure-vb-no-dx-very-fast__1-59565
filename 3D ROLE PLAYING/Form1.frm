VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Loading"
   ClientHeight    =   1230
   ClientLeft      =   75
   ClientTop       =   360
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1230
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3720
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   585
      ScaleWidth      =   585
      TabIndex        =   1
      Top             =   4560
      Width           =   615
      Begin VB.Shape Shape1 
         Height          =   1155
         Left            =   5280
         Top             =   7680
         Width           =   1155
      End
   End
   Begin VB.PictureBox backdrop 
      Height          =   735
      Left            =   2880
      Picture         =   "Form1.frx":063F
      ScaleHeight     =   675
      ScaleWidth      =   555
      TabIndex        =   0
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'THIS WAS CREATED AND COPYRIGHT 2005 BY DARWIN YU
'
'I WROTE EVERY LINE HERE EXCEPT FOR THE TRANSPARENT BITBLT FUNCTIONS WHICH IS USED FOR MAKING THE HANDS TRANSPARENT
'
'IF YOU'D LIKE TO USE PARTS OF THE CODE OR IMPROVE IT (GLADLY ACCEPTED), PLEASE EMAIL
'ME AT ibnzrokr@yahoo.com
'
'IF YOU IMPROVE THE CODE SIGNIFICANTLY, I WILL ADD YOUR NAME TO THE CREDITS SECTION
'ONCE THIS IS COMPLETE!
'
'I NEED SOMEONE TO FIGURE OUT A WAY TO ROTATE THE CAMERA...RIGHT NOW YOU CAN ONLY STRAFF
'
'IF YOU CAN FIND A WAY TO ROTATE THE MAIN MAP PICTURE QUICKLY, PLEASE NOTIFY ME!!!
'BECAUSE THIS IS PROBABLY THE BEST WAY TO ROTATE THE CAMERA. SO WHOEVER CAN ROTATE A LARGE
'IMAGE QUICKLY, YOU HAVE SIGNIFICANTLY HELPED ME OUT!!
'
'PLEASE VOTE! - REMEMBER, THIS IS A EARLY DRAFT - I DID THIS IN 2 HOURS.
'
'ALL TEXTURES ARE MINE - EXCEPT FOR THE HANDS (FROM HEXEN)
'
'THANKS!


''I will be making this into an upcoming RPG


Option Explicit

Private halfhEight               As Long 'half the form's height - used for resizing
Private Widthplus                As Long 'width of form


Private Sub DrawFrame(VerticalDisplacement As Long) 'GRABS DATA FROM PICTURE AND DRAWS THE WORLD
On Error Resume Next
  With picthedisplay 'GIVES THE SCENE A 3D LOOK
    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 260 + ((mY - halfhEight) * -1), Widthplus, 580 + VerticalDisplacement, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height 'DRAW THE FARTHEST POSSIBLE VIEW
    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 840 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, 700, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height 'DRAW A CLOSER VIEW
    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 960 + VerticalDisplacement + ((mY - halfhEight) * -1), Widthplus, halfhEight, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height 'DRAW THE CLOSEST VIEW
  End With
  
'    With picthedisplay
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 260, Widthplus, 580 + VerticalDisplacement, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 840 + VerticalDisplacement, Widthplus, 700, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'    picthedisplay.PaintPicture Picture1.Picture, 0, halfhEight + 960 + VerticalDisplacement, Widthplus, halfhEight, Shape1.Left, Shape1.Top, Shape1.Width, Shape1.Height
'  End With 'picThedisplay

End Sub



Private Sub Form_Load()
'LOAD GRAPHICS
Picture1.Picture = LoadPicture(App.Path & "\data\maps\main\world\w1.jpg")
backdrop.Picture = LoadPicture(App.Path & "\data\maps\main\world\mountain.jpg")
'___________________________________________


picthedisplay.Show 'DISPLAYS MAIN FORM
Form1.Hide 'HIDE THUS FORM
halfhEight = picthedisplay.Height / 2 'SET THE VALUE OF HALFHEIGHT
Widthplus = picthedisplay.Width + 120 'SET VALUE OF WIDTHPLUS
mY = halfhEight 'CENTERS VIEW
picthedisplay.PaintPicture backdrop.Picture, 0, 0, Widthplus, halfhEight + 360 'DRAW BG

DrawFrame 1 'DRAW THE WORLD
End Sub
