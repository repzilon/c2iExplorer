VERSION 5.00
Begin VB.Form FrmMenu 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicFond 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   1320
      Picture         =   "FrmMenu.frx":0000
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblProp 
      Caption         =   "Propriétés"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "FrmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim lx As Long, lY As Long, lWidth As Long, lHeight As Long
    
    lWidth = PicFond.Width
    lHeight = PicFond.Height
    For lx = 0 To Int(ScaleWidth / lWidth)
        For lY = 0 To ScaleHeight / lWidth
            PaintPicture PicFond.Picture, lx * lWidth, lY * lHeight, lWidth, lHeight, 0, 0
        Next
    Next

End Sub
