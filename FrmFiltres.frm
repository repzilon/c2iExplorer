VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmFiltres 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Filtres d'affichage"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicCont 
      AutoRedraw      =   -1  'True
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   915
      TabIndex        =   1
      Top             =   0
      Width           =   975
      Begin c2iExplorer.TButton TBConstantes 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":0000
         ImageAffichage  =   "FrmFiltres.frx":055E
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBVariable 
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":0ABC
         ImageAffichage  =   "FrmFiltres.frx":101A
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBPropriete 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":1578
         ImageAffichage  =   "FrmFiltres.frx":1AD6
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBMethode 
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":2034
         ImageAffichage  =   "FrmFiltres.frx":2592
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBEvenement 
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":2AF0
         ImageAffichage  =   "FrmFiltres.frx":304E
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBAPI 
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":35AC
         ImageAffichage  =   "FrmFiltres.frx":3B0A
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBAnnul 
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":4068
         ImageAffichage  =   "FrmFiltres.frx":45C6
         CouleurCache    =   128
         Decalage        =   10
      End
      Begin c2iExplorer.TButton TBOK 
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   840
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         ImageCache      =   "FrmFiltres.frx":4B24
         ImageAffichage  =   "FrmFiltres.frx":5082
         CouleurCache    =   128
         Decalage        =   10
      End
   End
   Begin VB.PictureBox PicFond 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   3120
      Picture         =   "FrmFiltres.frx":55E0
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSComctlLib.ImageList ImgLstFilter 
      Left            =   360
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   17
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":6129
            Key             =   "variable"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":6687
            Key             =   "api"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":6BE5
            Key             =   "n-api"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":7143
            Key             =   "evenement"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":76A1
            Key             =   "methode"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":7BFF
            Key             =   "n-constante"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":815D
            Key             =   "n-evenement"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":86BB
            Key             =   "n-methode"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":8C19
            Key             =   "n-propriete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":9177
            Key             =   "n-variable"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":96D5
            Key             =   "propriete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmFiltres.frx":9C33
            Key             =   "constante"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmFiltres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private bAfficheConstanteT As Boolean
Private bAfficheVariableT As Boolean
Private bAfficheProprieteT As Boolean
Private bAfficheMethodeT As Boolean
Private bAfficheEvenementT As Boolean
Private bAfficheAPIT As Boolean



Private Sub Form_Load()
PicCont.Width = ScaleWidth
PicCont.Height = ScaleHeight
PicCont.Left = 0
PicCont.Top = 0

bAfficheConstanteT = bAfficheConstante
bAfficheVariableT = bAfficheVariable
bAfficheProprieteT = bAffichePropriete
bAfficheMethodeT = bAfficheMethode
bAfficheEvenementT = bAfficheEvenement
bAfficheAPIT = bAfficheAPI

If Not bAfficheConstante Then
    Set TBConstantes.ImageAffichage = ImgLstFilter.ListImages("n-constante").Picture
End If
If Not bAfficheVariable Then
    Set TBVariable.ImageAffichage = ImgLstFilter.ListImages("n-variable").Picture
End If
If Not bAffichePropriete Then
    Set TBPropriete.ImageAffichage = ImgLstFilter.ListImages("n-propriete").Picture
End If
If Not bAfficheMethode Then
    Set TBMethode.ImageAffichage = ImgLstFilter.ListImages("n-methode").Picture
End If
If Not bAfficheEvenement Then
    Set TBEvenement.ImageAffichage = ImgLstFilter.ListImages("n-evenement").Picture
End If
If Not bAfficheAPI Then
    Set TBAPI.ImageAffichage = ImgLstFilter.ListImages("n-api").Picture
End If

End Sub


Private Sub Form_Resize()
    Dim lx As Long, lY As Long, lWidth As Long, lHeight As Long
    
    lWidth = PicFond.Width
    lHeight = PicFond.Height
    For lx = 0 To Int(PicCont.ScaleWidth / lWidth)
        For lY = 0 To PicCont.ScaleHeight / lWidth
            PicCont.PaintPicture PicFond.Picture, lx * lWidth, lY * lHeight, lWidth, lHeight, 0, 0
        Next
    Next

End Sub


Private Sub TBAnnul_Click()
Unload Me
End Sub

Private Sub TBAPI_Click()
bAfficheAPIT = Not bAfficheAPIT
If Not bAfficheAPIT Then
    Set TBAPI.ImageAffichage = ImgLstFilter.ListImages("n-api").Picture
Else
    Set TBAPI.ImageAffichage = ImgLstFilter.ListImages("api").Picture
End If

End Sub

Private Sub TBConstantes_Click()

bAfficheConstanteT = Not bAfficheConstanteT
If Not bAfficheConstanteT Then
    Set TBConstantes.ImageAffichage = ImgLstFilter.ListImages("n-constante").Picture
Else
    Set TBConstantes.ImageAffichage = ImgLstFilter.ListImages("constante").Picture
End If

End Sub

Private Sub TBEvenement_Click()
bAfficheEvenementT = Not bAfficheEvenementT
If Not bAfficheEvenementT Then
    Set TBEvenement.ImageAffichage = ImgLstFilter.ListImages("n-evenement").Picture
Else
    Set TBEvenement.ImageAffichage = ImgLstFilter.ListImages("evenement").Picture
End If

End Sub

Private Sub TBMethode_Click()
bAfficheMethodeT = Not bAfficheMethodeT
If Not bAfficheMethodeT Then
    Set TBMethode.ImageAffichage = ImgLstFilter.ListImages("n-methode").Picture
Else
    Set TBMethode.ImageAffichage = ImgLstFilter.ListImages("methode").Picture
End If

End Sub

Private Sub TBOK_Click()
bAfficheConstante = bAfficheConstanteT
bAfficheVariable = bAfficheVariableT
bAffichePropriete = bAfficheProprieteT
bAfficheMethode = bAfficheMethodeT
bAfficheEvenement = bAfficheEvenementT
bAfficheAPI = bAfficheAPIT


c2iExplorer.Connection.objUD.Init True

Unload Me
End Sub

Private Sub TBPropriete_Click()
bAfficheProprieteT = Not bAfficheProprieteT
If Not bAfficheProprieteT Then
    Set TBPropriete.ImageAffichage = ImgLstFilter.ListImages("n-propriete").Picture
Else
    Set TBPropriete.ImageAffichage = ImgLstFilter.ListImages("propriete").Picture
End If
End Sub


Private Sub TButton1_Click()

End Sub

Private Sub TBVariable_Click()
bAfficheVariableT = Not bAfficheVariableT
If Not bAfficheVariableT Then
    Set TBVariable.ImageAffichage = ImgLstFilter.ListImages("n-variable").Picture
Else
    Set TBVariable.ImageAffichage = ImgLstFilter.ListImages("variable").Picture
End If

End Sub
