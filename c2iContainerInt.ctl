VERSION 5.00
Begin VB.UserControl c2iContainerInt 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4995
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   PropertyPages   =   "c2iContainerInt.ctx":0000
   ScaleHeight     =   4710
   ScaleWidth      =   4995
   ToolboxBitmap   =   "c2iContainerInt.ctx":0026
   Begin VB.PictureBox picTitleFond 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   2280
      ScaleHeight     =   1095
      ScaleWidth      =   1575
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picFond 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   2880
      ScaleHeight     =   975
      ScaleWidth      =   1095
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox PicTitre 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   240
      Left            =   0
      Picture         =   "c2iContainerInt.ctx":0120
      ScaleHeight     =   240
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   0
      Width           =   4995
   End
   Begin VB.Shape shpGris 
      BorderColor     =   &H00000000&
      Height          =   1695
      Left            =   120
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Image ImgOuvert 
      Height          =   240
      Index           =   1
      Left            =   2160
      Picture         =   "c2iContainerInt.ctx":0222
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgFermée 
      Height          =   240
      Index           =   1
      Left            =   1560
      Picture         =   "c2iContainerInt.ctx":0324
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgFermée 
      Height          =   240
      Index           =   0
      Left            =   1560
      Picture         =   "c2iContainerInt.ctx":0426
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgOuvert 
      Height          =   240
      Index           =   0
      Left            =   2160
      Picture         =   "c2iContainerInt.ctx":0528
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "c2iContainerInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'//****************************************//
'//  Copyright c2i - Richard CLARK
'//  http://www.c2i.fr
'//  rc@c2i.fr
'//**************************************//

Private sCaption As String
Private iAlignement As AlignmentConstants
Private bOuvert As Boolean
Private sngHauteur As Single

Private bCapturé As Boolean

'//************ NOUVEAU AVEC LA VERSION 1.1 **********************//
Public Event BeforeOuvertureChange(ByRef Cancel As Boolean)
Public Event OuvertureChange()
Public Event Resize()
Public Event MouseTitleMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseTitleDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseTitleUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event TitleClick()
Public Event TitleDblClick()

Public Event MouseContainerMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseContainerDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseContainerUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ContainerClick()
Public Event ContainerDblClick()
'//**********************************//

Private bFermetureAutomatique As Boolean
'//************Généré par c2iExplorer*******************
Private bBorder As Boolean
Private bUserControlCaptured As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long


Private Function Dessine()
    PicTitre.Height = ImgFermée(0).Height

    If bOuvert Then                                        'il est fermé, on l'ouvre
        Height = sngHauteur
        PicTitre.Picture = ImgOuvert(0).Picture
    Else                                                   'il est ouvert, on le ferme
        Height = ImgFermée(0).Height
        PicTitre.Picture = ImgFermée(0).Picture
    End If
    PropertyChanged "Ouvert"
    PropertyChanged "Hauteur"

    With shpGris
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
    End With

    If bBorder Then
        shpGris.Visible = True
        PicTitre.BorderStyle = 1
    Else
        shpGris.Visible = False
        PicTitre.BorderStyle = 0
    End If
    DessineTitre
End Function

Private Function DessineFond()
    Dim iX As Long, iY As Long, iIncreX As Long, iIncreY As Long
    'dessin du container
    If picFond.Picture <> 0 Then
        iX = CInt(ScaleWidth / picFond.Width)
        iY = CInt(ScaleHeight / picFond.Height)
        For iIncreX = 0 To iX
            For iIncreY = 0 To iY
                PaintPicture picFond.Picture, picFond.Width * iIncreX, picFond.Height * iIncreY
            Next
        Next
    End If
    If picTitleFond.Picture <> 0 Then
        iX = CInt((ScaleWidth - ImgFermée(0).Width) / picTitleFond.Width)
        iY = CInt(ScaleHeight / picTitleFond.Height)
        For iIncreX = 0 To iX
            For iIncreY = 0 To iY
                PicTitre.PaintPicture picTitleFond.Picture, picTitleFond.Width * iIncreX + ImgFermée(0).Width, picTitleFond.Height * iIncreY
            Next
        Next
    End If
End Function


Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Property Get Ouvert() As Boolean
Attribute Ouvert.VB_ProcData.VB_Invoke_Property = "Standard"
    Ouvert = bOuvert
End Property

Property Let Ouvert(bOuvertA As Boolean)
    Dim bCancel As Boolean
    If bOuvertA <> bOuvert Then
        bCancel = False
        RaiseEvent BeforeOuvertureChange(bCancel)
        If Not bCancel Then
            bOuvert = bOuvertA
            Dessine
            RaiseEvent OuvertureChange
        End If
    End If
End Property

Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "Standard"
    Caption = sCaption
End Property

Property Let Caption(sCaptionA As String)
    sCaption = sCaptionA
    PropertyChanged "Caption"
    DessineTitre
End Property

Property Get Alignement() As AlignmentConstants
    Alignement = iAlignement
End Property

Property Let Alignement(iAlignA As AlignmentConstants)
    iAlignement = iAlignA
    PropertyChanged "Alignement"
    DessineTitre
End Property

Private Function DessineTitre()
    'dessin du titre
    PicTitre.Cls
    DessineFond

    Select Case iAlignement
        Case vbCenter
            If (PicTitre.ScaleWidth - PicTitre.ScaleX(ImgFermée(0).Width, vbTwips, vbPixels)) / 2 - PicTitre.TextWidth(sCaption) / 2 + PicTitre.ScaleX(ImgFermée(0).Width, vbTwips, vbPixels) >= ImgFermée(0).Width Then
                PicTitre.CurrentX = (PicTitre.ScaleWidth - PicTitre.ScaleX(ImgFermée(0).Width, vbTwips, vbPixels)) / 2 - PicTitre.TextWidth(sCaption) / 2 + PicTitre.ScaleX(ImgFermée(0).Width, vbTwips, vbPixels)
            Else
                PicTitre.CurrentX = ImgFermée(0).Width
            End If
        Case vbLeftJustify
            PicTitre.CurrentX = ImgFermée(0).Width
        Case vbRightJustify                                'vérif largueur texte
            If PicTitre.ScaleWidth - PicTitre.TextWidth(sCaption) > ImgFermée(0).Width Then
                PicTitre.CurrentX = PicTitre.ScaleWidth - PicTitre.TextWidth(sCaption)
            Else
                PicTitre.CurrentX = ImgFermée(0).Width
            End If
    End Select
    PicTitre.CurrentY = PicTitre.ScaleHeight / 2 - PicTitre.TextHeight(sCaption) / 2

    PicTitre.Print sCaption

End Function

Private Sub PicTitre_Click()
    RaiseEvent TitleClick
End Sub

Private Sub PicTitre_DblClick()
    RaiseEvent TitleDblClick
End Sub


Private Sub PicTitre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseTitleDown(Button, Shift, X, Y)
    If X < ImgOuvert(0).Width Then
        RaiseEvent MouseTitleDown(Button, Shift, X, Y)
        If bOuvert Then
            PicTitre.Picture = ImgOuvert(1).Picture
        Else
            PicTitre.Picture = ImgFermée(1).Picture
        End If
        DessineTitre
    End If
End Sub

Private Sub PicTitre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bFermetureAutomatique Then
        If Not bOuvert Then
            Ouvert = True
        End If
    End If

    UserControl_MouseMove Button, Shift, X, Y
    If X > ImgFermée(0).Width Then
        RaiseEvent MouseTitleMove(Button, Shift, X, Y)
    End If
End Sub

Private Sub PicTitre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If X > ImgFermée(0).Width Then
        RaiseEvent MouseTitleUp(Button, Shift, X, Y)
        Exit Sub
    Else
        RaiseEvent MouseTitleUp(Button, Shift, X, Y)
        If Not bFermetureAutomatique Then
            Ouvert = Not bOuvert
        End If
    End If
End Sub

Private Sub UserControl_Click()
    RaiseEvent ContainerClick
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent ContainerDblClick
End Sub


Private Sub UserControl_Initialize()
    sCaption = "c2iContainer"
    bOuvert = True
    sngHauteur = Height
    bCapturé = False
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bUserControlCaptured Then
        bUserControlCaptured = False
    End If
    RaiseEvent MouseContainerDown(Button, Shift, X, Y)
End Sub


Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRep As Long

    RaiseEvent MouseContainerMove(Button, Shift, X, Y)

    If bFermetureAutomatique Then
        If Not bUserControlCaptured Then
            bUserControlCaptured = True
            lngRep = SetCapture(hwnd)                      'on capture le curseur
        End If
        If X < 0 Or Y < 0 Or X > Width Or Y > Height Then  'si on sort du contrôle
            bUserControlCaptured = False
            lngRep = ReleaseCapture                        'on relache le curseur
            If bOuvert Then
                Ouvert = Not bOuvert
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseContainerUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'lecture des propriétés dans le "sac"
    On Error Resume Next
    iAlignement = PropBag.ReadProperty("Alignement", vbLeftJustify)
    Caption = PropBag.ReadProperty("Caption", "c2iContainer")
    Ouvert = PropBag.ReadProperty("Ouvert", True)
    sngHauteur = PropBag.ReadProperty("Hauteur")

    Set ImgOuvert(0).Picture = PropBag.ReadProperty("ImgOuvert0")
    Set ImgOuvert(1).Picture = PropBag.ReadProperty("ImgOuvert1")
    Set ImgFermée(0).Picture = PropBag.ReadProperty("ImgFermée0")
    Set ImgFermée(1).Picture = PropBag.ReadProperty("ImgFermée1")

    BackColor = PropBag.ReadProperty("CouleurFond")
    PicTitre.ForeColor = PropBag.ReadProperty("CouleurTexteTitre")
    PicTitre.BackColor = PropBag.ReadProperty("CouleurFondTitre")

    bFermetureAutomatique = PropBag.ReadProperty("Automatique")
    bBorder = PropBag.ReadProperty("Border")
    Set picFond.Picture = PropBag.ReadProperty("PicFond")
    Set picTitleFond.Picture = PropBag.ReadProperty("PicTitleFond")
    Set PicTitre.Font = PropBag.ReadProperty("Font")
    Dessine
End Sub

Private Sub UserControl_Resize()
    If Not bOuvert Then
        Dessine
    Else
        DessineTitre
        sngHauteur = Height
        With shpGris
            .Left = 0
            .Top = 0
            .Width = Width
            .Height = Height
        End With
    End If
    RaiseEvent Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Alignement", iAlignement, vbLeftJustify
    PropBag.WriteProperty "Caption", sCaption
    PropBag.WriteProperty "Ouvert", bOuvert, True
    PropBag.WriteProperty "Hauteur", sngHauteur

    PropBag.WriteProperty "ImgOuvert0", ImgOuvert(0).Picture
    PropBag.WriteProperty "ImgOuvert1", ImgOuvert(1).Picture
    PropBag.WriteProperty "ImgFermée0", ImgFermée(0).Picture
    PropBag.WriteProperty "ImgFermée1", ImgFermée(1).Picture
    PropBag.WriteProperty "CouleurFond", BackColor
    PropBag.WriteProperty "CouleurTexteTitre", PicTitre.ForeColor
    PropBag.WriteProperty "CouleurFondTitre", PicTitre.BackColor
    PropBag.WriteProperty "Automatique", bFermetureAutomatique
    PropBag.WriteProperty "Border", bBorder
    On Error Resume Next
    PropBag.WriteProperty "PicFond", picFond.Picture
    PropBag.WriteProperty "PicTitleFond", picTitleFond.Picture
    PropBag.WriteProperty "Font", PicTitre.Font
End Sub

Public Property Set PictureFerméeUp(ByVal picPictureFerméeUpA As Picture)
    Set ImgFermée(1).Picture = picPictureFerméeUpA
    PropertyChanged "ImgFermée1"
    Dessine
End Property

Public Property Get PictureFerméeUp() As Picture
    Set PictureFerméeUp = ImgFermée(1).Picture
End Property

Public Property Set PictureFerméeDown(ByVal picPictureFerméeDownA As Picture)
    Set ImgFermée(0).Picture = picPictureFerméeDownA
    PropertyChanged "ImgFermée0"
    Dessine
End Property

Public Property Get PictureFerméeDown() As Picture
    Set PictureFerméeDown = ImgFermée(0).Picture
End Property

Public Property Set PictureOuvertUp(ByVal picPictureOuvertUpA As Picture)
    Set ImgOuvert(0).Picture = picPictureOuvertUpA
    PropertyChanged "ImgOuvert0"
    Dessine
End Property

Public Property Get PictureOuvertUp() As Picture
    Set PictureOuvertUp = ImgOuvert(0).Picture
End Property

Public Property Set PictureOuvertDown(ByVal picPictureOuvertDownA As Picture)
    Set ImgOuvert(1).Picture = picPictureOuvertDownA
    PropertyChanged "ImgOuvert1"
    Dessine
End Property

Public Property Get PictureOuvertDown() As Picture
    Set PictureOuvertDown = ImgOuvert(1).Picture
End Property

Public Property Let CouleurFond(ByVal lngBackColorA As OLE_COLOR)
    BackColor = lngBackColorA
    PropertyChanged "CouleurFond"
End Property

Public Property Get CouleurFond() As OLE_COLOR
    CouleurFond = BackColor
End Property

Public Property Let CouleurFondTitre(ByVal lngCouleurFondTitreA As OLE_COLOR)
    PicTitre.BackColor = lngCouleurFondTitreA
    PropertyChanged "CouleurFondTitre"
    Dessine
End Property

Public Property Get CouleurFondTitre() As OLE_COLOR
    CouleurFondTitre = PicTitre.BackColor
End Property

Public Property Let CouleurTexteTitre(ByVal lngCouleurTexteTitreA As OLE_COLOR)
    PicTitre.ForeColor = lngCouleurTexteTitreA
    PropertyChanged "CouleurTexteTitre"
    Dessine
End Property

Public Property Get CouleurTexteTitre() As OLE_COLOR
    CouleurTexteTitre = PicTitre.ForeColor
End Property

Public Property Let Hauteur(ByVal sngHauteurA As Single)
    sngHauteur = sngHauteurA
    PropertyChanged "Hauteur"
    Dessine
End Property

Public Property Get Hauteur() As Single
    Hauteur = sngHauteur
End Property

Public Property Let Automatique(ByVal bFermetureAutomatiqueA As Boolean)
    bFermetureAutomatique = bFermetureAutomatiqueA
    PropertyChanged "Automatique"
End Property

Public Property Get Automatique() As Boolean
    Automatique = bFermetureAutomatique
End Property

Public Property Let Border(ByVal bBorderA As Boolean)
    bBorder = bBorderA
    PropertyChanged "Border"
    Dessine
End Property

Public Property Get Border() As Boolean
    Border = bBorder
End Property

Public Property Get ParentHwnd() As Long
    ParentHwnd = ContainerHwnd
End Property


Public Property Set PictureFond(ByVal picPictureFondA As Picture)
    Set picFond.Picture = picPictureFondA
    PropertyChanged "PicFond"
    Dessine
End Property

Public Property Get PictureFond() As Picture
    Set PictureFond = picFond.Picture
End Property


Public Property Set PictureTitleFond(ByVal picPictureTitleFondA As Picture)
    Set picTitleFond.Picture = picPictureTitleFondA
    PropertyChanged "PictureTitleFond"
    Dessine
End Property

Public Property Get PictureTitleFond() As Picture
    Set PictureTitleFond = picTitleFond.Picture
End Property

Public Property Get TitleHeight() As Long
    TitleHeight = ImgFermée(0).Height
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = PicTitre.Font
End Property

Public Property Set Font(FontA As StdFont)
    Set PicTitre.Font = FontA
    DessineTitre
    PropertyChanged "Font"
End Property
