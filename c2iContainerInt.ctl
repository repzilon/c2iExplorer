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
   LockControls    =   -1  'True
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
      Picture         =   "c2iContainerInt.ctx":0120
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgFermée 
      Height          =   240
      Index           =   1
      Left            =   1560
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgFermée 
      Height          =   240
      Index           =   0
      Left            =   1560
      Picture         =   "c2iContainerInt.ctx":01F1
      Top             =   480
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImgOuvert 
      Height          =   240
      Index           =   0
      Left            =   2160
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
' ***** BEGIN LICENSE BLOCK *****
' Version: MPL 1.1
'
' The contents of this file are subject to the Mozilla Public License Version
' 1.1 (the "License"); you may not use this file except in compliance with
' the License. You may obtain a copy of the License at
' http://www.mozilla.org/MPL/
'
' Software distributed under the License is distributed on an "AS IS" basis,
' WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License
' for the specific language governing rights and limitations under the
' License.
'
' The Original Code is c2iExplorer (for Visual Basic 6).
'
' The Initial Developer of the Original Code is
' c2i - Richard Clark Ingénierie Informatique (www.c2i.fr).
' Portions created by the Initial Developer are Copyright (C) 2000
' the Initial Developer. All Rights Reserved.
'
' Contributor(s):
' René Rhéaume (rener@mediom.qc.ca)
'
' ***** END LICENSE BLOCK *****

Option Explicit

'//****************************************//
'//  Copyright c2i - Richard CLARK
'//  http://www.c2i.fr
'//  rc@c2i.fr
'//**************************************//

' -------------------- DÉCLARATION CONSTANTES --------------------
' Par René Rhéaume, le 6 janvier 2002
Private Const conAlignement As String = "Alignement"
#If SlimC2iContainer = 0 Then
Private Const conAutomatique As String = "Automatique"
Private Const conBorder As String = "Border"
#End If
Private Const conCaption As String = "Caption"
Private Const conCouleurFond As String = "CouleurFond"
Private Const conCouleurFondTitre As String = "CouleurFondTitre"
Private Const conCouleurTexteTitre As String = "CouleurTexteTitre"
Private Const conFont As String = "Font"
Private Const conHauteur As String = "Hauteur"
#If ((SlimC2iContainer = 0) And (NoContainerDownImage = 0)) Then
Private Const conImgFermée0 As String = "ImgFermée0"
Private Const conImgOuvert1 As String = "ImgOuvert1"
#End If
Private Const conImgFermée1 As String = "ImgFermée1"
Private Const conImgOuvert0 As String = "ImgOuvert0"
#If SlimC2iContainer = 0 Then
Private Const conPicFond As String = "PicFond"
Private Const conPictureTitleFond As String = "PictureTitleFond"
#End If
Private Const conOuvert As String = "Ouvert"
Private Const conLibelle As String = "c2iContainer"

Private sCaption As String
Private iAlignement As AlignmentConstants
Private bOuvert As Boolean
Private sngHauteur As Single

'Private bCapturé As Boolean

'//************ NOUVEAU AVEC LA VERSION 1.1 **********************//
Public Event BeforeOuvertureChange(ByRef Cancel As Boolean)
Public Event OuvertureChange()
Public Event Resize()
#If SlimC2iContainer = 0 Then
Public Event MouseTitleMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseTitleDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#End If
Public Event MouseTitleUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#If SlimC2iContainer = 0 Then
Public Event TitleClick()
#End If
Public Event TitleDblClick()

#If SlimC2iContainer = 0 Then
Public Event MouseContainerMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseContainerDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseContainerUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event ContainerClick()
Public Event ContainerDblClick()
#End If
'//**********************************//

#If SlimC2iContainer = 0 Then
Private bFermetureAutomatique As Boolean
'//************Généré par c2iExplorer*******************
Private bBorder As Boolean
Private bUserControlCaptured As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Sub Dessine()
    PicTitre.Height = ImgFermée(0).Height

    If (bOuvert) Then                                      'il est fermé, on l'ouvre
        Height = sngHauteur
        PicTitre.Picture = ImgOuvert(0).Picture
    Else                                                   'il est ouvert, on le ferme
        Height = ImgFermée(0).Height
        PicTitre.Picture = ImgFermée(0).Picture
    End If
    PropertyChanged conOuvert
    PropertyChanged conHauteur
    
#If SlimC2iContainer = 0 Then
    With shpGris
        .Left = 0
        .Top = 0
        .Width = ScaleWidth
        .Height = ScaleHeight
        .Visible = bBorder
    End With

    PicTitre.BorderStyle = Abs(CLng(bBorder))
#End If
    DessineTitre
End Sub

'Procédure optimisée par René Rhéaume le 17 janvier 2003
' Mise en cache des dimensions et vérification contre les
' divisions par zéro.
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
#If SlimC2iContainer = 0 Then
Private Sub DessineFond(Optional ByVal blnTitreSeul As Boolean = conFaux)
    Dim iX As Long, iY As Long, iIncreX As Long, iIncreY As Long
    Dim sngLarg As Single, sngHaut As Single, sngLargImgFerme As Single
    'dessin du container
    If (blnTitreSeul = conFaux) Then
        If (picFond.Picture <> 0) Then
            sngLarg = picFond.Width
            sngHaut = picFond.Height
            If (sngLarg) Then iX = CInt(ScaleWidth / sngLarg)
            If (sngHaut) Then iY = CInt(ScaleHeight / sngHaut)
            For iIncreX = 0 To iX
                For iIncreY = 0 To iY
                    PaintPicture picFond.Picture, sngLarg * iIncreX, sngHaut * iIncreY
                Next
            Next
        End If
    End If
    If (picTitleFond.Picture <> 0) Then
        sngLarg = picTitleFond.Width
        sngHaut = picTitleFond.Height
        sngLargImgFerme = ImgFermée(0).Width
        If (sngLarg) Then iX = CInt((ScaleWidth - sngLargImgFerme) / sngLarg) Else: iX = 0
        If (sngHaut) Then iY = CInt(ScaleHeight / sngHaut) Else: iY = 0
        For iIncreX = 0 To iX
            For iIncreY = 0 To iY
                PicTitre.PaintPicture picTitleFond.Picture, _
                    sngLarg * iIncreX + sngLargImgFerme, sngHaut * iIncreY
            Next
        Next
    End If
End Sub
#End If

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Property Get Ouvert() As Boolean
    Ouvert = bOuvert
End Property

Property Let Ouvert(bOuvertA As Boolean)
    Dim bCancel As Boolean
    If (bOuvertA <> bOuvert) Then
        bCancel = conFaux
        RaiseEvent BeforeOuvertureChange(bCancel)
        If (Not bCancel) Then
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
    PropertyChanged conCaption
    DessineTitre
End Property

Property Get Alignement() As AlignmentConstants
    Alignement = iAlignement
End Property

Property Let Alignement(iAlignA As AlignmentConstants)
    iAlignement = iAlignA
    PropertyChanged conAlignement
    DessineTitre
End Property

'Procédure ajoutée par René Rhéaume le 30 juin 2002
' Ferme le conteneur sans déclencher d'événement
Public Sub SilentClose()
    If (bOuvert) Then
        bOuvert = conFaux
        Dessine
    End If
End Sub

' Optimisé par René Rhéaume le 6 janvier 2002
' Optimisé par René Rhéaume le 19 avril 2003
'  Mise en compilation conditionnelle d'éléments superflus
Private Sub DessineTitre()
    Dim sngLargImgFerm As Single
    Dim sngPosX As Single
    'dessin du titre
    With PicTitre
        .Cls
#If SlimC2iContainer = 0 Then
        DessineFond (conVrai)
#End If

        sngLargImgFerm = ImgFermée(0).Width
        Select Case iAlignement
            Case vbCenter
                sngPosX = (.ScaleWidth + .ScaleX(sngLargImgFerm, vbTwips, vbPixels) - .TextWidth(sCaption)) / 2
                If (sngPosX >= sngLargImgFerm) Then
                    .CurrentX = sngPosX
                Else
                    .CurrentX = sngLargImgFerm
                End If
            Case vbLeftJustify
                .CurrentX = sngLargImgFerm
            Case vbRightJustify                            'vérif largueur texte
                sngPosX = .ScaleWidth - .TextWidth(sCaption)
                If (sngPosX > sngLargImgFerm) Then
                    .CurrentX = sngPosX
                Else
                    .CurrentX = sngLargImgFerm
                End If
        End Select
        .CurrentY = (.ScaleHeight - .TextHeight(sCaption)) / 2
        PicTitre.Print sCaption
    End With
End Sub

#If SlimC2iContainer = 0 Then
Private Sub PicTitre_Click()
    RaiseEvent TitleClick
End Sub
#End If

Private Sub PicTitre_DblClick()
    RaiseEvent TitleDblClick
End Sub

#If SlimC2iContainer = 0 Then
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub PicTitre_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseTitleDown(Button, Shift, X, Y)
#If NoContainerDownImage = 0 Then
    If (X < ImgOuvert(0).Width) Then
        RaiseEvent MouseTitleDown(Button, Shift, X, Y)
        If (bOuvert) Then
            PicTitre.Picture = ImgOuvert(1).Picture
        Else
            PicTitre.Picture = ImgFermée(1).Picture
        End If
        DessineTitre
    End If
#End If
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub PicTitre_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case conFaux
        Case bFermetureAutomatique, bOuvert = conFaux
        Case Else
            Ouvert = conVrai
    End Select

    UserControl_MouseMove Button, Shift, X, Y
    If (X > ImgFermée(0).Width) Then
        RaiseEvent MouseTitleMove(Button, Shift, X, Y)
    End If
End Sub
#End If

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub PicTitre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseTitleUp(Button, Shift, X, Y)
    If (X <= ImgFermée(0).Width) Then
#If SlimC2iContainer = 0 Then
        If (Not bFermetureAutomatique) Then
            Ouvert = Not bOuvert
        End If
#Else
        Ouvert = Not bOuvert
#End If
    End If
End Sub

#If SlimC2iContainer = 0 Then
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_Click()
    RaiseEvent ContainerClick
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_DblClick()
    RaiseEvent ContainerDblClick
End Sub
#End If

Private Sub UserControl_Initialize()
    sCaption = conLibelle
    bOuvert = conVrai
    sngHauteur = Height
'    bCapturé = conFaux
End Sub

#If SlimC2iContainer = 0 Then
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Plus rapide ainsi
    bUserControlCaptured = conFaux
    RaiseEvent MouseContainerDown(Button, Shift, X, Y)
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRep As Long

    RaiseEvent MouseContainerMove(Button, Shift, X, Y)

    If (bFermetureAutomatique) Then
        If (Not bUserControlCaptured) Then
            bUserControlCaptured = conVrai
            lngRep = SetCapture(hWnd)                      'on capture le curseur
        End If
        'Équivalent à If X < 0 Or Y < 0 Or X > Width Or Y > Height Then
        Select Case conVrai
            Case X < 0, Y < 0, X > Width, Y > Height       'si on sort du contrôle
                bUserControlCaptured = conFaux
                lngRep = ReleaseCapture                    'on relache le curseur
                'If bOuvert Then
                '   Ouvert = Not bOuvert
                'End If
                bOuvert = conFaux
        End Select
    End If
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseContainerUp(Button, Shift, X, Y)
End Sub
#End If

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'lecture des propriétés dans le "sac"
    On Error Resume Next
    iAlignement = PropBag.ReadProperty(conAlignement, vbLeftJustify)
    Caption = PropBag.ReadProperty(conCaption, conLibelle)
    Ouvert = PropBag.ReadProperty(conOuvert, conVrai)
    sngHauteur = PropBag.ReadProperty(conHauteur)

    Set ImgOuvert(0).Picture = PropBag.ReadProperty(conImgOuvert0)
    Set ImgFermée(1).Picture = PropBag.ReadProperty(conImgFermée1)
#If ((SlimC2iContainer = 0) And (NoContainerDownImage = 0)) Then
    Set ImgFermée(0).Picture = PropBag.ReadProperty(conImgFermée0)
    Set ImgOuvert(1).Picture = PropBag.ReadProperty(conImgOuvert1)
#End If

    BackColor = PropBag.ReadProperty(conCouleurFond)
    PicTitre.ForeColor = PropBag.ReadProperty(conCouleurTexteTitre)
    PicTitre.BackColor = PropBag.ReadProperty(conCouleurFondTitre)
#If SlimC2iContainer = 0 Then
    bFermetureAutomatique = PropBag.ReadProperty(conAutomatique)
    bBorder = PropBag.ReadProperty(conBorder)
    Set picFond.Picture = PropBag.ReadProperty(conPicFond)
    Set picTitleFond.Picture = PropBag.ReadProperty(conPictureTitleFond)
#End If
    Set PicTitre.Font = PropBag.ReadProperty(conFont)
    Dessine
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_Resize()
    If Not bOuvert Then
        Dessine
    Else
        DessineTitre
        sngHauteur = Height
#If SlimC2iContainer = 0 Then
        With shpGris
            .Left = 0
            .Top = 0
            .Width = Width
            .Height = Height
        End With
#End If
    End If
    RaiseEvent Resize
End Sub

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty conAlignement, iAlignement, vbLeftJustify
    PropBag.WriteProperty conCaption, sCaption
    PropBag.WriteProperty conOuvert, bOuvert, conVrai
    PropBag.WriteProperty conHauteur, sngHauteur

    PropBag.WriteProperty conImgOuvert0, ImgOuvert(0).Picture
    PropBag.WriteProperty conImgFermée1, ImgFermée(1).Picture
#If ((SlimC2iContainer = 0) And (NoContainerDownImage = 0)) Then
    PropBag.WriteProperty conImgFermée0, ImgFermée(0).Picture
    PropBag.WriteProperty conImgOuvert1, ImgOuvert(1).Picture
#End If
    PropBag.WriteProperty conCouleurFond, BackColor
    PropBag.WriteProperty conCouleurTexteTitre, PicTitre.ForeColor
    PropBag.WriteProperty conCouleurFondTitre, PicTitre.BackColor
#If SlimC2iContainer = 0 Then
    PropBag.WriteProperty conAutomatique, bFermetureAutomatique
    PropBag.WriteProperty conBorder, bBorder
    On Error Resume Next
    PropBag.WriteProperty conPicFond, picFond.Picture
    PropBag.WriteProperty conPictureTitleFond, picTitleFond.Picture
#Else
    On Error Resume Next
#End If
    PropBag.WriteProperty conFont, PicTitre.Font
End Sub

Public Property Set PictureFerméeUp(ByVal picPictureFerméeUpA As Picture)
Attribute PictureFerméeUp.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set ImgFermée(1).Picture = picPictureFerméeUpA
    PropertyChanged conImgFermée1
    Dessine
End Property

Public Property Get PictureFerméeUp() As Picture
    Set PictureFerméeUp = ImgFermée(1).Picture
End Property

#If ((SlimC2iContainer = 0) And (NoContainerDownImage = 0)) Then
'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Set PictureFerméeDown(ByVal picPictureFerméeDownA As Picture)
Attribute PictureFerméeDown.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set ImgFermée(0).Picture = picPictureFerméeDownA
    PropertyChanged conImgFermée0
    Dessine
End Property

'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get PictureFerméeDown() As Picture
    Set PictureFerméeDown = ImgFermée(0).Picture
End Property
#End If

Public Property Set PictureOuvertUp(ByVal picPictureOuvertUpA As Picture)
Attribute PictureOuvertUp.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set ImgOuvert(0).Picture = picPictureOuvertUpA
    PropertyChanged conImgOuvert0
    Dessine
End Property

Public Property Get PictureOuvertUp() As Picture
    Set PictureOuvertUp = ImgOuvert(0).Picture
End Property

#If ((SlimC2iContainer = 0) And (NoContainerDownImage = 0)) Then
'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Set PictureOuvertDown(ByVal picPictureOuvertDownA As Picture)
Attribute PictureOuvertDown.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set ImgOuvert(1).Picture = picPictureOuvertDownA
    PropertyChanged conImgOuvert1
    Dessine
End Property

'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get PictureOuvertDown() As Picture
    Set PictureOuvertDown = ImgOuvert(1).Picture
End Property
#End If

Public Property Let CouleurFond(ByVal lngBackColorA As OLE_COLOR)
Attribute CouleurFond.VB_ProcData.VB_Invoke_PropertyPut = ";Apparence"
    BackColor = lngBackColorA
    PropertyChanged conCouleurFond
End Property

Public Property Get CouleurFond() As OLE_COLOR
    CouleurFond = BackColor
End Property

Public Property Let CouleurFondTitre(ByVal lngCouleurFondTitreA As OLE_COLOR)
Attribute CouleurFondTitre.VB_ProcData.VB_Invoke_PropertyPut = ";Apparence"
    PicTitre.BackColor = lngCouleurFondTitreA
    PropertyChanged conCouleurFondTitre
    Dessine
End Property

Public Property Get CouleurFondTitre() As OLE_COLOR
    CouleurFondTitre = PicTitre.BackColor
End Property

Public Property Let CouleurTexteTitre(ByVal lngCouleurTexteTitreA As OLE_COLOR)
Attribute CouleurTexteTitre.VB_ProcData.VB_Invoke_PropertyPut = ";Apparence"
    PicTitre.ForeColor = lngCouleurTexteTitreA
    PropertyChanged conCouleurTexteTitre
    Dessine
End Property

Public Property Get CouleurTexteTitre() As OLE_COLOR
    CouleurTexteTitre = PicTitre.ForeColor
End Property

Public Property Let Hauteur(ByVal sngHauteurA As Single)
    sngHauteur = sngHauteurA
    PropertyChanged conHauteur
    Dessine
End Property

Public Property Get Hauteur() As Single
    Hauteur = sngHauteur
End Property

#If SlimC2iContainer = 0 Then
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Let Automatique(ByVal bFermetureAutomatiqueA As Boolean)
    bFermetureAutomatique = bFermetureAutomatiqueA
    PropertyChanged conAutomatique
End Property

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get Automatique() As Boolean
    Automatique = bFermetureAutomatique
End Property

#If NoContainerDownImage = 0 Then
'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Let Border(ByVal bBorderA As Boolean)
    bBorder = bBorderA
    PropertyChanged conBorder
    Dessine
End Property

'Procédure optimisée par René Rhéaume le 18 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get Border() As Boolean
    Border = bBorder
End Property
#End If
#End If

'Problème avec cette propriété
'VB5 ne connaît pas ContainerHwnd (il le prend pour une variable non définie)
#If VersionVB = 6 Then
Public Property Get ParentHwnd() As Long
    ParentHwnd = ContainerHwnd
End Property
#End If

#If SlimC2iContainer = 0 Then
'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Set PictureFond(ByVal picPictureFondA As Picture)
Attribute PictureFond.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set picFond.Picture = picPictureFondA
    PropertyChanged conPicFond
    Dessine
End Property

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get PictureFond() As Picture
    Set PictureFond = picFond.Picture
End Property

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Set PictureTitleFond(ByVal picPictureTitleFondA As Picture)
Attribute PictureTitleFond.VB_ProcData.VB_Invoke_PropertyPutRef = ";Apparence"
    Set picTitleFond.Picture = picPictureTitleFondA
    PropertyChanged conPictureTitleFond
    Dessine
End Property

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get PictureTitleFond() As Picture
    Set PictureTitleFond = picTitleFond.Picture
End Property

'Procédure optimisée par René Rhéaume le 19 avril 2003
' Mise en compilation conditionnelle d'éléments superflus
Public Property Get TitleHeight() As Long
    TitleHeight = ImgFermée(0).Height
End Property
#End If

Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
    Set Font = PicTitre.Font
End Property

Public Property Set Font(FontA As StdFont)
    Set PicTitre.Font = FontA
    DessineTitre
    PropertyChanged conFont
End Property
