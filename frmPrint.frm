VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4830
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbPrint 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   4515
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   17
      MaskColor       =   132
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":000C
            Key             =   "tous"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":014D
            Key             =   "voir"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":028D
            Key             =   "quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   661
      ButtonWidth     =   847
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tous"
            Object.ToolTipText     =   "Aperçu général"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "voir"
            Object.ToolTipText     =   "Aperçu élément selectionné"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quit"
            Object.ToolTipText     =   "Quitter"
            ImageIndex      =   3
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   0
      Pattern         =   "[*]c2i*.*htm*;[*]c2i*.xml"
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblSelectedElement 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   3795
   End
   Begin VB.Label lblTitre 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblInfos 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' René Rhéaume  (rener@mediom.qc.ca)
' Pascal Martin (pmartin@stolz-sequipag.fr)
'
' ***** END LICENSE BLOCK *****

Option Explicit
Private Const conVoir As String = "voir"
Private Const conTous As String = "tous"
Private Const conQuit As String = "quit"
Private Const conDate As String = "Date"
Private Const conAttribImg As String = ".gif"" border=""0"" align=""middle"" hspace=""5"" />"
Private Const conBaliseImg As String = "<img src=""img/"
Private Const conFormatDate As String = "dddddd"
Private Const conComponents As String = "Components"
Private Const conMembers As String = "Members"
Private Const conNomForm As String = "frmPrint"
Private objVBPrjPrint As VBIDE.VBProject
Private objVBCmpPrint As VBIDE.VBComponent
Private intNoFichier As Integer

'Optimisation par René Rhéaume les 31 juillet 2001 et 5 janvier 2002
'Procédure modifiée par René Rhéaume le 14 juin 2002
' Ajout du cas «Exporter tout»
'Procédure modifiée par René Rhéaume le 19 juin 2002
' Correction de bogue si blnTous est à Faux
Private Sub OK(Optional ByVal blnTous As Boolean = conFaux)
    Dim strCheminModele As String
    Dim strFichierModele As String
    Dim blnCmpPresent As Boolean, blnPrjPresent As Boolean
    Dim blnExporterTout As Boolean
    strFichierModele = File1.Filename
    strCheminModele = File1.Path & conBS & strFichierModele

    If (LenB(strFichierModele) > 0) Then
        If (FichierExiste(strCheminModele)) Then
            blnPrjPresent = Not objVBPrjPrint Is Nothing
            blnCmpPresent = Not objVBCmpPrint Is Nothing
            blnExporterTout = blnTous Or (Not (blnPrjPresent Or blnCmpPresent))
            If (blnExporterTout) Then
                ExportHTML strCheminModele
            Else
                If (blnCmpPresent) Then
                    ExportHTMLComponent strCheminModele, objVBCmpPrint
                Else
                    ExportHTMLProject strCheminModele, objVBPrjPrint
                End If
            End If
            'Unload Me
        Else
            MsgBox mlgarFichExistePas(0) & strCheminModele & mlgarFichExistePas(1)
        End If
    End If
End Sub

Private Sub File1_DblClick()
    If (LenB(File1.Filename) > 0) Then OK
End Sub

'Modifié par René Rhéaume le 5 janvier 2002
'Modifié par René Rhéaume le 18 juin 2002
' Support multilingue
Private Sub Form_Load()
    Dim I As Long
    
    Me.Caption = LireChaineLocalisee(conNomForm, conL10nWCap, "Générer un rapport")
    If (blnMultilingueActive) Then
        For I = 1 To 5 Step 2
            Toolbar1.Buttons(I).ToolTipText = LireChaineLocalisee(conNomForm, _
                "Obj.Toolbar1.Buttons(" & I & ").ToolTipText", Toolbar1.Buttons(I).ToolTipText)
        Next I
    End If
    lblInfos = LireChaineLocalisee(conNomForm, "Obj.lblInfos.Caption", "Choisissez votre modèle :")
    lblTitre = LireChaineLocalisee(conNomForm, "Obj.lblTitre.Caption", "Élément sélectionné :")
    
    '    PositionForm Me, conVrai
    File1.Path = strCheminApp & "\html"
'    ExtraireImageList ImageList1, "frmPrint.ImageList1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '    PositionForm Me, conFaux
    Set objVBPrjPrint = Nothing
    Set objVBCmpPrint = Nothing
End Sub

'Modifié par René Rhéaume le 5 janvier 2002
'Procédure modifiée par René Rhéaume le 14 juin 2002
' OK gère maintenant le cas «Exporter tout»
'Procédure modifiée par René Rhéaume le 5 juillet 2002
' La procédure Quitter a été «inliné»
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case conVoir
            OK
        Case conQuit
            Unload Me
            'Quitter
        Case conTous
            OK conVrai
    End Select
End Sub

Public Sub ObjetSelectionnés(ByVal objVBPrj As VBIDE.VBProject, objVBCmp As VBIDE.VBComponent)
    Set objVBPrjPrint = objVBPrj
    Set objVBCmpPrint = objVBCmp
    lblSelectedElement = vbNullString
    If (Not objVBCmpPrint Is Nothing) Then
        lblSelectedElement = objVBCmpPrint.Name
    End If
    If (Not objVBPrjPrint Is Nothing) Then
        lblSelectedElement = objVBPrjPrint.Name
    End If
End Sub

'//-------------------------------------------------------------//
'// Contenu de l'ancien module MPrint ci-dessous ---------------//
'// Modifications apportées par René Rhéaume 1er août 2001 -----//
'// Remplacement du FileSystemObject et des TextStream par -----//
'// les fonctions d'entrée/sortie par défaut de Visual Basic ---//
'//-------------------------------------------------------------//

' Procédure modifiée par René Rhéaume le 3 mai 2002
'  Arrimage du correctif de Pascal Martin
' Procédure modifiée par René Rhéaume le 3 janvier 2003
'  Support des documents associés
Private Sub AnalyseHTMLCmp(ByVal objVBCmp As VBIDE.VBComponent, ByRef sM As String)
    Const conComponentName As String = "Component-Name"
    Const conComponentDescription As String = "Component-Description"
    Const conComponentNbrMember As String = "Component-NbrMember"
    Const conComponentNbrLines As String = "Component-NbrLines"
    
    Call RemplaceString(sM, "Component-Img", conBaliseImg & ExtraitIconeComponent(objVBCmp) & conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))
    Select Case objVBCmp.Type
        Case vbext_ct_ResFile, vbext_ct_RelatedDocument
            'Les fichiers de ressource ne possédent pas toutes les propriétés des autres composants
            Call RemplaceString(sM, conComponentName, objVBCmp.FileNames(1))
            Call RemplaceString(sM, conComponentDescription, vbNullString)
            Call RemplaceString(sM, conComponentNbrMember, conAZero)
            Call RemplaceString(sM, conComponentNbrLines, conAZero)
        Case Else
            Call RemplaceString(sM, conComponentName, objVBCmp.Name)
            Call RemplaceString(sM, conComponentDescription, objVBCmp.Description)
            Call RemplaceString(sM, conComponentNbrMember, CStr(objVBCmp.CodeModule.Members.Count))
            Call RemplaceString(sM, conComponentNbrLines, CStr(objVBCmp.CodeModule.CountOfLines))
    End Select
End Sub

' Procédure modifiée par René Rhéaume le 4 mai 2002
' Élimination du message d'erreur pour une description
Private Sub AnalyseHTMLMember(ByVal objVBMember As VBIDE.Member, ByRef sM As String)
    Dim lngC As Long, strDescription As String
    
    On Error Resume Next
    strDescription = objVBMember.Description
    Select Case (Err.Number)
        Case 32811
            strDescription = vbNullString
'        Case 0
'            'on ne fait rien
'        Case Else
'            GoTo AffichErr
    End Select
    Err.Clear
    
'    On Error GoTo AffichErr
    Call RemplaceString(sM, "Member-Name", objVBMember.Name)
    Call RemplaceString(sM, "Member-Description", strDescription)
    Call RemplaceString(sM, "Member-Img", conBaliseImg & ExtraitIconeMembre(objVBMember, lngC) & conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))
    Exit Sub
    
'AffichErr:
'    GererErrInattendue , "frmPrint.AnalyseHTMLMember"
'    Resume Next
End Sub

Private Sub AnalyseHTMLPrj(ByVal objVBPrj As VBIDE.VBProject, ByRef sM As String)
    On Error GoTo AffichErr
    Call RemplaceString(sM, "Project-Name", objVBPrj.Name)
    Call RemplaceString(sM, "Project-FileName", objVBPrj.Filename)
    Call RemplaceString(sM, "Project-BuildFileName", objVBPrj.BuildFileName)
    Call RemplaceString(sM, "Project-Description", objVBPrj.Description)
    Call RemplaceString(sM, "Project-Img", conBaliseImg & ExtraitIconeProjet(objVBPrj) & conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))
    Call RemplaceString(sM, "Project-NbrComponent", CStr(objVBPrj.VBComponents.Count))
    Exit Sub

AffichErr:
    GererErrInattendue , "frmPrint.AnalyseHTMLPrj"
End Sub

'Procédure optimisée par René Rhéaume le 30 juin 2002
'Procédure optimisée par René Rhéaume le 12 janvier 2003
Private Sub ExportHTML(ByVal sHTMLFileName As String)
    Dim objVBPrj As VBIDE.VBProject
    Dim sTete As String, sFin As String, sMiddle As String
    Dim strDate As String

    Screen.MousePointer = vbHourglass
    On Error GoTo GestErr
    SupprimerFichTemp

    'Nouvelle version utilisant les instructions I/O de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, "Projects") Then
        intNoFichier = OuvrirFichierAjout(c2iHTMLFile)
        If (intNoFichier <> -1) Then
            strDate = Format$(VBA.Date$, conFormatDate)
            RemplaceString sTete, conDate, strDate
            EcrireFichier sTete, intNoFichier
            
            For Each objVBPrj In VBInstance.VBProjects
                AddProject objVBPrj, sMiddle
            Next
            
            RemplaceString sFin, conDate, strDate
            EcrireFichier sFin, intNoFichier
            
            FermerFichier intNoFichier
        End If
    End If

    VoirFichApercu
    
SortieProc:
    Screen.MousePointer = vbDefault
    Exit Sub
GestErr:
    GererErrInattendue , "frmPrint.ExportHTML"
    Resume SortieProc
End Sub

'Procédure optimisée par René Rhéaume le 30 juin 2002
'Procédure optimisée par René Rhéaume le 12 janvier 2003
Private Sub ExportHTMLProject(ByVal sHTMLFileName As String, ByVal objVBPrj As VBIDE.VBProject)
    Dim objVBCmp As VBIDE.VBComponent
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    SupprimerFichTemp

    'Nouvelle version utilisant les instructions I/O de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, conComponents) Then
        intNoFichier = OuvrirFichierAjout(c2iHTMLFile)
        If (intNoFichier <> -1) Then
            AnalyseHTMLPrj objVBPrj, sTete
            EcrireFichier sTete, intNoFichier
    
            'barre de progression
            With frmPrint.pbPrint
                .Min = 0
                .Max = objVBPrj.VBComponents.Count
                .Value = 0
                For Each objVBCmp In objVBPrj.VBComponents
                    AddComponent objVBCmp, sMiddle
                    .Value = .Value + 1
                Next
            End With
            
            AnalyseHTMLPrj objVBPrj, sFin
            EcrireFichier sFin, intNoFichier
            
            FermerFichier intNoFichier
        End If
    End If
    
    VoirFichApercu

SortieProc:
    Set objVBCmp = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
GestErr:
    GererErrInattendue , "frmPrint.ExportHTMLProject"
    Resume SortieProc
End Sub

'Procédure optimisée par René Rhéaume le 30 juin 2002
'Procédure optimisée par René Rhéaume le 12 janvier 2003
Private Sub ExportHTMLComponent(ByVal sHTMLFileName As String, ByVal objVBCmp As VBIDE.VBComponent)
    Dim objVBMember As VBIDE.Member
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    SupprimerFichTemp

    'Nouvelle version utilisant les instructions I/O de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, conMembers) Then
        intNoFichier = OuvrirFichierAjout(c2iHTMLFile)
        If (intNoFichier <> -1) Then
            AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
            AnalyseHTMLCmp objVBCmp, sTete
            EcrireFichier sTete, intNoFichier
            
            'barre de progression
            With frmPrint.pbPrint
                .Min = 0
                .Max = objVBCmp.CodeModule.Members.Count
                .Value = 0
    
                For Each objVBMember In objVBCmp.CodeModule.Members
                    AddMember objVBMember, sMiddle
                    .Value = .Value + 1
                    If GetInputState() Then DoEvents
                Next
            End With
            
            AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
            AnalyseHTMLCmp objVBCmp, sFin
            EcrireFichier sFin, intNoFichier
            
            FermerFichier intNoFichier
        End If
    End If
    
    VoirFichApercu

SortieProc:
    Set objVBMember = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
GestErr:
    GererErrInattendue , "frmPrint.ExportHTMLComponent"
    Resume SortieProc
End Sub

'Modifié par René Rhéaume 1er août 2001
'Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
Private Sub AddProject(ByVal objVBPrj As VBIDE.VBProject, ByVal sMiddle As String)
    Dim objVBCmp As VBIDE.VBComponent
    Dim sTete As String, sFin As String
    'add project

    If ExtraitHTML(sTete, sFin, sMiddle, conComponents) Then
        AnalyseHTMLPrj objVBPrj, sTete
        EcrireFichier sTete, intNoFichier

        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBPrj.VBComponents.Count
            .Value = 0
            For Each objVBCmp In objVBPrj.VBComponents
                AddComponent objVBCmp, sMiddle
                .Value = .Value + 1
                If GetInputState() Then DoEvents
            Next
        End With
        AnalyseHTMLPrj objVBPrj, sFin
        EcrireFichier sFin, intNoFichier
    End If
    
    Set objVBCmp = Nothing
    
    Exit Sub
End Sub

' Modifié par René Rhéaume 1er août 2001
'  Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
' Procédure modifiée par René Rhéaume le 3 mai 2002
'  Arrimage du correctif de Pascal Martin
' Procédure modifiée par René Rhéaume le 3 janvier 2003
'  Support des documents associés
Private Sub AddComponent(ByVal objVBCmp As VBIDE.VBComponent, ByVal sMiddle As String)
    Dim objVBMember As VBIDE.Member
    Dim sTete As String, sFin As String

    If ExtraitHTML(sTete, sFin, sMiddle, conMembers) Then
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
        AnalyseHTMLCmp objVBCmp, sTete
        EcrireFichier sTete, intNoFichier
        Select Case objVBCmp.Type
            Case vbext_ct_ResFile, vbext_ct_RelatedDocument
                'les fichiers de ressource ne peuvent pas contenir de code
            Case Else
                For Each objVBMember In objVBCmp.CodeModule.Members
                    AddMember objVBMember, sMiddle
                Next
        End Select
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
        AnalyseHTMLCmp objVBCmp, sFin
        EcrireFichier sFin, intNoFichier
    End If

    Set objVBMember = Nothing
End Sub

'Modifié par René Rhéaume 1er août 2001
'Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
Private Sub AddMember(ByVal objVBMember As VBIDE.Member, ByVal sMiddle As String)
    If (AfficheMembre(objVBMember)) Then
        With objVBMember.Collection.Parent
            AnalyseHTMLPrj .Parent.Collection.Parent, sMiddle
            AnalyseHTMLCmp .Parent, sMiddle
            AnalyseHTMLMember objVBMember, sMiddle
        End With
        EcrireFichier sMiddle, intNoFichier
    End If
End Sub

'Fonction modifiée par René Rhéaume le 20 juin 2002
' Support multilingue
Private Function ExtraitHTML(sTete As String, sFin As String, sMiddle As String, ByVal sSearch As String) As Boolean
    Static mlgMsgFmtNonValable As String
    If (LenB(mlgMsgFmtNonValable) = 0) Then
        mlgMsgFmtNonValable = LireChaineLocalisee(conNomForm, _
            "Code.ExtraitHTML.mlgMsgFmtNonValable", "Format non valable")
    End If
    
    Const conHTMLCmmtEnd As String = " -->"
    Dim iDeb As Long, iFin As Long
    Dim sSearchDeb As String, sSearchFin As String
    Dim lngLenDeb As Long

    On Error GoTo Fin
    sSearchDeb = "<!-- " & sSearch & conHTMLCmmtEnd
    sSearchFin = "<!-- /" & sSearch & conHTMLCmmtEnd
    iDeb = InStr(1, sMiddle, sSearchDeb)
    iFin = InStr(1, sMiddle, sSearchFin)

    Select Case conVrai
        Case iDeb = 0, iFin = 0
'            ExtraitHTML = conFaux
            MsgBox mlgMsgFmtNonValable, vbExclamation
        Case Else
            sTete = Left$(sMiddle, iDeb - 1)
'            sFin = Right$(sMiddle, Len(sMiddle) - iFin - Len(sSearchFin) + 1)
            sFin = Mid$(sMiddle, iFin + Len(sSearchFin))
            lngLenDeb = Len(sSearchDeb)
            sMiddle = Mid$(sMiddle, iDeb + lngLenDeb, iFin - iDeb - lngLenDeb)
            ExtraitHTML = conVrai
    End Select

    Exit Function
Fin:
'    ExtraitHTML = conFaux
End Function

' Modifié par René Rhéaume le 18 janvier 2002
' Transformation de la fonction en Sub
Private Sub RemplaceString(ByRef sString As String, ByVal sSearch As String, ByVal sValeur As String)
    Dim iPos As Long
    Dim lngLenSearch As String
    sSearch = "[c2i" & sSearch & "]"
    lngLenSearch = Len(sSearch)

    iPos = InStr(1, sString, sSearch)
    Do Until iPos = 0
        iPos = InStr(1, sString, sSearch)
        If (iPos > 0) Then
            sString = Left$(sString, iPos - 1) & sValeur & Right$(sString, Len(sString) - iPos - lngLenSearch + 1)
        End If
    Loop
    '    RemplaceString = sString
End Sub

'Modifié par René Rhéaume le 16 juin 2002
' Optimisé la gestion d'erreur
'Déplacé de GestFichier par René Rhéaume le 18 mai 2002
' Optimisation de taille mémoire indiquée dans les manuels en ligne
' ~/Guide de l'utilisateur/Partie 2/Optimisation et compatibilité/Optimisation de la taille/Réduction de la taille du code
'Ajout par René Rhéaume le 1er août 2001
' Retourne 0 s'il n'y a pas eu d'erreur et -1 s'il y en a eu une.
' Retourne le contenu du fichier dans le 2e argument
'Fonction modifiée par René Rhéaume le 20 juin 2002
' Support multilingue
Private Function LireFichierTexte(ByVal chnNomFichier As String, ByRef ContenuFichier As String) As Integer
    Static mlgMsgImpossOuvrir As String
    If (LenB(mlgMsgImpossOuvrir) = 0) Then
        mlgMsgImpossOuvrir = LireChaineLocalisee(conNomForm, _
            "Code.LireFichierTexte.mlgMsgImpossOuvrir", "Impossible d'ouvrir le fichier :")
    End If
    On Error GoTo GestErr
    LireFichierTexte = 0
    Screen.MousePointer = vbHourglass

    ' Ouvre le fichier indiqué.
    Open chnNomFichier For Input As #1
    ContenuFichier = Input(LOF(1), 1)
    Close #1
    Screen.MousePointer = vbDefault
    Exit Function

GestErr:
    MsgBox mlgMsgImpossOuvrir & conEsp & chnNomFichier, vbExclamation
    LireFichierTexte = -1
    Screen.MousePointer = vbDefault
End Function

'Déplacé de GestFichier par René Rhéaume le 18 mai 2002
' Optimisation de taille mémoire indiquée dans les manuels en ligne
' ~/Guide de l'utilisateur/Partie 2/Optimisation et compatibilité/Optimisation de la taille/Réduction de la taille du code
' Ajout par René Rhéaume le 1er août 2001
' Retourne 0 s'il n'y a pas eu d'erreur et -1 s'il y en a eu une.
'Fonction modifiée par René Rhéaume le 20 juin 2002
' Support multilingue
'Fonction modifiée par René Rhéaume le 30 juin 2002
Private Function EcrireFichier(ByVal strContenu As String, ByVal intNoFich As Integer) _
                                As Boolean
    ' Instructions d'entrée/sortie de VB
    On Error GoTo GestErr

    ' Écrit le contenu du paramètre dans le fichier enregistré.
    ' Le caractère point-virgule (;) supprime le retour de chariot à la fin.
    Print #intNoFich, strContenu;
    EcrireFichier = conVrai

SortieProc:
    Exit Function
GestErr:
    GererErrFichier intNoFich, "frmPrint.EcrireFichier"
    EcrireFichier = conFaux
End Function

'Fonction ajoutée par René Rhéaume le 30 juin 2002
Private Function OuvrirFichierAjout(ByVal strNomFich As String) As Integer
    On Error GoTo GestErr
    ' Ouvre le fichier.
    OuvrirFichierAjout = FreeFile
    Open strNomFich For Append As #OuvrirFichierAjout
    
SortieFunc:
    Exit Function
GestErr:
    GererErrFichier OuvrirFichierAjout, "frmPrint.OuvrirFichierAjout"
    OuvrirFichierAjout = -1
End Function

'Fonction ajoutée par René Rhéaume le 30 juin 2002
Private Function FermerFichier(ByVal intNoFich As Integer) As Boolean
    On Error GoTo GestErr
    Close #intNoFich
    FermerFichier = conVrai

SortieFunc:
    Exit Function
GestErr:
    GererErrFichier intNoFich, "frmPrint.FermerFichier"
    FermerFichier = conFaux
End Function

'Procédure ajoutée par René Rhéaume le 30 juin 2002
Private Sub GererErrFichier(ByVal intNoFich As Integer, ByVal strSource As String)
    Static mlgMsgVerrouille As String
    If (LenB(mlgMsgVerrouille) = 0) Then
        mlgMsgVerrouille = LireChaineLocalisee(conNomForm, _
            "Code.GererErrFichier.mlgMsgVerrouille", _
            "» est verrouillé par une autre application ou est en lecture seule.")
    End If
    
    Select Case Err.Number
        Case 52
            MsgBox mlgarFichExistePas(0) & "#" & intNoFich & mlgMsgVerrouille, vbExclamation
        Case Else
            GererErrInattendue , strSource
    End Select
End Sub

'Procédure ajoutée par René Rhéaume le 30 juin 2002
Private Sub SupprimerFichTemp()
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If
End Sub

'Procédure ajoutée par René Rhéaume le 30 juin 2002
Private Sub VoirFichApercu()
    Static mlgarErrShellExec() As String
    Static blnLangOK As Boolean
    Dim I As Long
    If (Not blnLangOK) Then
        StrArray mlgarErrShellExec, "Impossible de voir l'aperçu", "Retour de ShellExecute : "
        If (blnMultilingueActive) Then
            For I = 0 To 1
                mlgarErrShellExec(I) = LireChaineLocalisee(conNomForm, _
                        "Code.VoirFichApercu.mlgarErrShellExec(" & I & conPF, _
                        mlgarErrShellExec(I))
            Next I
        End If
        blnLangOK = conVrai
    End If
    
    Dim lngRep As Long
    lngRep = ShellExecute(0, conOpen, c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    Select Case lngRep
        Case Is <= 32
'        Case 0, 2, 3, 5, 8, 11, 26, 27, 28, 29, 30, 31, 32
            GererErrInattendue mlgarErrShellExec(0) & vbCrLf & "  " & _
                mlgarErrShellExec(1) & lngRep, "frmPrint.VoirFichApercu"
    End Select
End Sub
