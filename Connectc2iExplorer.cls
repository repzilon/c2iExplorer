VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connectc2iExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "c2i Explorer VB5"
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

Implements IDTExtensibility

Private Const conClefVisible As String = "DisplayOnConnect"

'Public objWindowProp As Window 'c'est le container de notre document
'l'objet menu
Public objBouton As Office.CommandBarButton
'Public objBoutonPropriété As Office.CommandBarButton

'gestionnaire des clicks sur notre bouton
Public WithEvents BoutonBar As VBIDE.CommandBarEvents      'pour pouvoir gérer le click sur le bouton
Attribute BoutonBar.VB_VarHelpID = -1
'Public WithEvents BoutonBarProp As VBIDE.CommandBarEvents  'pour pouvoir gérer le click sur le bouton
'gestionnaire des évènement sur les components
Public WithEvents CmpEvent As VBIDE.VBComponentsEvents     'pour gérer les composants
Attribute CmpEvent.VB_VarHelpID = -1
'gestionnaire des projects
Public WithEvents prjEvent As VBIDE.VBProjectsEvents
Attribute prjEvent.VB_VarHelpID = -1
'gestionnaire des ajouts de références
Public WithEvents refEvent As VBIDE.ReferencesEvents
Attribute refEvent.VB_VarHelpID = -1
'Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)
    ' commentaire pour que la procédure soit compilée
End Sub

'Procédure modifiée par René Rhéaume le 29 avril 2002
' Meilleure conditions et vérification de l'existence
' de c2iExplorer.ini

'Procédure optimisée par René Rhéaume le 12 janvier 2003
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    '=================================================
    'connection du complément
    'on crée l'objet window et usrcomm
    '=================================================
'    Const conErrFichIntrouvable As Long = 53
    Const conNomMeth As String = "Connectc2iExplorer.IDTExtensibility_OnConnection"
    Dim strFichierLangue As String
    On Error GoTo GestErr
    ' Enregistre l'instance vb
    Set VBInstance = VBInst

    If (Not objWindow Is Nothing) Then
        Show ' En cours d'exécution, affiche seulement
        If (ConnectMode = vbext_cm_AfterStartup) Then
            ' Démarre à partir du Gestionaire de Compléments
            TerminerChargement
            AfficherExplorateur
        End If
    Else
        strCheminApp = App.Path
        'Ajout par René Rhéaume, 28 juillet 2001
        c2iINIFile = strCheminApp & conCheminRelatifINI
        Select Case conVrai
            Case FichierExiste(c2iINIFile), frmErreurINI.ObtenirActionINIAbsent
                ' Initialisation des variables globales
                c2iHTMLFile = strCheminApp & "\html\prev.htm"
                strFichierLangue = strCheminApp & "\Lang\" & _
                        LireChaineFichierINI(conSecGen, conValLang, vbNullString, c2iINIFile)
                If (ValiderFichierLangue(strFichierLangue)) Then
                    strFichLangueActuel = strFichierLangue
                    blnMultilingueActive = conVrai
                End If
                
                ChargerChainesGlobales
                'pour stocker les durées
                ReDim DuréePrj(0)
            
                'création de la fenêtre
                'objWindow est l'objet Window de vbe
                'objUD est le 'contenu' de la fenêtre
                With VBInstance.Windows
                    Set objWindow = .CreateToolWindow(AddInInst, "c2iExplorer.UDExplorer", conNomApp, sGuid, objUD)
                    Set objWindowBiblio = .CreateToolWindow(AddInInst, "c2iExplorer.UDBiblio", conNomApp, sGuidBiblio, objUDBiblio)
                    Set objWindowAddCode = .CreateToolWindow(AddInInst, "c2iExplorer.UDAddCode", conNomApp, sGuidAddCode, objUDAddCode)
                    Set objWindowClassViewer = .CreateToolWindow(AddInInst, "c2iExplorer.UDClassViewer", conNomApp, sGuidClassViewer, objUDClassViewer)
                    Set objWindowMsgBox = .CreateToolWindow(AddInInst, "c2iExplorer.UDMsgBox", conNomApp, sGuidMsgBox, objUDMsgBox)
                    Set objWindowCompteur = .CreateToolWindow(AddInInst, "c2iExplorer.UDCompteur", conNomApp, sGuidCompteur, objUDCompteur)
                End With
                Set Connection = Me
                TerminerChargement
            Case Else ' Abandon
                Set Connection = Nothing
                'Optimisé par René Rhéaume le 5 juillet 2002
'                Err.Raise conErrFichIntrouvable ', conNomApp & "." & conNomMeth, _
                            "c2iExplorer.ini est introuvable malgré l'intervention de l'utilsateur."
        End Select
    End If
    
    Exit Sub

GestErr:
'    If (Err.Number <> conErrFichIntrouvable) Then
        GererErrInattendue , conNomMeth
'    End If
End Sub

' Procédure modifiée par René Rhéaume le 29 avril 2002
' Élimination du message d'erreur si c2iExplorer.ini
' n'existe pas
'------------------------------------------------------
' Cette méthode supprime le complément de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    '=================================================
    'on se déconnecte, on supprime les objets inintéressants
    '=================================================
    On Error Resume Next
    If (Not objWindow Is Nothing) Then
        On Error GoTo GestErr
        If (objWindow.Visible) Then                            'sauvegarde de la visibilité
            WritePrivateProfileString conSecGen, conClefVisible, "1", c2iINIFile
        Else
            WritePrivateProfileString conSecGen, conClefVisible, conAZero, c2iINIFile
        End If
        
        objBouton.Delete
        'objBoutonPropriété.Delete
    
        Set Connection = Nothing
        Set objBouton = Nothing
        'Set objBoutonPropriété = Nothing
        Set objUD = Nothing
        Set objUDBiblio = Nothing
        Set objUDAddCode = Nothing
        Set objUDMsgBox = Nothing
        Set objUDCompteur = Nothing

        Set objWindow = Nothing
        Set objWindowBiblio = Nothing
        Set objWindowAddCode = Nothing
        Set objWindowMsgBox = Nothing
        Set objWindowCompteur = Nothing
        Set VBPrjSelected = Nothing
        Set VBInstance = Nothing
    End If

    Exit Sub
    
GestErr:
    GererErrInattendue , "Connectc2iExplorer.IDTExtensibility_OnDisconnection"
    Resume Next
End Sub

'Procédure modifiée par René Rhéaume le 22 janvier 2003
' On peut maintenant ancrer une fois la fenêtre de c2iExplorer et
' ne plus se soucier de la replacer à chaque démarrage de l'IDE.
Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    AfficherExplorateur
End Sub

' Procédure renommée par René Rhéaume le 12 janvier 2003
Private Sub TerminerChargement()
    '=================================================
    'On ajoute un bouton dans le menu Compléments
    '=================================================
    On Error GoTo GestErr

    #If VersionVB = 5 Then
        Set objBouton = VBInstance.CommandBars(1).Controls(9).Controls.Add
    #ElseIf VersionVB = 6 Then
        Set objBouton = VBInstance.CommandBars(1).Controls(11).Controls.Add
    #End If
    objBouton.Caption = conNomApp '"c2i Explorer"
    objBouton.Visible = conVrai
    'pour l'icone, on est obligé de passer par le presse papier
    Clipboard.SetData LoadResPicture(101, vbResBitmap)
    On Error Resume Next
    objBouton.PasteFace

    'Set objBoutonPropriété = VBInstance.CommandBars(1).Controls(11).Controls.Add
    'objBoutonPropriété.Caption = "c2i Propriétés"
    'objBoutonPropriété.Visible = conVrai
    'pour l'icone, on est obligé de passer par le presse papier
    'Clipboard.SetData LoadResPicture(102, vbResBitmap)
    'objBoutonPropriété.PasteFace

    'capture du click
    On Error GoTo GestErr
    With VBInstance.Events
        Set Me.BoutonBar = .CommandBarEvents(objBouton)
        'Set Me.BoutonBarProp = .CommandBarEvents(objBoutonPropriété)
        Set Me.CmpEvent = .VBComponentsEvents(Nothing)
        Set Me.prjEvent = .VBProjectsEvents
    End With
    
    Exit Sub

GestErr:
    GererErrInattendue , "Connectc2iExplorer.TerminerChargement"
End Sub

Private Sub Show()
    '=================================================
    'on montre le document par l'intermediaire de notre objet window
    '=================================================
    On Error GoTo GestErr

    objWindow.Visible = conVrai
    Exit Sub

GestErr:
    GererErrInattendue , "Connectc2iExplorer.Show"
End Sub

'Procédure ajoutée par René Rhéaume le 22 janvier 2003
Private Sub AfficherExplorateur()
    Dim lngAfficherExplorer As Long
    lngAfficherExplorer = GetPrivateProfileInt(conSecGen, conClefVisible, 0, strCheminApp & conCheminRelatifINI)
    Select Case lngAfficherExplorer
        Case 1
            Show
            objUD.Init
        Case Else
            WritePrivateProfileString conSecGen, conClefVisible, conAZero, strCheminApp & conCheminRelatifINI
    End Select
End Sub

Private Sub BoutonBar_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Show
End Sub

Private Sub CmpEvent_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    objUD.AjouteComponent VBComponent
End Sub

Private Sub CmpEvent_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    objUD.RetireComposant VBComponent
End Sub

Private Sub CmpEvent_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    objUD.RetireComposant VBComponent, OldName
    objUD.AjouteComponent VBComponent
End Sub

'Private Sub CmpEvent_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
'    Debug.Print VBComponent.Name
'End Sub

'Modifié par René Rhéaume le 23 février 2003
' Change de dossier lorsqu'on change de projet
Private Sub prjEvent_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Set refEvent = VBInstance.Events.ReferencesEvents(VBProject)
    Set VBPrjSelected = VBProject
    If (LenB(VBProject.Filename) > 0) Then
        VBInstance.LastUsedPath = VBProject.Filename
    End If
End Sub

Private Sub prjEvent_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Dim lngFinTabDurees As Long
    objUD.Init conVrai
    If (VBInstance.VBProjects.Count = 1) Then
        Set VBPrjSelected = VBProject
    End If
    lngFinTabDurees = UBound(DuréePrj)
    If (lngFinTabDurees = 0) Then
        DuréePrj(0).Durée = 0
        DuréePrj(0).Name = vbNullString
        DuréePrj(0).Filename = vbNullString
    End If
    lngFinTabDurees = lngFinTabDurees + 1
    ReDim Preserve DuréePrj(lngFinTabDurees)
    DuréePrj(lngFinTabDurees).Filename = VBProject.Filename
    DuréePrj(lngFinTabDurees).Name = VBProject.Name
    DuréePrj(lngFinTabDurees).Durée = timeGetTime / 1000
End Sub

Private Sub prjEvent_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Dim lngFinTabDurees As Long
    Dim I As Long
    lngFinTabDurees = UBound(DuréePrj)

    objUD.RetireProjet VBProject
    objUDCompteur.SauveDuree VBProject

    'suppression dans le tableau des projects pour la session
    For I = 0 To lngFinTabDurees
        If (VBProject.Name = DuréePrj(I).Name) Then
            DuréePrj(I).Filename = vbNullString
            DuréePrj(I).Name = vbNullString
            DuréePrj(I).Durée = 0
            Exit For
        End If
    Next I
End Sub

'Ajouté par René Rhéaume le 18 mai 2002
' Synchronise l'explorateur lorsque le projet est renommé
Private Sub prjEvent_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    objUD.Init conVrai
End Sub

Private Sub refEvent_ItemAdded(ByVal Reference As VBIDE.Reference)
    objUD.AjouteReference VBInstance.ActiveVBProject, Reference
End Sub

' Procédure ajoutée par René Rhéaume le 11 janvier 2003
Private Sub refEvent_ItemRemoved(ByVal Reference As VBIDE.Reference)
    objUD.RetireReference Reference
End Sub

'Sub AddToINI()
'  ' Ce code ajoute une entrée au fichier VbAddin.ini.
'  ' Ceci doit être exécuté depuis la fenêtre Exécution.
'  Debug.Print WritePrivateProfileString("Add-Ins32", "c2iExplorer.Connectc2iExplorer", conAZero, "vbaddin.ini")
'End Sub
