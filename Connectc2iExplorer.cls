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

Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    '=================================================
    'connection du complément
    'on crée l'objet window et usrcomm
    '=================================================
    On Error GoTo GestErr

    ' Enregistre l'instance vb
    Set VBInstance = VBInst

    If (Not objWindow Is Nothing) Then
        ' En cours d'exécution, affiche seulement
        Show
        If (ConnectMode = vbext_cm_AfterStartup) Then
            ' Démarre à partir du Gestionaire de Compléments
            AddToAddInCommandBar
        End If
        Exit Sub
    End If

    ' Initialisation des variables globales
    strCheminApp = App.Path
    c2iHTMLFile = strCheminApp & "\html\prev.htm"
    c2iDataFileNameOrigine = strCheminApp & "\data\base.c2i"
    'Ajout par René Rhéaume, 28 juillet 2001
    c2iINIFile = strCheminApp & conCheminRelatifINI
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

    AddToAddInCommandBar

    Exit Sub

GestErr:
    MsgBox "IDTExtensibility_OnConnection : " & Err.Description & " : " & CStr(Err.Number)    ' & vbCrLf & CStr(lngLigneOK)
End Sub

'------------------------------------------------------
' Cette méthode supprime le complément de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    '=================================================
    'on se déconnecte, on supprime les objets inintéressants
    '=================================================
    On Error GoTo GestErr

    If (objWindow.Visible) Then                            'sauvegarde de la visibilité
        WritePrivateProfileString conSecGen, conClefVisible, "1", c2iINIFile
    Else
        WritePrivateProfileString conSecGen, conClefVisible, "0", c2iINIFile
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

    Exit Sub
GestErr:
    MsgBox "IDTExtensibility_OnDisconnection : " + Err.Description, vbInformation
End Sub

Private Function AddToAddInCommandBar() As Office.CommandBarControl
    Dim lngAfficherExplorer As Long
    '=================================================
    'On ajoute un bouton dans le menu Compléments
    '=================================================
    On Error GoTo GestErr

    #If VersionVB = 5 Then
        Set objBouton = VBInstance.CommandBars(1).Controls(9).Controls.Add
    #ElseIf VersionVB = 6 Then
        Set objBouton = VBInstance.CommandBars(1).Controls(11).Controls.Add
    #End If
    objBouton.Caption = "c2i Explorer"
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

    lngAfficherExplorer = GetPrivateProfileInt(conSecGen, conClefVisible, 0, strCheminApp & conCheminRelatifINI)
    Select Case lngAfficherExplorer
        Case 1
            Show
            objUD.Init
        Case Else
            WritePrivateProfileString conSecGen, conClefVisible, "0", strCheminApp & conCheminRelatifINI
    End Select

    Exit Function

GestErr:
    MsgBox "AddToAddInCommandBar : " & Err.Description, vbInformation
End Function

Private Sub Show()
    '=================================================
    'on montre le document par l'intermediaire de notre objet window
    '=================================================
    On Error GoTo GestErr

    objWindow.Visible = conVrai
    Exit Sub

GestErr:
    MsgBox "BoutonBar_Click : " & Err.Description, vbInformation
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

Private Sub CmpEvent_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    Debug.Print VBComponent.Name
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    '    AddToAddInCommandBar
End Sub

Private Sub prjEvent_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Set refEvent = VBInstance.Events.ReferencesEvents(VBProject)
    Set VBPrjSelected = VBProject
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
    SauveDuree VBProject

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

Private Sub refEvent_ItemAdded(ByVal Reference As VBIDE.Reference)
    objUD.AjouteReference VBInstance.ActiveVBProject, Reference
End Sub

'Sub AddToINI()
'  ' Ce code ajoute une entrée au fichier VbAddin.ini.
'  ' Ceci doit être exécuté depuis la fenêtre Exécution.
'  Debug.Print WritePrivateProfileString("Add-Ins32", "c2iexplorer.Connectc2iExplorer", "0", "vbaddin.ini")
'End Sub
