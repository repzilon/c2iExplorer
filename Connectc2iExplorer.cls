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

    If Not objWindow Is Nothing Then
        ' En cours d'exécution, affiche seulement
        Show
        If ConnectMode = vbext_cm_AfterStartup Then
            ' Démarre à partir du Gestionaire de Compléments
            AddToAddInCommandBar
        End If
        Exit Sub
    End If

    'création de la fenêtre
    'objWindow est l'objet Window de vbe
    'objUD est le 'contenu' de la fenêtre
    Set objWindow = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDExplorer", "c2iExplorer", sGuid, objUD)
    Set objWindowBiblio = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDBiblio", "c2iExplorer", sGuidBiblio, objUDBiblio)
    Set objWindowAddCode = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDAddCode", "c2iExplorer", sGuidAddCode, objUDAddCode)
    Set objWindowClassViewer = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDClassViewer", "c2iExplorer", sGuidClassViewer, objUDClassViewer)
    Set objWindowMsgBox = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDMsgBox", "c2iExplorer", sGuidMsgBox, objUDMsgBox)
    Set objWindowCompteur = VBInstance.Windows.CreateToolWindow(AddInInst, "c2iExplorer.UDCompteur", "c2iExplorer", sGuidCompteur, objUDCompteur)
    Set Connection = Me

    AddToAddInCommandBar

    'pour stocker les durées
    ReDim DuréePrj(0)
    Exit Sub

GestErr:
    MsgBox "IDTExtensibility_OnConnection : " + Err.Description + " : " + CStr(Err.Number)

End Sub

'------------------------------------------------------
' Cette méthode supprime le complément de VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    '=================================================
    'on se déconnecte, on supprime les objets inintéressants
    '=================================================
    On Error GoTo GestErr

    If objWindow.Visible Then  'sauvegarde de la visibilité
        WritePrivateProfileString "General", "DisplayOnConnect", "1", c2iINIFile
    Else
        WritePrivateProfileString "General", "DisplayOnConnect", "0", c2iINIFile
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
    Set Me.BoutonBar = VBInstance.Events.CommandBarEvents(objBouton)
    'Set Me.BoutonBarProp = VBInstance.Events.CommandBarEvents(objBoutonPropriété)
    Set Me.CmpEvent = VBInstance.Events.VBComponentsEvents(Nothing)
    Set Me.prjEvent = VBInstance.Events.VBProjectsEvents

    lngAfficherExplorer = GetPrivateProfileInt("General", "DisplayOnConnect", 0, App.Path & "\data\c2iExplorer.ini")
    Select Case lngAfficherExplorer
        Case 1
            Show
            objUD.Init
        Case Else
            WritePrivateProfileString "General", "DisplayOnConnect", "0", App.Path & "\data\c2iExplorer.ini"
    End Select

    Exit Function

GestErr:
    MsgBox "AddToAddInCommandBar : " + Err.Description, vbInformation
End Function

Private Sub Show()
    '=================================================
    'on montre le document par l'intermediaire de notre objet window
    '=================================================
    On Error GoTo GestErr

    objWindow.Visible = conVrai
    Exit Sub

GestErr:
    MsgBox "BoutonBar_Click : " + "Err.Description, vbInformation"
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
    objUD.Init conVrai
    If VBInstance.VBProjects.Count = 1 Then
        Set VBPrjSelected = VBProject
    End If
    If UBound(DuréePrj) = 0 Then
        DuréePrj(0).Durée = 0
        DuréePrj(0).Name = ""
        DuréePrj(0).Filename = ""
    End If
    ReDim Preserve DuréePrj(UBound(DuréePrj) + 1)
    DuréePrj(UBound(DuréePrj)).Filename = VBProject.Filename
    DuréePrj(UBound(DuréePrj)).Name = VBProject.Name
    DuréePrj(UBound(DuréePrj)).Durée = timeGetTime / 1000
End Sub

Private Sub prjEvent_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    objUD.RetireProjet VBProject
    SauveDuree VBProject
    'suppression dans le tableau des projects pour la session
    Dim I As Long
    For I = 0 To UBound(DuréePrj)
        If VBProject.Name = DuréePrj(I).Name Then
            DuréePrj(I).Filename = ""
            DuréePrj(I).Name = ""
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
