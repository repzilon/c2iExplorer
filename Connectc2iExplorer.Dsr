VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connectc2iExplorer 
   ClientHeight    =   10050
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   9795
   _ExtentX        =   17277
   _ExtentY        =   17727
   _Version        =   393216
   Description     =   "Explorateur d'évènements pour vb"
   DisplayName     =   "c2iExplorer"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "Connectc2iExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


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

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    '=================================================
    'connection du complément
    'on crée l'objet window et usrcomm
    '=================================================

    On Error GoTo GestErr
    ' Enregistre l'instance vb
    Set VBInstance = Application

    If Not objWindow Is Nothing Then
        ' En cours d'exécution, affiche seulement
        Show
        If ConnectMode = ext_cm_AfterStartup Then
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
    MsgBox "AddinInstance_OnConnection : " + Err.Description + " : " + CStr(Err.Number)

End Sub

'------------------------------------------------------
' Cette méthode supprime le complément de VB
'------------------------------------------------------
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
    '=================================================
    'on se déconnecte, on supprime les objets initeressants
    '=================================================
    On Error GoTo GestErr

    If objWindow.Visible Then                              'sauvegarde de la visibilité
        SaveSetting "MesCompléments", App.Title, "DisplayOnConnect", "1"
    Else
        SaveSetting "MesCompléments", App.Title, "DisplayOnConnect", "0"
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
    MsgBox "AddinInstance_OnDisconnection : " + Err.Description, vbInformation
End Sub

Private Function AddToAddInCommandBar() As Office.CommandBarControl
    '=================================================
    'on ajoute un bouton dans la barre d'outil edition
    '=================================================
    On Error GoTo GestErr
    VBInstance.CommandBars(2).Visible = True

    Set objBouton = VBInstance.CommandBars(1).Controls(11).Controls.Add
    objBouton.Caption = "c2i Explorer"
    objBouton.Visible = True
    'pour l'icone, on est obligé de passer par le presse papier
    Clipboard.SetData LoadResPicture(101, vbResBitmap)
    objBouton.PasteFace

    'Set objBoutonPropriété = VBInstance.CommandBars(1).Controls(11).Controls.Add
    'objBoutonPropriété.Caption = "c2i Propriétés"
    'objBoutonPropriété.Visible = True
    'pour l'icone, on est obligé de passer par le presse papier
    'Clipboard.SetData LoadResPicture(102, vbResBitmap)
    'objBoutonPropriété.PasteFace

    'capture du click
    Set Me.BoutonBar = VBInstance.Events.CommandBarEvents(objBouton)
    'Set Me.BoutonBarProp = VBInstance.Events.CommandBarEvents(objBoutonPropriété)
    Set Me.CmpEvent = VBInstance.Events.VBComponentsEvents(Nothing)
    Set Me.prjEvent = VBInstance.Events.VBProjectsEvents

    If GetSetting("MesCompléments", App.Title, "DisplayOnConnect", "0") = "1" Then
        Show
        objUD.Init
    End If

    Exit Function

GestErr:
    MsgBox "AddInCommandBar : " + Err.Description, vbInformation

End Function

Private Sub Show()
    '=================================================
    'on montre le document par l'intermediaire de notre objet window
    '=================================================
    On Error GoTo GestErr

    objWindow.Visible = True
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

Private Sub prjEvent_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    Set refEvent = VBInstance.Events.ReferencesEvents(VBProject)
    Set VBPrjSelected = VBProject
End Sub

Private Sub prjEvent_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    objUD.Init True
    If VBInstance.VBProjects.Count = 1 Then
        Set VBPrjSelected = VBProject
    End If
    If UBound(DuréePrj) = 0 Then
        DuréePrj(0).Durée = 0
        DuréePrj(0).Name = ""
        DuréePrj(0).FileName = ""
    End If
    ReDim Preserve DuréePrj(UBound(DuréePrj) + 1)
    DuréePrj(UBound(DuréePrj)).FileName = VBProject.FileName
    DuréePrj(UBound(DuréePrj)).Name = VBProject.Name
    DuréePrj(UBound(DuréePrj)).Durée = timeGetTime / 1000
End Sub

Private Sub prjEvent_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    objUD.RetireProjet VBProject
    SauveDuree VBProject
    'suppression dans le tableau des projects pour la session
    Dim i As Long
    For i = 0 To UBound(DuréePrj)
        If VBProject.Name = DuréePrj(i).Name Then
            DuréePrj(i).FileName = ""
            DuréePrj(i).Name = ""
            DuréePrj(i).Durée = 0
            Exit For
        End If
    Next i
End Sub

Private Sub refEvent_ItemAdded(ByVal Reference As VBIDE.Reference)
    objUD.AjouteReference VBInstance.ActiveVBProject, Reference
End Sub
