Attribute VB_Name = "MExtraitIcone"
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
Public Connection As Connectc2iExplorer

'génération de notre GUID (avec l'utilitaire fournit avec vb dans le répertoire Tools
#If VersionVB = 6 Then
    Public Const sGuid As String = "5F9236AD-D207-11d2-8133-0040056BD863"    'Guid du usrcomm
    Public Const sGuidBiblio As String = "C20EBD8D-E5B7-11d2-A4C5-0040056BD863"
    Public Const sGuidAddCode As String = "E462A14F-5602-11d3-A5EA-0040056BD863"
    Public Const sGuidClassViewer As String = "E462A150-5602-11d3-A5EA-0040056BD863"
    Public Const sGuidMsgBox As String = "531FEF91-74B9-11d3-A642-0040056BD863"
    Public Const sGuidCompteur As String = "6B752121-C791-11d3-A7B4-0040056BD863"
#ElseIf VersionVB = 5 Then
    Public Const sGuid As String = "A9532500-998B-11d5-B256-80025919427B"    'Guid du usrcomm
    Public Const sGuidBiblio As String = "A9532501-998B-11d5-B256-80025919427B"
    Public Const sGuidAddCode As String = "A9532502-998B-11d5-B256-80025919427B"
    Public Const sGuidClassViewer As String = "A9532503-998B-11d5-B256-80025919427B"
    Public Const sGuidMsgBox As String = "A9532504-998B-11d5-B256-80025919427B"
    Public Const sGuidCompteur As String = "A9532505-998B-11d5-B256-80025919427B"
#End If

Public Const Mem_Privee As Long = vbBlack
Public Const Mem_Public As Long = vbBlue
Public Const Mem_Friend As Long = vbRed

Public Const conFaux As Boolean = False
Public Const conVrai As Boolean = True
Public Const conCheminRelatifINI As String = "\DATA\c2iExplorer.ini"
Public Const conDossierData As String = "\DATA"
Public Const conSecGen As String = "General"
Public Const conSecDurees As String = "Durees"
Public Const conValData As String = "Data"
Public Const conNomApp As String = "c2iExplorer"
Public Const conBS As String = "\"
Public Const conElement As String = "element"
Public Const conInconnu As String = "Inconnu"
Public Const conVariable As String = "variable"
Public Const conConstante As String = "constante"
Public Const conEvenement As String = "evenement"
Public Const conAPI As String = "api"
Public Const conMethode As String = "methode"
Public Const conPropriete As String = "propriete"
Public Const conActiveXDesigner As String = "ActiveXDesigner"
Public Const conClassModule As String = "ClassModule"
Public Const conDocObject As String = "DocObject"
Public Const conMSForm As String = "MSForm"
Public Const conPropPage As String = "PropPage"
Public Const conRelatedDocument As String = "RelatedDocument"
Public Const conResFile As String = "ResFile"
Public Const conStdModule As String = "StdModule"
Public Const conUserControl As String = "UserControl"
Public Const conVBChildForm As String = "VBChildForm"
Public Const conVBForm As String = "VBForm"
Public Const conVBMDIForm As String = "VBMDIForm"
Public Const conGet As String = "Get"
Public Const conLet As String = "Let"
Public Const conSet As String = "Set"
Public Const conLibelDescrpt As String = "Description"
Public Const conGlmt As String = """"
Public Const conMsgAucuneFenetreCodeActive As String = "Il n'y a pas de fenêtre de code active"
Public Const conFanionsOuvrir As Long = OFN_HIDEREADONLY + OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_EXPLORER + OFN_LONGNAMES
Public Const conAZero As String = "0"

'l'objet contenant notre UserControl
Public objUD As UDExplorer                                 'objet contenant le document créé
Public objUDBiblio As UDBiblio
Public objUDAddCode As UDAddCode
Public objUDClassViewer As UDClassViewer
Public objUDMsgBox As UDMsgBox
Public objUDCompteur As UDCompteur

'l'objet window, container de notre usercontrol
Public objWindow As Window                                 'c'est le container de notre document
Public objWindowBiblio As Window
Public objWindowAddCode As Window
Public objWindowClassViewer As Window
Public objWindowMsgBox As Window
Public objWindowCompteur As Window

Public VBInstance As VBIDE.VBE
Public VBPrjSelected As VBIDE.VBProject

Private Type InfoProjet
    Filename As String
    Name As String
    Durée As Long
End Type
Public DuréePrj() As InfoProjet

Public strCheminApp As String
Public c2iHTMLFile As String
Public c2iINIFile As String                                'Ajout par René Rhéaume, 28 juillet 2001
Public c2iDataFileNameOrigine As String
Public c2iCurrentDataFileName As String

'variables d'affichage dans le listview et le treeview
Public bAfficheConstante As Boolean
Public bAfficheVariable As Boolean
Public bAffichePropriete As Boolean
Public bAfficheMethode As Boolean
Public bAfficheEvenement As Boolean
Public bAfficheAPI As Boolean
Public bAffichePublic As Boolean
Public bAffichePrive As Boolean
Public bAfficheFriend As Boolean

Public Explorer As cExplorer

Public lngLanguage As c2iLanguage
Public Enum c2iLanguage
    c2i_Langue_Français = 1
    c2i_Langue_Anglais = 2
End Enum

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1

Public Declare Function GetInputState Lib "user32" () As Long

'Private Declare Function SetWindowPos Lib "user32" ( _
 '        ByVal hwnd As Long, _
 '        ByVal hWndInsertAfter As Long, _
 '        ByVal X As Long, _
 '        ByVal Y As Long, _
 '        ByVal cx As Long, _
 '        ByVal cy As Long, _
 '        ByVal wFlags As Long) As Long
'
'Private Const HWND_TOPMOST = -1
'Private Const HWND_NOTOPMOST = -2

Public Function ExtraitIconeProjet(ByVal objVBPrj As VBIDE.VBProject) As String

    Select Case objVBPrj.Type                              'image du projet
        Case vbext_pt_ActiveXControl
            ExtraitIconeProjet = "ActiveXControl"
        Case vbext_pt_ActiveXDll
            ExtraitIconeProjet = "ActiveXDll"
        Case vbext_pt_ActiveXExe
            ExtraitIconeProjet = "ActiveXExe"
        Case vbext_pt_StandardExe
            ExtraitIconeProjet = "VBproject"
        Case Else
            ExtraitIconeProjet = conInconnu
    End Select

End Function

Public Function ExtraitIconeMembre(ByVal objMember As VBIDE.Member, ByRef CouleurPortee As Long) As String
    Select Case objMember.Type
        Case vbext_mt_Const
            ExtraitIconeMembre = conConstante
        Case vbext_mt_Event
            ExtraitIconeMembre = conEvenement
        Case vbext_mt_Method
            'extraction de la ligne de code
            If objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Proc) <> 1 Then
                ExtraitIconeMembre = conMethode
            Else
                ExtraitIconeMembre = conAPI
            End If
        Case vbext_mt_Property
            ExtraitIconeMembre = conPropriete
        Case vbext_mt_Variable
            ExtraitIconeMembre = conVariable
        Case Else
            ExtraitIconeMembre = conInconnu
    End Select

    'détermination de la couleur suivant le type de membre
    Select Case objMember.Scope
        Case vbext_Friend
            CouleurPortee = Mem_Friend
        Case vbext_Private
            CouleurPortee = Mem_Privee
        Case vbext_Public
            CouleurPortee = Mem_Public
    End Select
End Function

Public Function ExtraitIconeComponent(ByVal VBCmp As VBIDE.VBComponent) As String

    Select Case VBCmp.Type
        Case vbext_ct_ActiveXDesigner
            ExtraitIconeComponent = conActiveXDesigner
        Case vbext_ct_ClassModule
            ExtraitIconeComponent = conClassModule
        Case vbext_ct_DocObject
            ExtraitIconeComponent = conDocObject
        Case vbext_ct_MSForm
            ExtraitIconeComponent = conMSForm
        Case vbext_ct_PropPage
            ExtraitIconeComponent = conPropPage
        Case vbext_ct_RelatedDocument
            ExtraitIconeComponent = conRelatedDocument
        Case vbext_ct_ResFile
            ExtraitIconeComponent = conResFile
        Case vbext_ct_StdModule
            ExtraitIconeComponent = conStdModule
        Case vbext_ct_UserControl
            ExtraitIconeComponent = conUserControl
        Case vbext_ct_VBForm
            If (VBCmp.Properties("MDIChild")) Then
                ExtraitIconeComponent = conVBChildForm
            Else
                ExtraitIconeComponent = conVBForm
            End If
        Case vbext_ct_VBMDIForm
            ExtraitIconeComponent = conVBMDIForm
        Case Else
            ExtraitIconeComponent = conInconnu
    End Select

End Function

'Public Function PositionForm(ByVal frmA As Form, Optional bDevant As Boolean = conVrai) As Long
'    If bDevant Then
'        PositionForm = SetWindowPos(frmA.hwnd, HWND_TOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
'    Else
'        PositionForm = SetWindowPos(frmA.hwnd, HWND_NOTOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
'    End If
'End Function

' Procédure modifiée par René Rhéaume le 24 avril 2002
' Support des adresses de courrier électronique
' en remplacement de UDBiblio.MailInternet
Public Sub ConnectionInternet(ByVal strAdresse As String, _
                                Optional ByVal blnCourriel As Boolean = conFaux)
    Const conOpen As String = "open"
    Dim lngRep As Long
    If (blnCourriel) Then
        lngRep = ShellExecute(0, conOpen, "mailto:" & strAdresse, vbNullString, vbNullString, SW_NORMAL)
    Else
        lngRep = ShellExecute(0, conOpen, strAdresse, vbNullString, vbNullString, SW_NORMAL)
    End If
    If (lngRep = 0) Then
        Select Case lngLanguage
            Case c2i_Langue_Anglais
                MsgBox "Impossible to connect to Internet", vbInformation
            Case Else
                MsgBox "Impossible d'établir la connection Internet", vbInformation
        End Select
    End If
End Sub

'Ajout par René Rhéaume en décembre 2001
'Routine d'émulation de la fonction Split de VB6 ultra-rapide
'Trouvé sur http://www.xbeat.net/vbspeed/
Public Sub SplitB(Expression$, ResultSplit$(), Optional Delimiter$ = " ")
    ' By Chris Lucas, cdl1051@earthlink.net, 20011208
    Dim c&, SLen&, DelLen&, tmp&, Results&()

    SLen = LenB(Expression) \ 2
    DelLen = LenB(Delimiter) \ 2

    ' Bail if we were passed an empty delimiter or an empty expression
    If SLen = 0 Or DelLen = 0 Then
        ReDim Preserve ResultSplit(0 To 0)
        ResultSplit(0) = Expression
        Exit Sub
    End If

    ' Count delimiters and remember their positions
    ReDim Preserve Results(0 To SLen)
    tmp = InStr(Expression, Delimiter)

    Do While tmp
        Results(c) = tmp
        c = c + 1
        tmp = InStr(Results(c - 1) + 1, Expression, Delimiter)
    Loop

    ' Size our return array
    ReDim Preserve ResultSplit(0 To c)

    ' Populate the array
    If c = 0 Then
        ' lazy man's call
        ResultSplit(0) = Expression
    Else
        ' typical call
        ResultSplit(0) = Left$(Expression, Results(0) - 1)
        For c = 0 To c - 2
            ResultSplit(c + 1) = Mid$(Expression, _
                    Results(c) + DelLen, _
                    Results(c + 1) - Results(c) - DelLen)
        Next c
        ResultSplit(c + 1) = Right$(Expression, SLen - Results(c) - DelLen + 1)
    End If
End Sub

'Procédure ajoutée par René Rhéaume le 8 juin 2002
' Gestionnaire centralisé des erreurs inattendues
Public Sub GererErrInattendue(Optional ByVal strMessage As String, _
                                Optional ByVal strSource As String)
    Const conCheminRelatifLog As String = "\DATA\c2iExplorer.log"
    Dim strCheminLog As String
    Dim strVersion As String
    
    strCheminLog = strCheminApp & conCheminRelatifLog
    If (LenB(strSource) = 0) Then
        strSource = Err.Source
    Else
        strSource = "c2iExplorer." & strSource
    End If
    If (LenB(strMessage) > 0) Then
        strMessage = "  " & strMessage & vbCrLf
    End If
    strMessage = strMessage & "  Erreur no " & Err.Number & vbCrLf & _
                    "  Description : " & Err.Description & vbCrLf & _
                    "  Source : " & strSource
    strVersion = App.Major & "." & App.Minor & "." & App.Revision
    
    With App
        .StartLogging strCheminLog, 2 'vbLogToFile
        .LogEvent vbCrLf & strMessage & vbCrLf & "  Version : " & strVersion & vbCrLf & _
                    "  Date/Heure : " & Now
    End With
    MsgBox strMessage, vbCritical
End Sub

Public Function AfficheMembre(ByVal objMember As VBIDE.Member) As Boolean
    '=========================================
    'Détermine si l'on affiche le membre sélectionné
    'en fonction du filtre général
    '=========================================

    AfficheMembre = conFaux                                'on n'affiche pas par default
    Select Case objMember.Type
        Case vbext_mt_Const
            AfficheMembre = bAfficheConstante
        Case vbext_mt_Event
            AfficheMembre = bAfficheEvenement
        Case vbext_mt_Method
            'extraction de la ligne de code
            If (objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Proc) <> 1) Then
                AfficheMembre = bAfficheMethode
            Else
                AfficheMembre = bAfficheAPI
            End If
        Case vbext_mt_Property
            AfficheMembre = bAffichePropriete
        Case vbext_mt_Variable
            AfficheMembre = bAfficheVariable
        Case Else
            AfficheMembre = conVrai
    End Select

    Select Case objMember.Scope
        Case vbext_Friend
            AfficheMembre = (AfficheMembre And bAfficheFriend)
        Case vbext_Private
            AfficheMembre = (AfficheMembre And bAffichePrive)
        Case vbext_Public
            AfficheMembre = (AfficheMembre And bAffichePublic)
    End Select

End Function

'Public Sub ExtraireImageList(ByVal imlAExtraire As ImageList, ByVal strDebutNomFich As String)
'    Dim lsiExtract As ListImage
'    For Each lsiExtract In imlAExtraire.ListImages
'        SavePicture lsiExtract.Picture, App.Path & conDossierData & _
'                    "\" & strDebutNomFich & "(" & lsiExtract.Key & ")[" & lsiExtract.Index & "].bmp"
'    Next
'    Set lsiExtract = Nothing
'End Sub
