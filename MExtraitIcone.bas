Attribute VB_Name = "MExtraitIcone"
' ***** BEGIN LICENSE BLOCK *****
' Version: MPL 1.1
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
Public Const sGuid As String = "A9532500-998B-11d5-B256-80025919427B    'Guid du usrcomm"
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

Public Type InfoProjet
    Filename As String
    Name As String
    Durée As Long
End Type
Public DuréePrj() As InfoProjet

Public c2iHTMLFile As String
Public c2iINIFile As String 'Ajout par René Rhéaume, 28 juillet 2001

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

Public c2iDataFileNameOrigine As String
Public c2iCurrentDataFileName As String
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
            ExtraitIconeProjet = "Inconnu"
    End Select

End Function

Public Function ExtraitIconeMembre(ByVal objMember As VBIDE.Member, ByRef CouleurPortee As Long) As String
    Select Case objMember.Type
        Case vbext_mt_Const
            ExtraitIconeMembre = "constante"
        Case vbext_mt_Event
            ExtraitIconeMembre = "evenement"
        Case vbext_mt_Method
            'extraction de la ligne de code
            If objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Proc) <> 1 Then
                ExtraitIconeMembre = "methode"
            Else
                ExtraitIconeMembre = "api"
            End If
        Case vbext_mt_Property
            ExtraitIconeMembre = "propriete"
        Case vbext_mt_Variable
            ExtraitIconeMembre = "variable"
        Case Else
            ExtraitIconeMembre = "Inconnu"
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
            ExtraitIconeComponent = "ActiveXDesigner"
        Case vbext_ct_ClassModule
            ExtraitIconeComponent = "ClassModule"
        Case vbext_ct_DocObject
            ExtraitIconeComponent = "DocObject"
        Case vbext_ct_MSForm
            ExtraitIconeComponent = "MSForm"
        Case vbext_ct_PropPage
            ExtraitIconeComponent = "PropPage"
        Case vbext_ct_RelatedDocument
            ExtraitIconeComponent = "RelatedDocument"
        Case vbext_ct_ResFile
            ExtraitIconeComponent = "ResFile"
        Case vbext_ct_StdModule
            ExtraitIconeComponent = "StdModule"
        Case vbext_ct_UserControl
            ExtraitIconeComponent = "UserControl"
        Case vbext_ct_VBForm
            If VBCmp.Properties("MDIChild") = conVrai Then
                ExtraitIconeComponent = "VBChildForm"
            Else
                ExtraitIconeComponent = "VBForm"
            End If
        Case vbext_ct_VBMDIForm
            ExtraitIconeComponent = "VBMDIForm"
        Case Else
            ExtraitIconeComponent = "Inconnu"
    End Select

End Function

Public Function GetDuree(ByVal sFileName As String) As Long
    Dim sM As String
    Dim I As Long, iNumPrj As Long

    iNumPrj = -1

    If sFileName = "" Then
        GetDuree = -1
        Exit Function
    End If

    On Error GoTo Suivant
    For I = 0 To UBound(DuréePrj)
        If DuréePrj(I).Filename = sFileName Then
            iNumPrj = I
            Exit For
        End If
    Next I

Suivant:
    On Error GoTo Fin

    ' Modification par René Rhéaume 2 août 2001
'    sM = GetSetting(App.EXEName, "Durees", sFileName, "")
    sM = LireChaineFichierINI("Durees", sFileName, "", c2iINIFile)
    If sM <> "" Then            'on la trouvé dans le fichier INI
        If iNumPrj = -1 Then
            GetDuree = CLng(sM)
        Else
            GetDuree = timeGetTime / 1000 - DuréePrj(iNumPrj).Durée + CLng(sM)
        End If
    Else                     'il n'est pas dans le fichier INI
        GetDuree = -1
    End If
    Exit Function

Fin:
    GetDuree = -1
End Function

'Public Function PositionForm(ByVal frmA As Form, Optional bDevant As Boolean = conVrai) As Long
'    If bDevant Then
'        PositionForm = SetWindowPos(frmA.hwnd, HWND_TOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
'    Else
'        PositionForm = SetWindowPos(frmA.hwnd, HWND_NOTOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
'    End If
'End Function

'Optimisé par René Rhéaume 2 septembre 2001
Public Sub TypePropriete(ByVal objCodeModule As VBIDE.CodeModule, ByRef lngStartLine As Long)
    'retourne la ligne de début de la propriété
    Dim lngDummy As Long, bTrouve As Boolean, sTexte As String
    'on stocke la valeur initiale
    lngDummy = lngStartLine

    bTrouve = conFaux
    On Error Resume Next
    Do Until bTrouve
        sTexte = objCodeModule.Lines(lngStartLine, 1)
        If Err.Number <> 0 Then Exit Do
        If lngStartLine > objCodeModule.CountOfLines Then Exit Do
        ' Si sTexte contient soit "Property Let", "Property Get" ou "Property Set"
        Select Case conVrai
            Case InStr(1, sTexte, "Property Let") > 0, InStr(1, sTexte, "Property Get") > 0, InStr(1, sTexte, "Property Set") > 0
                If Left$(sTexte, 1) <> "'" Then
                    bTrouve = conVrai
                    Exit Do
                End If
        End Select
        If RechercheString(sTexte, "End Property") Then
            bTrouve = conFaux
            Exit Do
        End If
        lngStartLine = lngStartLine + 1
    Loop
    'si bTrouve, on se casse
    If bTrouve Then Exit Sub
    'sinon, il faut faire une recherche arrière
    lngStartLine = lngDummy                                'on recommence du début
    Do Until bTrouve
        sTexte = objCodeModule.Lines(lngStartLine, 1)
        If Err.Number <> 0 Then Exit Do
        ' Si sTexte contient soit "Property Let", "Property Get" ou "Property Set"
        Select Case conVrai
            Case InStr(1, sTexte, "Property Let") > 0, InStr(1, sTexte, "Property Get") > 0, InStr(1, sTexte, "Property Set") > 0
                If Left$(sTexte, 1) <> "'" Then
                    bTrouve = conVrai
                    Exit Do
                End If
        End Select
        If RechercheString(sTexte, "End Property") Then
            lngStartLine = lngStartLine + 1
            Exit Do
        End If
        lngStartLine = lngStartLine - 1
    Loop

End Sub

Public Function RechercheString(ByVal sTexte As String, sM As String) As Boolean
'    RechercheString = conFaux
'    If InStr(1, sTexte, sM) > 0 Then                       'recherche
        If Left$(sTexte, Len(sM)) = sM Then
            RechercheString = conVrai
        End If
'    End If
End Function

Public Sub ConnectionInternet(ByVal sURL As String)
    Dim lngRep As Long
    lngRep = ShellExecute(0, "open", sURL, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then
        Select Case lngLanguage
            Case c2i_Langue_Français
                MsgBox "Impossible d'établir la connection Internet", vbInformation
            Case Else
                MsgBox "Impossible to connect to Internet", vbInformation
        End Select
    End If
End Sub

Public Sub SauveDuree(objPrj As VBProject)
    Dim lngDuree As Long

    On Error GoTo Fin
    lngDuree = GetDuree(objPrj.Filename)
    If lngDuree <> -1 Then
        ' Modification par René Rhéaume 2 août 2001
        Call WritePrivateProfileString("Durees", CStr(objPrj.Filename), CStr(lngDuree), c2iINIFile)
'        SaveSetting App.EXEName, "Durees", objPrj.Filename, lngDuree
    End If

Fin:
End Sub

'Ajout par René Rhéaume le 28 juillet 2001
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
