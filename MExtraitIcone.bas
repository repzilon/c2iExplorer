Attribute VB_Name = "MExtraitIcone"
Option Explicit
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Connection As Connectc2iExplorer

'génération de notrel Guid (avec l'utilitaire fournit avec vb dans le répertoire Tools
Public Const sGuid As String = "5F9236AD-D207-11d2-8133-0040056BD863"    'Guid du usrcomm
Public Const sGuidBiblio As String = "C20EBD8D-E5B7-11d2-A4C5-0040056BD863"
Public Const sGuidAddCode As String = "E462A14F-5602-11d3-A5EA-0040056BD863"
Public Const sGuidClassViewer As String = "E462A150-5602-11d3-A5EA-0040056BD863"
Public Const sGuidMsgBox As String = "531FEF91-74B9-11d3-A642-0040056BD863"
Public Const sGuidCompteur As String = "6B752121-C791-11d3-A7B4-0040056BD863"

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
    FileName As String
    Name As String
    Durée As Long
End Type

Public DuréePrj() As InfoProjet

Public c2iHTMLFile As String
Public c2iHeadFile As String
Public c2iProjectFile As String
Public c2iComponentFile As String
Public c2iTableFile As String
Public c2iMemberFile As String
Public c2iDescriptionFile As String
Public c2iLineFile As String

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

'1=Français
'2=Anglais
Public lngLanguage As Long

Public Enum c2iLanguage
    c2i_Langue_Français = 1
    c2i_Langue_Anglais = 2
End Enum

Public Const Mem_Privee As Long = vbBlack
Public Const Mem_Public As Long = vbBlue
Public Const Mem_Friend As Long = vbRed

Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
        ByVal hwnd As Long, _
        ByVal wMsg As Long, _
        ByVal wParam As Long, _
        lParam As Any) _
        As Long

Public Const WM_SETREDRAW = &HB
Public Declare Function GetLastError Lib "kernel32" () As Long

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

Public Function ExtraitIconeComponent(VBCmp As VBIDE.VBComponent) As String

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
            If VBCmp.Properties("MDIChild") = True Then
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

Public Function GetDuree(sFileName As String) As Long
    Dim sM As String
    Dim i As Long, iNumPrj As Long

    iNumPrj = -1

    If sFileName = "" Then
        GetDuree = -1
        Exit Function
    End If

    On Error GoTo Suivant
    For i = 0 To UBound(DuréePrj)
        If DuréePrj(i).FileName = sFileName Then
            iNumPrj = i
            Exit For
        End If
    Next i

Suivant:
    On Error GoTo Fin

    sM = GetSetting(App.EXEName, "Durees", sFileName, "")
    If sM <> "" Then                                       'on la trouvé dans la base de registre
        If iNumPrj = -1 Then
            GetDuree = CLng(sM)
        Else
            GetDuree = timeGetTime / 1000 - DuréePrj(iNumPrj).Durée + CLng(sM)
        End If
    Else                                                   'il n'est pas dans la base de registre
        GetDuree = -1
    End If
    Exit Function

Fin:
    GetDuree = -1
End Function

Public Function PositionForm(ByVal frmA As Form, Optional bDevant As Boolean = True) As Long
    If bDevant Then
        PositionForm = SetWindowPos(frmA.hwnd, HWND_TOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
    Else
        PositionForm = SetWindowPos(frmA.hwnd, HWND_NOTOPMOST, frmA.Left \ Screen.TwipsPerPixelX, frmA.Top \ Screen.TwipsPerPixelY, frmA.Width \ Screen.TwipsPerPixelX, frmA.Height \ Screen.TwipsPerPixelY, 0)
    End If
End Function




Public Function TypePropriete(ByVal objCodeModule As VBIDE.CodeModule, ByRef lngStartLine As Long)
    'retourne la ligne de début de la propriété
    Dim lngDummy As Long, bTrouve As Boolean, sTexte As String
    'on stocke la valeur initiale
    lngDummy = lngStartLine

    bTrouve = False
    On Error Resume Next
    Do Until bTrouve
        sTexte = objCodeModule.Lines(lngStartLine, 1)
        If Err.Number <> 0 Then Exit Do
        If lngStartLine > objCodeModule.CountOfLines Then Exit Do
        If InStr(1, sTexte, "Property Let") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If InStr(1, sTexte, "Property Get") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If InStr(1, sTexte, "Property Set") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If RechercheString(sTexte, "End Property") Then
            bTrouve = False
            Exit Do
        End If
        lngStartLine = lngStartLine + 1
    Loop
    'si bTrouve, on se casse
    If bTrouve Then Exit Function
    'sinon, il faut faire une recherche arrière
    lngStartLine = lngDummy                                'on recommence du début
    Do Until bTrouve
        sTexte = objCodeModule.Lines(lngStartLine, 1)
        If Err.Number <> 0 Then Exit Do
        If InStr(1, sTexte, "Property Let") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If InStr(1, sTexte, "Property Get") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If InStr(1, sTexte, "Property Set") > 0 Then
            If Left(sTexte, 1) <> "'" Then
                bTrouve = True
                Exit Do
            End If
        End If
        If RechercheString(sTexte, "End Property") Then
            lngStartLine = lngStartLine + 1
            Exit Do
        End If
        lngStartLine = lngStartLine - 1
    Loop

End Function

Public Function RechercheString(ByVal sTexte As String, sM As String) As Boolean
    RechercheString = False
    If InStr(1, sTexte, sM) > 0 Then                       'recherche
        If Left(sTexte, Len(sM)) = sM Then
            RechercheString = True
            Exit Function
        End If
    End If

End Function


Public Function ConnectionInternet(ByVal sURL As String)
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
End Function

Public Function SauveDuree(objPrj As VBProject)
    Dim sM As String, lngDuree As Long

    On Error GoTo Fin
    lngDuree = GetDuree(objPrj.FileName)
    If lngDuree <> -1 Then
        SaveSetting App.EXEName, "Durees", objPrj.FileName, lngDuree
    End If

Fin:
End Function


