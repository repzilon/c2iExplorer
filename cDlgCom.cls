VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cDlgCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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
' René Rhéaume (rener@mediom.qc.ca)
'
' ***** END LICENSE BLOCK *****

Option Explicit
'// ********** ELABORE AVEC C2IEXPLORER ******************
'// richard.clark@wanadoo.fr
'// http://pro.wanadoo.fr/c2i
'//*****************************************************

'//*********************************
'Déclarations pour boite répertoire
'//*********************************
'Private Const BIF_RETURNONLYFSDIRS = 1
'Private Const BIF_DONTGOBELOWDOMAIN = 2

'Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
'Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

'Private Type BrowseInfo
'    hwndOwner      As Long
'    pIDLRoot       As Long
'    pszDisplayName As Long
'    lpszTitle      As Long
'    ulFlags        As Long
'    lpfnCallback   As Long
'    lParam         As Long
'    iImage         As Long
'End Type
'//*********************************
Private sFileName As String                                'nom du fichier
Private sFileTitle As String                               'titre du fichier
Private sDefaultEx As String                               'extension par default
Private sInitDir As String                                 'répertoire d'ouverture
Private lngFilterIndex As Long                             'index du filtre
Private sFilter As String                                  'filtre
Private lnghdcOwner As Long                                'propriétaire de la boite de dialogue
Private lngFlags As Long                                   'les flags
Private sDialogTitle As String                             'titre de la boite de dialogue
Private sFolder As String                                  'le dossier du fichier

'//*********************************
'Déclarations pour boite open et save
'//*********************************
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Enum ErrorConstants
    OF_CANCEL = &H800
    OF_CREATE = &H1000
    OF_DELETE = &H200
    OF_EXIST = &H4000
    OF_PROMPT = &H2000
    OF_READ = &H0
    OF_READWRITE = &H2
    OF_REOPEN = &H8000
    OF_SHARE_COMPAT = &H0
    OF_SHARE_DENY_NONE = &H40
    OF_SHARE_DENY_READ = &H30
    OF_SHARE_DENY_WRITE = &H20
    OF_SHARE_EXCLUSIVE = &H10
    OF_VERIFY = &H400
    OF_WRITE = &H1
End Enum
Public Enum FileOpenConstants
    OFN_ALLOWMULTISELECT = &H200
    OFN_CREATEPROMPT = &H2000
    OFN_ENABLEHOOK = &H20
    OFN_ENABLETEMPLATE = &H40
    OFN_ENABLETEMPLATEHANDLE = &H80
    OFN_EXPLORER = &H80000
    OFN_EXTENSIONDIFFERENT = &H400
    OFN_FILEMUSTEXIST = &H1000
    OFN_HIDEREADONLY = &H4
    OFN_LONGNAMES = &H200000
    OFN_NOCHANGEDIR = &H8
    OFN_NODEREFERENCELINKS = &H100000
    OFN_NOLONGNAMES = &H40000
    OFN_NONETWORKBUTTON = &H20000
    OFN_NOREADONLYRETURN = &H8000
    OFN_NOTESTFILECREATE = &H10000
    OFN_NOVALIDATE = &H100
    OFN_OVERWRITEPROMPT = &H2
    OFN_PATHMUSTEXIST = &H800
    OFN_READONLY = &H1
    OFN_SHAREAWARE = &H4000
    OFN_SHAREFALLTHROUGH = 2
    OFN_SHARENOWARN = 1
    OFN_SHAREWARN = 0
    OFN_SHOWHELP = &H10
End Enum

Public Function ShowOpen() As Boolean
    ShowOpen = ShowOpenSave(conVrai)
End Function

'Fonction modifiée par René Rhéaume le 20 juin 2002
' Support multilingue
'Fonction optimisée par René Rhéaume le 5 juillet 2002
Private Function ShowOpenSave(bOpen As Boolean) As Boolean
    Static mlgDebutDefaultFilter As String
    If (LenB(mlgDebutDefaultFilter) = 0) Then
        mlgDebutDefaultFilter = LireChaineLocalisee("cDlgCom", _
            "Code.ShowOpenSave.mlgDebutDefaultFilter", "Tous")
    End If
    
    Const conLongMaxNomFich As Integer = 260
    Dim tOpFile As OPENFILENAME, s As String
    Dim lngLenS As Long

    With tOpFile
        .lStructSize = Len(tOpFile)
        If (LenB(sFileTitle) = 0) Then sFileTitle = vbNullChar
        If (LenB(sFilter) = 0) Then sFilter = mlgDebutDefaultFilter & " (*.*)|*.*"
        If (lngFilterIndex = 0) Then lngFilterIndex = 1
        .hwndOwner = lnghdcOwner
        .lpstrInitialDir = sInitDir
        .lpstrDefExt = sDefaultEx
        If (LenB(sDialogTitle) > 0) Then .lpstrTitle = sDialogTitle

        ' Pour faire un style type Windows, remplacer les | par des nuls
        Dim I As Long
        s = sFilter
        lngLenS = Len(s)
        For I = 1 To lngLenS
            If (Mid$(s, I, 1) = "|") Then
                Mid$(s, I, 1) = vbNullChar
            End If
        Next
        ' Mettre un double nul à la fin
        s = s & (vbNullChar & vbNullChar)
        .lpstrFilter = s

        .nFilterIndex = lngFilterIndex

        ' Bourrer les tampons fichier et "title" fichier jusqu'à la longueur maximum du chemin
        s = sFileName & String$(conLongMaxNomFich - Len(sFileName), vbNullChar)
        .lpstrFile = s
        .nMaxFile = conLongMaxNomFich
        s = sFileTitle & String$(conLongMaxNomFich - Len(sFileTitle), vbNullChar)
        .lpstrFileTitle = s
        .nMaxFileTitle = conLongMaxNomFich

        ' Passer les drapeaux, en éliminant les drapeaux non VB
        .Flags = lngFlags And &H1FF1F

        ' Optimisé en taille par René Rhéaume le 5 janvier 2002
        If bOpen Then
            ShowOpenSave = CBool(GetOpenFileName(tOpFile))
        Else
            ShowOpenSave = CBool(GetSaveFileName(tOpFile))
        End If
        If (ShowOpenSave) Then
            sFileName = Left$(.lpstrFile, InStr(.lpstrFile, vbNullChar) - 1)
            sFileTitle = Left$(.lpstrFileTitle, InStr(.lpstrFileTitle, vbNullChar) - 1)
            sFolder = Left$(sFileName, Len(sFileName) - Len(sFileTitle) - 1)
            lngFlags = .Flags
        Else
            sFileName = vbNullString
            sFileTitle = vbNullString
            sFolder = vbNullString
            lngFlags = 0
        End If
    End With
End Function

Public Function ShowSave() As Boolean
    ShowSave = ShowOpenSave(conFaux)
End Function

'Public Function ShowOpenFolder() As Boolean
'Dim lngRet As Long, sDummy As String
'Dim tBrowseInfo As BrowseInfo
'
'
'With tBrowseInfo
'    .hwndOwner = lnghdcOwner
'    .lpszTitle = lstrcat(sDialogTitle, vbNullString)
'    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
'End With
'
'lngRet = SHBrowseForFolder(tBrowseInfo)
'
'If lngRet Then
'    sDummy = Space$(255)
'    SHGetPathFromIDList lngRet, sDummy
'    sFolder = Left$(sDummy, InStr(sDummy, vbNullChar) - 1)
'    ShowOpenFolder = conVrai
'Else
'    ShowOpenFolder = conFaux
'End If
'
'
'End Function

'//****************LES PROPRIÉTÉS**************************

Property Get Folder() As String
    Folder = sFolder
End Property

Property Get Filename() As String
    Filename = sFileName
End Property

Property Let Filename(sFileNameA As String)
    sFileName = sFileNameA
End Property

Property Get DialogTitle() As String
    DialogTitle = sDialogTitle
End Property

Property Let DialogTitle(sDialogTitleA As String)
    sDialogTitle = sDialogTitleA
End Property

Property Get Flags() As FileOpenConstants
    Flags = lngFlags
End Property

Property Let Flags(lngFlagsA As FileOpenConstants)
    lngFlags = lngFlagsA
End Property

Property Get hdcOwner() As Long
    hdcOwner = lnghdcOwner
End Property

Property Let hdcOwner(lnghdcOwnerA As Long)
    lnghdcOwner = lnghdcOwnerA
End Property

Property Get Filter() As String
    Filter = sFilter
End Property

Property Let Filter(sFilterA As String)
    sFilter = sFilterA
End Property

Property Get FilterIndex() As Long
    FilterIndex = lngFilterIndex
End Property

Property Let FilterIndex(lngFilterIndexA As Long)
    lngFilterIndex = lngFilterIndexA
End Property

Property Get InitDir() As String
    InitDir = sInitDir
End Property

Property Let InitDir(sInitDirA As String)
    sInitDir = sInitDirA
End Property

Property Get FileTitle() As String
    FileTitle = sFileTitle
End Property

Property Get DefaultEx() As String
    DefaultEx = sDefaultEx
End Property

Property Let DefaultEx(sDefaultExA As String)
    sDefaultEx = sDefaultExA
End Property
