Attribute VB_Name = "GestFichier"
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
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

'Module d'exploitation des fichiers

'Fonction modifiée par René Rhéaume
'Nouvelle version le 2 juin 2002
' Inspiré de la fonction de Richard Clark dans MFileOp sur c2i.fr
Public Function FichierExiste(ByVal sNom As String) As Boolean
    On Error GoTo GestErr
    Select Case conFaux
        Case (LenB(sNom)), (LenB(Dir(sNom))), Not (GetAttr(sNom) And vbDirectory)
        Case Else
            FichierExiste = conVrai
    End Select
    Exit Function

GestErr:
'   FichierExiste = conFaux
End Function

'Mise en commentaire par René Rhéaume le 26 juillet 2001
'Ces fonctions étaient inutilisées
'Public Function HourGlass(obj As Object)
'    Static ordMouse As Integer, fOn As Boolean
'    On Error Resume Next
'    If Not fOn Then
'        ' Sauver le pointeur et activer le sablier hourglass
'        ordMouse = obj.MousePointer
'        obj.MousePointer = vbHourglass
'        fOn = conVrai
'    Else
'        ' Restaurer le pointeur
'        obj.MousePointer = ordMouse
'        fOn = conFaux
'    End If
'    Debug.Assert Err = 0
'End Function
'
'Function NormalizePath(sPath As String) As String
'    NormalizePath = sPath
'    If Right$(sPath, 1) <> conBS Then NormalizePath = sPath & conBS
'End Function
'
'Function GetExtPos(sSpec As String) As Integer
'    Dim iLast As Integer, iExt As Integer
'    iLast = Len(sSpec)
'
'    ' Analyser à reculons pour trouver l'extension ou le nom de base
'    For iExt = iLast + 1 To 1 Step -1
'        Select Case Mid$(sSpec, iExt, 1)
'        Case "."
'            ' Le premier . à partir de la droite est le début de l'extension
'            Exit For
'        Case conBS
'            ' Le premier \ à partir de la droite est le début du nom
'            iExt = iLast + 1
'            Exit For
'        End Select
'    Next
'
'    ' Un retour négatif indique l'absence d'extension, mais
'    ' voici le nom de base pour que l'appelant n'ait pas à ré-analyser.
'    GetExtPos = iExt
'End Function
'
'Public Function VérifFileName(sFile As String, sExt As String) As String
'
''// vérification si l'extension du fichier est bien enregistrée
'If Len(sExt) <> 3 Then
'    Debug.Print "VérifFileName Erreur"
'End If
'
'VérifFileName = sFile
'
'If UCase(Right$(sFile, 4)) <> "." + UCase(sExt) Then
'    VérifFileName = sFile + "." + sExt
'End If
'
'End Function
'
'Public Function RepertoireExiste(sChemin As String, sRep As String) As Boolean
'Dim MonNom As String
'
'RepertoireExiste = conFaux
'MonNom = Dir(sChemin, vbDirectory)    ' Sélectionne la première entrée.
'Do While MonNom <> vbNullString       ' Début de la boucle.
'    ' Ignore le répertoire courant et le répertoire supérieur.
'    If MonNom <> "." And MonNom <> ".." Then
'        ' Utilise une comparaison bit à bit pour s'assurer que MonNom
'        ' est un répertoire.
'        If (GetAttr(sChemin & MonNom) And vbDirectory) = vbDirectory Then
'            If UCase(MonNom) = UCase(sRep) Then
'                RepertoireExiste = conVrai 'on a trouvé le répertoire
'                Exit Function
'            End If
'        End If
'    End If
'    MonNom = Dir        ' Sélectionne l'entrée suivante.
'Loop
'End Function

'Ajout par René Rhéaume le 28 juillet 2001
Public Function LireChaineFichierINI(ByVal chnSection As String, _
        ByVal chnClef As String, ByVal chnValeurDefaut As String, _
        ByVal chnFichierIni As String) As String

    Dim chnTemp As String, lngLongueurChaine As Long
    chnTemp = String$(512, 0)
    lngLongueurChaine = GetPrivateProfileString(chnSection, chnClef, chnValeurDefaut, chnTemp, Len(chnTemp), chnFichierIni)
    LireChaineFichierINI = Left$(chnTemp, lngLongueurChaine)
End Function

'Ajout par René Rhéaume le 28 juillet 2001
Public Function LireSectionFichierINI(ByVal chnSection As String, _
        ByVal chnFichierIni As String) As String

    Dim chnTemp As String, lngLongueurChaine As Long
    chnTemp = String$(2048, 0)
    lngLongueurChaine = GetPrivateProfileSection(chnSection, chnTemp, Len(chnTemp), chnFichierIni)
    LireSectionFichierINI = Left$(chnTemp, lngLongueurChaine)
End Function
