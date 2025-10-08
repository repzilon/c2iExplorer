Attribute VB_Name = "modFonctions"
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
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Const conFaux As Boolean = False
Public Const conVrai As Boolean = True
Public Const conDossierData As String = "\DATA"
Public Const conNomApp As String = "c2iExplorer"

Public strCheminApp As String
Public c2iINIFile As String
Public strFichLangueActuel As String
Public blnMultilingueActive As Boolean

'Procédure ajoutée par René Rhéaume le 18 juin 2002
' Environ 12% plus rapide qu'Array
Public Sub StrArray(ByRef strarDest() As String, ParamArray vntarChaines() As Variant)
    Dim lngIdxFin As Long, I As Long
    lngIdxFin = UBound(vntarChaines)
    ReDim strarDest(0 To lngIdxFin) As String
    For I = 0 To lngIdxFin
        strarDest(I) = vntarChaines(I)
    Next I
End Sub

'Procédure ajoutée par René Rhéaume le 8 juin 2002
' Gestionnaire centralisé des erreurs inattendues
Public Sub GererErrInattendue(Optional ByVal strMessage As String, _
                                Optional ByVal strSource As String)
    Const conCheminRelatifLog As String = conDossierData & "\c2iExplorer.log"
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

'Fonction ajoutée par René Rhéaume le 18 juin 2002
' Tiré de VBSpeed http://www.xbeat.net/vbspeed
Private Function IsSameString(String1 As String, String2 As String) As Boolean
  If (LenB(String1) = LenB(String2)) Then
    IsSameString = (InStrB(1, String1, String2, vbBinaryCompare) <> 0)
  End If
End Function

'Fonction modifiée par René Rhéaume
'Nouvelle version le 2 juin 2002
' Inspiré de la fonction de Richard Clark dans MFileOp sur c2i.fr
Private Function FichierExiste(ByVal sNom As String) As Boolean
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

'Ajout par René Rhéaume le 28 juillet 2001
Public Function LireChaineFichierINI(ByVal chnSection As String, _
        ByVal chnClef As String, ByVal chnValeurDefaut As String, _
        ByVal chnFichierIni As String) As String

    Dim chnTemp As String, lngLongueurChaine As Long
    chnTemp = String$(512, 0)
    lngLongueurChaine = GetPrivateProfileString(chnSection, chnClef, chnValeurDefaut, chnTemp, Len(chnTemp), chnFichierIni)
    LireChaineFichierINI = Left$(chnTemp, lngLongueurChaine)
End Function

'Fonction ajoutée par René Rhéaume le 18 juin 2002
' Fait partie du système multilingue
Public Function ValiderFichierLangue(ByVal strFichierLangue As String) As Boolean
    Const conSecInfo As String = "@Info"
    Dim strVersion As String
    If (FichierExiste(strFichierLangue)) Then
        If (IsSameString(LireChaineFichierINI(conSecInfo, "ProgramName", vbNullString, strFichierLangue), conNomApp)) Then
            strVersion = App.Major & "." & App.Minor & "." & App.Revision
            If (IsSameString(LireChaineFichierINI(conSecInfo, "ProgramVer", vbNullString, strFichierLangue), strVersion)) Then
                ValiderFichierLangue = conVrai
             End If
        End If
    End If
End Function

'Ajouté par René Rhéaume le 18 juin 2002
' Fait partie du système multilingue
'Modifié par René Rhéaume le 30 juin 2002
Public Function LireChaineLocalisee(ByVal strModule As String, _
ByVal strClef As String, ByVal strValeurDefaut As String) As String
    LireChaineLocalisee = LireChaineFichierINI(strModule, strClef, strValeurDefaut, strFichLangueActuel)
End Function
