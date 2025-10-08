Attribute VB_Name = "modImporterDurees"
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
' The Original Code is Utilitaire d'importation des durées de développement.
'
' The Initial Developer of the Original Code is
' René Rhéaume <rener@mediom.qc.ca>.
' Portions created by the Initial Developer are Copyright (C) 2003
' the Initial Developer. All Rights Reserved.
'
' ***** END LICENSE BLOCK *****

Option Explicit

Sub Main()
    DemarrerSupportMultilingue
    ImporterDurees
End Sub

'Ajouté par René Rhéaume 25 août 2001
' Importe les durées de la base de registres vers le fichier INI
' Utile pour les versions VB6 seulement
' Retourne 0 s'il n'y a pas d'erreur
'Fonction modifiée par René Rhéaume le 20 juin 2002
' Support multilingue
'Fonction modifiée par René Rhéaume le 14 septembre 2002
' Retrait du code de déboguage
'Fonction modifiée par René Rhéaume le 26 juin 2003
' - Transformation en procédure
' - Supprime la clé après l'importation
'Procédure déplacée par René Rhéaume le 2 juillet 2003
Private Sub ImporterDurees()
    Static mlgarErrImportation() As String
    Static blnLangOK As Boolean
    Dim J As Long
    If (Not blnLangOK) Then
        StrArray mlgarErrImportation, "Une erreur est survenue lors de l'importation des durées", _
            "de la base de registres vers le fichier INI."
        If (blnMultilingueActive) Then
            For J = 0 To 1
                mlgarErrImportation(J) = LireChaineLocalisee("modImporterDurees", _
                    "Code.ImporterDurees.mlgarErrImportation(" & J & ")", mlgarErrImportation(J))
            Next J
        End If
        blnLangOK = conVrai
    End If

    Const conSecDurees As String = "Durees"
    Dim tbSettings As Variant, I As Long
    Dim sFileName As String
    Dim chnDuree As String                ' Un peu bizarre, mais pas de conversions inutiles
    Dim lngFinTab As String

    On Error GoTo GestErr
    tbSettings = GetAllSettings(conNomApp, conSecDurees)

    If (Not IsEmpty(tbSettings)) Then
        lngFinTab = UBound(tbSettings)
        For I = 0 To lngFinTab
            sFileName = tbSettings(I, 0)
            chnDuree = tbSettings(I, 1)
            Call WritePrivateProfileString(conSecDurees, sFileName, chnDuree, c2iINIFile)
        Next I
    End If
    ' On peut maintenant supprimer la clé devenue obsolète.
    DeleteSetting App.EXEName

Fin:
    Exit Sub
GestErr:
    GererErrInattendue mlgarErrImportation(0) & vbNewLine & _
                        mlgarErrImportation(1), "UDCompteur.ImporterDurees"
End Sub

' Procédure ajoutée par René Rhéaume le 30 juin 2003
Private Sub DemarrerSupportMultilingue()
    Const conSecGen As String = "General"
    Const conValLang As String = "Language"
    Const conCheminRelatifINI As String = conDossierData & "\c2iExplorer.ini"
    Dim strFichierLangue As String
    
    strCheminApp = App.Path
    c2iINIFile = strCheminApp & conCheminRelatifINI
    strFichierLangue = strCheminApp & "\Lang\" & _
        LireChaineFichierINI(conSecGen, conValLang, vbNullString, c2iINIFile)
    If (ValiderFichierLangue(strFichierLangue)) Then
        strFichLangueActuel = strFichierLangue
        blnMultilingueActive = conVrai
    End If
End Sub
