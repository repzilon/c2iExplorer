VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3480
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAutosaveLibrary 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cboLangPack 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CheckBox chkUseLangPack 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox txtSaveInterval 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "0"
      Top             =   840
      Width           =   375
   End
   Begin VB.CheckBox chkSaveDurations 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label lblNote 
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   2100
      Width           =   3255
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   120
      X2              =   3420
      Y1              =   1215
      Y2              =   1215
   End
   Begin VB.Line linSep 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   120
      X2              =   3420
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblSaveInterval 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'Procédure ajoutée par René Rhéaume le 21 juin 2002
'Procédure modifiée par René Rhéaume le 16 avril 2003
Private Sub Form_Load()
    ChargerChainesLocales
    RemplirCombo
    SelectionnerCombo
    chkAutosaveLibrary.Value = Abs(blnSauvAutoBiblCode)
    chkSaveDurations.Value = Abs(bSauve)
    txtSaveInterval.Text = CStr(lngDureeSauve)
    chkSaveDurations_Click
    chkUseLangPack_Click
    Show
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
Private Sub chkSaveDurations_Click()
    Dim blnActiver As Boolean
    blnActiver = CBool(chkSaveDurations.Value)
    lblSaveInterval.Enabled = blnActiver
    txtSaveInterval.Enabled = blnActiver
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
Private Sub chkUseLangPack_Click()
    cboLangPack.Enabled = CBool(chkUseLangPack.Value)
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
'Procédure modifiée par René Rhéaume le 16 avril 2003
' Sauvegarde automatique de la bibliothèque de code optionnelle
Private Sub cmdOK_Click()
    Dim intSauverBiblio As Integer
    Dim intSauverDurees As Integer
    Dim lngIntervalle As Long
    Dim strFichierLangue As String
    
    intSauverBiblio = chkAutosaveLibrary.Value
    intSauverDurees = chkSaveDurations.Value
    lngIntervalle = CLng(txtSaveInterval.Text)
    If (chkUseLangPack.Value = 1) Then
        strFichierLangue = cboLangPack.List(cboLangPack.ListIndex)
    End If
    
    WritePrivateProfileString conSecGen, conValSauvBibl, CStr(intSauverBiblio), c2iINIFile
    WritePrivateProfileString conSecGen, conValSauveDur, CStr(intSauverDurees), c2iINIFile
    WritePrivateProfileString conSecGen, conValDurSavNtrv, CStr(lngIntervalle), c2iINIFile
    WritePrivateProfileString conSecGen, conValLang, strFichierLangue, c2iINIFile
    
    blnSauvAutoBiblCode = CBool(intSauverBiblio)
    bSauve = CBool(intSauverDurees)
    lngDureeSauve = lngIntervalle
    
    Unload Me
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
Private Sub txtSaveInterval_KeyPress(KeyAscii As Integer)
' 48-57 «0» - «9»
'  8    Retour arrière
'  9    Tabulation
' 13    Entrée
' 27    Echap
    Dim blnCarControle As Boolean
    blnCarControle = (KeyAscii = 27) + (KeyAscii = 13) + (KeyAscii = 8) + (KeyAscii = 9)
    If (Not blnCarControle) Then
        Select Case KeyAscii
            Case Is < 48, Is > 57
                KeyAscii = 0
        End Select
    End If
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
'Procédure modifiée par René Rhéaume le 18 avril 2003
' Utilise la fonction Dir au lieu des API
' Légèrement plus lent mais plus petit et plus simple
Private Sub RemplirCombo()
    Dim strCheminLang As String
    Dim strFichLang As String
    
'    cboLangPack.AddItem vbNullString
    
    strCheminLang = App.Path & "\Lang\"
    strFichLang = Dir(strCheminLang & "*.lng")
    Do While (LenB(strFichLang) > 0)
        If (ValiderFichierLangue(strCheminLang & strFichLang)) Then
            cboLangPack.AddItem strFichLang
        End If
        strFichLang = Dir
    Loop
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
'Procédure modifiée par René Rhéaume le 16 avril 2003
Private Sub ChargerChainesLocales()
    Const conNomForm As String = "frmOptions"
    
    Caption = LireChaineLocalisee(conNomForm, conL10nWCap, "Options")
    chkAutosaveLibrary.Caption = LireChaineLocalisee(conNomForm, _
        "Obj.chkAutosaveLibrary.Caption", "Sauvegarder automatiquement la biblio")
    chkSaveDurations.Caption = LireChaineLocalisee(conNomForm, _
        "Obj.chkSaveDurations.Caption", "Sauvegarder les durées")
    lblSaveInterval.Caption = LireChaineLocalisee(conNomForm, _
        "Obj.lblSaveInterval.Caption", "Intervalle de sauvegarde des durées :")
    txtSaveInterval.ToolTipText = LireChaineLocalisee(conNomForm, _
        "Obj.txtSaveInterval.ToolTipText", "En minutes")
    chkUseLangPack.Caption = LireChaineLocalisee(conNomForm, _
        "Obj.chkUseLangPack.Caption", "Utiliser le fichier de traduction suivant :")
    lblNote.Caption = LireChaineLocalisee(conNomForm, _
        "Obj.lblNote.Caption", _
        "Le changement de langue s'appliquera au prochain démarrage de Visual Basic.")
    cmdOK.Caption = LireChaineLocalisee(conNomForm, "Obj.cmdOK.Caption", "&OK")
    cmdCancel.Caption = LireChaineLocalisee(conNomForm, "Obj.cmdCancel.Caption", "&Annuler")
End Sub

'Procédure ajoutée par René Rhéaume le 21 juin 2002
'Procédure modifiée par René Rhéaume le 3 janvier 2003
' Les changements au niveau de la langue persistent à chaque affichage
Private Sub SelectionnerCombo()
    Dim strFichierLangue As String
    Dim lngNbElem As String
    Dim lngIdx As Long
    Dim blnTrouve As Boolean
    
'    If (blnMultilingueActive) Then
        strFichierLangue = LireChaineFichierINI(conSecGen, conValLang, vbNullString, c2iINIFile)
        lngNbElem = cboLangPack.ListCount
        
        Do Until ((blnTrouve) Or (lngIdx >= lngNbElem))
            If (IsSameString(cboLangPack.List(lngIdx), strFichierLangue)) Then
                cboLangPack.ListIndex = lngIdx
                chkUseLangPack.Value = 1
                blnTrouve = conVrai
            Else
                lngIdx = lngIdx + 1
            End If
        Loop
'    End If
End Sub
