VERSION 5.00
Begin VB.Form frmDuree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modification de la durée"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDuree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1545
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Fermer"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "&Appliquer"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   660
      Width           =   735
   End
   Begin VB.TextBox txtMn 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   660
      Width           =   735
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblProjectName 
      BackStyle       =   0  'Transparent
      Caption         =   "NomProjet"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label lblTemps 
      BackStyle       =   0  'Transparent
      Caption         =   "s"
      Height          =   255
      Index           =   2
      Left            =   3240
      TabIndex        =   4
      Top             =   660
      Width           =   255
   End
   Begin VB.Label lblTemps 
      BackStyle       =   0  'Transparent
      Caption         =   "mn"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.Label lblTemps 
      BackStyle       =   0  'Transparent
      Caption         =   "h"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   660
      Width           =   375
   End
End
Attribute VB_Name = "frmDuree"
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

' Procédure modifiée par René Rhéaume le 2 août 2001, 18 août 2001
' et 5 janvier 2002
Private Sub cmdAppliquer_Click()
    Dim lngTime As Long, iNumPrj As Long
    Dim bTrouve As Boolean
    Dim lngFinDureePrj As Long
    Dim strNomProjet As String

    Call GestionLostFocus(txtH, 1000)
    Call GestionLostFocus(txtMn)
    Call GestionLostFocus(txtS)
    lngTime = CLng(txtH) * 3600 + CLng(txtMn) * 60 + CLng(txtS)
    strNomProjet = lblProjectName.Caption
    '    bTrouve = conFaux
    '    iNumPrj = 0
    lngFinDureePrj = UBound(DuréePrj)

    Do Until ((bTrouve) Or (iNumPrj > lngFinDureePrj))
        If (DuréePrj(iNumPrj).Filename = strNomProjet) Then
            bTrouve = conVrai
        Else
            iNumPrj = iNumPrj + 1
        End If
    Loop

    If (bTrouve) Then
        lngTime = lngTime - timeGetTime / 1000 + DuréePrj(iNumPrj).Durée
    End If

    '    SaveSetting App.EXEName, conSecDurees, strNomProjet, CStr(lngTime)
    Call WritePrivateProfileString(conSecDurees, strNomProjet, CStr(lngTime), c2iINIFile)
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub txtH_LostFocus()
    Call GestionLostFocus(txtH, 1000)
End Sub

Private Sub txtMn_LostFocus()
    Call GestionLostFocus(txtMn)
End Sub

Private Sub txtS_LostFocus()
    Call GestionLostFocus(txtS)
End Sub

' Procédure ajoutée par René Rhéaume le 5 janvier 2002
' ============================ ALGORITHME ============================
' - DÉCLARATION CONSTANTES
'     zeroANum <-- «0» : alphanumérique
' - DÉCLARATION VARIABLES
'     temps    <--  «» : alphanumérique
'     val      <--   0 : numérique
'     max      <--  60 : numérique
' - ENTRÉE
'     Lire temps, max
' - TRAITEMENT
'     Si (temps contient du numérique = Vrai) Alors
'       val <-- Convertir en numérique(temps)
'       Si ((val < 0) OU (val > max)) Alors
'         temps <-- zeroANum
'       Sinon
'         temps <-- Convertir en alphanumérique(val)
'       FinSi
'     Sinon
'       temps <-- zeroANum
'     FinSi
' - SORTIE
'     Afficher temps
' ============================ PROCÉDURE ============================
Private Sub GestionLostFocus(ByRef txtTemps As TextBox, Optional ByVal lngMax As Long = 60)
    ' -------------------- DÉCLARATION VARIABLES ---------------------
    Dim lngVal As Long
    Dim strTemps As String
    ' ENTRÉE FAITE PAR LES PARAMÈTRES DE LA PROCÉDURE
    ' --------------------- TRAITEMENT ET SORTIE ---------------------
    ' Vérification de l'argument lngMax
    If (lngMax < 0) Then
        Err.Raise 6
    End If
    ' Vérification générale
    strTemps = txtTemps.Text
    If (IsNumeric(strTemps)) Then
        lngVal = CLng(Int(Val(strTemps)))
        Select Case conVrai
            Case (lngVal < 0), lngVal > lngMax
                strTemps = conAZero
            Case Else
                strTemps = CStr(lngVal)
        End Select
    Else
        strTemps = conAZero
    End If
    txtTemps.Text = strTemps
End Sub

