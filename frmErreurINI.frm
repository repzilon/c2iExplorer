VERSION 5.00
Begin VB.Form frmErreurINI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Le fichier c2iExplorer.ini n'a pas été trouvé"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErreurINI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8610
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbandonner 
      Cancel          =   -1  'True
      Caption         =   "&Abandonner"
      Height          =   375
      Left            =   5018
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdRegenerer 
      Caption         =   "&Recréer"
      Default         =   -1  'True
      Height          =   375
      Left            =   3698
      TabIndex        =   2
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdChercher 
      Caption         =   "&Chercher"
      Height          =   375
      Left            =   2378
      TabIndex        =   1
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      Height          =   1575
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   7635
   End
   Begin VB.Image imgIcone 
      Height          =   585
      Left            =   120
      Picture         =   "frmErreurINI.frx":000C
      Top             =   120
      Width           =   585
   End
End
Attribute VB_Name = "frmErreurINI"
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
Private intRetourMessage As Integer
Private strCheminData As String

Private Sub cmdAbandonner_Click()
    Unload Me
End Sub

Private Sub cmdChercher_Click()
    On Error GoTo Fin
'    Const conTitreDialogue As String = "Sélectionner"
'    Const conDebutFiltre As String = "Paramètres de c2iExplorer"

    Dim cCommonDialog As cDlgCom
    Set cCommonDialog = New cDlgCom
    With cCommonDialog
        .DefaultEx = ".ini"
'        .DialogTitle = conTitreDialogue & " c2iExplorer.ini"
'        .Filter = conDebutFiltre & " (c2iExplorer.ini)|c2iExplorer.ini"
        .DialogTitle = "Sélectionner c2iExplorer.ini"
        .Filter = "Paramètres de c2iExplorer (c2iExplorer.ini)|c2iExplorer.ini"
        .FilterIndex = 1
        .Flags = conFanionsOuvrir
        .InitDir = strCheminApp & conDossierData
        .ShowOpen
        If (.Filename <> vbNullString) Then
            ' Copier le fichier vers le bon endroit
            MkDir strCheminData
            FileCopy .Filename, strCheminApp & conCheminRelatifINI
            intRetourMessage = 1
            Unload Me
        End If
    End With

Fin:
    Set cCommonDialog = Nothing
End Sub

Private Sub cmdRegenerer_Click()
    MkDir strCheminData
    Open (strCheminApp & conCheminRelatifINI) For Output As #1
    Print #1, "[General]"
    Print #1, "DefaultType=6"
    Print #1, "Data="
    Print #1, "SaveDurations=1"
    Print #1, "DurationSaveInterval=2"
    Print #1, "Language="
    Print #1, "DisplayOnConnect=0"
    Print #1, vbNullString
    Print #1, "[Types]"
    Print #1, "Collection=o;m"
    Print #1, "Object=o;obj"
    Print #1, "Form=o;frm"
    Print #1, "Control=o;ctl"
    Print #1, "Picture=o;pic"
    Print #1, "Integer=v;int"
    Print #1, "Long=v;lng"
    Print #1, "Boolean=v;bln"
    Print #1, "String=v;str"
    Print #1, "Single=v;sng"
    Print #1, "Double=v;dbl"
    Print #1, "Byte=v;byt"
    Print #1, "Date=v;dtm"
    Print #1, "Variant=v;vnt"
    Print #1, "Currency=v;cur"
    Print #1, vbNullString
    Print #1, "[Durees]"
    Close #1
    intRetourMessage = 2
    Unload Me
End Sub

Private Sub Form_Load()
    Const conMessage_0 As String = "c2iExplorer n'a pas trouvé le fichier «c2iExplorer.ini» dans le dossier «"
    Const conMessage_1 As String = "»."
    Const conMessage_2 As String = "- Si vous avez une copie de ce fichier sur votre système, cliquez sur «Chercher» afin de la localiser et de la copier au bon endroit."
    Const conMessage_3 As String = "- Vous pouvez également regénérer le fichier INI par défaut en cliquant sur «Recréer»."
    Const conMessage_4 As String = "- Finalement, cliquez sur «Abandonner» si vous ne souhaitez pas démarrer c2iExplorer. Une fois sorti de Visual Basic, vous pourrez réinstaller le complément."
    intRetourMessage = 0
    strCheminData = strCheminApp & conDossierData
    lblMessage.Caption = conMessage_0 & strCheminData & conMessage_1 & (vbCrLf & vbCrLf) & _
                            conMessage_2 & vbCrLf & conMessage_3 & vbCrLf & conMessage_4
End Sub

Public Function ObtenirActionINIAbsent() As Integer
    Me.Show vbModal
    ObtenirActionINIAbsent = intRetourMessage
End Function
