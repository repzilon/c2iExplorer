VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBiblio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Détails"
   ClientHeight    =   5985
   ClientLeft      =   1965
   ClientTop       =   405
   ClientWidth     =   6660
   Icon            =   "frmBiblio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   5985
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgTB 
      Left            =   480
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   50
      ImageHeight     =   17
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":030A
            Key             =   "putinlib"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":0D74
            Key             =   "quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   661
      ButtonWidth     =   1508
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgTB"
      HotImageList    =   "imgTB"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Ajouter l'élément à la catégorie"
            ImageKey        =   "putinlib"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quit"
            Object.ToolTipText     =   "Quitter"
            ImageKey        =   "quit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Framecode 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      Begin VB.TextBox txtDescription 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   600
         Width           =   5175
      End
      Begin VB.TextBox txtDeclaration 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1560
         Width           =   5175
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
      Begin VB.TextBox txtHead 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2520
         Width           =   5175
      End
      Begin VB.TextBox txtInline 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3480
         Width           =   5175
      End
      Begin VB.TextBox txtEnd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   4440
         Width           =   5175
      End
      Begin VB.Label lblDescription 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   45
      End
      Begin VB.Label lblNom 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   45
      End
      Begin VB.Label lblEndProc 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   4440
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInlineProc 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   45
      End
      Begin VB.Label lblHeadProc 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblDeclare 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   1560
         Width           =   45
      End
   End
End
Attribute VB_Name = "frmBiblio"
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
Private Const conQuit As String = "quit"
Private Const conAdd As String = "add"
Private mObjE As cElement

' Modifié par René Rhéaume le 5 janvier 2002
Private Sub Form_Load()
    '    PositionForm Me
    Select Case lngLanguage
        Case c2i_Langue_Anglais
            lblNom = "Name"
            lblDescription = conLibelDescrpt
            lblDeclare = "Declaration"
            lblHeadProc = "Procedure header"
            lblInlineProc = "Code"
            lblEndProc = "Procedure footer"
            tbMain.Buttons(conAdd).ToolTipText = "Add to the selected folder"
            tbMain.Buttons(conQuit).ToolTipText = "Quit"
        Case Else
            lblNom = "Nom"
            lblDescription = conLibelDescrpt
            lblDeclare = "Déclaration"
            lblHeadProc = "En-tête de la procédure"
            lblInlineProc = "Code"
            lblEndProc = "Fin de la procédure"
            tbMain.Buttons(conAdd).ToolTipText = "Ajouter à la catégorie sélectionnée"
            tbMain.Buttons(conQuit).ToolTipText = "Quitter"
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '    PositionForm Me, conFaux
    objUDBiblio.bFrmBiblio = conFaux
    Set mObjE = Nothing
End Sub

Public Property Set Element(ByVal objE As cElement)
    If (Not objE Is mObjE) Then
        Set mObjE = objE
        With mObjE
            txtName = .Titre
            txtDescription = .Description
            txtDeclaration = .DeclareProc
            txtHead = .HeadProc
            txtInline = .InLineProc
            txtEnd = .EndProc
        End With
    End If
End Property

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case conQuit
            Unload Me
        Case conAdd
            AddCode
    End Select
End Sub

' Optimisé par René Rhéaume le 18 janvier 2002
Private Sub AddCode()
    Dim mNodeParent As Node, sKey As String
    Dim objE As cElement

    On Error Resume Next
    Set mNodeParent = objUDBiblio.TVEx.SelectedItem
    Select Case conVrai
        Case mNodeParent Is Nothing, Trim$(txtName) = vbNullString
            Exit Sub
    End Select

    If (mNodeParent.Image = conElement) Then
        Set mNodeParent = mNodeParent.Parent
    End If

    sKey = mNodeParent.FullPath & "/" & Trim$(txtName)
    Set objE = Explorer.Elements(sKey)
    If (objE Is Nothing) Then
        'on ajoute un nouvel élément
        Set objE = Explorer.Elements.Add(Trim$(txtDeclaration), _
                Trim$(txtHead), Trim$(txtInline), Trim$(txtEnd), _
                Trim$(txtDescription), sKey, sKey)
        objUDBiblio.AddElement objE
    Else
        objE.DeclareProc = Trim$(txtDeclaration)
        objE.HeadProc = Trim$(txtHead)
        objE.InLineProc = Trim$(txtInline)
        objE.EndProc = Trim$(txtEnd)
        objE.Description = Trim$(txtDescription)
    End If
    Set mNodeParent = Nothing
    Set objE = Nothing
End Sub
