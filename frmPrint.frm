VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   4830
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbPrint 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   4515
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   1
      Min             =   1e-4
      Scrolling       =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   17
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":000C
            Key             =   "quit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":056A
            Key             =   "tous"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0AC8
            Key             =   "voir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   661
      ButtonWidth     =   847
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tous"
            ImageKey        =   "tous"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "voir"
            ImageKey        =   "voir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quit"
            ImageKey        =   "quit"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      Pattern         =   "c2i*.htm*;c2i*.xml"
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblSelectedElement 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   3795
   End
   Begin VB.Label lblTitre 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblInfos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrint"
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
Private Const conVoir As String = "voir"
Private Const conTous As String = "tous"
Private Const conQuit As String = "quit"
Private Const conDate As String = "Date"
Private Const conAttribImg As String = ".gif' border='0' align='middle' hspace='5'>"
Private Const conBaliseImg As String = "<img src='img/"
Private Const conFormatDate As String = "dddddd"
Private Const conComponents As String = "Components"
Private Const conMembers As String = "Members"
Private objVBPrjPrint As VBIDE.VBProject
Private objVBCmpPrint As VBIDE.VBComponent

'Optimisation par René Rhéaume les 31 juillet 2001 et 5 janvier 2002
Private Sub OK()
    Dim strCheminModele As String
    Dim strFichierModele As String
    Dim blnCmpPresent As Boolean
    strFichierModele = File1.Filename
    strCheminModele = File1.Path & conBS & strFichierModele

    If (strFichierModele <> vbNullString) Then
        If (FichierExiste(strCheminModele)) Then
            blnCmpPresent = Not objVBCmpPrint Is Nothing
            If (objVBPrjPrint Is Nothing) Then
                If (blnCmpPresent = conFaux) Then
                    ExportHTML strCheminModele
                End If
            Else
                If (blnCmpPresent) Then
                    ExportHTMLComponent strCheminModele, objVBCmpPrint
                Else
                    ExportHTMLProject strCheminModele, objVBPrjPrint
                End If
                '            Unload Me
            End If
        Else
            MsgBox "Le fichier [" & strCheminModele & "] n'existe pas."
        End If
    End If
End Sub

Private Sub Quitter()
    Unload Me
End Sub

Private Sub File1_DblClick()
    If (File1.Filename <> vbNullString) Then OK
End Sub

' Modifié par René Rhéaume le 5 janvier 2002
Private Sub Form_Load()
    Select Case lngLanguage
        Case c2i_Langue_Français
            Me.Caption = "Impression"
            Toolbar1.Buttons(conVoir).ToolTipText = "Aperçu élément selectionné"
            Toolbar1.Buttons(conTous).ToolTipText = "Aperçu général"
            Toolbar1.Buttons(conQuit).ToolTipText = "Quitter"
            lblInfos = "Choisissez votre modèle :"
            lblTitre = "Élément sélectionné :"
        Case Else
            Me.Caption = "Print"
            Toolbar1.Buttons(conVoir).ToolTipText = "Preview"
            Toolbar1.Buttons(conTous).ToolTipText = "General Preview"
            Toolbar1.Buttons(conQuit).ToolTipText = "Quit"
            lblInfos = "Choose your model :"
            lblTitre = "Selected element :"
    End Select

    '    PositionForm Me, conVrai
    File1.Path = strCheminApp & "\html"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '    PositionForm Me, conFaux
    Set objVBPrjPrint = Nothing
    Set objVBCmpPrint = Nothing
End Sub

'Modifié par René Rhéaume le 5 janvier 2002
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim chnFichierModele As String
    Select Case Button.Key
        Case conVoir
            OK
        Case conQuit
            Quitter
        Case conTous
            chnFichierModele = File1.Path & conBS & File1.Filename
            If (chnFichierModele <> vbNullString) Then
                If FichierExiste(chnFichierModele) Then
                    ExportHTML chnFichierModele
                    '                    Unload Me
                Else
                    MsgBox "Le fichier [" & chnFichierModele & "] n'existe pas."
                End If
            End If
    End Select
End Sub

Public Sub ObjetSelectionnés(ByVal objVBPrj As VBIDE.VBProject, objVBCmp As VBIDE.VBComponent)
    Set objVBPrjPrint = objVBPrj
    Set objVBCmpPrint = objVBCmp
    lblSelectedElement = vbNullString
    If Not objVBCmpPrint Is Nothing Then
        lblSelectedElement = objVBCmpPrint.Name
    End If
    If Not objVBPrjPrint Is Nothing Then
        lblSelectedElement = objVBPrjPrint.Name
    End If
End Sub

'//-------------------------------------------------------------//
'// Contenu de l'ancien module MPrint ci-dessous ---------------//
'// Modifications apportées par René Rhéaume 1er août 2001 -----//
'// Remplacement du FileSystemObject et des TextStream par -----//
'// les fonctions d'entrée/sortie par défaut de Visual Basic ---//
'//-------------------------------------------------------------//

Private Sub AnalyseHTMLCmp(ByVal objVBCmp As VBIDE.VBComponent, sM As String)
    Call RemplaceString(sM, "Component-Name", objVBCmp.Name)
    Call RemplaceString(sM, "Component-Description", objVBCmp.Description)
    Call RemplaceString(sM, "Component-Img", conBaliseImg + ExtraitIconeComponent(objVBCmp) + conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))
    Call RemplaceString(sM, "Component-NbrMember", CStr(objVBCmp.CodeModule.Members.Count))
    Call RemplaceString(sM, "Component-NbrLines", CStr(objVBCmp.CodeModule.CountOfLines))
End Sub

Private Sub AnalyseHTMLMember(ByVal objVBMember As VBIDE.Member, sM As String)
    Dim lngC As Long

    On Error Resume Next
    Call RemplaceString(sM, "Member-Name", objVBMember.Name)
    Call RemplaceString(sM, "Member-Description", objVBMember.Description)
    Call RemplaceString(sM, "Member-Img", conBaliseImg & ExtraitIconeMembre(objVBMember, lngC) & conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))

End Sub

Private Sub AnalyseHTMLPrj(ByVal objVBPrj As VBIDE.VBProject, sM As String)
    On Error GoTo AffichErr
    Call RemplaceString(sM, "Project-Name", objVBPrj.Name)
    Call RemplaceString(sM, "Project-FileName", objVBPrj.Filename)
    Call RemplaceString(sM, "Project-BuildFileName", objVBPrj.BuildFileName)
    Call RemplaceString(sM, "Project-Description", objVBPrj.Description)
    Call RemplaceString(sM, "Project-Img", conBaliseImg & ExtraitIconeProjet(objVBPrj) & conAttribImg)
    Call RemplaceString(sM, conDate, Format$(VBA.Date$, conFormatDate))
    Call RemplaceString(sM, "Project-NbrComponent", CStr(objVBPrj.VBComponents.Count))
    Exit Sub

AffichErr:
    MsgBox conErrNo & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbExclamation, "AnalyseHTMLPrj"
End Sub

Public Sub ExportHTML(ByVal sHTMLFileName As String)
    '    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim lngRep As Long, objVBPrj As VBIDE.VBProject
    Dim sTete As String, sFin As String, sMiddle As String
    Dim strDate As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    'Ancienne version
    '    'add projects
    '    Set fso = New FileSystemObject
    '    Set txtOut = fso.CreateTextFile(c2iHTMLFile)
    '    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    '    sMiddle = txtIn.ReadAll
    '    txtIn.Close
    '    Set txtIn = Nothing
    '    If ExtraitHTML(sTete, sFin, sMiddle, "Projects") Then
    '        sTete = RemplaceString(sTete, conDate, Format$(VBA.Date$, conFormatDate))
    '        txtOut.Write sTete
    '        For Each objVBPrj In VBInstance.VBProjects
    '            AddProject objVBPrj, txtOut, sMiddle
    '        Next
    '        sFin = RemplaceString(sFin, conDate, Format$(VBA.Date$, conFormatDate))
    '        txtOut.Write sFin
    '    End If
    '
    '    Set txtOut = Nothing
    '    Set fso = Nothing

    'Nouvelle version utilisant les instructions I/O de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, "Projects") Then
        strDate = Format$(VBA.Date$, conFormatDate)
        RemplaceString sTete, conDate, strDate
        Call EcrireFichier(sTete, c2iHTMLFile)
        For Each objVBPrj In VBInstance.VBProjects
            AddProject objVBPrj, sMiddle
        Next
        RemplaceString sFin, conDate, strDate
        Call EcrireFichier(sFin, c2iHTMLFile)
    End If

    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If (lngRep = 0) Then GoTo GestErr

    Exit Sub
GestErr:
    '    txtOut.Close
    '    Set txtOut = Nothing
    '    Set fso = Nothing
    MsgBox conErrNo & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbExclamation
End Sub

Public Sub ExportHTMLProject(ByVal sHTMLFileName As String, ByVal objVBPrj As VBIDE.VBProject)
    '    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim lngRep As Long, objVBCmp As VBIDE.VBComponent
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    'Ancienne version
    '    'add projects
    '    Set fso = New FileSystemObject
    '    Set txtOut = fso.CreateTextFile(c2iHTMLFile)
    '    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    '    sMiddle = txtIn.ReadAll
    '    txtIn.Close
    '    Set txtIn = Nothing
    '
    '    If ExtraitHTML(sTete, sFin, sMiddle, conComponents) Then
    '        AnalyseHTMLPrj objVBPrj, sTete
    '        txtOut.Write sTete
    '
    '        'barre de progression
    '        With frmPrint.pbPrint
    '            .Min = 0
    '            .Max = objVBPrj.VBComponents.Count
    '            .Value = 0
    '            For Each objVBCmp In objVBPrj.VBComponents
    '                AddComponent objVBCmp, sMiddle
    '                .Value = .Value + 1
    '            Next
    '        End With
    '        AnalyseHTMLPrj objVBPrj, sFin
    '        txtOut.Write sFin
    '    End If
    '    Set txtOut = Nothing
    '    Set fso = Nothing

    'Nouvelle version utilisant les instructions I/O de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, conComponents) Then
        AnalyseHTMLPrj objVBPrj, sTete
        Call EcrireFichier(sTete, c2iHTMLFile)

        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBPrj.VBComponents.Count
            .Value = 0
            For Each objVBCmp In objVBPrj.VBComponents
                AddComponent objVBCmp, sMiddle
                .Value = .Value + 1
            Next
        End With
        AnalyseHTMLPrj objVBPrj, sFin
        Call EcrireFichier(sFin, c2iHTMLFile)
    End If

    Set objVBCmp = Nothing
    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then GoTo GestErr

    Exit Sub
GestErr:
    '    txtOut.Close
    '    Set txtOut = Nothing
    '    Set fso = Nothing
    Screen.MousePointer = vbDefault
    MsgBox conErrNo & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbExclamation, "ExportHTMLProject"
End Sub

Public Sub ExportHTMLComponent(ByVal sHTMLFileName As String, ByVal objVBCmp As VBIDE.VBComponent)
    '    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim lngRep As Long, objVBMember As VBIDE.Member
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    ' Ancienne version
    '    'add projects
    '    Set fso = New FileSystemObject
    '    Set txtOut = fso.CreateTextFile()
    '    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    '    sMiddle = txtIn.ReadAll
    '    txtIn.Close
    '    Set txtIn = Nothing
    '
    '    If ExtraitHTML(sTete, sFin, sMiddle, conMembers) Then
    '        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
    '        AnalyseHTMLCmp objVBCmp, sTete
    '        txtOut.Write sTete
    '        'barre de progression
    '        With frmPrint.pbPrint
    '            .Min = 0
    '            .Max = objVBCmp.CodeModule.Members.Count
    '            .Value = 0
    '
    '            For Each objVBMember In objVBCmp.CodeModule.Members
    '                AddMember objVBMember, sMiddle
    '                .Value = .Value + 1
    '                DoEvents
    '            Next
    '        End With
    '        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
    '        AnalyseHTMLCmp objVBCmp, sFin
    '        txtOut.Write sFin
    '    End If
    '    Set txtOut = Nothing
    '    Set fso = Nothing

    'Nouvelle version utilisant les instructions E/S de fichier de VB
    Call LireFichierTexte(sHTMLFileName, sMiddle)
    If ExtraitHTML(sTete, sFin, sMiddle, conMembers) Then
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
        AnalyseHTMLCmp objVBCmp, sTete
        Call EcrireFichier(sTete, c2iHTMLFile)
        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBCmp.CodeModule.Members.Count
            .Value = 0

            For Each objVBMember In objVBCmp.CodeModule.Members
                AddMember objVBMember, sMiddle
                .Value = .Value + 1
                If GetInputState() Then DoEvents
            Next
        End With
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
        AnalyseHTMLCmp objVBCmp, sFin
        Call EcrireFichier(sFin, c2iHTMLFile)
    End If

    Set objVBMember = Nothing
    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then GoTo GestErr

    Exit Sub
GestErr:
    '    txtOut.Close
    '    Set txtOut = Nothing
    '    Set fso = Nothing
    Screen.MousePointer = vbDefault
    MsgBox conErrNo & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source, vbExclamation
End Sub

'Modifié par René Rhéaume 1er août 2001
'Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
Private Sub AddProject(ByVal objVBPrj As VBIDE.VBProject, ByVal sMiddle As String)
    Dim objVBCmp As VBIDE.VBComponent
    Dim sTete As String, sFin As String
    'add project

    If ExtraitHTML(sTete, sFin, sMiddle, conComponents) Then
        AnalyseHTMLPrj objVBPrj, sTete
        Call EcrireFichier(sTete, c2iHTMLFile)

        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBPrj.VBComponents.Count
            .Value = 0
            For Each objVBCmp In objVBPrj.VBComponents
                AddComponent objVBCmp, sMiddle
                .Value = .Value + 1
                If GetInputState() Then DoEvents
            Next
        End With
        AnalyseHTMLPrj objVBPrj, sFin
        Call EcrireFichier(sFin, c2iHTMLFile)
    End If

    Set objVBCmp = Nothing
End Sub

'Modifié par René Rhéaume 1er août 2001
'Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
Private Sub AddComponent(ByVal objVBCmp As VBIDE.VBComponent, ByVal sMiddle As String)
    Dim objVBMember As VBIDE.Member
    Dim sTete As String, sFin As String

    If ExtraitHTML(sTete, sFin, sMiddle, conMembers) Then

        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
        AnalyseHTMLCmp objVBCmp, sTete
        Call EcrireFichier(sTete, c2iHTMLFile)
        For Each objVBMember In objVBCmp.CodeModule.Members
            AddMember objVBMember, sMiddle
        Next
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
        AnalyseHTMLCmp objVBCmp, sFin
        Call EcrireFichier(sFin, c2iHTMLFile)
    End If
    Set objVBMember = Nothing
End Sub

'Modifié par René Rhéaume 1er août 2001
'Retrait du paramètre TextStream et nouvel appel pour écriture de fichier
Private Sub AddMember(ByVal objVBMember As VBIDE.Member, ByVal sMiddle As String)

    If (AfficheMembre(objVBMember)) Then
        With objVBMember.Collection.Parent
            AnalyseHTMLPrj .Parent.Collection.Parent, sMiddle
            AnalyseHTMLCmp .Parent, sMiddle
            AnalyseHTMLMember objVBMember, sMiddle
        End With
        Call EcrireFichier(sMiddle, c2iHTMLFile)
    End If

End Sub

Private Function ExtraitHTML(sTete As String, sFin As String, sMiddle As String, ByVal sSearch As String) As Boolean
    Const conHTMLCmmtEnd As String = " -->"
    Dim iDeb As Long, iFin As Long
    Dim sSearchDeb As String, sSearchFin As String

    On Error GoTo Fin
    sSearchDeb = "<!-- " & sSearch & conHTMLCmmtEnd
    sSearchFin = "<!-- /" & sSearch & conHTMLCmmtEnd
    iDeb = InStr(1, sMiddle, sSearchDeb)
    iFin = InStr(1, sMiddle, sSearchFin)

    Select Case conVrai
        Case iDeb = 0, iFin = 0
            ExtraitHTML = conFaux
            MsgBox "Format non valable", vbExclamation
            Exit Function
    End Select

    sTete = Left$(sMiddle, iDeb - 1)
    sFin = Right$(sMiddle, Len(sMiddle) - iFin - Len(sSearchFin) + 1)
    sMiddle = Mid$(sMiddle, iDeb + Len(sSearchDeb), iFin - iDeb - Len(sSearchDeb))

    ExtraitHTML = conVrai
    Exit Function
Fin:
    ExtraitHTML = conFaux
End Function

' Modifié par René Rhéaume le 18 janvier 2002
' Transformation de la fonction en Sub
Private Sub RemplaceString(ByRef sString As String, ByVal sSearch As String, ByVal sValeur As String)
    Dim iPos As Long
    sSearch = "[c2i" & sSearch & "]"

    iPos = InStr(1, sString, sSearch)
    Do Until iPos = 0
        iPos = InStr(1, sString, sSearch)
        If (iPos > 0) Then
            sString = Left$(sString, iPos - 1) & sValeur & Right$(sString, Len(sString) - iPos - Len(sSearch) + 1)
        End If
    Loop
    '    RemplaceString = sString
End Sub
