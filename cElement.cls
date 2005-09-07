VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'//************Généré par c2iExplorer*******************
Private sNom As String
Private sDescription As String
Private sEndProc As String
Private sInLineProc As String
Private sHeadProc As String
Private sDeclareProc As String

Friend Property Let Nom(ByVal sNomA As String)
    Const conLongueurMax As Long = 255
    If (Len(sNomA) > conLongueurMax) Then
        sNom = Left$(sNomA, conLongueurMax)
    Else
        sNom = sNomA
    End If
End Property

Friend Property Get Nom() As String
    Nom = sNom
End Property

Friend Property Get Titre() As String
    Dim iPos As Long, iDeb As Long

    iDeb = 1
    iPos = InStr(iDeb, sNom, conSepChemBiblio)
    Do Until (iPos = 0)
        iDeb = iPos
        iPos = InStr(iPos + 1, sNom, conSepChemBiblio)
    Loop
    Titre = Right$(sNom, Len(sNom) - iDeb)
End Property

Friend Property Let Description(ByVal sDescriptionA As String)
    Const conLongueurMax As Long = 255
    If Len(sDescriptionA) > conLongueurMax Then
        sDescription = Left$(sDescriptionA, conLongueurMax)
    Else
        sDescription = sDescriptionA
    End If
End Property

Friend Property Get Description() As String
    Description = sDescription
End Property

Friend Property Let EndProc(ByVal sEndProcA As String)
    sEndProc = sEndProcA
End Property

Friend Property Get EndProc() As String
    EndProc = sEndProc
End Property

Friend Property Let InLineProc(ByVal sInLineProcA As String)
    sInLineProc = sInLineProcA
End Property

Friend Property Get InLineProc() As String
    InLineProc = sInLineProc
End Property

Friend Property Let HeadProc(ByVal sHeadProcA As String)
    sHeadProc = sHeadProcA
End Property

Friend Property Get HeadProc() As String
    HeadProc = sHeadProc
End Property

'Private Sub Class_Initialize()
'    InitClass 4, 1
'End Sub
'
'Private Sub Class_Terminate()
'    InitClass 4, -1
'End Sub

Friend Property Let DeclareProc(ByVal sDeclareProcA As String)
    sDeclareProc = sDeclareProcA
End Property

Friend Property Get DeclareProc() As String
    DeclareProc = sDeclareProc
End Property

Friend Property Get Ident() As Long
    Dim iPos As Long

    iPos = 1
    iPos = InStr(iPos, sNom, conSepChemBiblio)

    On Error Resume Next
    Do Until (iPos = 0)
        Ident = Ident + 1
        iPos = InStr(iPos + 1, sNom, conSepChemBiblio)
    Loop
End Property

'Modification par René Rhéaume le 9 mars 2002
'Réduction de taille
Friend Property Get NomRepertoire(ByVal iIdent As Long) As String
    'c2i/API/GRAPHICS/BitBlt
    'Renvoi API pour iIdent=1
    'Ident=3
    NomRepertoire = ChercherNomDossier(iIdent, conFaux)
'    MsgBox "NomRepertoire = " & NomRepertoire, vbInformation
End Property

'Modification par René Rhéaume le 9 mars 2002
'Réduction de taille
Friend Property Get CheminComplet(ByVal iIdent As Long) As String
    'c2i/API/GRAPHICS/BitBlt
    'Renvoi c2i/API pour iIdent=1
    'Ident=3
    CheminComplet = ChercherNomDossier(iIdent, conVrai)
'    MsgBox "CheminComplet = " & CheminComplet, vbInformation
End Property

'Ajout par René Rhéaume le 9 mars 2002
'Regroupement du traitement de NomRepertoire et de CheminComplet
Private Function ChercherNomDossier(ByVal iIdent As Long, ByVal blnCheminComplet As Boolean) As String
    Dim iPos As Long, iNum As Long, iPosNext As Long
    
    Select Case conVrai
        Case iIdent > Ident - 1, Ident < 0
            ChercherNomDossier = vbNullString
        Case Else
'            On Error Resume Next
            iPos = 1
            iPosNext = InStr(iPos, sNom, conSepChemBiblio)
            iNum = 0
            Do Until (iNum = iIdent)
                iPos = iPosNext + 1
                iPosNext = InStr(iPos, sNom, conSepChemBiblio)
                iNum = iNum + 1
            Loop
            If (blnCheminComplet) Then
                ChercherNomDossier = Left$(sNom, iPosNext - 1)
            Else
                ChercherNomDossier = Mid$(sNom, iPos, iPosNext - iPos)
            End If
    End Select
End Function
