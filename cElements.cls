VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cElement"
Attribute VB_Ext_KEY = "Member0" ,"cElement"
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
'
' ***** END LICENSE BLOCK *****

Option Explicit
'variable locale pour la collection
Private mCol As Collection

Friend Function Add(DeclareProc As String, HeadProc As String, _
                    InLineProc As String, EndProc As String, _
                    Description As String, Nom As String, _
                    Optional sKey As String) As cElement
    'créer un nouvel objet
    Dim objNewMember As cElement
    Set objNewMember = New cElement

    'définir les propriétés passées à la méthode
    objNewMember.HeadProc = HeadProc
    objNewMember.InLineProc = InLineProc
    objNewMember.EndProc = EndProc
    objNewMember.Description = Description
    objNewMember.Nom = Nom
    objNewMember.DeclareProc = DeclareProc
    If (LenB(sKey) = 0) Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    'renvoyer l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Friend Property Get Item(vntIndexKey As Variant) As cElement
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Friend Property Get Count() As Long
    Count = mCol.Count
End Property

Friend Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    '    InitClass 3, 1
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    '    InitClass 3, -1
End Sub
