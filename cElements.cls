VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
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
Option Explicit

'variable locale pour la collection
Private mCol As Collection

Public Function Add(DeclareProc As String, HeadProc As String, InLineProc As String, EndProc As String, Description As String, Nom As String, Optional sKey As String) As cElement
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
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'renvoyer l'objet créé
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cElement
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    InitClass 3, 1
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
    InitClass 3, -1
End Sub

