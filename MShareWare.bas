Attribute VB_Name = "MShareWare"
Option Explicit

Public bc2i_ShareWare As Boolean

Public Function TestShareWare() As Boolean
Dim sUtilisateur As String, sSociete As String, sPassWord As String
Dim a As String, b As String, i As Long, j As Long

bc2i_ShareWare = False

sUtilisateur = GetSetting("MesCompléments", App.Title, "Utilisateur", "")
sSociete = GetSetting("MesCompléments", App.Title, "Société", "")
sPassWord = GetSetting("MesCompléments", App.Title, "PassWord", "")

a = Left(Trim(sUtilisateur), 1)
b = Left(Trim(sSociete), 1)
i = Left(CStr(Len(Trim(sUtilisateur))), 1)
j = Left(CStr(Len(Trim(sSociete))), 1)

If Mid(sPassWord, 2, 1) <> a Then bc2i_ShareWare = True
If Mid(sPassWord, 11, 1) <> b Then bc2i_ShareWare = True
If Mid(sPassWord, 5, 1) <> CStr(i) Then bc2i_ShareWare = True
If Mid(sPassWord, 14, 1) <> CStr(j) Then bc2i_ShareWare = True

If Mid(sPassWord, 6, 1) <> "-" Then bc2i_ShareWare = True
If Mid(sPassWord, 10, 1) <> "-" Then bc2i_ShareWare = True

End Function
