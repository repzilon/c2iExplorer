VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiblio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'//************Généré par c2iExplorer*******************
Private mElements As cElements
Private sName As String
Private sURL As String
Private sEMail As String
Private sDescription As String

Public Property Get Elements() As cElements
Set Elements = mElements
End Property

Public Property Let Nom(ByVal sNameA As String)
If Len(sNameA) > 255 Then
    sNameA = Left(sNameA, 255)
End If
If InStr(1, sNameA, "\") > 0 Then
    If lngLanguage = c2i_Langue_Français Then
        MsgBox "Nom de bibliothèque incorrecte", vbInformation
    Else
        MsgBox "The name of the library is incorrect", vbInformation
    End If
    Exit Property
End If
sName = sNameA
End Property

Public Property Get Nom() As String
Nom = sName
End Property

Public Property Let URL(ByVal sURLA As String)
If Len(sURLA) > 255 Then
    sURLA = Left(sURLA, 255)
End If
sURL = sURLA
End Property

Public Property Get URL() As String
URL = sURL
End Property

Public Property Let EMail(ByVal sEMailA As String)
If Len(sEMailA) > 255 Then
    sEMailA = Left(sEMailA, 255)
End If
sEMail = sEMailA
End Property

Public Property Get EMail() As String
EMail = sEMail
End Property

Public Property Let Description(ByVal sDescriptionA As String)
If Len(sDescriptionA) > 255 Then
    sDescriptionA = Left(sDescriptionA, 255)
End If
sDescription = sDescriptionA
End Property

Public Property Get Description() As String
Description = sDescription
End Property


Private Sub Class_Initialize()
Set mElements = New cElements
InitClass 2, 1
End Sub

Private Sub Class_Terminate()
Set mElements = Nothing
InitClass 2, -1
End Sub
