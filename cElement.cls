VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'//************Généré par c2iExplorer*******************
Private sNom As String
Private sDescription As String
Private sEndProc As String
Private sInLineProc As String
Private sHeadProc As String
'//************Généré par c2iExplorer*******************
Private sDeclareProc As String


Public Property Let Nom(ByVal sNomA As String)
    If Len(sNomA) > 255 Then
        sNom = Left(sNomA, 255)
    Else
        sNom = sNomA
    End If
End Property

Public Property Get Nom() As String
    Nom = sNom
End Property

Public Property Get Titre() As String
    Dim iPos As Long, iDeb As Long

    iDeb = 1
    iPos = InStr(iDeb, sNom, "/")
    Do Until iPos = 0
        iDeb = iPos
        iPos = InStr(iPos + 1, sNom, "/")
    Loop
    Titre = Right(sNom, Len(sNom) - iDeb)

End Property

Public Property Let Description(ByVal sDescriptionA As String)
    If Len(sDescriptionA) > 255 Then
        sDescription = Left(sDescriptionA, 255)
    Else
        sDescription = sDescriptionA
    End If
End Property

Public Property Get Description() As String
    Description = sDescription
End Property

Public Property Let EndProc(ByVal sEndProcA As String)
    sEndProc = sEndProcA
End Property

Public Property Get EndProc() As String
    EndProc = sEndProc
End Property

Public Property Let InLineProc(ByVal sInLineProcA As String)
    sInLineProc = sInLineProcA
End Property

Public Property Get InLineProc() As String
    InLineProc = sInLineProc
End Property

Public Property Let HeadProc(ByVal sHeadProcA As String)
    sHeadProc = sHeadProcA
End Property

Public Property Get HeadProc() As String
    HeadProc = sHeadProc
End Property

Private Sub Class_Initialize()
    InitClass 4, 1
End Sub

Private Sub Class_Terminate()
    InitClass 4, -1
End Sub
Public Property Let DeclareProc(ByVal sDeclareProcA As String)
    sDeclareProc = sDeclareProcA
End Property

Public Property Get DeclareProc() As String
    DeclareProc = sDeclareProc
End Property

Public Property Get Ident() As Long
    Dim iPos As Long

    iPos = 1
    iPos = InStr(iPos, sNom, "/")

    On Error Resume Next
    Do Until iPos = 0
        Ident = Ident + 1
        iPos = InStr(iPos + 1, sNom, "/")
    Loop

End Property

Public Property Get NomRepertoire(ByVal iIdent As Long) As String
    Dim iPos As Long, iNum As Long, iPosNext As Long
    'c2i/API/GRAPHICS/BitBlt
    'Renvoi API pour iIdent=1
    'Ident=3

    If iIdent > Ident - 1 Or Ident < 0 Then
        NomRepertoire = ""
        Exit Property
    End If

    On Error Resume Next
    iPos = 1
    iPosNext = InStr(iPos, sNom, "/")
    iNum = 0
    NomRepertoire = Mid(sNom, iPos, iPosNext - iPos)
    Do Until iNum = iIdent
        iPos = iPosNext + 1
        iPosNext = InStr(iPos, sNom, "/")
        NomRepertoire = Mid(sNom, iPos, iPosNext - iPos)
        iNum = iNum + 1
    Loop

End Property

Public Property Get CheminComplet(ByVal iIdent As Long) As String
    Dim iPos As Long, iNum As Long, iPosNext As Long
    'c2i/API/GRAPHICS/BitBlt
    'Renvoi c2i/API pour iIdent=1
    'Ident=3

    If iIdent > Ident - 1 Or Ident < 0 Then
        CheminComplet = ""
        Exit Property
    End If

    On Error Resume Next
    iPos = 1
    iPosNext = InStr(iPos, sNom, "/")
    iNum = 0
    Do Until iNum = iIdent
        iPos = iPosNext + 1
        iPosNext = InStr(iPos, sNom, "/")
        iNum = iNum + 1
    Loop

    CheminComplet = Left(sNom, iPosNext - 1)
End Property
