Attribute VB_Name = "MPrint"
Option Explicit

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_NORMAL = 1

Private Function AnalyseHTMLCmp(ByVal objVBCmp As VBIDE.VBComponent, sM As String)
    sM = RemplaceString(sM, "Component-Name", objVBCmp.Name)
    sM = RemplaceString(sM, "Component-Description", objVBCmp.Description)
    sM = RemplaceString(sM, "Component-Img", "<IMG SRC='img/" + ExtraitIconeComponent(objVBCmp) + ".gif' BORDER=0 Align='middle' hspace='5'>")
    sM = RemplaceString(sM, "Date", FormatDateTime(Date, vbGeneralDate))
    sM = RemplaceString(sM, "Component-NbrMember", CStr(objVBCmp.CodeModule.Members.Count))
    sM = RemplaceString(sM, "Component-NbrLines", CStr(objVBCmp.CodeModule.CountOfLines))
End Function

Private Function AnalyseHTMLMember(ByVal objVBMember As VBIDE.Member, sM As String)
    Dim lngC As Long

    On Error Resume Next
    sM = RemplaceString(sM, "Member-Name", objVBMember.Name)
    sM = RemplaceString(sM, "Member-Description", objVBMember.Description)
    sM = RemplaceString(sM, "Member-Img", "<IMG SRC='img/" + ExtraitIconeMembre(objVBMember, lngC) + ".gif' BORDER=0 Align='middle' hspace='5'>")
    sM = RemplaceString(sM, "Date", FormatDateTime(Date, vbGeneralDate))

End Function


Private Function AnalyseHTMLPrj(ByVal objVBPrj As VBIDE.VBProject, sM As String)
    sM = RemplaceString(sM, "Project-Name", objVBPrj.Name)
    sM = RemplaceString(sM, "Project-FileName", objVBPrj.FileName)
    sM = RemplaceString(sM, "Project-BuildFileName", objVBPrj.BuildFileName)
    sM = RemplaceString(sM, "Project-Description", objVBPrj.Description)
    sM = RemplaceString(sM, "Project-Img", "<IMG SRC='img/" + ExtraitIconeProjet(objVBPrj) + ".gif' BORDER=0 Align='middle' hspace='5'>")
    sM = RemplaceString(sM, "Date", FormatDateTime(Date, vbGeneralDate))
    sM = RemplaceString(sM, "Project-NbrComponent", CStr(objVBPrj.VBComponents.Count))
End Function

Public Function ExportHTML(ByVal sHTMLFileName As String)
    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim sEnTete As String, lngRep As Long
    Dim objVBPrj As VBIDE.VBProject
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    'add projects
    Set fso = New FileSystemObject
    Set txtOut = fso.CreateTextFile(c2iHTMLFile)
    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    sMiddle = txtIn.ReadAll
    txtIn.Close
    Set txtIn = Nothing
    If ExtraitHTML(sTete, sFin, sMiddle, "Projects") Then
        sTete = RemplaceString(sTete, "Date", FormatDateTime(Date, vbGeneralDate))
        txtOut.Write sTete
        For Each objVBPrj In VBInstance.VBProjects
            AddProject objVBPrj, txtOut, sMiddle
        Next
        sFin = RemplaceString(sFin, "Date", FormatDateTime(Date, vbGeneralDate))
        txtOut.Write sFin
    End If

    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then GoTo GestErr

    Exit Function
GestErr:
    txtOut.Close
    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description
End Function

Public Function ExportHTMLProject(ByVal sHTMLFileName As String, ByVal objVBPrj As VBIDE.VBProject)
    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim sEnTete As String, lngRep As Long
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    'add projects
    Set fso = New FileSystemObject
    Set txtOut = fso.CreateTextFile(c2iHTMLFile)
    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    sMiddle = txtIn.ReadAll
    txtIn.Close
    Set txtIn = Nothing

    Dim objVBCmp As VBIDE.VBComponent

    If ExtraitHTML(sTete, sFin, sMiddle, "Components") Then
        AnalyseHTMLPrj objVBPrj, sTete
        txtOut.Write sTete

        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBPrj.VBComponents.Count
            .Value = 0

            For Each objVBCmp In objVBPrj.VBComponents
                AddComponent objVBCmp, txtOut, sMiddle
                .Value = .Value + 1
            Next
        End With
        AnalyseHTMLPrj objVBPrj, sFin
        txtOut.Write sFin
    End If

    Set objVBCmp = Nothing

    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then GoTo GestErr

    Exit Function
GestErr:
    txtOut.Close
    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description

End Function

Public Function ExportHTMLComponent(ByVal sHTMLFileName As String, ByVal objVBCmp As VBIDE.VBComponent)
    Dim fso As FileSystemObject, txtOut As TextStream, txtIn As TextStream
    Dim sEnTete As String, lngRep As Long
    Dim sTete As String, sFin As String, sMiddle As String

    Screen.MousePointer = vbHourglass

    On Error GoTo GestErr
    If FichierExiste(c2iHTMLFile) Then
        Kill c2iHTMLFile
    End If

    'add projects
    Set fso = New FileSystemObject
    Set txtOut = fso.CreateTextFile(c2iHTMLFile)
    Set txtIn = fso.OpenTextFile(sHTMLFileName)
    sMiddle = txtIn.ReadAll
    txtIn.Close
    Set txtIn = Nothing

    Dim objVBMember As VBIDE.Member

    If ExtraitHTML(sTete, sFin, sMiddle, "Members") Then

        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
        AnalyseHTMLCmp objVBCmp, sTete
        txtOut.Write sTete
        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBCmp.CodeModule.Members.Count
            .Value = 0

            For Each objVBMember In objVBCmp.CodeModule.Members
                AddMember objVBMember, txtOut, sMiddle
                .Value = .Value + 1
                DoEvents
            Next
        End With
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
        AnalyseHTMLCmp objVBCmp, sFin
        txtOut.Write sFin
    End If

    Set objVBMember = Nothing

    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault

    lngRep = ShellExecute(0, "open", c2iHTMLFile, vbNullString, vbNullString, SW_NORMAL)
    If lngRep = 0 Then GoTo GestErr

    Exit Function
GestErr:
    txtOut.Close
    Set txtOut = Nothing
    Set fso = Nothing
    Screen.MousePointer = vbDefault
    Debug.Print Err.Description

End Function

Private Function AddProject(ByVal objVBPrj As VBIDE.VBProject, ByVal txtOut As TextStream, ByVal sMiddle As String)
    Dim objVBCmp As VBIDE.VBComponent
    Dim sTete As String, sFin As String
    'add project

    If ExtraitHTML(sTete, sFin, sMiddle, "Components") Then
        AnalyseHTMLPrj objVBPrj, sTete
        txtOut.Write sTete

        'barre de progression
        With frmPrint.pbPrint
            .Min = 0
            .Max = objVBPrj.VBComponents.Count
            .Value = 0
            For Each objVBCmp In objVBPrj.VBComponents
                AddComponent objVBCmp, txtOut, sMiddle
                .Value = .Value + 1
                DoEvents
            Next
        End With
        AnalyseHTMLPrj objVBPrj, sFin
        txtOut.Write sFin
    End If

    Set objVBCmp = Nothing
End Function

Private Function AddComponent(ByVal objVBCmp As VBIDE.VBComponent, ByVal txtOut As TextStream, ByVal sMiddle As String)
    Dim objVBMember As VBIDE.Member
    Dim sTete As String, sFin As String

    If ExtraitHTML(sTete, sFin, sMiddle, "Members") Then

        AnalyseHTMLPrj objVBCmp.Collection.Parent, sTete
        AnalyseHTMLCmp objVBCmp, sTete
        txtOut.Write sTete
        For Each objVBMember In objVBCmp.CodeModule.Members
            AddMember objVBMember, txtOut, sMiddle
        Next
        AnalyseHTMLPrj objVBCmp.Collection.Parent, sFin
        AnalyseHTMLCmp objVBCmp, sFin
        txtOut.Write sFin
    End If

    Set objVBMember = Nothing

End Function

Private Function AddMember(ByVal objVBMember As VBIDE.Member, ByVal txtOut As TextStream, ByVal sMiddle As String)

    If AffMembre(objVBMember) Then
        AnalyseHTMLPrj objVBMember.Collection.Parent.Parent.Collection.Parent, sMiddle
        AnalyseHTMLCmp objVBMember.Collection.Parent.Parent, sMiddle
        AnalyseHTMLMember objVBMember, sMiddle

        txtOut.Write sMiddle
    End If

End Function


Private Function AffMembre(objMember As VBIDE.Member) As Boolean
    '=========================================
    'Détermine si l'on affiche le membre sélectionné
    'en fonction du filtre général
    '=========================================
    '

    AffMembre = False                                      'on n'affiche pas par default
    Select Case objMember.Type
        Case vbext_mt_Const
            AffMembre = bAfficheConstante
        Case vbext_mt_Event
            AffMembre = bAfficheEvenement
        Case vbext_mt_Method
            'extraction de la ligne de code
            If objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Proc) <> 1 Then
                AffMembre = bAfficheMethode
            Else
                AffMembre = bAfficheAPI
            End If
        Case vbext_mt_Property
            AffMembre = bAffichePropriete
        Case vbext_mt_Variable
            AffMembre = bAfficheVariable
        Case Else
            AffMembre = True
    End Select

    Select Case objMember.Scope
        Case vbext_Friend
            If Not bAfficheFriend Then AffMembre = False
        Case vbext_Private
            If Not bAffichePrive Then AffMembre = False
        Case vbext_Public
            If Not bAffichePublic Then AffMembre = False
    End Select

End Function


Private Function ExtraitHTML(sTete As String, sFin As String, sMiddle As String, ByVal sSearch As String) As Boolean
    Dim iDeb As Long, iFin As Long
    Dim sSearchDeb As String, sSearchFin As String

    On Error GoTo Fin
    sSearchDeb = "<!-- " + sSearch + " -->"
    sSearchFin = "<!-- /" + sSearch + " -->"
    iDeb = InStr(1, sMiddle, sSearchDeb)
    iFin = InStr(1, sMiddle, sSearchFin)

    If iDeb = 0 Or iFin = 0 Then
        ExtraitHTML = False
        MsgBox "Format non valable", vbExclamation
        Exit Function
    End If

    sTete = Left(sMiddle, iDeb - 1)
    sFin = Right(sMiddle, Len(sMiddle) - iFin - Len(sSearchFin) + 1)
    sMiddle = Mid(sMiddle, iDeb + Len(sSearchDeb), iFin - iDeb - Len(sSearchDeb))

    ExtraitHTML = True
    Exit Function
Fin:
    ExtraitHTML = False
End Function

Private Function RemplaceString(sString As String, ByVal sSearch As String, ByVal sValeur As String)
    Dim iPos As Long


    sSearch = "[c2i" + sSearch + "]"

    iPos = InStr(1, sString, sSearch)
    Do Until iPos = 0
        iPos = InStr(1, sString, sSearch)
        If iPos > 0 Then
            sString = Left(sString, iPos - 1) + sValeur + Right(sString, Len(sString) - iPos - Len(sSearch) + 1)
        End If
    Loop
    RemplaceString = sString
End Function
