Attribute VB_Name = "GestFichier"
Option Explicit

'Module d'exploitation des fichiers
Public Function FichierExiste(sNom As String) As Boolean
        On Error Resume Next
        If GetAttr(sNom) = vbDirectory Then
                FichierExiste = False
                Exit Function
        End If
        Call FileLen(sNom)
        FichierExiste = (Err = 0)
End Function

Public Function HourGlass(obj As Object)
    Static ordMouse As Integer, fOn As Boolean
    On Error Resume Next
    If Not fOn Then
        ' Sauver le pointeur et activer le sablier hourglass
        ordMouse = obj.MousePointer
        obj.MousePointer = vbHourglass
        fOn = True
    Else
        ' Restaurer le pointeur
        obj.MousePointer = ordMouse
        fOn = False
    End If
    Debug.Assert Err = 0
End Function


Function NormalizePath(sPath As String) As String
    NormalizePath = sPath
    If Right$(sPath, 1) <> "\" Then NormalizePath = sPath & "\"
End Function

Function GetExtPos(sSpec As String) As Integer
    Dim iLast As Integer, iExt As Integer
    iLast = Len(sSpec)
    
    ' Analyser à reculons pour trouver l'extension ou le nom de base
    For iExt = iLast + 1 To 1 Step -1
        Select Case Mid$(sSpec, iExt, 1)
        Case "."
            ' Le premier . à partir de la droite est le début de l'extension
            Exit For
        Case "\"
            ' Le premier \ à partir de la droite est le début du nom
            iExt = iLast + 1
            Exit For
        End Select
    Next

    ' Un retour négatif indique l'absence d'extension, mais
    ' voici le nom de base pour que l'appelant n'ait pas à ré-analyser.
    GetExtPos = iExt
End Function

Public Function VérifFileName(sFile As String, sExt As String) As String

'// vérification si l'extension du fichier est bien enregistrée
If Len(sExt) <> 3 Then
    Debug.Print "VérifFileName Erreur"
End If

VérifFileName = sFile

If UCase(Right(sFile, 4)) <> "." + UCase(sExt) Then
    VérifFileName = sFile + "." + sExt
End If

End Function


Public Function RepertoireExiste(sChemin As String, sRep As String) As Boolean
Dim MonNom As String

RepertoireExiste = False
MonNom = Dir(sChemin, vbDirectory)    ' Sélectionne la première
                    ' entrée.
Do While MonNom <> ""       ' Début de la boucle.
    ' Ignore le répertoire courant et le répertoire supérieur.
    If MonNom <> "." And MonNom <> ".." Then
        ' Utilise une comparaison bit à bit pour s'assurer que MonNom
        ' est un répertoire.
        If (GetAttr(sChemin & MonNom) And vbDirectory) = vbDirectory Then
            If UCase(MonNom) = UCase(sRep) Then
                RepertoireExiste = True 'on a trouvé le répertoire
                Exit Function
            End If
        End If
    End If
    MonNom = Dir        ' Sélectionne l'entrée
                    ' suivante.
Loop
End Function

