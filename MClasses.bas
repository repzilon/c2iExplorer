Attribute VB_Name = "MClasses"
Option Explicit


Private iNClasses(5) As Long

Public Function InitClass(i As Integer, iNum As Integer)

    iNClasses(i) = iNClasses(i) + iNum

End Function

Public Function NombreInstancesClasses() As Long
    Dim iTotal As Long, i As Long

    iTotal = 0
    For i = 0 To 4
        iTotal = iTotal + iNClasses(i)
    Next i
    Debug.Print "Total : " + CStr(iTotal) + " instances."
    NombreInstancesClasses = iTotal
End Function

Public Function DétailInstancesClasses()
    Dim i As Integer
    Dim iTotal As Long

    Debug.Print "******************************* " + CStr(Time) + " **************"
    Debug.Print "cExplorer", CStr(iNClasses(0))
    Debug.Print "cBiblios", CStr(iNClasses(1))
    Debug.Print "cBiblio", CStr(iNClasses(2))
    Debug.Print "cElements", CStr(iNClasses(3))
    Debug.Print "cElement", CStr(iNClasses(4))

    iTotal = 0
    For i = 0 To 4
        iTotal = iTotal + iNClasses(i)
    Next i
    Debug.Print "******************************* " + CStr(iTotal) + " **************"
End Function



