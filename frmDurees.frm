VERSION 5.00
Begin VB.Form frmDurees 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "c2iExplorer - Compteurs"
   ClientHeight    =   3750
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmdurees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim tbSettings As Variant, i As Long, lItem As ListItem

tbSettings = GetAllSettings(App.EXEName, "Durees")
For i = 0 To UBound(tbSettings)
    Set lItem = lv.ListItems.Add(, tbSettings(i, 0), tbSettings(i, 0))
    lItem.SubItems(1) = tbSettings(i, 1)
    'Debug.Print tbSettings(i, 0); tbSettings(i, 1)
Next i
Set lItem = Nothing
End Sub
