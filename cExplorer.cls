VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'//************Généré par c2iExplorer*******************
Private mElements As cElements

Public Property Get Elements() As cElements
    Set Elements = mElements
End Property


Private Sub Class_Initialize()
    Set mElements = New cElements
    InitClass 0, 1
End Sub

Private Sub Class_Terminate()
    Set mElements = Nothing
    InitClass 0, -1
End Sub
