VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
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

'//************Généré par c2iExplorer*******************
Private mElements As cElements

Public Property Get Elements() As cElements
    Set Elements = mElements
End Property

Private Sub Class_Initialize()
    Set mElements = New cElements
    '    InitClass 0, 1
End Sub

Private Sub Class_Terminate()
    Set mElements = Nothing
    '    InitClass 0, -1
End Sub

