VERSION 5.00
Begin VB.Form frmDuree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modification de la durée"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   Icon            =   "frmDuree.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   2  'Custom
   ScaleHeight     =   1545
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Fermer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "&Appliquer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   660
      Width           =   735
   End
   Begin VB.TextBox txtMn 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   660
      Width           =   735
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   735
   End
   Begin VB.Label lblProjectName 
      Caption         =   "NomProjet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3510
   End
   Begin VB.Label Label3 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   660
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "mn"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   0
      Top             =   660
      Width           =   375
   End
End
Attribute VB_Name = "frmDuree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
' René Rhéaume (rener@mediom.qc.ca)
'
' ***** END LICENSE BLOCK *****

Option Explicit

Private Sub cmdAppliquer_Click()
    Dim lngTime As Long, iNumPrj As Long
    Dim bTrouve As Boolean

    ' Modification par René Rhéaume 18 août 2001
    Call txtH_LostFocus
    Call txtMn_LostFocus
    Call txtS_LostFocus
    lngTime = CLng(txtH) * 3600 + CLng(txtMn) * 60 + CLng(txtS)

    bTrouve = conFaux
    For iNumPrj = 0 To UBound(DuréePrj)
        If DuréePrj(iNumPrj).Filename = lblProjectName Then
            bTrouve = conVrai
            Exit For
        End If
    Next
    If bTrouve Then
        lngTime = lngTime - timeGetTime / 1000 + DuréePrj(iNumPrj).Durée
    End If

    'Modification par René Rhéaume 2 août 2001
'    SaveSetting App.EXEName, "Durees", lblProjectName, CStr(lngTime)
    Call WritePrivateProfileString("Durees", lblProjectName, CStr(lngTime), c2iINIFile)
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub txtH_LostFocus()
    Dim lngVal As Long

    On Error GoTo Fin
    If Not IsNumeric(txtH) Then
        Err.Raise 13
    End If
    lngVal = CLng(Val(txtH))
    If CStr(lngVal) <> txtH Then
        Err.Raise 13
    End If
    If lngVal < 0 Or lngVal > 1000 Then
        Err.Raise 13
    End If

    Exit Sub

Fin:
    txtH = "0"
End Sub

Private Sub txtMn_LostFocus()
    Dim lngVal As Long

    On Error GoTo Fin
    If Not IsNumeric(txtMn) Then
        Err.Raise 13
    End If
    lngVal = CLng(Val(txtMn))
    If CStr(lngVal) <> txtMn Then
        Err.Raise 13
    End If
    If lngVal < 0 Or lngVal > 60 Then
        Err.Raise 13
    End If

    Exit Sub

Fin:
    txtMn = "0"
End Sub

Private Sub txtS_LostFocus()
    Dim lngVal As Long

    On Error GoTo Fin
    If Not IsNumeric(txtS) Then
        Err.Raise 13
    End If
    lngVal = CLng(Val(txtS))
    If CStr(lngVal) <> txtS Then
        Err.Raise 13
    End If
    If lngVal < 0 Or lngVal > 60 Then
        Err.Raise 13
    End If

    Exit Sub

Fin:
    txtS = "0"
End Sub
