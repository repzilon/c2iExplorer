VERSION 5.00
Begin VB.Form frmDuree 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Modification de la durée"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   3765
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quitter"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdAppliquer 
      Caption         =   "&Appliquer"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2520
      TabIndex        =   5
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtMn 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblProjectName 
      AutoSize        =   -1  'True
      Caption         =   "NomProjet"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "s"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "mn"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "h"
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   375
   End
End
Attribute VB_Name = "frmDuree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sProjectName

Private Sub cmdAppliquer_Click()
    Dim lngTime As Long, iNumPrj As Long
    Dim bTrouve As Boolean

    lngTime = CLng(txtH) * 3600 + CLng(txtMn) * 60 + CLng(txtS)

    bTrouve = False
    For iNumPrj = 0 To UBound(DuréePrj)
        If DuréePrj(iNumPrj).FileName = lblProjectName Then
            bTrouve = True
            Exit For
        End If
    Next
    If bTrouve Then
        lngTime = lngTime - timeGetTime / 1000 + DuréePrj(iNumPrj).Durée
    End If

    SaveSetting App.EXEName, "Durees", lblProjectName, CStr(lngTime)
    Unload Me
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub txtH_Validate(Cancel As Boolean)
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
    Cancel = True
    txtH = "0"

End Sub

Private Sub txtMn_Validate(Cancel As Boolean)
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
    Cancel = True
    txtMn = "0"

End Sub

Private Sub txtS_Validate(Cancel As Boolean)
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
    Cancel = True
    txtS = "0"
End Sub
