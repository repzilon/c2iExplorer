VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClassViewer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Classe"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9765
   Icon            =   "frmClassViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   528
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   651
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicSource 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   661
      ButtonWidth     =   847
      ButtonHeight    =   609
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgListView 
      Left            =   6240
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   128
      ImageWidth      =   100
      ImageHeight     =   14
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":030A
            Key             =   "methode"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":0674
            Key             =   "propgetlet"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":09DE
            Key             =   "propget"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":0D48
            Key             =   "proplet"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":10B2
            Key             =   "proppublic"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvProp 
      Height          =   4935
      Left            =   3120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   8705
      View            =   2
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      Icons           =   "imgListView"
      SmallIcons      =   "imgListView"
      ColHdrIcons     =   "imgListView"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgTBGene 
      Left            =   6720
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   17
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":141C
            Key             =   "print"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClassViewer.frx":197A
            Key             =   "actu"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInt 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Private mFunction as Long"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblClassName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cPersonnels"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "frmClassViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjCmp As VBIDE.VBComponent

Private Sub Form_Load()
Set lvProp.SmallIcons = imgListView

lvProp.View = lvwReport

Set tbMain.ImageList = ImgTBGene
Set tbMain.HotImageList = ImgTBGene

MakeToolBar
PositionForm Me, True

End Sub

Private Function MakeToolBar()
Dim objButton As Button

Set objButton = tbMain.Buttons.Add(, , , tbrSeparator)

Set objButton = tbMain.Buttons.Add(, , , , "actu")
objButton.ToolTipText = "Actualisation"

Set objButton = tbMain.Buttons.Add(, , , tbrSeparator)

Set objButton = tbMain.Buttons.Add(, , , , "print")
objButton.ToolTipText = "Imprimer"


End Function

Private Sub Form_Resize()
On Error Resume Next
With lblInt
    .Height = ScaleHeight - .Top - 10
End With
With lvProp
    .Width = ScaleWidth - .Left
    .Height = lblInt.Height
    .ColumnHeaders(1).Width = .Width * 0.9
End With
With PicSource
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight - tbMain.Height
End With
End Sub

Public Function Init(ByVal objCmp As VBIDE.VBComponent)
Dim objMember As VBIDE.Member, lngColor As Long, lItem As ListItem
Dim iGet As Long, iLet As Long, iSet As Long

lvProp.ListItems.Clear
lblInt = ""

Set mobjCmp = objCmp

lblClassName = mobjCmp.Name
For Each objMember In mobjCmp.CodeModule.Members
    Select Case ExtraitIconeMembre(objMember, lngColor)
        Case "methode"
            If lngColor <> Mem_Privee Then
                Set lItem = lvProp.ListItems.Add(, , objMember.Name, "methode", "methode")
                lItem.ForeColor = lngColor
            End If
        Case "propriete"
            If lngColor <> Mem_Privee Then
                On Error Resume Next
                iGet = objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Get)
                If iGet <> 0 Then
                    Set lItem = lvProp.ListItems.Add(, , objMember.Name, "propget", "propget")
                    lItem.ForeColor = lngColor
                End If
                iLet = objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Let)
                If iLet <> 0 Then
                    If iGet <> 0 Then
                        lItem.Icon = "propgetlet"
                        lItem.SmallIcon = "propgetlet"
                    Else
                        Set lItem = lvProp.ListItems.Add(, , objMember.Name, "proplet", "proplet")
                        lItem.ForeColor = lngColor
                    End If
                End If
                iSet = objMember.Collection.Parent.ProcBodyLine(objMember.Name, vbext_pk_Set)
                If iSet <> 0 Then
                    If iGet <> 0 Then
                        lItem.Icon = "propgetlet"
                        lItem.SmallIcon = "propgetlet"
                    Else
                        Set lItem = lvProp.ListItems.Add(, , objMember.Name, "proplet", "proplet")
                        lItem.ForeColor = lngColor
                    End If
                End If
            End If

        Case "variable"
            If lngColor = Mem_Public Then
                Set lItem = lvProp.ListItems.Add(, , objMember.Name, "proppublic", "proppublic")
                lItem.ForeColor = lngColor
            Else
                lblInt = lblInt + objMember.Name + vbCrLf
            End If
End Select
Next
Set objMember = Nothing

Set lvProp.SelectedItem = Nothing

End Function

Private Sub Form_Unload(Cancel As Integer)
PositionForm Me, False
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Image
    Case "actu"
        Init mobjCmp
    Case "print"
        Imprime
End Select
End Sub

Private Function Imprime()
Dim lngDC As Long

lngDC = GetDC(Me.hwnd)
BitBlt PicSource.hdc, 0, 0, Me.ScaleWidth, Me.ScaleHeight - tbMain.Height, lngDC, 0, tbMain.Height, vbSrcCopy
ReleaseDC Me.hwnd, lngDC
PicSource.Refresh
Printer.PaintPicture PicSource.Image, 0, 0
Printer.EndDoc
End Function
