VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Impression"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pbPrint 
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   17
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0000
            Key             =   "quit"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":055E
            Key             =   "tous"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrint.frx":0ABC
            Key             =   "voir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   661
      ButtonWidth     =   847
      ButtonHeight    =   609
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tous"
            ImageKey        =   "tous"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "voir"
            ImageKey        =   "voir"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quit"
            ImageKey        =   "quit"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   0
      Pattern         =   "c2i*.htm*"
      TabIndex        =   0
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label lblSelectedElement 
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblTitre 
      Caption         =   "Elément sélectionné :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label lblInfos 
      Caption         =   "Choisissez votre modèle :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private objVBPrjPrint As VBIDE.VBProject
Private objVBCmpPrint As VBIDE.VBComponent

Private Sub OK()

    If FichierExiste(File1.Path + "\" + File1.FileName) Then
        If objVBPrjPrint Is Nothing And objVBCmpPrint Is Nothing Then
            ExportHTML File1.Path + "\" + File1.FileName
            Unload Me
            Exit Sub
        End If
        If Not objVBCmpPrint Is Nothing Then
            ExportHTMLComponent File1.Path + "\" + File1.FileName, objVBCmpPrint
            Unload Me
            Exit Sub
        End If
        ExportHTMLProject File1.Path + "\" + File1.FileName, objVBPrjPrint
        Unload Me
    End If

End Sub

Private Sub Quitter()
    Unload Me
End Sub

Private Sub File1_DblClick()
    If File1.FileName <> "" Then OK
End Sub

Private Sub Form_Load()
    Select Case lngLanguage
        Case c2i_Langue_Français
            Me.Caption = "Impression"
            Toolbar1.Buttons("voir").ToolTipText = "Aperçu élément selectionné"
            Toolbar1.Buttons("tous").ToolTipText = "Aperçu général"
            Toolbar1.Buttons("quit").ToolTipText = "Quitter"
            lblInfos = "Choisissez votre modèle :"
            lblTitre = "Elément sélectionné :"
        Case Else
            Me.Caption = "Print"
            Toolbar1.Buttons("voir").ToolTipText = "Preview"
            Toolbar1.Buttons("tous").ToolTipText = "General Preview"
            Toolbar1.Buttons("quit").ToolTipText = "Quit"
            lblInfos = "Choose your model :"
            lblTitre = "Selected element :"
    End Select

    PositionForm Me, True
    File1.Path = App.Path + "\html"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PositionForm Me, False
    Set objVBPrjPrint = Nothing
    Set objVBCmpPrint = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "voir"
            OK
        Case "quit"
            Quitter
        Case "tous"
            If FichierExiste(File1.Path + "\" + File1.FileName) Then
                ExportHTML File1.Path + "\" + File1.FileName
                Unload Me
            End If
    End Select
End Sub

Public Function ObjetSelectionnés(ByVal objVBPrj As VBIDE.VBProject, objVBCmp As VBIDE.VBComponent)
    Set objVBPrjPrint = objVBPrj
    Set objVBCmpPrint = objVBCmp
    lblSelectedElement = ""
    If Not objVBCmpPrint Is Nothing Then
        lblSelectedElement = objVBCmpPrint.Name
    End If
    If Not objVBPrjPrint Is Nothing Then
        lblSelectedElement = objVBPrjPrint.Name
    End If
End Function
