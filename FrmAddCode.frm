VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmAddCode 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ajout de code"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.TreeView tv 
      Height          =   3015
      Left            =   3480
      TabIndex        =   6
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5318
      _Version        =   393217
      Indentation     =   88
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Text            =   "txtNameFunction"
      Top             =   0
      Width           =   3495
   End
   Begin VB.Frame FramePortee 
      Caption         =   "Portée"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   1575
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1215
      Begin VB.OptionButton opPortee 
         Caption         =   "Friend"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton opPortee 
         Caption         =   "Public"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton opPortee 
         Caption         =   "Private"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin MSComctlLib.ImageList ImgLst 
      Left            =   2280
      Top             =   3120
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
            Picture         =   "FrmAddCode.frx":0000
            Key             =   "quitter"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":055E
            Key             =   "ok"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   4320
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   635
      ButtonWidth     =   609
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ok"
            Object.ToolTipText     =   "Appliquer"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "quitter"
            Object.ToolTipText     =   "Annuler"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImgLSttv 
      Left            =   2760
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":0ABC
            Key             =   "Set"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":0BCE
            Key             =   "Let"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":0CE0
            Key             =   "Get"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":0DF2
            Key             =   "methode"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAddCode.frx":1144
            Key             =   "propriete"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmAddCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim objButton As Button

Set tb.ImageList = ImgLst
With tb
    For Each objButton In .Buttons
        objButton.Image = objButton.Key
    Next
End With

Set objButton = Nothing

Dim mNode As Node
Set tv.ImageList = ImgLSttv

'SUB
Set mNode = tv.Nodes.Add(, , "Sub", "Procédure", "methode")
Set mNode = tv.Nodes.Add(mNode, tvwChild, "SubArgument", "Argument", "propriete")
tv.Nodes.Add mNode, tvwChild, "SubValeur", "Valeur", "propriete"

'FUNCTION
Set mNode = tv.Nodes.Add(, , "Function", "Fonction", "methode")
'Set mNode = tv.Nodes.Add(mNode, tvwChild, "SubArgument", "Argument", "propriete")
'tv.Nodes.Add mNode, tvwChild, "SubValeur", "Valeur", "propriete"

Set mNode = tv.Nodes.Add(, , "Property", "Propriété", "propriete")

Set mNode = Nothing
'With ComboFunction
'    .AddItem "Byte"
'    .AddItem "Boolean"
'    .AddItem "Integer"
'    .AddItem "Long"
'    .AddItem "Single"
'    .AddItem "Double"
'    .AddItem "Currency"
'    .AddItem "Decimal"
'    .AddItem "Date"
'    .AddItem "Object"
'    .AddItem "String"
'    .AddItem "Variant"
'    .ListIndex = 3
'End With
End Sub
