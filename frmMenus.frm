VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenus 
   Caption         =   "Form1"
   ClientHeight    =   2970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   2970
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   2520
      TabIndex        =   1
      Top             =   0
      Width           =   3135
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   840
         Width           =   375
      End
      Begin VB.TextBox txtDescription 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label lblExe 
         Height          =   975
         Left            =   240
         TabIndex        =   5
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Exécutable"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4895
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

