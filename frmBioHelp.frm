VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmBioHelp 
   Caption         =   "  BioHelp"
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   7200
   Icon            =   "frmBioHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4170
   ScaleWidth      =   7200
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      ToolTipText     =   "Click to Exit"
      Top             =   3720
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox rtfBioHelp 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmBioHelp.frx":030A
   End
End
Attribute VB_Name = "frmBioHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()

    Unload Me
    
End Sub
