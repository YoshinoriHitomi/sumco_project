VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "spr32x30.ocx"
Begin VB.Form Form9 
   Caption         =   "ó¨ìÆí‚é~àÍóóÅ@"
   ClientHeight    =   7860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8010
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   ScaleHeight     =   7860
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows ÇÃä˘íËíl
   Begin VB.CommandButton Command1 
      Caption         =   "ï¬Ç∂ÇÈ"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   7320
      Width           =   1575
   End
   Begin VB.Frame FraTitle 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Label lblTitle 
         Caption         =   "ó¨ìÆí‚é~àÍóóÅ@"
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "ÇlÇr ÉSÉVÉbÉN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   2550
      End
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7575
      _Version        =   196608
      _ExtentX        =   13361
      _ExtentY        =   9763
      _StockProps     =   64
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
      MaxCols         =   9
      MaxRows         =   6
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "Form10.frx":0000
      VisibleCols     =   1
      VisibleRows     =   2
   End
   Begin VB.Label Label17 
      BorderStyle     =   1  'é¿ê¸
      Caption         =   "7017-038Y0-000"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1365
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "ÉuÉçÉbÉNID"
      BeginProperty Font 
         Name            =   "ÇlÇr ÇoÉSÉVÉbÉN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   900
      Width           =   1095
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
