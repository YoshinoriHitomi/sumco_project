VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_2 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "f_cmbc039_2(CW750) - 300mm結晶操業システム"
   ClientHeight    =   10875
   ClientLeft      =   1875
   ClientTop       =   2820
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ Ｐゴシック"
      Size            =   14.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   1018
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtDKTmpMid 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   795
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   8580
      Width           =   480
   End
   Begin VB.TextBox txtANTempMid 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   94
      Top             =   9285
      Width           =   600
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   3705
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   82
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   4695
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   81
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   5685
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   80
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   6675
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   79
      Top             =   7425
      Width           =   990
   End
   Begin VB.TextBox txtRRGMid 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   8655
      Width           =   960
   End
   Begin VB.TextBox txtCutPosMid 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   8070
      Width           =   600
   End
   Begin VB.TextBox txtSXLMid 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   465
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   8040
      Width           =   600
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   7665
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   75
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   8655
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   74
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   9645
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   73
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   7
      Left            =   10635
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   72
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   8
      Left            =   11625
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   71
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   9
      Left            =   12615
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   70
      Top             =   7425
      Width           =   990
   End
   Begin VB.TextBox txtKisei 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   6795
      Width           =   600
   End
   Begin VB.TextBox txtRoJdg 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   12600
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   3105
      Width           =   600
   End
   Begin VB.TextBox txtDKTmp 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   4275
      Width           =   480
   End
   Begin VB.TextBox txtDKTmp 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   2
      Left            =   780
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   6135
      Width           =   480
   End
   Begin VB.PictureBox pic_Png 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   900
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   58
      Top             =   9255
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CheckBox chk_Png 
      Caption         =   "PNG保存"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1290
      TabIndex        =   57
      Top             =   9315
      Width           =   1095
   End
   Begin VB.CommandButton CmdChangeWF_EP 
      Caption         =   "WF  >>"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   53
      Tag             =   "WF"
      Top             =   1605
      Width           =   855
   End
   Begin VB.TextBox txtANTempTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   4890
      Width           =   600
   End
   Begin VB.TextBox txtANTempTail 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   6825
      Width           =   600
   End
   Begin FPSpread.vaSpread sprWarp 
      Height          =   930
      Left            =   9435
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   825
      Width           =   3615
      _Version        =   196608
      _ExtentX        =   6376
      _ExtentY        =   1640
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   5
      MaxRows         =   4
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_2.frx":0000
      UserResize      =   0
      VisibleCols     =   5
      VisibleRows     =   1
   End
   Begin VB.TextBox txtCutPosTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   3720
      Width           =   600
   End
   Begin FPSpread.vaSpread spdHinbanTop 
      Height          =   465
      Left            =   1305
      TabIndex        =   41
      Top             =   1095
      Width           =   7740
      _Version        =   196608
      _ExtentX        =   13653
      _ExtentY        =   820
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   8
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":06E5
      UserResize      =   0
      VisibleCols     =   4
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdHinbanCen 
      Height          =   435
      Left            =   1305
      TabIndex        =   39
      Top             =   1905
      Width           =   11610
      _Version        =   196608
      _ExtentX        =   20479
      _ExtentY        =   767
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   10
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":0C3D
      UserResize      =   0
      VisibleCols     =   9
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdKensaTop 
      Height          =   1815
      Left            =   3720
      TabIndex        =   37
      Top             =   3450
      Width           =   10155
      _Version        =   196608
      _ExtentX        =   17912
      _ExtentY        =   3201
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayRowHeaders=   0   'False
      MaxCols         =   12
      MaxRows         =   7
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_2.frx":126A
      VisibleCols     =   8
      VisibleRows     =   7
   End
   Begin VB.TextBox txtRRGTail 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2625
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   6240
      Width           =   960
   End
   Begin VB.TextBox txtCutPosTail 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   5655
      Width           =   600
   End
   Begin VB.TextBox txtSXLTail 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   5625
      Width           =   600
   End
   Begin VB.TextBox txtJHAll 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   4905
      Width           =   600
   End
   Begin VB.TextBox txtRRGTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   4320
      Width           =   960
   End
   Begin FPSpread.vaSpread spdMeasTop 
      Height          =   1155
      Left            =   1320
      TabIndex        =   25
      Top             =   3705
      Width           =   1050
      _Version        =   196608
      _ExtentX        =   1852
      _ExtentY        =   2037
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":1B1D
      VisibleCols     =   1
      VisibleRows     =   5
   End
   Begin VB.TextBox txtSXLTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3720
      Width           =   600
   End
   Begin VB.TextBox txtSxlID 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   5745
      Locked          =   -1  'True
      MaxLength       =   13
      TabIndex        =   21
      Top             =   765
      Width           =   1335
   End
   Begin VB.TextBox txtJfName 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   780
      Width           =   1212
   End
   Begin VB.TextBox txtStaffID 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   17
      Top             =   780
      Width           =   972
   End
   Begin VB.Frame fraF 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   30
      TabIndex        =   13
      Top             =   9540
      Width           =   15195
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]　　　WFﾏｯﾌﾟ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   11448
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11]　　前画面"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   12680
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F８]　　＊＊＊"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   8984
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F７]　　＊＊＊"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   7752
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F６]　 廃棄"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   6520
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F５]　　再抜試"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   5288
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F４]　　＊＊＊"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F９]　　＊＊＊"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F１]　　ﾒｲﾝﾒﾆｭｰ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F２]　　ｻﾌﾞﾒﾆｭｰ"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1592
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F３]　　＊＊＊"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]　　実行"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   13920
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame FraTitle 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   15225
      Begin VB.Label lblvers 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   13740
         TabIndex        =   46
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label lblTime 
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   13740
         TabIndex        =   43
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  '中央揃え
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3360
         TabIndex        =   42
         Top             =   225
         Width           =   6870
      End
      Begin VB.Label lblTitle 
         Caption         =   "ＷＦセンター総合判定"
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   105
         TabIndex        =   12
         Top             =   255
         Width           =   4575
      End
   End
   Begin FPSpread.vaSpread spdMeasTail 
      Height          =   1155
      Left            =   1320
      TabIndex        =   33
      Top             =   5640
      Width           =   1050
      _Version        =   196608
      _ExtentX        =   1852
      _ExtentY        =   2037
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":1DF7
      VisibleCols     =   1
      VisibleRows     =   5
   End
   Begin FPSpread.vaSpread spdKensaTail 
      Height          =   1575
      Left            =   3720
      TabIndex        =   38
      Top             =   5370
      Width           =   10155
      _Version        =   196608
      _ExtentX        =   17912
      _ExtentY        =   2778
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      MaxCols         =   12
      MaxRows         =   7
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_2.frx":20D1
      VisibleCols     =   8
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdHinbanTail 
      Height          =   435
      Left            =   1305
      TabIndex        =   40
      Top             =   2325
      Width           =   11610
      _Version        =   196608
      _ExtentX        =   20479
      _ExtentY        =   767
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   11
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":29B8
      UserResize      =   0
      VisibleCols     =   7
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdHinbanTail2 
      Height          =   435
      Left            =   1305
      TabIndex        =   52
      Top             =   2940
      Width           =   10545
      _Version        =   196608
      _ExtentX        =   18600
      _ExtentY        =   767
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   13
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":3044
      UserResize      =   0
      VisibleCols     =   13
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdHinbanCenEpi 
      Height          =   465
      Left            =   1305
      TabIndex        =   54
      Top             =   1905
      Width           =   9255
      _Version        =   196608
      _ExtentX        =   16325
      _ExtentY        =   820
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   7
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":351B
      UserResize      =   0
      VisibleCols     =   7
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread spdHinbanHed 
      Height          =   210
      Left            =   1305
      TabIndex        =   59
      Top             =   2745
      Width           =   10545
      _Version        =   196608
      _ExtentX        =   18600
      _ExtentY        =   370
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   3
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":3AA6
      UserResize      =   0
      VisibleCols     =   3
   End
   Begin FPSpread.vaSpread spdMeasMid 
      Height          =   1155
      Left            =   1320
      TabIndex        =   83
      Top             =   8055
      Width           =   1050
      _Version        =   196608
      _ExtentX        =   1852
      _ExtentY        =   2037
      _StockProps     =   64
      BackColorStyle  =   1
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_2.frx":3DCE
      VisibleCols     =   1
      VisibleRows     =   5
   End
   Begin FPSpread.vaSpread spdKensaMid 
      Height          =   1575
      Left            =   3705
      TabIndex        =   84
      Top             =   7725
      Width           =   10155
      _Version        =   196608
      _ExtentX        =   17912
      _ExtentY        =   2778
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   2
      DisplayColHeaders=   0   'False
      DisplayRowHeaders=   0   'False
      MaxCols         =   12
      MaxRows         =   7
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_2.frx":4116
      VisibleCols     =   8
      VisibleRows     =   1
   End
   Begin VB.Label Label19 
      Caption         =   "DK温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   255
      TabIndex        =   96
      Top             =   8640
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "AN温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2595
      TabIndex        =   95
      Top             =   9060
      Width           =   720
   End
   Begin VB.Label lblMSMP_SUU 
      Caption         =   "実績／必要数："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13875
      TabIndex        =   93
      Top             =   8415
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.Label lblMSMP_JOSU 
      Caption         =   "許容枚数："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13905
      TabIndex        =   92
      Top             =   8985
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lblMSMP_TANI 
      Caption         =   "抜試単位："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13890
      TabIndex        =   91
      Top             =   7800
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblMSMP_FLG 
      Caption         =   "中間抜試（製品保証）"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   135
      TabIndex        =   90
      Top             =   7455
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Label lblMidMsg 
      Caption         =   "中間抜試メッセージ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3705
      TabIndex        =   89
      Top             =   9360
      Width           =   6870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1008
      X2              =   8
      Y1              =   488
      Y2              =   488
   End
   Begin VB.Label Label8 
      Caption         =   "RRG"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2625
      TabIndex        =   88
      Top             =   8430
      Width           =   465
   End
   Begin VB.Label Label9 
      Caption         =   "比抵抗"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1425
      TabIndex        =   87
      Top             =   7815
      Width           =   705
   End
   Begin VB.Label Label10 
      Caption         =   "再抜試位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2535
      TabIndex        =   86
      Top             =   7830
      Width           =   1050
   End
   Begin VB.Label Label16 
      Caption         =   "Mid 位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   225
      TabIndex        =   85
      Top             =   7830
      Width           =   1065
   End
   Begin VB.Label lblIchi 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7425
      TabIndex        =   69
      Top             =   1590
      Width           =   1635
   End
   Begin VB.Label Label7 
      Caption         =   "結晶位置(枚数)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6060
      TabIndex        =   68
      Top             =   1650
      Width           =   1425
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "払出規制"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   300
      TabIndex        =   66
      Top             =   6555
      Width           =   900
   End
   Begin VB.Label Label5 
      Caption         =   "認定炉判定"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12480
      TabIndex        =   64
      Top             =   2865
      Width           =   1005
   End
   Begin VB.Label Label11 
      Caption         =   "DK温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   63
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label Label14 
      Caption         =   "DK温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   62
      Top             =   6195
      Width           =   735
   End
   Begin VB.Label lblMukesaki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7785
      TabIndex        =   56
      Top             =   765
      Width           =   450
   End
   Begin VB.Label Label3 
      Caption         =   "向先"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7305
      TabIndex        =   55
      Top             =   825
      Width           =   390
   End
   Begin VB.Label Label2 
      Caption         =   "AN温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   51
      Top             =   4650
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "AN温度"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   48
      Top             =   6600
      Width           =   705
   End
   Begin VB.Label Label22 
      Caption         =   "RRG"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   34
      Top             =   6015
      Width           =   465
   End
   Begin VB.Label Label21 
      Caption         =   "比抵抗"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   31
      Top             =   5400
      Width           =   705
   End
   Begin VB.Label Label20 
      Caption         =   "再抜試位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2550
      TabIndex        =   30
      Top             =   5415
      Width           =   1050
   End
   Begin VB.Label Label17 
      Caption         =   "Ｔａｉｌ 位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   29
      Top             =   5415
      Width           =   1065
   End
   Begin VB.Label Label15 
      Caption         =   "RRG"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2640
      TabIndex        =   26
      Top             =   4080
      Width           =   465
   End
   Begin VB.Label Label13 
      Caption         =   "比抵抗"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1440
      TabIndex        =   23
      Top             =   3465
      Width           =   705
   End
   Begin VB.Label Label12 
      Caption         =   "再抜試位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2520
      TabIndex        =   22
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label36 
      Caption         =   "仮SXL－ID"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4665
      TabIndex        =   20
      Top             =   825
      Width           =   1020
   End
   Begin VB.Label Label35 
      Caption         =   "担当者コード"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "評価仕様"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1365
      TabIndex        =   16
      Top             =   1635
      Width           =   900
   End
   Begin VB.Label Label25 
      Caption         =   "Ｔｏｐ 位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   240
      TabIndex        =   15
      Top             =   3480
      Width           =   1050
   End
   Begin VB.Label Label26 
      Caption         =   "実行偏析"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   480
      TabIndex        =   14
      Top             =   4950
      Width           =   810
   End
End
Attribute VB_Name = "f_cmbc039_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private MsComment As String                'コメント   07/10/05 miyatake 承認機能追加
'>>>>> add 2011/07/14 Marushita
''  ウィンドウの表示位置・状態変更
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
         ByVal cy As Long, ByVal wFlags As Long) As Long

'ウインドウ画像のデバイスコンテキスト取得
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'デバイスコンテキストの解放
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
''BitBlt
'Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, _
'    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
'    ByVal ySrc As Long, ByVal dwRop As Long) As Long
'StretchBlt
Private Declare Function StretchBlt Lib "gdi32" ( _
    ByVal hDestDC As Long, _
    ByVal nDestLeft As Long, ByVal nDestTop As Long, _
    ByVal nDestWidth As Long, ByVal nDestHeight As Long, _
    ByVal hSrcDC As Long, _
    ByVal nSrcLeft As Long, ByVal nSrcTop As Long, _
    ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, _
    ByVal dwRop As Long) As Long

Private Const SWP_NOSIZE = &H1              ''サイズを指定しない
Private Const SWP_NOMOVE = &H2              ''位置を指定しない
Private Const HWND_TOPMOST = -1             ''常に手前
Private Const HWND_NOTOPMOST = -2           ''最前面表示解除
Private Const SRCCOPY = &HCC0020
Private Const SCALEPER = 85                 ''縮小％
'<<<<< add 2011/07/14 Marushita

'*******************************************************************************
'*    関数名        : CmdChangeWF_EP_Click
'*
'*    処理概要      : 1.ＷＦ⇔エピ切換えボタンクリック処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub CmdChangeWF_EP_Click()
    Dim i       As Integer              'Add 2011/03/23 SMPK Miyata

'    On Error Resume Next
    
    CmdChangeWF_EP.Enabled = False
    
    '各種情報クリア
    CtrlEnabled txtSXLTop, CTRL_DISABLE, True       'SXLトップ位置
    CtrlEnabled txtCutPosTop, CTRL_DISABLE, True    '最カット位置（トップ）
    CtrlEnabled txtRRGTop, CTRL_DISABLE, True       'RRG（トップ）
    CtrlEnabled txtSXLTail, CTRL_DISABLE, True      'SXLテイル位置
    CtrlEnabled txtCutPosTail, CTRL_DISABLE, True   '最カット位置（テイル）
    CtrlEnabled txtRRGTail, CTRL_DISABLE, True      'RRG（テイル）
    CtrlEnabled txtJHAll, CTRL_DISABLE, True        '実行偏析全体
    CtrlEnabled txtANTempTop, CTRL_DISABLE, True    'AN温度（トップ）
    CtrlEnabled txtANTempTail, CTRL_DISABLE, True   'AN温度（テイル）
    CtrlEnabled txtRoJdg, CTRL_DISABLE, True        '認定炉判定     *2008/08/28 kameda
    CtrlEnabled txtKisei, CTRL_DISABLE, True        '払出規制       *2010/02/15 kameda
'Add Start 2011/03/23 SMPK Miyata
    CtrlEnabled txtSXLMid, CTRL_DISABLE_SKY, True      'SXL中間位置
    CtrlEnabled txtCutPosMid, CTRL_DISABLE_SKY, True   '最カット位置（中間）
    CtrlEnabled txtRRGMid, CTRL_DISABLE_SKY, True      'RRG（中間）
'Add End   2011/03/23 SMPK Miyata
'Add Start 2011/08/25 Y.Hitomi
    CtrlEnabled txtANTempMid, CTRL_DISABLE_SKY, True   'AN温度（中間)
    CtrlEnabled txtDKTmpMid, CTRL_DISABLE_SKY, True    'DK温度（中間)
'Add End   2011/08/25 Y.Hitomi

    '比抵抗情報クリア
    SpCtrlBlockEnabled Me.spdMeasTop, 1, 1, 1, 5, CTRL_DISABLE, True
    SpCtrlBlockEnabled Me.spdMeasTail, 1, 1, 1, 5, CTRL_DISABLE, True
'Add Start 2011/03/23 SMPK Miyata
    SpCtrlBlockEnabled spdMeasMid, 1, 1, spdMeasMid.MaxCols, spdMeasMid.MaxRows, CTRL_DISABLE_SKY, True
'Add End   2011/03/23 SMPK Miyata

    '実績情報クリア
    SpCtrlBlockEnabled Me.spdKensaTop, 1, -1, 12, -1, CTRL_DISABLE
    SpCtrlBlockEnabled Me.spdKensaTail, 1, -1, 12, -1, CTRL_DISABLE
'Add Start 2011/07/21 Y.Hitomi
    SpCtrlBlockEnabled Me.spdKensaMid, 1, -1, 12, -1, CTRL_DISABLE_SKY
'Add End   2011/07/21 Y.Hitomi

    '' WF → EP
    If CmdChangeWF_EP.Tag = "WF" Then
        '仕様表示
        Call sub_cmbc061_2_ChangeHinSpec(1)
        '実績情報表示
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlTop039
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlTail039

        CmdChangeWF_EP.Tag = "EP"
        CmdChangeWF_EP.Caption = "エピ >>"
        EPSiyouSansyouFlg = True
        cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039)
        cmdF(5).Enabled = (txtJfName.text <> "")
    '' EP → WF
    Else
        '仕様表示
        Call sub_cmbc061_2_ChangeHinSpec(0)
        '比抵抗情報表示
        sub_PutRs
        'AN温度：DKANの3～6桁がAN温度
        Me.txtANTempTop.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTop039, WFRES).DKAN, 3, 4), "0") 'AN温度
        'チェックNGの時は背景色を変える
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "TOP") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    If Not (typ_CType.JudgAntnp(SxlTop039)) Then
                        CtrlEnabled Me.txtANTempTop, CTRL_DISABLE_WARNING, False  'AN温度
                    End If
                End If
            End If
        End If
        'AN温度：DKANの3～6桁がAN温度
        Me.txtANTempTail.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTail039, WFRES).DKAN, 3, 4), "0") 'AN温度
        'チェックNGの時は背景色を変える
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "BOT") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTail039).WFRESRS1CW = "1" Then
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled Me.txtANTempTail, CTRL_DISABLE_WARNING, False  'AN温度
                    End If
                End If
            End If
        End If
        '実績情報表示
        sub_PutRslt typ_CType.typ_rslt(), SxlTop039
        sub_PutRslt typ_CType.typ_rslt(), SxlTail039

        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "ＷＦ >>"
    End If
'Add Start 2011/03/23 SMPK Miyata
    '中間位置選択ボタン数ループ
    For i = optPosSelMid.LBound To optPosSelMid.UBound
        If optPosSelMid(i).Value = True Then
            '中間抜試サンプル位置ボタンクリック処理を行う
            Call optPosSelMid_Click(i)
            Exit For
        End If
    Next i
'Add End   2011/03/23 SMPK Miyata

    CmdChangeWF_EP.Enabled = True
End Sub

'*******************************************************************************
'*    関数名        : Form_Unload
'*
'*    処理概要      : 1.Form_Unload処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    Cancel        ,I  ,Integer
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Unload WFCJudgDialog
End Sub



'Add Start 2011/03/09 SMPK Miyata
'*******************************************************************************
'*    関数名        : optPosSelMid_Click
'*
'*    処理概要      : 1.中間抜試サンプル位置ボタンクリック
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub optPosSelMid_Click(Index As Integer)
    
    If CmdChangeWF_EP.Tag = "WF" Then
        '比抵抗値表示(中間抜試)
        Call sub_PutRsMid(Index + 1)
        
        sub_PutRslt typ_CType.typ_rslt(), SxlMidl039 + Index
    Else
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlMidl039 + Index
    End If
    
End Sub
'Add End   2011/03/09 SMPK Miyata

'*******************************************************************************
'*    関数名        : txtJfName_Change
'*
'*    処理概要      : 1.担当者が変更になった場合、再抜試と実行ボタンの権限を変更
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtJfName_Change()
    Dim StopFlgF5 As Boolean
    Dim StopFlgF12 As Boolean
    
    'add 流動停止チェック追加 SETkimizuka Start
    StopFlgF5 = CheckXODY4(WATCH_PROCCD_NUKISI, "", txtSxlId.text)
    StopFlgF12 = CheckXODY4(WATCH_PROCCD, "", txtSxlId.text)
    'add 流動停止チェック追加 SETkimizuka End
    
    
'upd 流動停止チェック追加 SETkimizuka Start
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
''    cmdF(5).Enabled = (txtJfName.Text <> "")
'    cmdF(5).Enabled = ((txtJfName.text <> "") _
'                        And ((typ_CType.typ_si.HEPHS = False) _
'                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)))
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    cmdF(5).Enabled = ((txtJfName.text <> "") _
                        And ((typ_CType.typ_si.HEPHS = False) _
                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)) And (StopFlgF5 = True))
'upd 流動停止チェック追加 SETkimizuka End
'    cmdF(6).Enabled = (txtJfName.Text <> "")
''2001/12/18 S.Sano    cmdF(12).Enabled = ((txtJfName.Text <> "") And TotalJudg)
'ＷＦサンプル処理変更 2003.05.20 yakimura
'    cmdF(12).Enabled = ((txtJfName.Text <> "") And (TotalJudg Or bPPlus Or bNPlus)) ''2001/12/18 S.Sano
'upd 流動停止チェック追加 SETkimizuka Start
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
''    cmdF(12).Enabled = ((txtJfName.Text <> "") And TotalJudg039)
'    '' エピ仕様、実績、判定結果NGを参照済みの場合は実行ボタンを押下可能にする
'    cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039 _
'                        And ((typ_CType.typ_si.HEPHS = False) _
'                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)))
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'ＷＦサンプル処理変更 2003.05.20 yakimura
    '' エピ仕様、実績、判定結果NGを参照済み、流動停止の場合は実行ボタンを押下可能にする
    cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039 _
                        And ((typ_CType.typ_si.HEPHS = False) _
                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)) And (StopFlgF12 = True))
'upd 流動停止チェック追加 SETkimizuka End

    '流動停止の場合はメッセージ表示する 2010/06/16 SETsw kubota
    If StopFlgF12 = False Then
        Call MsgOut(0, PROCD_WFC_SOUGOUHANTEI & "工程での流動停止品です。(F12不可)", DEBUG_DISP)
    End If
    If StopFlgF5 = False Then
        Call MsgOut(0, left$(WATCH_PROCCD_NUKISI, Len(WATCH_PROCCD_NUKISI) - 1) & "工程での流動停止品です。(F5,F12不可)", DEBUG_DISP)
    End If

End Sub

'*******************************************************************************
'*    関数名        : txtStaffID_Change
'*
'*    処理概要      : 1.担当コード変更処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtStaffID_Change()
    If STAFFIDBUFF <> Trim(txtStaffID.text) Then
        txtJfName.text = ""
    End If
End Sub

'*******************************************************************************
'*    関数名        : txtStaffID_KeyDown
'*
'*    処理概要      : 1.担当者コード入力チェック処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    KeyCode       ,I  ,Integer　,キーコード
'*                    Shift         ,I  ,Integer  ,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtStaffID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim FuncAns As FUNCTION_RETURN

'    '' 画面表示メッセージクリア
'    lblMsg.Caption = ""
    
    If KeyCode = vbKeyReturn And txtStaffID.Locked <> True Then
        '' 画面表示メッセージクリア
        lblMsg.Caption = ""
        FuncAns = StaffIDCheck(txtStaffID, txtJfName, lblMsg)
    End If
End Sub

'*******************************************************************************
'*    関数名        : cmdF_Click
'*
'*    処理概要      : 1.ファンクションボタンがクリックされたら、各処理に分岐する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    Index       ,I  ,Integer　,コントロール配列の添字
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub cmdF_Click(intIndex As Integer)
    
    Dim sErrMsg   As String
    
    '' 処理分岐
    Select Case intIndex
        Case 1          '' Ｆ１キー（メインメニュー）
            '' プログラム終了処理
             GotoMainMenu
        Case 2          '' Ｆ２キー（サブメニュー）
            '' サブメニューに戻る
            GotoSubMenu
        Case 11 ''Ｆ11（前画面）
            '' 前画面に戻る
            intModoru = 1
            Unload Me
            f_cmbc039_1.Visible = True
            CloseFormProc f_cmbc039_1, f_cmbc039_2
        Case 5 '' Ｆ5キー（再抜試）
            '>>>>> add 2011/07/14 Marushita
            'キャプチャの保存
            Call saveCapture_BMP(Me, pic_Png)
            '<<<<< add 2011/07/14 Marushita
            
            '' 流動監視チェック add 09/03/17 SETkimizuka
            If CheckXODY4(WATCH_PROCCD_NUKISI, "", txtSxlId.text) = False Then
                lblMsg.Caption = Y4_STOP_ERR
                Exit Sub
            End If
            
            '' 実行処理を行う
            typ_CType.StrStaffId = txtStaffID.text
            typ_CType.strStaffName = txtJfName.text
            If fnc_ExecutionProcess(intIndex) = FUNCTION_RETURN_FAILURE Then
                Exit Sub
            End If
                    
            '' 再カット画面に遷移
            CloseFormProc f_cmbc039_3, f_cmbc039_2
        Case 10
            'WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙからﾃﾞｰﾀを取得
            If SelWFmap(vbNullString, SelectSxlID039, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_2.lblMsg.Caption = sErrMsg
                Exit Sub
            End If
            
            'ｽﾌﾟﾚｯﾄﾞにﾃﾞｰﾀを表示
            If SetWFmapData = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_4.lblMsg.Caption = sErrMsg
'Chg Start 2011/03/11 SMPK Miyata
'                f_cmbc039_4.sprExamine.MaxRows = 0
                f_cmbc039_4.sprWfmapView.MaxRows = 0
'Chg End   2011/03/11 SMPK Miyata

                Exit Sub
            End If
            f_cmbc039_4.txtSxlId.text = SelectSxlID039
            f_cmbc039_4.Show
        Case 12       '' Ｆ12キー（実行）
            '' 担当者IDのチェック
            If f_cmzcChkUser.CanExec(Me.Name, txtStaffID.text) = False Then
                lblMsg.Caption = GetMsgStr("EUSR0")
                Exit Sub
            End If
            
            '' 流動監視チェック add 09/03/17 SETkimizuka
            If CheckXODY4(WATCH_PROCCD_ENT, "", txtSxlId.text) = False Then
                lblMsg.Caption = Y4_STOP_ERR
                Exit Sub
            End If
            
            If MsgBox(GetMsgStr("PIN01"), vbOKCancel, "WF総合判定") = vbOK Then
                
                ' 承認機能追加による修正  2007/10/05 miyatake ===================> START
                '' コメント入力
                If Me.chk_Png = 1 Then
                    If f_comment.GetComment(MsComment) <> vbOK Then
                        Exit Sub
                    End If
                    Call SetForceForegroundWindow(Me.hwnd)
                End If
                ' 承認機能追加による修正  2007/10/05 miyatake ===================> START
                
                BeginProcess '' プロセス開始
                '' 実行処理を行う
                If fnc_ExecutionProcess(intIndex) = FUNCTION_RETURN_FAILURE Then
                    EndProcess '' プロセス終了
                    Exit Sub
                End If
                EndProcess '' プロセス終了
                        
                '' 前画面に戻る
                intModoru = 2
                Unload Me
                f_cmbc039_1.Visible = True
            End If
    End Select
End Sub

'*******************************************************************************
'*    関数名        : Form_KeyDown
'*
'*    処理概要      : 1.キーボード押下処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    KeyCode     ,I  ,Integer　,キーコード
'*                    Shift       ,I  ,Integer　,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '' 画面表示メッセージクリア
    lblMsg.Caption = ""
    '' ファンクションキーが有効なら
    If KeyCode >= 112 And KeyCode <= 123 Then
        If cmdF(KeyCode - 111).Enabled = True Then
            '' ファンクションキー押下処理を実行する
            Call cmdF_Click(KeyCode - 111)
        End If
    End If
    
#If JudgDebug Then
    If (Shift = vbShiftMask + vbCtrlMask) And (KeyCode = 68) Then
        WFCJudgDialog.Visible = (Not WFCJudgDialog.Visible)
    End If
#End If
End Sub

'*******************************************************************************
'*    関数名        : Form_Load
'*
'*    処理概要      : 1.Form_Load処理
'*                    2.Warp判定用ﾃﾞｰﾀ取得
'*                    3.振替可否チェック（仕様）
'*                    4.Warp/合成角度情報表示
'*                    5.規格情報表示
'*                    6.比抵抗情報表示
'*                    7.実績情報表示
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Load()
    Me.Hide
    Me.Show
    DoEvents

    Load WFCJudgDialog
    CtrlEnabled txtStaffID, CTRL_ENABLE, True       '担当者コード
    CtrlEnabled txtJfName, CTRL_DISABLE, True       '担当者名
    CtrlEnabled txtSxlId, CTRL_DISABLE, True      'ブロックID
    txtStaffID.text = typ_AType.StrStaffId ' スタッフID
    txtJfName.text = typ_AType.strStaffName ' スタッフ名
    txtSxlId.text = SelectSxlID039 ' ブロックIDの表示
    SpCtrlInit spdKensaTop, 0
    SpCtrlInit spdKensaTail, 0
    sprWarp.MaxRows = 0             '05/12/15 ooba
    
    'Add Start 2011/04/28 SMPK Miyata (初期表示時にちらつく防止)
    '表示画面クリア
    sub_InitDisp
    'Add End   2011/04/28 SMPK Miyata

    ' 現在日時の表示
    '' 処理時間セット
    SetPresentTime lblTime

    ' バージョン情報の表示
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    '' フォーム位置セット
    CenterForm Me
    
    f_cmbc039_2.Enabled = False
    BeginProcess '' プロセス開始
    lblMsg.Caption = GetMsgStr(PWAIT)
    DoEvents
        
    'Del Start 2011/04/28 SMPK Miyata (初期表示時にちらつく防止) 処理を上へ移動
    ''表示画面クリア
    'sub_InitDisp
    'Del End   2011/04/28 SMPK Miyata
    
    Dim intErrCode As Integer
    Dim strErrMsg As String
    Dim intRet As Integer
    
'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    Dim wkXsdcw     As typ_XSDCW
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------
'>>>>> add start 2011/06/30 Marushita
    Dim iMinMidCnt      As Integer       '中間抜試の必要数
    Dim iRstMidCnt      As Integer       '中間抜試の件数
    Dim iMSMPTANI       As Integer       '中間抜試単位(mm)
'<<<<< add end 2011/06/30 Marushita
        
    'Add Start 2011/09/29 Y.Hitomi
    Dim sSXLIDFLG       As Integer       'ＳＸＬＩＤ確定可否フラグ
    'Add End   2011/09/29 Y.Hitomi
        
'↓変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
' このプロシージャ内でクリア用の変数を宣言すると、VBの容量制限に引っかかるので､
' 別プロシージャでクリアする｡
    'typ_Ctypeを初期化
'    Dim clear_typeC As typ_AllTypesC
'    typ_CType = clear_typeC
    Call Crear_type_Siyou_Spv
'↑変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------

'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    typ_CType.JudgDkTmp(SxlTop) = JUDG_OK
    typ_CType.JudgDkTmp(SxlTail) = JUDG_OK
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------
    
    Call InitHensu2(typ_CType)   '2003-11-01 SystemBrain 追加
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    Call InitHensu2_EP(typ_CType_EP)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    '2009/10/15 Kameda
    ReDim gNinteiro_Data(0)
'Del Start 2012/07/09 Y.Hitomi
'        2011/05/31 Kameda
'    ReDim tbl_chk2_5.MLTJDG(1)
'Del End 2012/07/09 Y.Hitomi
    
    '画面情報設定

    ''Warp判定対応　06/01/11 ooba START ==================================>
    'Warp判定用ﾃﾞｰﾀ取得
    If fnc_LoadData_Warp() = FUNCTION_RETURN_FAILURE Then
        f_cmbc039_2.Enabled = True
        f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        EndProcess
        Exit Sub
    End If
    
'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    ' DK温度(実績)取得
    typ_CType.DkTmpJsk(SxlTop) = GetWfDKTmpCode(False, typ_AType.typ_Param.WFSMP(SxlTop))
    typ_CType.DkTmpJsk(SxlTail) = GetWfDKTmpCode(False, typ_AType.typ_Param.WFSMP(SxlTail))
    ' DK温度(仕様)取得
    wkXsdcw.HINBCW = typ_AType.typ_Param.HINBCA
    wkXsdcw.REVNUMCW = typ_AType.typ_Param.REVNUMCA
    wkXsdcw.FACTORYCW = typ_AType.typ_Param.FACTORYCA
    wkXsdcw.OPECW = typ_AType.typ_Param.OPECA
    typ_CType.DkTmpSiyo = GetWfDKTmpCode(True, wkXsdcw)
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------

    ReDim tWarpMeasG(0)
    ReDim tKakuMeasG(0)
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    tMapHinG = tMapHin(1)
    tNew_Hinban = tMapHinG.HIN
    ''Warp判定対応　06/01/11 ooba END ====================================>
    
    '呼び出す関数が違っていたから、変更したよ。2003/10/08 SystemBrain MM
    If funChkFurikaeShiyou(PROCD_WFC_SOUGOUHANTEI, txtSxlId.text, tOld_Hinban, tNew_Hinban, _
                           intErrCode, strErrMsg, typ_b, typ_CType, 0) < FUNCTION_RETURN_SUCCESS Then
        f_cmbc039_2.Enabled = True
        f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
        lblMsg.Caption = strErrMsg
        EndProcess '' プロセス終了
'        Exit Function
        'Add Start 2011/05/10 SMPK Miyata
        lblMidMsg.Caption = typ_CType.sMidErrMsg
        'Add End   2011/05/10 SMPK Miyata
        Exit Sub
    End If
    'Add Start 2011/05/10 SMPK Miyata
    lblMidMsg.Caption = typ_CType.sMidErrMsg
    'Add End   2011/05/10 SMPK Miyata

' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    EPSiyouSansyouFlg = False
    If typ_CType.typ_si.HEPHS = True Then
        '' エピ先行評価項目に判定NGがある場合は、切換えボタンを赤色で表示する
        '' その場合、エピ判定結果を参照するまで実行ボタンを押下不可とする
        If RET_3_4 > 0 Then
            f_cmbc039_2.CmdChangeWF_EP.backColor = vbRed
        Else
            f_cmbc039_2.CmdChangeWF_EP.backColor = &H8000000F
            EPSiyouSansyouFlg = True
        End If
        f_cmbc039_2.spdHinbanCenEpi.Visible = False
        f_cmbc039_2.CmdChangeWF_EP.Enabled = True
    Else
        f_cmbc039_2.CmdChangeWF_EP.backColor = &H8000000F
        f_cmbc039_2.spdHinbanCenEpi.Visible = False
        f_cmbc039_2.CmdChangeWF_EP.Enabled = False
        EPSiyouSansyouFlg = True
    End If
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    
    tMapHin(1).WARPFLG = tMapHinG.WARPFLG       'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ　06/01/12 ooba
    tMapHin(1).KAKUFLG = tMapHinG.KAKUFLG       '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ　06/01/12 ooba
    
    If intErrCode = 0 Then
        TotalJudg039 = True
    Else
        TotalJudg039 = False
    End If
    
    '振替ﾁｪｯｸ未実施品番のWarp/合成角度判定　06/01/11 ooba START ========================>
    lblMsg.Caption = ""
    Dim i, j    As Integer
    For i = 1 To UBound(tMapHin)
        '振替ﾁｪｯｸ実施の確認
        tMapHinG = tMapHin(i)
        For j = 1 To 2
            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                intRet = funChkFurikaeShiyou("CW763", txtSxlId.text, tMapHinG.HIN, _
                                             tMapHinG.HIN, intErrCode, strErrMsg, _
                                             typ_b, typ_CType, 0)

                tMapHin(i).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                '判定NG
                If intRet = 1 Then
                    TotalJudg039 = False
                '振替ﾁｪｯｸｴﾗｰ
                ElseIf intRet < 0 Then
                    f_cmbc039_2.Enabled = True
                    f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                    f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                    lblMsg.Caption = strErrMsg
                    EndProcess
                    Exit Sub
                End If
            End If
        Next j
    Next i
    'Warp判定NGでも仕様がない場合は総合判定OKとする
    'Nr濃度追加　06/06/08 ooba
    '2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
    With typ_CType.typ_si
        If Not fnc_CheckHWS(.HWFRHWYS) And _
           Not fnc_CheckHWS(.HWFONHWS) And _
           Not fnc_CheckHWS(.HWFBM1HS) And _
           Not fnc_CheckHWS(.HWFBM2HS) And _
           Not fnc_CheckHWS(.HWFBM3HS) And _
           Not fnc_CheckHWS(.HWFOF1HS) And _
           Not fnc_CheckHWS(.HWFOF2HS) And _
           Not fnc_CheckHWS(.HWFOF3HS) And _
           Not fnc_CheckHWS(.HWFOF4HS) And _
           Not fnc_CheckHWS(.HWFDSOHS) And _
           Not fnc_CheckHWS(.HWFMKHWS) And _
           Not (fnc_CheckHWS(.HWFSPVHS) Or fnc_CheckHWS(.HWFDLHWS) Or fnc_CheckHWS(.HWFNRHS)) And _
           Not fnc_CheckHWS(.HWFOS1HS) And _
           Not fnc_CheckHWS(.HWFOS2HS) And _
           Not fnc_CheckHWS(.HWFOS3HS) And _
           Not fnc_CheckHWS(.HWFZOHWS) And _
           Not fnc_CheckHWS(.HWFDENHS) And _
           Not fnc_CheckHWS(.HWFLDLHS) And _
           Not fnc_CheckHWS(.HWFDVDHS) And _
           Not fnc_CheckHWS(.HEPBM1HS) And _
           Not fnc_CheckHWS(.HEPBM2HS) And _
           Not fnc_CheckHWS(.HEPBM3HS) And _
           Not fnc_CheckHWS(.HEPOF1HS) And _
           Not fnc_CheckHWS(.HEPOF2HS) And _
           Not fnc_CheckHWS(.HEPOF3HS) Then
            TotalJudg039 = True
        End If
    End With
    'Warp判定NGの場合ｴﾗｰﾒｯｾｰｼﾞ表示
    For i = 1 To UBound(tWarpMeasG)
        If tWarpMeasG(i).EXISTFLG >= 0 Then
            If Not tWarpMeasG(i).Judg Then
                lblMsg.Caption = "Warp判定エラー　品番振替を行ってください。"
                Exit For
            End If
        End If
    Next i
    'Warp/合成角度情報表示
    Call WarpKakuDisp(Me)
    '振替ﾁｪｯｸ未実施品番のWarp/合成角度判定　06/01/11 ooba END ==========================>
                
'ここからtyp_Aからtyp_Cを使用すること
    
    '規格情報表示
    sub_PutSeihinTop        '上段
    sub_PutSeihinCenter     '中段
    sub_PutSeihinTail       '下段
' 06/08/15 Add エピ先行評価追加対応 SMP)hama -s-
    If typ_CType.typ_si.HEPHS = True Then
        sub_PutSeihinEpi        'エピ
    End If
' 06/08/15 Add エピ先行評価追加対応 SMP)hama -e-
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'SP表示 → SPV(Fe)、SPV(拡散長)、SPV(Nr)表示による変更
    sub_PutSeihinTail2      '下段2
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    'typ_rtInit
    
    '比抵抗情報表示
    sub_PutRs

    '実績情報表示
    'sub_PutRslt typ_AType.typ_rslt(), SxlTop039
    sub_PutRslt typ_CType.typ_rslt(), SxlTop039
    
    'sub_PutRslt typ_AType.typ_rslt(), SxlTail039
    sub_PutRslt typ_CType.typ_rslt(), SxlTail039

'Add Start 2011/03/09 SMPK Miyata
    sub_PutRslt typ_CType.typ_rslt(), SxlMidl039

    '中間抜試サンプル位置ボタン設定
    sub_SampleMidlePosBtnSet
    
'Add End   2011/03/09 SMPK Miyata
    
'Add Start 2011/08/10 Y.Hitomi
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
        
        If typ_CType.typ_si.MSMPFLG = "1" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "中間抜試(製品保証)"
        ElseIf typ_CType.typ_si.MSMPFLG = "2" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "中間抜試(製作参考)"
        ElseIf typ_CType.typ_si.MSMPFLG = "3" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "中間抜試(製作保証)"
        End If
        
        '中間抜試単位(中間抜試許容値(枚数)/(mm))
        lblMSMP_TANI.Visible = True
        '中間抜試単位(mm)を取得
        If getMSMPTANI(tNew_Hinban, iMSMPTANI) = FUNCTION_RETURN_FAILURE Then
            iMSMPTANI = 0
        End If
        lblMSMP_TANI.Caption = "抜試単位：" & vbCrLf & _
        CInt(typ_CType.typ_si.MSMPTANIMAI) & "枚"
        '中間抜試必要数
        lblMSMP_SUU.Visible = True
        '中間抜試の必要数 = (SXLのWF枚数 - 中間抜試許容値(枚数)) / 中間抜試単位(枚数)
        iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
        'マイナスの場合、０とする
        If iMinMidCnt < 0 Then iMinMidCnt = 0
        '中間抜試の件数
        iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
        lblMSMP_SUU.Caption = "実績/必要数：" & vbCrLf & _
        CInt(iRstMidCnt) & "/" & CInt(iMinMidCnt) & "枚"
        '中間抜試単位(枚数)
        lblMSMP_JOSU.Visible = True
        lblMSMP_JOSU.Caption = "許容枚数：" & vbCrLf & _
        CInt(typ_CType.typ_si.MSMPCONSTMAI) & "枚"
    Else
        lblMSMP_FLG.Visible = False
        lblMSMP_TANI.Visible = False
        lblMSMP_SUU.Visible = False
        lblMSMP_JOSU.Visible = False
    End If
'Add End 2011/08/10 Y.Hitomi
    

'end 共通モジュールへ変更になる

    'ｻﾝﾌﾟﾙ異常処理追加　06/10/19 ooba
    For i = 1 To 2
        With typ_CType.typ_Param.WFSMP(i)
            If .WFRESRS1CW = "2" Or _
               .WFRESOICW = "2" Or _
               .WFRESB1CW = "2" Or _
               .WFRESB2CW = "2" Or _
               .WFRESB3CW = "2" Or _
               .WFRESL1CW = "2" Or _
               .WFRESL2CW = "2" Or _
               .WFRESL3CW = "2" Or _
               .WFRESL4CW = "2" Or _
               .WFRESDSCW = "2" Or _
               .WFRESDZCW = "2" Or _
               .WFRESSPCW = "2" Or _
               .WFRESDO1CW = "2" Or _
               .WFRESDO2CW = "2" Or _
               .WFRESDO3CW = "2" Or _
               .WFRESAOICW = "2" Or _
               .WFRESGDCW = "2" Or _
               .EPRESB1CW = "2" Or _
               .EPRESB2CW = "2" Or _
               .EPRESB3CW = "2" Or _
               .EPRESL1CW = "2" Or _
               .EPRESL2CW = "2" Or _
               .EPRESL3CW = "2" Then
               
                f_cmbc039_2.Enabled = True
                txtStaffID.Locked = True
                lblMsg.Caption = "サンプル異常 (" & .REPSMPLIDCW & ")"
                EndProcess
                Exit Sub
            End If
        End With
    Next i
        
    '認定炉判定    *2008/08/28 kameda    mod 2009/10/15 Kameda
    'If gNinteiro_Data.JUDGRO = "0" Then
    If gNinteiro_Data(1).JUDGRO = "0" Then
        txtRoJdg.text = "OK"
    'ElseIf gNinteiro_Data.JUDGRO = "-1" Then
    ElseIf gNinteiro_Data(1).JUDGRO = "-1" Then
        TotalJudg039 = False
        txtRoJdg.text = "NG"
        CtrlEnabled txtRoJdg, CTRL_DISABLE_WARNING, False
    End If
        
    '払出し規制追加   *2010/02/15 Kameda
    PutAllData_Haraidashi
    
    'マルチ引上げ適用判定  2011/05/31 Kameda
    If tbl_chk2_5.MLTJDG(1) = "-1" Then
        TotalJudg039 = False
        'Add Start 2012/07/09 Y.Hitomi
        lblMsg.Caption = "マルチ適用不可品番の為、流動できません。"
        'lblMsg.Caption = "マルチ引上げ適用エラー"
        'Add End 2012/07/09 Y.Hitomi
    End If
    
    
     '>>>>> Mod Start 2012/09/07 SETsw Marushita WF10枚以下を流動可とする
'    'Add Start 2010/08/26 Y.Hitomi WF10枚以下は、流動不可とする
'    With typ_CType.typ_Param
'        If .COUNT <= 10 Then
'            TotalJudg039 = False
'            lblMsg.Caption = "WF枚数が10枚以下です。"
'            CtrlEnabled lblIchi, CTRL_DISABLE_WARNING, False
'        End If
'    End With
'    'Add End  2010/08/26 Y.Hitomi
     '<<<<< Mod End 2012/09/07 SETsw Marushita WF10枚以下を流動可とする
    
    'Add Start 2011/09/28 Y.Hitomi SXLID確定可否フラグチェック対応
    If getSXLIDFLG(tNew_Hinban, sSXLIDFLG) = FUNCTION_RETURN_SUCCESS Then
        If sSXLIDFLG = "1" Then
            TotalJudg039 = False
            lblMsg.Caption = "SXLID確定不可品番の為、流動できません。"
        End If
    Else
        TotalJudg039 = False
        lblMsg.Caption = "SXLID確定可否チェックエラー"
    End If
    'Add End  2011/09/28 Y.Hitomi
    
    EndProcess '' プロセス終了
    f_cmbc039_2.Enabled = True
    
    lblMukesaki.Caption = sCmbMukeName
    
    ' 承認機能追加による修正  07/10/05 miyatake ===================> START
    ''PNG保存チェックボックスON/OFF
    If Trim(GetCodeFieldA9(SWS_CHK_KEY1, SWS_CHK_KEY2, SWS_CHK_KEY3, SWS_CHK_COLUMN)) = SWS_CHK_VALUE_ON Then
        Me.chk_Png = 1
    ElseIf Trim(GetCodeFieldA9(SWS_CHK_KEY1, SWS_CHK_KEY2, SWS_CHK_KEY3, SWS_CHK_COLUMN)) = SWS_CHK_VALUE_OFF Then
        Me.chk_Png = 0
    Else
        Me.chk_Png = 0
    End If
    ' 承認機能追加による修正  07/10/05 miyatake ===================> END
    
    '結晶位置表示  2010/02/15 Kameda
    lblIchi.Caption = GetXtalPos(txtSxlId.text)
    
    'フォーカスセット（担当者）
    txtStaffID.SetFocus
End Sub

'*******************************************************************************
'*    関数名        : Sub_SetParamData
'*
'*    処理概要      : 1.前画面からの引数を設定する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub Sub_SetParamData()
    Call Sub_S_SetParamData
End Sub

'*******************************************************************************
'*    関数名        : sub_InitDisp
'*
'*    処理概要      : 1.画面の初期化処理
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Public Sub sub_InitDisp()
    intEnCmd = 0                                    'ボタン表示標準
    lblMsg.Caption = ""
    lblMidMsg.Caption = ""                          'Add 2011/05/10 SMPK Miyata

    CtrlEnabled txtSXLTop, CTRL_DISABLE, True       'SXLトップ位置
    CtrlEnabled txtCutPosTop, CTRL_DISABLE, True    '最カット位置（トップ）
    CtrlEnabled txtRRGTop, CTRL_DISABLE, True       'RRG（トップ）
    CtrlEnabled txtSXLTail, CTRL_DISABLE, True      'SXLテイル位置
    CtrlEnabled txtCutPosTail, CTRL_DISABLE, True   '最カット位置（テイル）
    CtrlEnabled txtRRGTail, CTRL_DISABLE, True      'RRG（テイル）
    CtrlEnabled txtJHAll, CTRL_DISABLE, True        '実行偏析全体
    CtrlEnabled txtRoJdg, CTRL_DISABLE, True        '認定炉判定 *2008/08/28 kameda
    CtrlEnabled txtKisei, CTRL_DISABLE, True        '払出規制　 *2010/02/15 kameda
'Add Start 2011/03/09 SMPK Miyata
    CtrlEnabled txtSXLMid, CTRL_DISABLE_GRAY, True      'SXL中間抜試位置
    CtrlEnabled txtCutPosMid, CTRL_DISABLE_GRAY, True   '最カット位置（中間抜試）
    CtrlEnabled txtRRGMid, CTRL_DISABLE_GRAY, True      'RRG（中間抜試）
'Add End   2011/03/09 SMPK Miyata
'Add Start 2011/08/25 Y.Hitomi
    CtrlEnabled txtANTempMid, CTRL_DISABLE_GRAY, True   'AN温度（中間)
    CtrlEnabled txtDKTmpMid, CTRL_DISABLE_GRAY, True    'DK温度（中間)
'Add End   2011/08/25 Y.Hitomi

    Call InitHensu(typ_AType)
    
    With f_cmbc039_2
        '規格シート上段
        ''2001/07/27 修正
        SpCtrlInit .spdHinbanTop, 1
    '↓変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        'AN温度追加
        SpCtrlBlockEnabled .spdHinbanTop, 1, 1, 5, 2, CTRL_DISABLE
    '↑変更 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        '規格シート中段
        ''2001/07/27 修正
        SpCtrlInit .spdHinbanCen, 1
        SpCtrlBlockEnabled .spdHinbanCen, 1, 1, 9, 2, CTRL_DISABLE
        '規格シート下段
        ''2001/07/27 修正
        SpCtrlInit .spdHinbanTail, 1
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 7, 2, CTRL_DISABLE
    '*** UPDATE ↓ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数追加
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 10, 2, CTRL_DISABLE    'GD仕様表示追加　05/02/04 ooba
'↓変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'SP表示 → SPV(Fe)、SPV(拡散長)、SPV(Nr)表示による変更
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 11, 2, CTRL_DISABLE
        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 10, 2, CTRL_DISABLE
'↑変更 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
    '*** UPDATE ↑ Y.SIMIZU 2005/10/1 GDﾗｲﾝ数追加
'↓追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'SP表示 → SPV(Fe)、SPV(拡散長)、SPV(Nr)表示による変更
        '規格シート下段2
        SpCtrlInit .spdHinbanTail2, 1
'        SpCtrlBlockEnabled .spdHinbanTail2, 1, 1, 3, 2, CTRL_DISABLE
        SpCtrlBlockEnabled .spdHinbanTail2, 1, 1, 13, 2, CTRL_DISABLE   '08/03/12 ooba
'↑追加 SPV判定処理追加 2006/06/12 SMP)kondoh ---------------
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        '規格シートエピ
        SpCtrlInit .spdHinbanCenEpi, 1
        SpCtrlBlockEnabled .spdHinbanCenEpi, 1, 1, 6, 2, CTRL_DISABLE
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        
        '比抵抗(TOP)
        ''2001/07/27 修正
        SpCtrlInit .spdMeasTop, 5
        SpCtrlBlockEnabled .spdMeasTop, 1, 1, 1, 5, CTRL_DISABLE
        
        '比抵抗(TAIL)
        ''2001/07/27 修正
        SpCtrlInit .spdMeasTail, 5
        SpCtrlBlockEnabled .spdMeasTail, 1, 1, 1, 5, CTRL_DISABLE

'Add Start 2011/03/10 SMPK Miyata
        '比抵抗(MIDLE)
        ''2001/07/27 修正
        SpCtrlInit .spdMeasMid, 5
        SpCtrlBlockEnabled .spdMeasMid, 1, 1, .spdMeasMid.MaxCols, .spdMeasMid.MaxRows, CTRL_DISABLE_GRAY, True
'Add End   2011/03/10 SMPK Miyata

        '判定実績(TOP)
        ''2001/07/27 修正
        SpCtrlInit .spdKensaTop, 0
        
        '判定実績(TAIL)
        ''2001/07/27 修正
        SpCtrlInit .spdKensaTail, 0
        
        '判定実績(MIDLE)
        SpCtrlInit spdKensaMid, 0           'Add 2011/03/10 SMPK Miyata

    End With
End Sub

'*******************************************************************************
'*    関数名        : fnc_ExecutionProcess
'*
'*    処理概要      : 1.入力画面においての入力された値を登録する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    Index       ,I  ,Integer　,Cmdボタン配列の添字
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_ExecutionProcess(Index As Integer) As FUNCTION_RETURN
    Dim udtImgData              As typImgData           '07/10/05 miyatake START ================>
    Dim udtImgData_Detail(0)    As typImgData_Detail
    Dim sErrMsg                 As String
    
    udtImgData.detail = udtImgData_Detail     '07/10/05 miyatake END ==================>
    
    '' パラメータ初期化
    fnc_ExecutionProcess = FUNCTION_RETURN_FAILURE

    '' パラメータ判定処理を行う
    If StaffIDCheck(txtStaffID, txtJfName, lblMsg) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

    typ_CType.StrStaffId = Trim(txtStaffID.text) ' スタッフID
    typ_CType.strStaffName = Trim(txtJfName.text) ' スタッフ名
    typ_CType.typ_Param.SXLID = SelectSxlID039 ' ブロックID
     
    ''データ登録を行う
    Select Case Index
        Case 12
            BeginProcess '' プロセス開始
    
            If TotalJudg039 Then
                OraDB.BeginTrans
                If RegWfSogoRsltOK() <> FUNCTION_RETURN_SUCCESS Then
                    OraDB.Rollback
                    EndProcess '' プロセス終了
                    Exit Function
                End If
                Debug.Print "新DB書込み処理開始"
                If MakeParameter(WF_HANTEI_FORM) <> FUNCTION_RETURN_SUCCESS Then
                    OraDB.Rollback
                    Debug.Print "新DB書込み処理異常終了"
                    Call clearType  '構造体初期化
                    EndProcess '' プロセス終了
                    Exit Function
                End If
    '            OraDB.Rollback
    
                ' 承認機能追加による修正  07/10/05 miyatake ===================> START
                ''PNGファイル作成
                If Me.chk_Png = 1 Then
                    udtImgData.xtal = BlkNow.XTALC2
                    udtImgData.STAFFID = txtStaffID
                    udtImgData.SXLID = txtSxlId
                    udtImgData.memo = MsComment
'                    If FileCreate_PNG(PROCD_WFC_SOUGOUHANTEI, udtImgData, Me, sErrMsg, Nothing, pic_Png) = FUNCTION_RETURN_FAILURE Then
                    If FileCreate_PNG(PROCD_WFC_SOUGOUHANTEI, udtImgData, Me, sErrMsg, Nothing, pic_Png) = False Then 'upd 09/02/04 SETmiyatake
                        OraDB.Rollback
                        lblMsg.Caption = sErrMsg
                        Debug.Print "PNGファイル作成処理異常終了"
                        Call clearType  '構造体初期化
                        EndProcess '' プロセス終了
                        Exit Function
                    End If
                End If
                ' 承認機能追加による修正  07/10/05 miyatake ===================> END
    
                Call clearType  '構造体初期化
                
                OraDB.CommitTrans
                Debug.Print "新DB書込み処理正常終了"
                
                ' 承認機能追加による修正  07/10/05 miyatake ===================> START
                If Me.chk_Png = 1 Then
                    ''PNGファイル送信
                    Call FileReSend_PNG(PROCD_WFC_SOUGOUHANTEI)
                End If
                ' 承認機能追加による修正  07/10/05 miyatake ===================> END
            Else
                EndProcess '' プロセス終了
                lblMsg.Caption = GetMsgStr(TJE01)
                Exit Function
            End If
    End Select
    
    '' 処理正常終了
    fnc_ExecutionProcess = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    関数名        : fnc_LoadData_Warp
'*
'*    処理概要      : 1.Warp/合成角度判定用ﾃﾞｰﾀ取得
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_LoadData_Warp() As FUNCTION_RETURN

    Dim i, j, k, m, n       As Integer
    Dim RET                 As FUNCTION_RETURN
    Dim udtWarpMapData()    As type_DBDRV_Nukisi
    Dim udtTmp_Y018()       As typ_WarpKakuData     '標準測定ﾃﾞｰﾀ(TBCMY018)取得用
    
    fnc_LoadData_Warp = FUNCTION_RETURN_FAILURE
    
    ReDim tSXLID(0)
    tSXLID(0).SXLID = txtSxlId.text
    '関連ﾌﾞﾛｯｸID取得
    If DBDRV_BLOCKIDGET() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("ESXL2")
        Exit Function
    End If
    
    'WFﾏｯﾌﾟﾃﾞｰﾀ取得
    If DBDRV_WARPMAPGET(udtWarpMapData()) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EGET2", "Y011")
        Exit Function
    End If
    
    'Warp/合成角度ﾃﾞｰﾀ取得
    ReDim tWarpInitG(0)
    ReDim tKakuInitG(0)
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    ReDim tKakuXInitG(0)
    ReDim tKakuYInitG(0)
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    bMapWarpFlg = False
    
    For i = 0 To UBound(tSXLID)
        '合成角度ﾃﾞｰﾀ取得
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "ORIENT", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '合成角度ﾃﾞｰﾀｾｯﾄ
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuInitG(m + n)
            
            For j = 1 To n
                tKakuInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        
        'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
        '横(X)角度ﾃﾞｰﾀ取得
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "XKAKU", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '横(X)角度ﾃﾞｰﾀｾｯﾄ
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuXInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuXInitG(m + n)

            For j = 1 To n
                tKakuXInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        
        '縦(Y)角度ﾃﾞｰﾀ取得
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "YKAKU", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '縦(Y)角度ﾃﾞｰﾀｾｯﾄ
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuYInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuYInitG(m + n)

            For j = 1 To n
                tKakuYInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応

        'Warpﾃﾞｰﾀ取得
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "WARP", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        'Warpﾃﾞｰﾀｾｯﾄ
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tWarpInitG)
            n = UBound(udtTmp_Y018)
            k = 0
            Call fnc_MapWarpChk(udtTmp_Y018())
            For j = 1 To n
                'WFﾏｯﾌﾟに紐付かないﾃﾞｰﾀはｾｯﾄしない
                If udtTmp_Y018(j).EXISTFLG <> -1 Then
                    k = k + 1
                    ReDim Preserve tWarpInitG(m + k)
                    tWarpInitG(m + k) = udtTmp_Y018(j)
                Else
                    bMapWarpFlg = True
                End If
            Next j
        End If
    Next i
    
    'WFﾏｯﾌﾟ上の品番情報取得
    ReDim tMapHin(0)
    m = 0
    For i = 1 To UBound(udtWarpMapData) Step 2
        If udtWarpMapData(i).hinban <> vbNullString And _
           Trim(udtWarpMapData(i).hinban) <> "Z" And _
           Trim(udtWarpMapData(i).hinban) <> "G" Then
           
            m = m + 1
            ReDim Preserve tMapHin(m)
            '品番
            tMapHin(m).HIN.hinban = udtWarpMapData(i).hinban
            tMapHin(m).HIN.mnorevno = udtWarpMapData(i).REVNUM
            tMapHin(m).HIN.factory = udtWarpMapData(i).factory
            tMapHin(m).HIN.opecond = udtWarpMapData(i).opecond
            'ﾌﾞﾛｯｸID
            tMapHin(m).BLOCKID = udtWarpMapData(i).LOTID
            'ﾌﾞﾛｯｸ内連番(Start)
            tMapHin(m).BLKSEQ_S = CInt(udtWarpMapData(i).BLOCKSEQ)
            'ﾌﾞﾛｯｸ内連番(End)
            tMapHin(m).BLKSEQ_E = CInt(udtWarpMapData(i + 1).BLOCKSEQ)
            '振替ﾁｪｯｸﾌﾗｸﾞ
            tMapHin(m).WARPFLG = False
            tMapHin(m).KAKUFLG = False
        End If
    Next i
    
    fnc_LoadData_Warp = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    関数名        : fnc_MapWarpChk
'*
'*    処理概要      : 1.WFﾏｯﾌﾟとWarp実績の紐付きﾁｪｯｸ
'*
'*    パラメータ    : 変数名      ,IO ,型 　　　　　      ,説明
'*                    udtChkWarp()  ,I  ,typ_WarpKakuData   ,標準測定ﾃﾞｰﾀ(Warp実績)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Sub fnc_MapWarpChk(udtChkWarp() As typ_WarpKakuData)

    Dim i, j, k, m, n   As Integer
    
    k = 1
    m = UBound(sWrpLOTID)
    n = UBound(udtChkWarp)
    
    For i = 1 To n
        udtChkWarp(i).EXISTFLG = 0        'WFﾏｯﾌﾟ範囲外の実績も判定対象とする。
        For j = k To m
            'WFﾏｯﾌﾟとWarp実績のﾌﾞﾛｯｸID／ﾌﾞﾛｯｸ内連番が一致すれば紐付き有り
            If udtChkWarp(i).BLOCKID = sWrpLOTID(1) And _
               udtChkWarp(i).WAFID < iWrpBLOCKSEQ(1) Then
            
                udtChkWarp(i).EXISTFLG = 0
                k = j
                Exit For
            ElseIf udtChkWarp(i).BLOCKID = sWrpLOTID(m) And _
                   udtChkWarp(i).WAFID > iWrpBLOCKSEQ(m) Then
            
                udtChkWarp(i).EXISTFLG = 0
                k = j
                Exit For
            ElseIf udtChkWarp(i).BLOCKID = sWrpLOTID(j) And _
                   udtChkWarp(i).WAFID = iWrpBLOCKSEQ(j) Then
            
                udtChkWarp(i).EXISTFLG = 1
                k = j + 1
                Exit For
            End If
        Next j
    Next i
End Sub

'*******************************************************************************
'*    関数名        : fnc_CheckHWS
'*
'*    処理概要      : 1.処理方法をチェックして検査の有無を返す
'*
'*    パラメータ    : 変数名      ,IO ,型 　　　,説明
'*                    sHWS  　　　,I  ,String 　,処理方法
'*
'*    戻り値        : Boolean 検査の有無
'*
'*******************************************************************************
Private Function fnc_CheckHWS(ByVal sHWS As String) As Boolean
    If sHWS = "H" Or sHWS = "S" Then
        fnc_CheckHWS = True
    Else
        fnc_CheckHWS = False
    End If
End Function

'*******************************************************************************
'*    関数名        : spdKensaTail_TopLeftChange
'*
'*    処理概要      : 1.実績表示一覧TOP/BOT間で、横スクロールを連動させる
'*
'*    パラメータ    : 変数名      ,IO ,型 　　　,説明
'*                    oldLeft  　 ,I  ,String 　,スクロール前の最左列の列番号
'*                    oldTop   　 ,I  ,String 　,スクロール前の最上行の行番号
'*                    NewLeft  　 ,I  ,String 　,スクロール後の最左列の列番号
'*                    NewTop   　 ,I  ,String 　,スクロール後の最上行の行番号
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub spdKensaTail_TopLeftChange(ByVal oldLeft As Long, ByVal oldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdKensaTop.LeftCol = spdKensaTail.LeftCol
End Sub

'*******************************************************************************
'*    関数名        : spdKensaTop_TopLeftChange
'*
'*    処理概要      : 1.実績表示一覧TOP/BOT間で、横スクロールを連動させる
'*
'*    パラメータ    : 変数名      ,IO ,型 　　　,説明
'*                    oldLeft  　 ,I  ,String 　,スクロール前の最左列の列番号
'*                    oldTop   　 ,I  ,String 　,スクロール前の最上行の行番号
'*                    NewLeft  　 ,I  ,String 　,スクロール後の最左列の列番号
'*                    NewTop   　 ,I  ,String 　,スクロール後の最上行の行番号
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub spdKensaTop_TopLeftChange(ByVal oldLeft As Long, ByVal oldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdKensaTail.LeftCol = spdKensaTop.LeftCol
End Sub

'*******************************************************************************
'*    関数名        : sub_PutSeihinTop
'*
'*    処理概要      : 1.製品シート表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSeihinTop()
    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ

    With f_cmbc039_2
    '↓変更 熱処理判断処理追加
    '2.1.3 AN温度 実績反映チェック追加
        'AN温度追加
        For i = 1 To 5
            .spdHinbanTop.col = i
            .spdHinbanTop.row = 1
            Select Case i
            Case 1
                '品番
                '品番12桁表示-------Start SystemBrain 2003/10/05
                .spdHinbanTop.Value = typ_CType.typ_Param.hinban & Format(typ_CType.typ_Param.REVNUM, "00") & typ_CType.typ_Param.factory & typ_CType.typ_Param.opecond
                '.spdHinbanTop.Value = typ_CType.typ_Param.HINBCA
            Case 2
                'タイプ
                .spdHinbanTop.Value = typ_CType.typ_si.HWFTYPE
            Case 3
                '方位
                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDIR
            Case 4
                '結晶ドープ
                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDOP
                
                '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
                '2.1.3 AN温度 実績反映チェック追加
                'AN温度追加
            Case 5
                'AN温度
                .spdHinbanTop.Value = typ_CType.typ_si.HWFANTNP
            End Select
        Next i
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_PutSeihinCenter
'*
'*    処理概要      : 1.製品シート表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSeihinCenter()
    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ

    'CENTER側
    With f_cmbc039_2
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
''        For i = 1 To 9
        For i = 1 To 10
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
            .spdHinbanCen.col = i
            .spdHinbanCen.row = 1

            Select Case i
            Case 1
                '比抵抗
                .spdHinbanCen.Value = toRsStr_nl(typ_CType.typ_si.HWFRMIN, typ_CType.typ_si.HWFRMAX)
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFR = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata

            Case 2
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
                'DK温度
                .spdHinbanCen.Value = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpSiyo))
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFR = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 3
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'Oi
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFONMIN, "0.00") & " - " & DBData2DispData_nl(typ_CType.typ_si.HWFONMAX, "0.00")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFO = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 3
            Case 4
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'BMD1
                'べき乗数変更
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFBM1AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HWFBM1AX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 4
            Case 5
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'BMD2
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFBM2AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HWFBM2AX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 5
            Case 6
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'BMD3
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFBM3AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HWFBM3AX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 6
            Case 7
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'OSF1
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOF1AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HWFOF1MX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 7
            Case 8
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'OSF2
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOF2AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HWFOF2MX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 8
            Case 9
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'OSF3
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOF3AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HWFOF3MX, "0.0")
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
'            Case 9
            Case 10
'--------------- 2008/08/25 UPDATE  END   By Systech ---------------
                'OSF4
'                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOF4AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HWFOF4MX, "0.0")
                'Change 2010/01/17 SIRD対応　Y.Hitomi
                'SIRD
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSIRDMX, "##0")
            End Select
        Next i
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_PutSeihinTail
'*
'*    処理概要      : 1.製品シート表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSeihinTail()
    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ

    'TAIL側
    With f_cmbc039_2
        For i = 1 To 11
            .spdHinbanTail.col = i
            .spdHinbanTail.row = 1
            Select Case i
                Case 1
                    'DS
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDSOMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDSOMX, "0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFDS = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 2
                    'DZ
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFMKMIN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFMKMAX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFDZ = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 3
                    'SP
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVMX, "0.00") & " , " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDLMIN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0")
                Case 4
                    'D1
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOS1MN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFOS1MX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFDOI = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 5
                    'D2
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOS2MN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFOS2MX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFDOI = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 6
                    'D3
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFOS3MN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_AType.typ_si.HWFOS3MX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFDOI = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                ''残存酸素仕様表示追加
                Case 7
                    'AO
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFZOMIN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFZOMAX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFAOI = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                ''GD仕様表示追加
                'GDﾗｲﾝ数追加
                Case 8
                    'GDﾗｲﾝ数
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFGDLINE, "")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 9
                    'Den
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDENMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDENMX, "0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 10
                    'L/DL
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech Start
''                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFLDLMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLMX, "0")
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFLDLMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLMX, "0") & " , " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLRMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLRMX, "0")
'' 2008/10/01 L/DL,OSF判定ﾛｼﾞｯｸ追加 UPD By Systech End
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 11
                    'DVD2
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDVDMNN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDVDMXN, "0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            End Select
        Next i
    End With

End Sub

'*******************************************************************************
'*    関数名        : sub_PutSeihinTail2
'*
'*    処理概要      : 1.製品シート表示
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSeihinTail2()
    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ

    'TAIL2側
    With f_cmbc039_2
''        For i = 1 To 3
''            .spdHinbanTail2.col = i
''            .spdHinbanTail2.row = 1
''            Select Case i
''            Case 1
''                'SP(Fe)
''                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVMX, "0.00") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUG, "0.00") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUR, "0.000") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFSPVSTD, "0.000")
''            Case 2
''                'SP(拡散長)
''                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMIN, "0.0") & " - " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0.0") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFDLPUG, "0.00") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFDLPUR, "0.000")
''            Case 3
''                'SP(Nr)
''                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRMX, "0.00") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFNRPUG, "0.00") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFNRPUR, "0.000") & " , " & _
''                                      DBData2DispData_nl(typ_CType.typ_si.HWFNRSTD, "0.000")
''            End Select
''        Next i
        
        'SPV表示変更　08/03/13 ooba START ===============================================>
        For i = 1 To 13
            .spdHinbanTail2.col = i
            .spdHinbanTail2.row = 1
            Select Case i
            'SP(Fe)
            Case 1      '上限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVMX, "0.00")
            Case 2      'PUA限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUG, "0.00")
            Case 3      'PUA率
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUR, "0.000")
            Case 4      '標準偏差
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVSTD, "0.000")
            'SP(拡散長)
            Case 5      '下限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMIN, "0.0")
            Case 6      '上限
            
                '保証方法ｺｰﾄﾞが「L」(AVE+MIN)以外の場合
                If typ_CType.typ_si.HWFDLHWT <> "L" Then
                    .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0.0")
                End If
            Case 7      'AVE下限(上限)
            
                '保証方法ｺｰﾄﾞが「L」(AVE+MIN)の場合は上限をAVE下限とする
                If typ_CType.typ_si.HWFDLHWT = "L" Then
                    .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0.0")
                End If
            Case 8      'PUA限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLPUG, "0.00")
            Case 9      'PUA率
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLPUR, "0.000")
            'SP(Nr)
            Case 10     '上限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRMX, "0.00")
            Case 11     'PUA限
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRPUG, "0.00")
            Case 12     'PUA率
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRPUR, "0.000")
            Case 13     '標準偏差
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRSTD, "0.000")
            End Select
        Next i
        'SPV表示変更　08/03/13 ooba END =================================================>
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_PutSeihinEpi
'*
'*    処理概要      : 1.製品シート表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSeihinEpi()
    Dim i As Integer, j As Integer      ' ﾙｰﾌﾟ ｶｳﾝﾀ

    'エピ側
    With f_cmbc039_2
        'Chg Start 2011/04/28 SMPK Miyata　(OSF3Eまで表示していないので修正)
'        For i = 1 To 6
        For i = 1 To 7
        'Chg End   2011/04/28 SMPK Miyata
            .spdHinbanCenEpi.col = i
            .spdHinbanCenEpi.row = 1
            Select Case i
            Case 1
                'BMD1E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPBM1AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HEPBM1AX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 2
                'BMD2E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPBM2AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HEPBM2AX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 3
                'BMD3E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPBM3AN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HEPBM3AX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 4
                'BMD3E(外周)　09/05/07 ooba
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPBM3GSAN, "0.0") & " - " & DBData2DispData_nl(typ_CType.typ_si.HEPBM3GSAX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPBM = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 5
                'OSF1E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPOF1AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HEPOF1MX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 6
                'OSF2E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPOF2AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HEPOF2MX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            Case 7
                'OSF3E
                .spdHinbanCenEpi.Value = DBData2DispData_nl(typ_CType.typ_si.HEPOF3AX, "0.00") & " , " & DBData2DispData_nl(typ_CType.typ_si.HEPOF3MX, "0.0")
                .spdHinbanCenEpi.backColor = IIf(typ_CType.typ_si.MSMPFLGEPOF = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
            End Select
        Next i
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_PutRs
'*
'*    処理概要      : 1.比抵抗値表示(TOP,TAIL)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutRs()
    '比抵抗値表示(TOP側)
    sub_PutRsTop

    '比抵抗値表示(TAIL側)
    sub_PutRsTail

'Add Start 2011/03/09 SMPK Miyata
    '比抵抗値表示(MIDLE側)
    Call sub_PutRsMid(1)
'Add End   2011/03/09 SMPK Miyata

End Sub

'*******************************************************************************
'*    関数名        : sub_PutRsTop
'*
'*    処理概要      : 1.比抵抗値表示(TOP)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutRsTop()
    Dim blJudg  As Boolean  '判定結果
    Dim dblScut As Double   '再カット位置
    Dim dblCoef As Double   '実行偏析

    dblScut = typ_CType.dblScut(SxlTop039)
    dblCoef = typ_CType.COEF(SxlTop039)

    With f_cmbc039_2
        '' WF検査指示（Rs)*****************************************************************
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        'DK温度
        .txtDKTmp(SxlTop).text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(SxlTop)))
        If Not typ_CType.JudgDkTmp(SxlTop) Then
            .txtDKTmp(SxlTop).backColor = COLOR_NG
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        '保証方法ﾁｪｯｸ追加
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "TOP") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then

                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '位置

                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG

                '↓追加 熱処理判断処理追加
                '2.1.3 AN温度 実績反映チェック追加
                    'AN温度追加
                    '項目：DKANの3～6桁がAN温度
                    .txtANTempTop.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTop039, WFRES).DKAN, 3, 4), "0") 'AN温度
                    
                    'チェックNGの時は背景色を変える
                    If Not (typ_CType.JudgAntnp(SxlTop039)) Then
                        CtrlEnabled .txtANTempTop, CTRL_DISABLE_WARNING, False  'AN温度
                    End If
                    
                    'ＷＦサンプル処理変更
                    If Not (typ_CType.JudgRrg(SxlTop039)) Then
                        CtrlEnabled .txtRRGTop, CTRL_DISABLE_WARNING, False  'RRG
                    End If
                    
                    If dblCoef = -1 Or dblCoef = -9999 Then

                        .txtJHAll.text = ""         '実行偏析ブロック
                    Else
                        .txtJHAll.text = DBData2DispData(dblCoef, "0.000")         '実行偏析ブロック
                    End If

                    '再カット位置
                    'ＷＦサンプル処理変更
                    If typ_CType.JudgRes(SxlTop039) Then '2002/03/04 S.Sano
                        .txtCutPosTop.text = "OK"
                    Else
                        Select Case dblScut
                            Case -9999
                                .txtCutPosTop.text = ""
                            Case Is <= typ_CType.typ_Param.INGOTPOS
                                .txtCutPosTop.text = typ_CType.typ_Param.INGOTPOS
                            Case Is >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                                .txtCutPosTop.text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                            Case Else
                                .txtCutPosTail.text = DBData2DispData(dblScut, "0")
                        End Select
                        CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP再カット
                        intEnCmd = 1
                    End If
                        
                    '比抵抗
                    If UBound(typ_CType.typ_y013top) > 0 Then
                        With .spdMeasTop
                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA1)
                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA2)
                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA3)
                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA4)
                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA5)
                        End With
                        RsSpreadSet .spdMeasTop, 1 '2002/01/25 S.Sano
                    Else
                        With .spdMeasTop
                            .col = 1
                            .row = 1:
                            .CellType = CellTypeStaticText
                            .Value = "仕様有"
                            
                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "検査有"
                            
                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "実績無"
                        End With
                    End If
                End If
            Else
                .txtSXLTop.text = ""            '位置
                .txtRRGTop.text = ""            'RRG
                .txtJHAll.text = ""
                
                '再カット位置
                'ＷＦサンプル処理変更
                .txtCutPosTop.text = "NG"
                CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP再カット

                '比抵抗
                With .spdMeasTop
                    .col = 1
                    .row = 1:
                    .CellType = CellTypeStaticText
                    .Value = "仕様有"
                    
                    .row = 2:
                    .CellType = CellTypeStaticText
                    .Value = "検査無"
                    .row = 3: .Value = ""
                    .row = 4: .Value = ""
                    .row = 5: .Value = ""
                End With
            End If
        Else
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '位置
                    
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG
                    If dblCoef = -1 Or dblCoef = -9999 Then
                        .txtJHAll.text = ""         '実行偏析ブロック
                    Else
                        .txtJHAll.text = DBData2DispData(dblCoef, "0.000")         '実行偏析ブロック
                    End If

                    '再カット位置
                    .txtCutPosTop.text = "OK"

                    '比抵抗
                    If UBound(typ_CType.typ_y013top) > 0 Then
                        With .spdMeasTop
                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA1)
                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA2)
                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA3)
                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA4)
                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA5)
                        End With
                        RsSpreadSet .spdMeasTop, 1
                    Else
                        With .spdMeasTop
                            .col = 1
                            .row = 1:
                                    .CellType = CellTypeStaticText
                                    .Value = "仕様無"
                            .row = 2:
                                    .CellType = CellTypeStaticText
                                    .Value = "検査有"
                            .row = 3:
                                    .CellType = CellTypeStaticText
                                    .Value = "実績無"
                        End With
                    End If
                Else
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '位置
                    
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG
                    .txtJHAll.text = ""
                    
                    '再カット位置
                    .txtCutPosTop.text = "OK"
                    
                    '比抵抗
                    With .spdMeasTop
                        .col = 1
                        .row = 1:
                        .CellType = CellTypeStaticText
                        .Value = "仕様無"
                        
                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "検査有"
                        
                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "実績無"
                        
                        .row = 4: .Value = ""
                        .row = 5: .Value = ""
                    End With
                End If
            End If
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_PutRsTail
'*
'*    処理概要      : 1.比抵抗値表示(TAIL)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutRsTail()
    Dim blJudg  As Boolean  '判定結果
    Dim dblScut As Double   '再カット位置

    dblScut = typ_CType.dblScut(SxlTail039)

    With f_cmbc039_2
        '' WF検査指示（Rs)*****************************************************************
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        'DK温度
        .txtDKTmp(SxlTail).text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(SxlTail)))
        If Not typ_CType.JudgDkTmp(SxlTail) Then
            .txtDKTmp(SxlTail).backColor = COLOR_NG
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        '保証方法ﾁｪｯｸ追加
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "BOT") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then

                If typ_CType.typ_Param.WFSMP(SxlTail039).WFRESRS1CW = "1" Then
                    .txtSXLTail.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH, "0")           '位置
                    
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.000000")  'RRG
                    
                    '↓追加 熱処理判断処理追加
                    '2.1.3 AN温度 実績反映チェック追加
                    'AN温度追加
                    '項目：DKANの3～6桁がAN温度
                    .txtANTempTail.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTail039, WFRES).DKAN, 3, 4), "0") 'AN温度
                    'チェックNGの時は背景色を変える
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled .txtANTempTail, CTRL_DISABLE_WARNING, False  'AN温度
                    End If

                    'ＷＦサンプル処理変更
                    If Not (typ_CType.JudgRrg(SxlTail039)) Then
                        CtrlEnabled .txtRRGTail, CTRL_DISABLE_WARNING, False  'RRG
                    End If


                    'ＷＦサンプル処理変更
                    If typ_CType.JudgRes(SxlTail039) Then
                        .txtCutPosTail.text = "OK"
                    Else
                        Select Case dblScut
                            Case -9999
                                .txtCutPosTail.text = ""
                            Case Is <= typ_CType.typ_Param.INGOTPOS
                                .txtCutPosTail.text = typ_CType.typ_Param.INGOTPOS
                            Case Is >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                                .txtCutPosTail.text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                            Case Else
                                .txtCutPosTail.text = DBData2DispData(dblScut, "0")
                        End Select
                        
                        CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail再カット
                        intEnCmd = 1
                    End If

                    '比抵抗
                    If UBound(typ_CType.typ_y013tail) > 0 Then
                        With .spdMeasTail
                            .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA1)
                            .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA2)
                            .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA3)
                            .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA4)
                            .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA5)
                        End With
                        RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
                    Else
                        With .spdMeasTail
                            .col = 1
                            .row = 1:
                            .CellType = CellTypeStaticText
                            .Value = "仕様有"
                            
                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "検査有"
                            
                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "実績無"
                        End With
                    End If
                End If
            Else
                .txtSXLTail.text = ""            '位置
                .txtRRGTail.text = ""            'RRG
                .txtJHAll.text = ""
                '再カット位置

                'ＷＦサンプル処理変更
                .txtCutPosTail.text = "NG"
                CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'Tail再カット

                '比抵抗
                With .spdMeasTail
                    .col = 1
                    .row = 1:
                            .CellType = CellTypeStaticText
                            .Value = "仕様有"
                    .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "検査無"
                    .row = 3: .Value = ""
                    .row = 4: .Value = ""
                    .row = 5: .Value = ""
                End With
            End If
        Else
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then
                .txtSXLTail.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH, "0")            '位置
                
                'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                '.txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.00")  'RRG
                .txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.000000")  'RRG

                '再カット位置
                .txtCutPosTail.text = "OK"

                '比抵抗
                If UBound(typ_CType.typ_y013tail) > 0 Then
                    With .spdMeasTail
                        .SetFloat 1, 1, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA1)
                        .SetFloat 1, 2, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA2)
                        .SetFloat 1, 3, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA3)
                        .SetFloat 1, 4, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA4)
                        .SetFloat 1, 5, CDbl(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA5)
                    End With
                    RsSpreadSet .spdMeasTail, 1 '2002/01/25 S.Sano
                Else
                    With .spdMeasTail
                        .col = 1
                        
                        .row = 1:
                        .CellType = CellTypeStaticText
                        .Value = "仕様無"
                        
                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "検査有"
                        
                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "実績無"
                    End With
                End If
            End If
        End If
    End With
End Sub

'Add Start 2011/03/09 SMPK Miyata
'*******************************************************************************
'*    関数名        : sub_PutRsMid
'*
'*    処理概要      : 1.比抵抗値表示(中間抜試)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                  : iMidNo        ,I  ,Integer  ,中間抜試No(1-10)
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutRsMid(iMidNo As Integer)
    Dim blJudg  As Boolean  '判定結果
    Dim dblScut As Double   '再カット位置
    Dim tt      As Integer  'Top Tail Midle判定用

    tt = SxlMidl + iMidNo - 1

    If tt < 1 Or tt > UBound(typ_CType.typ_Param.WFSMP) Then
        Exit Sub
    End If
    
    dblScut = typ_CType.dblScut(tt)
    
    With f_cmbc039_2
        
        '保証方法ﾁｪｯｸ(品ＷＦ比抵抗検査頻度＿抜) 有りの場合
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "MID") Then
            '状態FLG(Rs)が 1：通常、2：反映、3：推定の場合
            If InStr("123", typ_CType.typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
                '実績FLG1(Rs)が1：実績ありの場合
                If typ_CType.typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                    
                    'Add Start 2011/08/25 Y.Hitomi
                    'DK温度
                    txtDKTmpMid.text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(tt)))
                    If Not typ_CType.JudgDkTmp(tt) Then
                        txtDKTmpMid.backColor = COLOR_NG
                    End If
                    txtSXLMid.backColor = COLOR_SKY         'SXL中間位置
                    txtCutPosMid.backColor = COLOR_SKY      '最カット位置（中間）
                    txtRRGMid.backColor = COLOR_SKY         'RRG（中間）
                    txtANTempMid.backColor = COLOR_SKY      'AN温度（中間）
                    txtDKTmpMid.backColor = COLOR_SKY       'DK温度（中間）
                    SpCtrlBlockEnabled spdMeasMid, 1, 1, spdMeasMid.MaxCols, spdMeasMid.MaxRows, CTRL_DISABLE_SKY, True
                    'Add End   2011/08/25 Y.Hitomi

                    'Mid 位置
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")

                    'Add Start 2011/08/11 Y.Hitomi
                    'AN温度追加 ：DKANの3～6桁がAN温度
                    .txtANTempMid.text = DBData2DispData(Mid(typ_CType.typ_y013(tt, WFRES).DKAN, 3, 4), "0") 'AN温度
                    'チェックNGの時は背景色を変える
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled .txtANTempMid, CTRL_DISABLE_WARNING, False  'AN温度
                    End If
                    'Add End   2011/08/11 Y.Hitomi
                    
                    'ＷＦサンプル処理変更
                    If Not (typ_CType.JudgRrg(tt)) Then
                        'RRG色変更
                        CtrlEnabled .txtRRGMid, CTRL_DISABLE_WARNING, False
                    End If

                    '中間抜試 再カット位置
                    'ＷＦサンプル処理変更
                    If typ_CType.JudgRes(tt) Then
                        .txtCutPosMid.text = "OK"
                    Else
                        Select Case dblScut
                            Case -9999
                                .txtCutPosMid.text = ""
                            Case Is <= typ_CType.typ_Param.INGOTPOS
                                .txtCutPosMid.text = typ_CType.typ_Param.INGOTPOS
                            Case Is >= typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                                .txtCutPosMid.text = typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH
                            Case Else
                                .txtCutPosMid.text = DBData2DispData(dblScut, "0")
                        End Select
                        CtrlEnabled .txtCutPosMid, CTRL_DISABLE_WARNING, False
                        intEnCmd = 1
                    End If

                    '比抵抗
                    If UBound(typ_CType.typ_y013midl_ary(iMidNo).typ_y013midl) > 0 Then
                        With .spdMeasMid
                            .SetFloat 1, 1, val(typ_CType.typ_y013(tt, WFRES).MESDATA1)
                            .SetFloat 1, 2, val(typ_CType.typ_y013(tt, WFRES).MESDATA2)
                            .SetFloat 1, 3, val(typ_CType.typ_y013(tt, WFRES).MESDATA3)
                            .SetFloat 1, 4, val(typ_CType.typ_y013(tt, WFRES).MESDATA4)
                            .SetFloat 1, 5, val(typ_CType.typ_y013(tt, WFRES).MESDATA5)
                        End With
                        RsSpreadSet .spdMeasMid, 1
                    Else
                        With .spdMeasMid
                            .col = 1
                            .row = 1:
                            .CellType = CellTypeStaticText
                            .Value = "仕様有"

                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "検査有"

                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "実績無"
                        End With
                    End If
                End If
            Else
                '状態FLG(Rs)が検査なし(1：通常、2：反映、3：推定以外)の場合

                .txtSXLMid.text = ""            '位置
                .txtRRGMid.text = ""            'RRG

                '中間抜試 再カット位置
                'ＷＦサンプル処理変更
'                .txtCutPosMid.text = "NG"
'                CtrlEnabled .txtCutPosMid, CTRL_DISABLE_WARNING, False
'
            End If
        Else
            '保証方法ﾁｪｯｸ(品ＷＦ比抵抗検査頻度＿抜) なしの場合

            '状態FLG(Rs)が 1：通常、2：反映、3：推定の場合
            If InStr("123", typ_CType.typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
                '実績FLG1(Rs)が1：実績ありの場合
                If typ_CType.typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then

                    'Mid 位置
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")
                    
                    '再カット位置
                    .txtCutPosMid.text = "OK"

                    '比抵抗
                    If UBound(typ_CType.typ_y013midl_ary(tt).typ_y013midl) > 0 Then
                        With .spdMeasMid
                            .SetFloat 1, 1, val(typ_CType.typ_y013(tt, WFRES).MESDATA1)
                            .SetFloat 1, 2, val(typ_CType.typ_y013(tt, WFRES).MESDATA2)
                            .SetFloat 1, 3, val(typ_CType.typ_y013(tt, WFRES).MESDATA3)
                            .SetFloat 1, 4, val(typ_CType.typ_y013(tt, WFRES).MESDATA4)
                            .SetFloat 1, 5, val(typ_CType.typ_y013(tt, WFRES).MESDATA5)
                        End With
                        RsSpreadSet .spdMeasMid, 1
                    Else
                        With .spdMeasMid
                            .col = 1
                            .row = 1:
                                    .CellType = CellTypeStaticText
                                    .Value = "仕様無"
                            .row = 2:
                                    .CellType = CellTypeStaticText
                                    .Value = "検査有"
                            .row = 3:
                                    .CellType = CellTypeStaticText
                                    .Value = "実績無"
                        End With
                    End If
                Else
                    'Mid 位置
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRGの小数桁数を6桁(7桁目四捨五入)に変更 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")

                    '再カット位置
                    .txtCutPosMid.text = "OK"

                    '比抵抗
                    With .spdMeasMid
                        .col = 1
                        .row = 1:
                        .CellType = CellTypeStaticText
                        .Value = "仕様無"

                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "検査有"

                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "実績無"

                        .row = 4: .Value = ""
                        .row = 5: .Value = ""
                    End With
                End If
            End If
        End If
    End With
End Sub
'Add End   2011/03/09 SMPK Miyata

'*******************************************************************************
'*    関数名        : sub_PutRslt
'*
'*    処理概要      : 1.実績値表示(TOP)
'*
'*    パラメータ    : 変数名        ,IO ,型           ,説明
'*                    udt_rslt()    ,I  ,typ_ALLRSLT  ,実績情報構造体
'*                    tt            ,I  ,Integer      ,TopTail判定用
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutRslt(udt_rslt() As typ_ALLRSLT, tt As Integer)
    Dim i, j        As Integer
    Dim spdVa       As vaSpread
    Dim lngSpMaxLine   As Long

'    '最大行数取得
    lngSpMaxLine = 0
    Do While udt_rslt(tt, lngSpMaxLine).OKNG <> ""
        lngSpMaxLine = lngSpMaxLine + 1
    Loop

    If tt = SxlTop039 Then
        Set spdVa = f_cmbc039_2.spdKensaTop
'Chg Start 2011/03/09 SMPK Miyata
'    Else
'        Set spdVa = f_cmbc039_2.spdKensaTail
    ElseIf tt = SxlTail039 Then
        Set spdVa = f_cmbc039_2.spdKensaTail
    Else
        Set spdVa = f_cmbc039_2.spdKensaMid
'Chg End   2011/03/09 SMPK Miyata
    End If

    SpCtrlInit spdVa, lngSpMaxLine
'Chg Start 2011/04/26 SMPK Miyata
'    SpCtrlBlockEnabled spdVa, 1, 1, lngSpMaxLine, 5, CTRL_DISABLE
    If tt = SxlTop039 Or tt = SxlTail039 Then
        SpCtrlBlockEnabled spdVa, 1, 1, spdVa.MaxCols, lngSpMaxLine, CTRL_DISABLE
    Else
        SpCtrlBlockEnabled spdVa, 1, 1, spdVa.MaxCols, lngSpMaxLine, CTRL_DISABLE_SKY
    End If
'Chg End   2011/04/26 SMPK Miyata

    i = 1
    Do While udt_rslt(tt, i - 1).OKNG <> ""
        With udt_rslt(tt, i - 1)
            spdVa.row = i
            For j = 1 To 12
                spdVa.col = j
                Select Case j
                    Case 1
                        '位置
                        spdVa.Value = DBData2DispData(CVar(.pos), "0")
                    Case 2
                        '内容
                        If left(.NAIYO, 3) = "BMD" Then
                            spdVa.Value = .NAIYO & "(×E4)"
                        Else
                            spdVa.Value = .NAIYO
                        End If
                    Case 3
                        '情報１
                        spdVa.Value = .INFO1
                    Case 4
                        '情報２
                        spdVa.Value = .INFO2
                    Case 5
                        '情報３
                        spdVa.Value = .INFO3
                    Case 6
                        '情報４
                        spdVa.Value = .INFO4
                    Case 7
                        '情報５
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO5
                    Case 8
                        '情報７
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO6
                    Case 9
                        '情報８
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO7
                    Case 10
                        '情報８
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO8
                    Case 11
                        '判定
                        'ＷＦサンプル処理変更
                        If .OKNG = "NG" Then
                            SpCtrlEnabled spdVa, spdVa.col, spdVa.row, CTRL_DISABLE_WARNING
                            intEnCmd = 1
'Add Start 2011/03/23 SMPK Miyata
'                        Else
'                            SpCtrlEnabled spdVa, spdVa.col, spdVa.row, CTRL_DISABLE
'Add End   2011/03/23 SMPK Miyata
                        End If
                        spdVa.Value = .OKNG
                    Case 12
                        '位置
                        spdVa.Value = CStr(DBData2DispData(.SMPLID, "0"))
                End Select
            Next j
        End With
        i = i + 1
    Loop

    'ソート処理
    If i <> 1 Then
        With spdVa
            .MaxRows = i - 1                      '　品番（行数）
            .row = 1                            ' セルブロックを設定
            .col = 1
            .row2 = i - 1
            .col2 = 12
            .SortBy = SS_SORT_BY_ROW

            .SortKey(1) = 11                    ' 第１ソートキーを設定

            ' 昇順に並べ替え
            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            .Action = SS_ACTION_SORT
        End With
    End If
End Sub

'*******************************************************************************************
'*    関数名        : sub_PutRslt_EP
'*
'*    処理概要      : 1.WF仕様⇔エピ仕様の表示切替
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　udt_rslt 　　 ,I  ,typ_ALLRSLT     ,実績情報構造体
'*　　　　　　　　　　tt       　　 ,I  ,Integer         ,TopTail判定用
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_PutRslt_EP(udt_rslt() As typ_ALLRSLT_EX, tt As Integer)
    Dim i, j            As Integer
    Dim spdVa           As vaSpread
    Dim lngSpMaxLine    As Long

    ''最大行数取得
    lngSpMaxLine = 0
    Do While udt_rslt(tt, lngSpMaxLine).OKNG <> ""
        lngSpMaxLine = lngSpMaxLine + 1
    Loop

    If tt = SxlTop039 Then
        Set spdVa = f_cmbc039_2.spdKensaTop
'Chg Start 2011/03/23 SMPK Miyata
'    Else
'        Set spdVa = f_cmbc039_2.spdKensaTail
    ElseIf tt = SxlTail039 Then
        Set spdVa = f_cmbc039_2.spdKensaTail
    Else
        Set spdVa = f_cmbc039_2.spdKensaMid

'Chg End   2011/03/23 SMPK Miyata
    End If

    SpCtrlInit spdVa, lngSpMaxLine
    
'Chg Start 2011/08/25 Y.Hitomi
'        SpCtrlBlockEnabled spdVa, 1, 1, lngSpMaxLine, 5, CTRL_DISABLE
    If tt = SxlTop039 Or tt = SxlTail039 Then
        SpCtrlBlockEnabled spdVa, 1, 1, lngSpMaxLine, 5, CTRL_DISABLE
    Else
        SpCtrlBlockEnabled spdVa, 1, 1, lngSpMaxLine, 5, CTRL_DISABLE_SKY
    End If
'Chg End  2011/08/25 Y.Hitomi

    i = 1
    Do While udt_rslt(tt, i - 1).OKNG <> ""
        With udt_rslt(tt, i - 1)
            spdVa.row = i
            For j = 1 To 12
                spdVa.col = j
                Select Case j
                Case 1
                    '位置
                    spdVa.Value = DBData2DispData(CVar(.pos), "0")
                Case 2
                    '内容
                    spdVa.Value = .NAIYO
                Case 3
                    '情報１
                    spdVa.Value = .INFO1
                Case 4
                    '情報２
                    spdVa.Value = .INFO2
                Case 5
                    '情報３
                    spdVa.Value = .INFO3
                Case 6
                    '情報４
                    spdVa.Value = .INFO4
                Case 7
                    '情報５
                    spdVa.Value = .INFO5
                Case 8
                    '情報７
                    spdVa.Value = .INFO6
                Case 9
                    '情報８
                    spdVa.Value = .INFO7
                Case 10
                    '情報８
                    spdVa.Value = .INFO8
                Case 11
                        If .OKNG = "NG" Then
                            SpCtrlEnabled spdVa, spdVa.col, spdVa.row, CTRL_DISABLE_WARNING
                            intEnCmd = 1
                        End If
                        spdVa.Value = .OKNG
                Case 12
                    '位置
                    spdVa.Value = CStr(DBData2DispData(.SMPLID, "0"))
                End Select
            Next j
        End With
        i = i + 1
    Loop

    'ソート処理
    If i <> 1 Then
        With spdVa
            .MaxRows = i - 1                    '　品番（行数）
            .row = 1                            ' セルブロックを設定
            .col = 1
            .row2 = i - 1
            .col2 = 12
            .SortBy = SS_SORT_BY_ROW
            .SortKey(1) = 11                    ' 第１ソートキーを設定
            
            ' 昇順に並べ替え
            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            .Action = SS_ACTION_SORT
        End With
    End If
End Sub

''*******************************************************************************
''*    関数名        : RegWfSogoRsltOK
''*
''*    処理概要      : 1.総合判定実績挿入
''*                    2.WF_GD実績(TBCMJ015)更新処理
''*                    3.SXL管理更新
''*                    4.WFサンプル管理更新
''*
''*    パラメータ    : 変数名        ,IO ,型           ,説明
''*
''*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
''*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
''*
''*******************************************************************************
'Private Function RegWfSogoRsltOK() As FUNCTION_RETURN
'    Dim udt_soz         As typ_TBCMW005                             ' WF総合判定実績
'    Dim udt_sxl         As type_DBDRV_scmzc_fcmlc001c_UpdSXL1       ' SXL管理
'    Dim udt_WFSmp(2)    As type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
'    Dim i               As Long
'    Dim intCnt          As Integer
'
'    'WF総合判定実績
'    With udt_soz
'        .CRYNUM = typ_CType.typ_Param.CRYNUM                                ' 結晶番号
'        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' インゴット位置
'        .CRYLEN = typ_CType.typ_Param.LENGTH                                ' 長さ
'        .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI                                 ' 管理工程コード
'        .PROCCODE = PROCD_WFC_SOUGOUHANTEI                                  ' 工程コード
'        .SXLID = NtoS(typ_CType.typ_Param.SXLID)                                  ' SXLID
'        .CODE = "0"                                                         ' 区分コード
'        .TSTAFFID = typ_CType.strStaffID                                    ' 登録社員ID
'    End With
'
'    'WF総合判定実績挿入
'    If DBDRV_scmzc_fcmlc001c_InsWfSougou(udt_soz) <> FUNCTION_RETURN_SUCCESS Then
'        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "W005")
'        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    '' WF_GD実績(TBCMJ015)更新処理
'    If UBound(typ_J015_WFGDUpd) > 0 Then
'        'ﾃﾞｰﾀ数分UPDATE
'        For intCnt = 1 To UBound(typ_J015_WFGDUpd)
'            If DBDRV_scmzc_fcmlc001c_UpdGDdata(typ_J015_WFGDUpd(intCnt), typ_CType.strStaffID) _
'                                        <> FUNCTION_RETURN_SUCCESS Then
'                f_cmbc039_2.lblMsg.Caption = GetMsgStr("EAPLY") & "J015"
'                RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'                Exit Function
'            End If
'        Next
'    End If
'
'    'SXL管理
'    With udt_sxl
'        .CRYNUM = NtoS(typ_AType.typ_Param.CRYNUMCA)                        ' 結晶番号
'        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' 結晶内開始位置
'        .NOWPROC = PROCD_SXL_KAKUTEI                                        ' 現在工程
'        .LASTPASS = PROCD_WFC_SOUGOUHANTEI                                  ' 最終通過工程
'    End With
'
'    'SXL管理更新
'    If DBDRV_scmzc_fcmlc001c_UpdSXL1(udt_sxl) <> FUNCTION_RETURN_SUCCESS Then
'        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E042")
'        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    'WFサンプル管理が存在する場合は確定区分コードに1を立てる
'    'エピ先行評価追加対応
'    If (UBound(typ_CType.typ_y013top) <> 0 Or UBound(typ_CType_EP.typ_y022top) <> 0) _
'        And (UBound(typ_CType.typ_y013tail) <> 0 Or UBound(typ_CType_EP.typ_y022tail) <> 0) Then
'
'        'WFサンプル管理
'        udt_WFSmp(1).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' 結晶番号
'        udt_WFSmp(1).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTop039).INPOSCW    ' 結晶内開始位置
'        udt_WFSmp(1).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTop039).SMPKBNCW     ' サンプル区分
'        udt_WFSmp(2).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' 結晶番号
'        udt_WFSmp(2).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTail039).INPOSCW   ' 結晶内開始位置
'        udt_WFSmp(2).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTail039).SMPKBNCW    ' サンプル区分
'
'        'WFサンプル管理更新
'        If DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(udt_WFSmp) <> FUNCTION_RETURN_SUCCESS Then
'            f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E044")
'            RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'    End If
'End Function

'*******************************************************************************************
'*    関数名        : sub_cmbc061_2_ChangeHinSpec
'*
'*    処理概要      : 1.WF仕様⇔エピ仕様の表示切替
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　intCategory　 ,I  ,Integer         ,表示カテゴリ(0:WF仕様,1:エピ仕様)
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_cmbc061_2_ChangeHinSpec(Optional intCategory As Integer = 0)
    Dim i       As Long

    On Error Resume Next

    With f_cmbc039_2
        .spdHinbanCen.ReDraw = False
        .spdHinbanCenEpi.ReDraw = False
        .spdHinbanTail.ReDraw = False
        .spdHinbanTail2.ReDraw = False
        .spdHinbanHed.ReDraw = False                '08/03/12 ooba
        .spdHinbanTop.ReDraw = False

        Select Case intCategory
            Case 0          ' WF仕様データの表示
                'WFデータのスプレッドを表示
                .spdHinbanCen.Visible = True
                .spdHinbanTail.Visible = True
                .spdHinbanTail2.Visible = True
                .spdHinbanHed.Visible = True        '08/03/12 ooba
                .spdHinbanCenEpi.Visible = False
    
                'AN温度
                f_cmbc039_2.spdHinbanTop.Value = typ_CType.typ_si.HWFANTNP
            Case 1          ' エピ仕様データの表示
                'エピデータのスプレッドを表示
                .spdHinbanCen.Visible = False
                .spdHinbanTail.Visible = False
                .spdHinbanTail2.Visible = False
                .spdHinbanHed.Visible = False       '08/03/12 ooba
                .spdHinbanCenEpi.Visible = True
    
                'AN温度(製品仕様エピデータ1)
                f_cmbc039_2.spdHinbanTop.Value = typ_CType.typ_si.HEPANTNP
            Case Else
        End Select

        .spdHinbanCen.ReDraw = True
        .spdHinbanCenEpi.ReDraw = True
        .spdHinbanTail.ReDraw = True
        .spdHinbanTail2.ReDraw = True
        .spdHinbanHed.ReDraw = True                 '08/03/12 ooba
        .spdHinbanTop.ReDraw = True
    End With
End Sub
'*******************************************************************************************
'*    関数名        : sub_SampleMidlePosBtnSet
'*
'*    処理概要      : 1.中間抜試サンプル位置ボタン設定
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　:　　           ,   ,                ,
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_SampleMidlePosBtnSet()
    Dim i       As Long
    Dim k       As Long
    Dim blnOK   As Boolean

    'ボタン数ループ
    For i = optPosSelMid.LBound To optPosSelMid.UBound

        '中間抜試があるかサンプル管理有無で判断
        If SxlMidl + i <= UBound(typ_CType.typ_Param.WFSMP) Then
            '中間抜試が有りの場合
            
            'ボタン設定
            With optPosSelMid(i)
                .Enabled = True
                .Caption = typ_CType.typ_Param.WFSMP(SxlMidl + i).INPOSCW
                
                
                '検査NG項目があるか検索
                blnOK = True
                ' 比抵抗判定
                'Cng Start 2011/07/06 Y.Hitomi
                If typ_CType.JudgRes(SxlMidl + i) = False And txtSXLMid.text <> "" Then
                'If typ_CType.JudgRes(SxlMidl + i) = False Then
                'Cng End   2011/07/06 Y.Hitomi
                    blnOK = False
                End If
                k = 0
                Do While typ_CType.typ_rslt(SxlMidl + i, k).OKNG <> ""
                       
                    If typ_CType.typ_rslt(SxlMidl + i, k).OKNG <> "OK" Then
                        blnOK = False
                        Exit Do
                    End If
                    k = k + 1
                Loop
                
                '表示位置の色設定　検査OK：黒　検査NG：赤
                If blnOK = True Then
                    .ForeColor = vbBlack
                Else
                    .ForeColor = vbRed
                End If
                If i = optPosSelMid.LBound Then .Value = True
                
            End With
        Else
            '中間抜試が無しの場合
            
            'ボタン設定
            With optPosSelMid(i)
                .Enabled = False
                .Caption = ""
                .ForeColor = vbBlack
            End With
        
        End If
    Next

End Sub

'>>>>> add 2011/07/13 Marushita
'画面キャプチャ処理
Public Function saveCapture_BMP(ByRef frm As Form, ByRef picData As PictureBox) As Boolean
    
    Dim lRetVal As Long
    Dim lDC As Long
    
On Error GoTo Err:
    
    '手前に表示
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    DoEvents
    
    ' ウィンドウをアクティブにする
    Call SetForceForegroundWindow(frm.hwnd)
    DoEvents
    
    'ハンドルからデバイスコンテキストを取得
    lDC = GetDC(frm.hwnd)
    
    picData.AutoRedraw = True
    picData.Width = frm.ScaleWidth + 10
    picData.Height = frm.ScaleHeight + 29
    
    'lRetVal = BitBlt(picData.hdc, 0, 0, picData.Width, picData.Height, lDC, -3, -22, SRCCOPY)
    lRetVal = StretchBlt(picData.hdc, 0, 0, picData.Width * CLng(SCALEPER) / 100, picData.Height * CLng(SCALEPER) / 100, lDC, 0, 0, picData.Width, picData.Height, SRCCOPY)
    
    DoEvents
    'クリップボード内にビットマップ形式のデータがあるか調べる
    If lRetVal <> 0 Then
        'ファイル名を生成
        SavePicture pic_Png.Image, App.Path & CAP_FNAME
    Else
        '失敗
        Call MsgBox("失敗")
    End If
    
    'DC開放
    Call ReleaseDC(frm.hwnd, lDC)
    
    '手前に表示を解除
    SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    
    saveCapture_BMP = True
    Exit Function
    
Err:
    Call MsgOut(0, "画面キャプチャ保存に失敗しました" & vbCrLf _
                 & Err.Number & ":" & Err.Description, ERR_DISP)
    
    '手前に表示を解除
    SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    
    saveCapture_BMP = False

End Function
'<<<<< add 2011/07/13 Marushita
