VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '固定(実線)
   Caption         =   "f_cmbc039_3(CW750) - 300mm結晶操業システム"
   ClientHeight    =   10875
   ClientLeft      =   0
   ClientTop       =   750
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   9
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   1018
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdDisp 
      Caption         =   "WFC総合判定参照表示"
      Height          =   375
      Left            =   600
      TabIndex        =   63
      Top             =   9000
      Width           =   2175
   End
   Begin VB.PictureBox pic_Png 
      Height          =   135
      Left            =   2250
      ScaleHeight     =   75
      ScaleWidth      =   315
      TabIndex        =   59
      Top             =   8535
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chk_Png 
      Caption         =   "PNG保存"
      Height          =   375
      Left            =   690
      TabIndex        =   58
      Top             =   8535
      Width           =   1335
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
      Left            =   14280
      TabIndex        =   49
      Tag             =   "WF"
      Top             =   2160
      Width           =   855
   End
   Begin FPSpread.vaSpread sprWarp 
      Height          =   1140
      Left            =   3120
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   7935
      Width           =   4095
      _Version        =   196608
      _ExtentX        =   7218
      _ExtentY        =   2011
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
      MaxCols         =   6
      MaxRows         =   4
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_3.frx":0000
      UserResize      =   0
      VisibleCols     =   5
      VisibleRows     =   1
   End
   Begin FPSpread.vaSpread vaSpread5 
      Height          =   240
      Left            =   360
      TabIndex        =   45
      Top             =   2520
      Width           =   1560
      _Version        =   196608
      _ExtentX        =   2752
      _ExtentY        =   423
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      MaxRows         =   0
      OperationMode   =   1
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":0739
      UserResize      =   0
      VisibleCols     =   1
      ClipboardOptions=   0
   End
   Begin VB.PictureBox pic_check 
      BorderStyle     =   0  'なし
      Height          =   210
      Index           =   0
      Left            =   10380
      Picture         =   "f_cmbc039_3.frx":09A9
      ScaleHeight     =   210
      ScaleWidth      =   195
      TabIndex        =   37
      Top             =   8610
      Width           =   195
   End
   Begin VB.PictureBox pic_check 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'なし
      FillColor       =   &H00C0C0C0&
      Height          =   285
      Index           =   1
      Left            =   10335
      Picture         =   "f_cmbc039_3.frx":0D1B
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   8775
      Width           =   315
   End
   Begin VB.TextBox txtBotRsltR 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   8175
      Width           =   1095
   End
   Begin VB.TextBox txtTopRsltR 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7935
      Width           =   1095
   End
   Begin VB.TextBox txtCryP 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   8745
      Width           =   855
   End
   Begin VB.TextBox txtBlkP 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   8475
      Width           =   855
   End
   Begin VB.TextBox txtBlkID 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   8205
      Width           =   855
   End
   Begin VB.TextBox txtTarget 
      Alignment       =   1  '右揃え
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   8310
      TabIndex        =   17
      Top             =   7935
      Width           =   855
   End
   Begin VB.TextBox txtKSXLID 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   1170
      Width           =   1260
   End
   Begin VB.TextBox txtCryNum 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   1500
      Width           =   1200
   End
   Begin VB.TextBox txtJfName 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   2670
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1170
      Width           =   1500
   End
   Begin VB.TextBox txtStaffID 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   825
   End
   Begin VB.Frame fraF 
      Height          =   1095
      Left            =   30
      TabIndex        =   1
      Top             =   9540
      Width           =   15195
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11]　　前画面"
         Height          =   735
         Index           =   11
         Left            =   12680
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]　　ｻﾝﾌﾟﾙ"
         Height          =   735
         Index           =   10
         Left            =   11448
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F８]　　削除"
         Height          =   735
         Index           =   8
         Left            =   8984
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F７]　　挿入"
         Height          =   735
         Index           =   7
         Left            =   7752
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F６]　　WFﾏｯﾌﾟ"
         Height          =   735
         Index           =   6
         Left            =   6520
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F５]　　振替"
         Enabled         =   0   'False
         Height          =   735
         Index           =   5
         Left            =   5288
         Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F４]　　候補"
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F９]　　結晶ｲﾒｰｼﾞ表示"
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F１]　　＊＊＊"
         Height          =   735
         Index           =   1
         Left            =   360
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F２]　　ｻﾌﾞﾒﾆｭｰ"
         Height          =   735
         Index           =   2
         Left            =   1592
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F３]　　ｷｬﾝｾﾙ"
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]　　実行"
         Height          =   735
         Index           =   12
         Left            =   13905
         TabIndex        =   35
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
      Top             =   0
      Width           =   15225
      Begin VB.Label lblvers 
         Height          =   195
         Left            =   13740
         TabIndex        =   44
         Top             =   480
         Width           =   1440
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
         TabIndex        =   4
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
         Left            =   1800
         TabIndex        =   3
         Top             =   210
         Width           =   8535
      End
      Begin VB.Label lblTitle 
         Caption         =   "再抜試指示"
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
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4575
      End
   End
   Begin FPSpread.vaSpread sprSpec 
      Height          =   885
      Left            =   75
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2520
      Width           =   15150
      _Version        =   196608
      _ExtentX        =   26723
      _ExtentY        =   1561
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   30
      MaxCols         =   33
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      ScrollBars      =   2
      SpreadDesigner  =   "f_cmbc039_3.frx":121D
      UserResize      =   0
      VisibleCols     =   32
      VisibleRows     =   5
   End
   Begin FPSpread.vaSpread vaSpread8 
      Height          =   240
      Left            =   75
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   3990
      Width           =   3870
      _Version        =   196608
      _ExtentX        =   6826
      _ExtentY        =   423
      _StockProps     =   64
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      MaxRows         =   0
      RetainSelBlock  =   0   'False
      ScrollBars      =   0
      ShadowColor     =   12632256
      ShadowDark      =   10070188
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":3267
      UserResize      =   0
      VisibleCols     =   1
   End
   Begin FPSpread.vaSpread vaSpread7 
      Height          =   480
      Left            =   75
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   4215
      Width           =   3870
      _Version        =   196608
      _ExtentX        =   6826
      _ExtentY        =   847
      _StockProps     =   64
      AllowCellOverflow=   -1  'True
      ColsFrozen      =   3
      DisplayRowHeaders=   0   'False
      MaxCols         =   7
      MaxRows         =   1
      ScrollBars      =   0
      ShadowColor     =   12632256
      ShadowDark      =   10070188
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":34AB
      UserResize      =   0
      VisibleCols     =   3
   End
   Begin FPSpread.vaSpread vaSpread6 
      Height          =   480
      Left            =   5730
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   4215
      Width           =   8640
      _Version        =   196608
      _ExtentX        =   15240
      _ExtentY        =   847
      _StockProps     =   64
      ColsFrozen      =   15
      DisplayRowHeaders=   0   'False
      MaxCols         =   25
      MaxRows         =   1
      ScrollBars      =   0
      ShadowColor     =   12632256
      ShadowDark      =   10070188
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":382F
      UserResize      =   0
      VisibleCols     =   15
   End
   Begin FPSpread.vaSpread vaSpread9 
      Height          =   255
      Left            =   5730
      TabIndex        =   54
      Top             =   3990
      Width           =   8640
      _Version        =   196608
      _ExtentX        =   15240
      _ExtentY        =   450
      _StockProps     =   64
      ColsFrozen      =   1
      DisplayRowHeaders=   0   'False
      MaxCols         =   2
      ScrollBars      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":433D
      UserResize      =   0
      VisibleCols     =   1
      VisibleRows     =   500
   End
   Begin FPSpread.vaSpread sprExamine 
      Height          =   3510
      Left            =   75
      TabIndex        =   55
      Top             =   3990
      Width           =   15090
      _Version        =   196608
      _ExtentX        =   26617
      _ExtentY        =   6191
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   12
      DisplayRowHeaders=   0   'False
      MaxCols         =   47
      MaxRows         =   4
      ScrollBars      =   2
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_3.frx":5B4B
      UserResize      =   0
      VisibleCols     =   12
      VisibleRows     =   4
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
      Height          =   255
      Left            =   10755
      TabIndex        =   62
      Top             =   1650
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblNukishi 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   21.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   8235
      TabIndex        =   61
      Top             =   1335
      Width           =   2295
   End
   Begin VB.Label lblKanren 
      Caption         =   "関連ブロック"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   60
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label22 
      Caption         =   "向先"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5130
      TabIndex        =   57
      Top             =   1530
      Width           =   1455
   End
   Begin VB.Label lblMukesaki 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      Caption         =   "4棟"
      Height          =   255
      Left            =   6600
      TabIndex        =   56
      Top             =   1515
      Width           =   390
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      Caption         =   "評価結果受信済"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10170
      TabIndex        =   48
      Top             =   7890
      Width           =   1500
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ｻﾝﾌﾟﾙ無(結晶実績)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12705
      TabIndex        =   46
      Top             =   8745
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H0000FFFF&
      Caption         =   "ｻﾝﾌﾟﾙ無(反映)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12705
      TabIndex        =   43
      Top             =   8445
      Width           =   1575
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "ｻﾝﾌﾟﾙ有"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   12705
      TabIndex        =   42
      Top             =   8145
      Width           =   1575
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ｻﾝﾌﾟﾙ無"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12705
      TabIndex        =   41
      Top             =   7845
      Width           =   1575
   End
   Begin VB.Label lbl_check 
      AutoSize        =   -1  'True
      Caption         =   "：1SXL"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   1
      Left            =   10560
      TabIndex        =   38
      Top             =   8655
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lbl_check 
      AutoSize        =   -1  'True
      Caption         =   "：SXL分割"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   2
      Left            =   10560
      TabIndex        =   40
      Top             =   8835
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl_check 
      AutoSize        =   -1  'True
      Caption         =   "ﾁｪｯｸﾎﾞｯｸｽ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   10380
      TabIndex        =   39
      Top             =   8340
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label Label14 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "結晶P"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7500
      TabIndex        =   22
      Top             =   8745
      Width           =   810
   End
   Begin VB.Label Label13 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "ﾌﾞﾛｯｸP"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7500
      TabIndex        =   20
      Top             =   8475
      Width           =   810
   End
   Begin VB.Label Label7 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "ﾌﾞﾛｯｸID"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7500
      TabIndex        =   18
      Top             =   8205
      Width           =   810
   End
   Begin VB.Label Label6 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "ねらいρ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   7500
      TabIndex        =   16
      Top             =   7935
      Width           =   810
   End
   Begin VB.Label Label5 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "B側実績ρ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   690
      TabIndex        =   14
      Top             =   8175
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '中央揃え
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H80000005&
      BackStyle       =   0  '透明
      BorderStyle     =   1  '実線
      Caption         =   "T側実績ρ"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   690
      TabIndex        =   12
      Top             =   7935
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "仮ＳＸＬ－ＩＤ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5130
      TabIndex        =   8
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "結晶番号"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   435
      TabIndex        =   10
      Top             =   1530
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "担当者コード"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   435
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "f_cmbc039_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'                                       2001/06/14
'===============================================================================
' 再抜試指示画面
' 概要    :
'===============================================================================

Private orgXl As c_cmzcXl                         ' 読み込み時点の結晶情報
Private tblHinNum() As tFullHinban                ' 品番テーブル
Private InMaxRow As Integer                       ' 最大行保持
Private bSampFlag As Boolean                      ' サンプル取得フラグ
Private CutCntFlg As Integer                      ' サンプルのないSXL
Private orgSXL As c_cmzcSxls                      ' 初期状態のSXL構成
Private giFKeyFlg As Integer                      'サンプル・実行ボタンフラグ
Private bJituChkFlg As Boolean                    ' 実測ﾃﾞｰﾀﾁｪｯｸﾌﾗｸﾞ
Private iBetuRow()  As Integer                    ' 共有(別)ｻﾝﾌﾟﾙ行
Private CpyCrySmpl  As typ_CpyJisseki             ' 結晶実績引継ぎﾃﾞｰﾀ

' 振替チェック追加による修正
Private Type tbl_MotoHin
    MOTOICHIS As Integer                          ' 振替元開始位置
    MOTOICHIE As Integer                          ' 振替元終了位置
    MOTOHIN As tFullHinban                        ' 振替元品番
End Type
Private MotoHinban() As tbl_MotoHin               ' 振替元品番データ

Private Type tbl_FuriNaiyou
    FURIUMU As Integer                            ' 振替有無(0:無し、1:有り)
    ICHI    As Integer                            ' 位置
    MOTOHIN As tFullHinban                        ' 振替元品番
    SAKIHIN As tFullHinban                        ' 振替先品番
    TREPID As String                              ' 代表サンプルID(TOP)
    BREPID As String                              ' 代表サンプルID(BOT)
End Type
Private FurikaeNaiyou() As tbl_FuriNaiyou         ' 振替内容設定データ
Private FurikaeNaiyouWK() As tbl_FuriNaiyou       ' 振替内容設定データ(Work用)

Private Type tbl_TokusaiBan
    ICHI     As Integer                           ' 位置
    MOTOHIN  As tFullHinban                       ' 振替元品番
    SAKIHIN  As tFullHinban                       ' 振替先品番
    BANGOU   As String                            ' 特採番号
    RIYUU    As String                            ' 特採理由
    ERRRIYUU As String                            ' エラー理由
    TREPID As String                              '代表サンプルID(TOP)
    BREPID As String                              '代表サンプルID(BOT)
End Type

Private TokusaiBangou() As tbl_TokusaiBan         ' 特採番号データ
Private TokuCnt As Integer                        ' 特採番号データカウンタ
Private TokuCntWK As Integer                      ' 特採番号データカウンタ(Work用)

Private FurikaeRireki() As typ_XSDCE_Update       ' 振替履歴データ

Private bTokuKengenFlag As Boolean                ' 特採権限フラグ
Private tblKns() As typ_XSDCW                     ' 検査項目をとっておくための構造体
Private tKensa() As typ_XSDCW                     ' 検査項目取得用
Private tWafk() As typeSprWFmap                   ' ウエハーセンター入庫情報のデータをとっておく

Private sComment As String                        ' コメント   07/10/05 miyatake 承認機能追加
Private bSampleBtn        As Boolean              ' サンプルボタン状態 2010/12/08 Marushita

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
    CmdChangeWF_EP.Enabled = False

    If CmdChangeWF_EP.Tag = "WF" Then
        Call sub_cmbc039_3_ChangeHinSpec(1)
        CmdChangeWF_EP.Tag = "EP"
        CmdChangeWF_EP.Caption = "エピ >>"

    Else
        Call sub_cmbc039_3_ChangeHinSpec(0)
        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "ＷＦ >>"
    End If

    CmdChangeWF_EP.Enabled = True
End Sub

'>>>>> 2011/07/14 Marushita
'WFC総合判定画面表示ボタン押下時にキャプチャ画面を表示する
Private Sub cmdDisp_Click()
    '画面表示
    f_hanteiS.picDisp.Picture = LoadPicture(App.Path & CAP_FNAME)
    Call f_hanteiS.Show
End Sub
'<<<<< 2011/07/14 Marushita

'*******************************************************************************
'*    関数名        : cmdF_Click
'*
'*    処理概要      : 1.ファンクションボタンがクリックされたら、各処理に分岐する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    intIndex    ,I  ,Integer　,コントロール配列の添字
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub cmdF_Click(intIndex As Integer)
    Dim sErrMsg     As String
    Dim sBlkId      As String
    Dim sSXLID      As String
    Dim blnsflg     As Boolean

    lblMsg.Caption = ""
    Tokusai = ""

    '' 処理分岐
    Select Case intIndex
        Case 2          '' Ｆ２キー（サブメニュー）
            '>>>>> 2011/07/14 Marushita
            '総合判定判定表示中は閉じる
            If f_hanteiS.Visible = True Then
                Unload f_hanteiS
            End If
            '<<<<< 2011/07/14 Marushita
            GotoSubMenu
        Case 3          '' Ｆ３キー（キャンセル）
            ' ボタン連打を防ぐ為
            cmdF(3).Enabled = False

            Call sub_DispClear
            lblMsg.Caption = GetMsgStr("PWAIT")
            Call sub_LoadAndDisp

            '2003/04/21 Fキーフラグ追加　0:初期状態 1:サンプルボタン 2:実行ボタン
            giFKeyFlg = 0 '初期状態
            ReDim FurikaeNaiyouWK(0)
            TokuCntWK = TokuCnt
            ReDim Preserve TokusaiBangou(TokuCntWK)

            ' ボタン連打を防ぐ為
            cmdF(3).Enabled = True
        Case 4
            '候補ボタン
            Call sub_FurikaeKouho
        Case 5
            '特採ボタン
            Call sub_TokusaiInput
            Tokusai = "1"
            lblMsg.Caption = TBN_Msg
        Case 6
'Add Start 2011/03/11 SMPK Miyata
            'WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙからﾃﾞｰﾀを取得
            If SelWFmap(vbNullString, SelectSxlID039, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_2.lblMsg.Caption = sErrMsg
                Exit Sub
            End If
            
            'ｽﾌﾟﾚｯﾄﾞにﾃﾞｰﾀを表示
            If SetWFmapData = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_4.lblMsg.Caption = sErrMsg
                f_cmbc039_4.sprWfmapView.MaxRows = 0
                Exit Sub
            End If

            'サンプルボタン押下時は画面の状態を状態と色で反映する
            If giFKeyFlg = 1 Then
                fnc_DispChangeDataWfmap '状態をMapに表示
            End If

            f_cmbc039_4.txtSxlId.text = SelectSxlID039
            f_cmbc039_4.Show
'Add End   2011/03/11 SMPK Miyata

'Del Start 2011/03/11 SMPK Miyata
'            cmbSprChg.ListIndex = 0
'            With sprWfmapView
'                .col = 30
'                .ColHidden = True
'            End With
'
'            If bWfmapView = False Then
'                '条件を取得
'                sSXLID = Trim(txtKSXLID.text)
'                lblSxlId.Caption = sSXLID
'                sBlkId = vbNullString
'
'                'WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙからﾃﾞｰﾀを取得
'                If sub_SelWFmap(sBlkId, sSXLID, sErrMsg) = FUNCTION_RETURN_FAILURE Then
'                    lblMsg.Caption = sErrMsg
'                    sprExamine.MaxRows = 0
'                    Exit Sub
'                End If
'                sprWfmapView.ReDraw = False
'                If fnc_SetWFmapData = FUNCTION_RETURN_FAILURE Then
'                    lblMsg.Caption = sErrMsg
'                    sprExamine.MaxRows = 0
'                    Exit Sub
'                End If
'
'                sprWfmapView.ReDraw = True
'                Me.top = 0
'                bWfmapView = True
'                Me.Height = 11280
'            ElseIf bWfmapView = True Then
'                Me.top = 1500
'                bWfmapView = False
'                Me.Height = 8580
'                Exit Sub
'            End If
'
'            'サンプルボタン押下時は画面の状態を状態と色で反映する
'            If giFKeyFlg = 1 Then
'                fnc_DispChangeDataWfmap '状態をMapに表示
'            End If
'
'            'ｽﾌﾟﾚｯﾄﾞﾃﾞｰﾀソート
'            With sprWfmapView
'                .BlockMode = True
'                .col = 1
'                .col2 = .MaxCols
'                .row = 1
'                .row2 = .MaxRows
'                .SortBy = SortByRow
'                .SortKey(1) = 8
'                .SortKeyOrder(1) = SortKeyOrderAscending
'                .Action = ActionSort
'                .BlockMode = False
'            End With
'Del End   2011/03/11 SMPK Miyata

        Case 7          '' Ｆ７キー（行挿入）
            sprExamine.SetFocus
            Call sprExamine_KeyDown(vbKeyF7, 0)
        Case 8          '' Ｆ８キー（行削除）
            sprExamine.SetFocus
            Call sprExamine_KeyDown(vbKeyF8, 0)
        Case 9          '' Ｆ９キー（結晶イメージ表示）
            If sub_DispSample(blnsflg) = True Then
                If blnsflg = True Then 'データチェックでエラーになった状態ではフラグを立てないように修正
                    giFKeyFlg = 1
                End If
                Call sub_DrawImage
            End If

            ' 振替チェック追加による修正
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
'----- サンプルボタン２回押下不可対応　2010/12/08 Marushita
            If bSampleBtn = True And giFKeyFlg = 1 Then
                bSampleBtn = False
                cmdF(7).Enabled = False
                cmdF(8).Enabled = False
                cmdF(10).Enabled = False
            End If
        Case 10         '' Ｆ10キー（サンプル）
            If sub_DispSample(blnsflg) = False Then
                Exit Sub
            End If

            If blnsflg = True Then 'データチェックでエラーになった状態ではフラグを立てないように修正
                giFKeyFlg = 1
            End If

            '2003/04/21 Fキーフラグ追加　0:初期状態 1:サンプルボタン 2:実行ボタン
            If fnc_DispHinSpec(1) = False Then
                Exit Sub
            End If

            ' 振替チェック追加による修正
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
'----- サンプルボタン２回押下不可対応　2010/12/08 Marushita
            If bSampleBtn = True And giFKeyFlg = 1 Then
                bSampleBtn = False
                cmdF(7).Enabled = False
                cmdF(8).Enabled = False
                cmdF(10).Enabled = False
            End If
        Case 11         '' Ｆ11キー（前画面）
            '>>>>> 2011/07/14 Marushita
            '総合判定参照表示中は閉じる
            If f_hanteiS.Visible = True Then
                Unload f_hanteiS
            End If
            '<<<<< 2011/07/14 Marushita
            CloseFormProc f_cmbc039_2, f_cmbc039_3
        Case 12         '' Ｆ12キー（実行）
            '' 担当者IDのチェック
            If f_cmzcChkUser.CanExec(Me.Name, txtStaffID.text) = False Then
                lblMsg.Caption = GetMsgStr("EUSR0")
                Exit Sub
            End If

            '≪SaveData≫の前に移動
            giFKeyFlg = 2 '実行ボタン

            Call sub_SaveData

            ' 振替チェック追加による修正
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
    End Select
End Sub

'*******************************************************************************
'*    関数名        : cmdF_GotFocus
'*
'*    処理概要      : 1.コントロールがフォーカスを取得した時に実行する処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    Index       ,I  ,Integer　,コントロール配列の添字
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub cmdF_GotFocus(Index As Integer)
    '' 行削除／行挿入ファンクションキーでなければ
    If Index <> 7 And Index <> 8 Then
        '' 行削除／行挿入ファンクションキーを無効にする
        cmdF(7).Enabled = False
        cmdF(8).Enabled = False
    End If

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
    Dim intIndex As Integer

    '' ファンクションキーが有効なら
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        '' 画面表示メッセージクリア
        lblMsg.Caption = ""

        intIndex = KeyCode - (vbKeyF1 - 1)
        If KeyCode = vbKeyF5 Then
            Call cmdF_Click(intIndex)
        End If
        If cmdF(intIndex).Visible = True And cmdF(intIndex).Enabled = True Then
            '' ファンクションキー押下処理を実行する
            If KeyCode <> vbKeyF7 And KeyCode <> vbKeyF8 Then
                Call cmdF_Click(intIndex)
            End If
        End If
    End If
End Sub

'*******************************************************************************
'*    関数名        : Form_Load
'*
'*    処理概要      : 1.SXL構成の初期状態を保存する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Load()
    Dim sxl As c_cmzcSxl

    Call Pic_Disp(0)

    '' 画面クリア
    Call sub_DispClear

'Del Start 2011/03/11 SMPK Miyata
'    '' フォーム位置セット
'    CenterForm Me
'    Me.Height = 8580
'Del End   2011/03/11 SMPK Miyata
    Me.Show

    ' バージョン情報の表示
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision


    '品番を1列追加したことによる列の変更
    'エピ先行評価追加対応
    sprExamine.ColsFrozen = 35

    CutCntFlg = 0
    sprExamine.col = 5
    sprExamine.TypeComboBoxWidth = 1

    '' データ初期処理
    lblMsg.Caption = GetMsgStr("PWAIT")
    DoEvents
    Call sub_InitData

    '' データのロードと表示
    lblMsg.Caption = GetMsgStr("PWAIT")
    DoEvents
    Call sub_LoadAndDisp

    '' SXL構成の初期状態を保存する
    Set orgSXL = New c_cmzcSxls
    Set sxl = New c_cmzcSxl

    With tblSXL
        sxl.CRYNUM = .CRYNUM
        sxl.INGOTPOS = .INGOTPOS
        sxl.LENGTH = .COUNT
        sxl.hinban = .hinban
        sxl.REVNUM = .REVNUM
        sxl.factory = .factory
        sxl.opecond = .opecond
    End With
    orgSXL.Add sxl
    Set sxl = Nothing

'2003/04/21 Fキーフラグ追加　0:初期状態 1:サンプルボタン 2:実行ボタン
    giFKeyFlg = 0 '初期状態

    ReDim tblWafInd(0)
    ReDim tblNukishi(0) '抜試変更用構造体
    ' 特採権限追加による修正
            '' 担当者IDのチェック
            If f_cmzcChkUser.CanExec("TOKUSAI", txtStaffID.text) Then
                bTokuKengenFlag = True
            Else
                bTokuKengenFlag = False
            End If

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
End Sub

'*******************************************************************************
'*    関数名        : Form_Unload
'*
'*    処理概要      : 1.フォームがアンロードされる時に実行する処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    Cancel      ,I  ,Integer
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Set orgSXL = Nothing

    '' 結晶図ウィンドウのアンロード
    Unload f_cmzc003a
End Sub

'*******************************************************************************
'*    関数名        : sprExamine_Change
'*
'*    処理概要      : 1.品番変更時に向け先を再取得
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    col         ,I  ,Long     ,選択列
'*                    Row         ,I  ,Long     ,選択行
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprExamine_Change(ByVal col As Long, ByVal row As Long)
    Dim blAns           As Boolean
    Dim vTemp           As Variant
    Dim vTemp1          As Variant
    Dim vTemp2          As String
    Dim udtFullHinban   As tFullHinban

    '' 位置と品番が変更の場合
    If (col = 2) Or (col = 4) Then
        '' 特採情報クリア
        Call sub_ClearTokusai
    End If

    If col = 2 Then
        lblMsg.Caption = vbNullString
        blAns = sprExamine.GetText(col, row, vTemp)
        vTemp = Trim(vTemp)
        Select Case ChkString(vTemp, 8, 8)
        Case CHK_NG, CHK_NULL
            If Trim$(vTemp) <> "Z" Then
                lblMsg.Caption = GetMsgStr(EHIN1)
            End If
            Exit Sub
        End Select
        vTemp2 = vTemp
        If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr(EHIN0)
            Exit Sub
        End If
    End If

    Dim sHinban   As String
    Dim sMuke     As String
    Dim sNewMuke  As String
    '' 品番が変更の場合
    If (col = 2) Then
        sprExamine.col = 2
        sprExamine.row = row
        sHinban = sprExamine.text

        sprExamine.col = 2
        sprExamine.row = row + 1

        sprExamine.text = sCmbMukeName

        '' 向先情報取得
        sNewMuke = GetMukesaki(sHinban)
        sprExamine.text = sNewMuke

        If InStr(1, sNewMuke, left(sBaseMukesaki, 1), vbTextCompare) = 0 Then
            lblMsg.Caption = "向先が変更されています。"
            sprExamine.backColor = vbRed
        Else
            lblMsg.Caption = ""
            sprExamine.backColor = vbWhite
        End If
    End If
    'サンプルボタン押下後データが書き換えられた場合、実行できなくするため2003/04/25 okazaki
    bSampFlag = False
End Sub

'*******************************************************************************
'*    関数名        : sprExamine_ComboSelChange
'*
'*    処理概要      : 1.コンボボックスの選択項目が変更された時に実行する処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    col         ,I  ,Long     ,コンボボックスの列番号
'*                    Row         ,I  ,Long     ,コンボボックスの行番号
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprExamine_ComboSelChange(ByVal col As Long, ByVal row As Long)
    With sprExamine
        If col = 2 And row = 1 Then
            .row = 1
            .col = 2
            If Trim$(.text) = "Z" Then
                .col = 9
                .Lock = False
                .backColor = COLOR_OK
            Else
                .col = 9
                .Lock = True
                .TypeComboBoxCurSel = 0
                .backColor = COLOR_DISABLE
            End If
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : sprExamine_GotFocus
'*
'*    処理概要      : 1.コントロールがフォーカスを取得した時に実行する処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprExamine_GotFocus()
    '' 行削除／行挿入ファンクションキーを有効にする
'----- サンプルボタン押下不可対応　2010/12/08 Marushita
    If bSampleBtn = False Then
    Else
        If CutCntFlg = 0 Then
            cmdF(7).Enabled = True
            cmdF(8).Enabled = True
        End If
    End If
    'If CutCntFlg = 0 Then
    '    cmdF(7).Enabled = True
    '    cmdF(8).Enabled = True
    'End If
End Sub

'*******************************************************************************
'*    関数名        : sprExamine_KeyDown
'*
'*    処理概要      : 1.抜試指示一覧を編集する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    KeyCode     ,I  ,Integer　,キーコード
'*                    Shift       ,I  ,Integer　,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprExamine_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vDeleteFlag     As Variant
    Dim intResult       As Integer
    Dim lngNewRow       As Long
    Dim lngNewRow2      As Long
    Dim blRtn           As Boolean
    Dim vGetHinban      As Variant
    Dim intRowCnt       As Integer
    Dim intColCnt       As Integer
    Dim sDataNum        As String
    Dim vBColor         As Variant
    Dim sColor          As String
    Dim blZhinFlg(2)    As Boolean      'Z品番行の有無
    Dim lngsRow         As Long         '処理開始行
    Dim lngeRow         As Long         '処理終了行

    '' キーコード分岐
    Select Case KeyCode
    Case vbKeyF7        '' F7キー（行挿入）
        'F7キーが使用不可のとき処理を抜ける　2010/12/14　Marushita
        If cmdF(7).Enabled = False Then
            Exit Sub
        End If
        Call sub_F_InsertRow

        '' サンプルフラグオフ
        bSampFlag = False
        cmdF(12).Enabled = False

        '' 特採情報クリア
        Call sub_ClearTokusai
    Case vbKeyF8        '' F8キー（行削除）
        'F8キーが使用不可のとき処理を抜ける　2010/12/14　Marushita
        If cmdF(8).Enabled = False Then
            Exit Sub
        End If
        vDeleteFlag = ""
        'アクティブセルの位置により２行削除(フラグの立っているものは不可)
        '品番を1列追加したことによる列の変更-------start iida 2003/09/06
'        If (sprExamine.GetText(29, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        ''残存酸素検査項目追加による変更　03/12/15 ooba
'        If (sprExamine.GetText(30, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        'GD追加による変更　05/02/17 ooba
        'If (sprExamine.GetText(31, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        If (sprExamine.GetText(37, sprExamine.ActiveRow, vDeleteFlag) = True) Then
            If (vDeleteFlag = "1") Or (vDeleteFlag = "3") Then
                Exit Sub
            End If
        End If
        With sprExamine
            intResult = .ActiveRow Mod 2
            If intResult = 0 Then     '偶数行
                lngNewRow = .ActiveRow
                lngNewRow2 = .ActiveRow + 1
            Else                      '奇数
                lngNewRow = .ActiveRow - 1
                lngNewRow2 = .ActiveRow
            End If

            .GetText 2, lngNewRow2, vGetHinban
            If intResult = 0 Then 'カーソル行が偶数の時は１行上の品番を消去したい場合なので品番を書き換える
                .SetText 2, lngNewRow - 1, vGetHinban
            End If
            .row = lngNewRow
            .row2 = lngNewRow2
            .col = (-1)
            .BlockMode = True
            .FormulaSync = True
            .Action = ActionDeleteRow

            '最大行数削除
            .MaxRows = sprExamine.MaxRows - 2

            '品番の色を適用する
            .GetText 2, lngNewRow - 1, vGetHinban

            .row = lngNewRow - 1
            .col = 4
           vBColor = .backColor

            If (vGetHinban = "Z" Or vGetHinban = "Ｚ") And vBColor = &H8080FF Then
                .row = lngNewRow - 1
                .row2 = lngNewRow - 1
                .col = 4
                .col2 = .MaxCols
                .backColor = &H8080FF
                .row = lngNewRow
                .row2 = lngNewRow
                .col = 5
                .col2 = .MaxCols
                .backColor = &H8080FF
                .BlockMode = True
                .row = lngNewRow - 1
                .row2 = lngNewRow
                .col = 11
                .col2 = .MaxCols
                .ForeColor = &H8080FF
                .BlockMode = False
            Else
                .row = lngNewRow - 1
                .row2 = lngNewRow - 1
                .col = 4
                .col2 = 10
                .backColor = &H80FF80
                .ForeColor = vbBlack
                .row = lngNewRow
                .row2 = lngNewRow
                .col = 5
                .col2 = 10
                .backColor = &H80FF80
                .ForeColor = vbBlack

                '' Z行のみの処理とする=================================>
                '削除行の上がZ品番かどうか
                .col = 11: .row = lngNewRow - 1
                If .backColor = &H8080FF Then blZhinFlg(1) = True Else blZhinFlg(1) = False
                '削除行の下がZ品番かどうか
                .col = 11: .row = lngNewRow
                If .backColor = &H8080FF Then blZhinFlg(2) = True Else blZhinFlg(2) = False
                'Z品番行の場合
                If blZhinFlg(1) Or blZhinFlg(2) Then
                    '処理開始行設定
                    If blZhinFlg(1) Then lngsRow = lngNewRow - 1 Else lngsRow = lngNewRow
                    '処理終了行設定
                    If blZhinFlg(2) Then lngeRow = lngNewRow Else lngeRow = lngNewRow - 1
                    .row = lngsRow
                    .row2 = lngeRow
                    .col = 10
                    .col2 = .MaxCols
                    .ForeColor = vbBlack
                    .row = lngsRow
                    .row2 = lngeRow
                    .col = 11
                    .col2 = .MaxCols
                    .BlockMode = True
                    .backColor = vbWhite
                    .BlockMode = False

                    ''検査項目欄の背景再表示
                    For intRowCnt = lngsRow To lngeRow
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        For intColCnt = 11 To 35
                            .row = intRowCnt
                            .col = intColCnt
                            sDataNum = .text
                            If (sDataNum <> "") Then
                                .row = intRowCnt
                                .col = intColCnt
                                .ForeColor = vbBlack
                                .backColor = vbBlack
                            Else

                            End If
                        Next intColCnt
                    Next intRowCnt
                End If
            End If

            .row = .MaxRows - 1
            .col = 2
            sColor = .text
            If sColor <> "Z" Then
                .row = .MaxRows
                .col = 4
                .backColor = &H80FF80

            End If
            .BlockMode = False
        End With

        '' サンプルフラグオフ
        bSampFlag = False
        cmdF(12).Enabled = False

        '' 特採情報クリア
        Call sub_ClearTokusai
    End Select
End Sub

'*******************************************************************************
'*    関数名        : sprExamine_KeyPress
'*
'*    処理概要      : 1.検査指示サンプルを入力する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    KeyAscii    ,I  ,Integer　,文字コード
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprExamine_KeyPress(KeyAscii As Integer)

    Dim sSampID1 As String
    Dim sSampID2 As String
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer

    With sprExamine
        i = .ActiveRow
        j = .ActiveCol

        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        If j < 11 Or j > 35 Then
            If j = 4 Or j = 3 Then
                bSampFlag = False
            End If
            Exit Sub
        Else
            KeyAscii = 0    'add  2003/05/07 hitec)matsumoto
        End If
        If KeyAscii <> vbKeyBack And (KeyAscii < vbKey3 Or KeyAscii > vbKey4) Then
            KeyAscii = 0
            Exit Sub
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : sprSpec_GotFocus
'*
'*    処理概要      : 1.コントロールがフォーカスを取得した時
'*                      行削除／行挿入ファンクションキーを無効にする
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sprSpec_GotFocus()
    '' 行削除／行挿入ファンクションキーを有効にする
    cmdF(7).Enabled = False
    cmdF(8).Enabled = False
End Sub

'*******************************************************************************
'*    関数名        : txtStaffID_KeyDown
'*
'*    処理概要      : 1.担当者コードKeyDown処理
'*                      （特採機能追加による修正）
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    KeyCode     ,I  ,Integer　,キーコード
'*                    Shift       ,I  ,Integer  ,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtStaffID_KeyDown(KeyCode As Integer, Shift As Integer)
    '特採機能追加による修正
    If f_cmzcChkUser.CanExec("TOKUSAI", txtStaffID.text) Then
        bTokuKengenFlag = True
    Else
        bTokuKengenFlag = False
    End If
End Sub

'*******************************************************************************
'*    関数名        : txtTarget_KeyDown
'*
'*    処理概要      : 1.ねらいρでリターンキー押下
'*                      （ねらいρからのﾌﾞﾛｯｸID、ﾌﾞﾛｯｸP、結晶Pの取得）
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    KeyCode     ,I  ,Integer　,キーコード
'*                    Shift       ,I  ,Integer  ,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtTarget_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim udtTmpResPosCal As type_ResPosCal
    Dim udtCof          As type_Coefficient
    Dim dblMenseki      As Double
    Dim dblTopWght      As Double
    Dim intPos          As Integer
    Dim lngWgtCharge    As Long     '偏析計算用パラメータ
    Dim dblWgtTop       As Double   '偏析計算用パラメータ
    Dim dblWgtTopCut    As Double   '偏析計算用パラメータ
    Dim dblDM           As Double   '偏析計算用パラメータ
    Dim sBlkId          As String
    Dim intBlkPos       As Integer
    Dim dblCalcPos      As Double

    lblMsg.Caption = ""
    txtBlkID = ""
    txtBlkP = ""
    txtCryP.text = ""

    '' リターンキー押下時のみ処理続行
    If KeyCode = vbKeyReturn Then
        If ChkTextBox(txtTarget, CHK_NUMBER, 5, 5) = FUNCTION_RETURN_FAILURE Then
            '' エラーメッセージを表示する
            lblMsg.Caption = GetMsgStr("EINPM")
            Exit Sub
        End If
        txtTarget.text = toRsStr(val(txtTarget.text))
        udtCof.TOPSMPLPOS = tblSXL.INGOTPOS
        udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT
        udtCof.TOPRES = tblTotal.typ_y013(1, WFRES).MESDATA5
        udtCof.BOTRES = tblTotal.typ_y013(2, WFRES).MESDATA5
        
        ''偏析計算用パラメータ取得 マルチ引上対応 参照関数変更 2008/05/22 SETsw Nakada
        If GetCoeffParams_new(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
'        If GetCoeffParams(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
            Debug.Print "偏析計算用パラメータの取得に失敗した"
        End If
        dblMenseki = AreaOfCircle(dblDM)
        dblTopWght = dblWgtTop + dblWgtTopCut
        udtCof.DUNMENSEKI = dblMenseki                      ' 断面積
        udtCof.CHARGEWEIGHT = lngWgtCharge                  ' チャージ量
        udtCof.TOPWEIGHT = dblTopWght                       ' トップ重量
        udtCof.TOPSMPLPOS = tblSXL.INGOTPOS                 ' トップ位置
        udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT  ' ボトム位置

        udtTmpResPosCal.COEFFICIENT = CoefficientCalculation(udtCof)
        udtTmpResPosCal.DUNMENSEKI = udtCof.DUNMENSEKI
        udtTmpResPosCal.CHARGEWEIGHT = udtCof.CHARGEWEIGHT
        udtTmpResPosCal.TOPWEIGHT = udtCof.TOPWEIGHT
        udtTmpResPosCal.TOPSMPLPOS = udtCof.TOPSMPLPOS
        udtTmpResPosCal.TOPRES = udtCof.TOPRES
        udtTmpResPosCal.target = val(txtTarget.text)

        '' 偏析計算関数の呼び出し
        dblCalcPos = PosCalculation(udtTmpResPosCal)
        If dblCalcPos <= -9999 Or dblCalcPos > 9999 Then
            lblMsg.Caption = GetMsgStr("ECLC2")
            Exit Sub
        End If
        intPos = dblCalcPos

        '' 計算位置がSXLの中に含まれている場合
        txtCryP.text = intPos
        If (tblSXL.INGOTPOS <= intPos) And (intPos < tblSXL.INGOTPOS + tblSXL.COUNT) Then
            Call orgXl.GetBlkPos(intPos, sBlkId, intBlkPos)
            txtBlkID = Right(Trim$(sBlkId), 3)
            txtBlkP = ""
            If intBlkPos <> -9999 Then txtBlkP = intBlkPos
        End If
    End If
End Sub

'*******************************************************************************
'*    関数名        : sub_Top_Btm_TEIKOU
'*
'*    処理概要      : 1.画面範囲のねらいρを算出
'*                      （結晶Pのからρを取得）
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    ingot     　,I  ,Integer　,結晶P
'*                    dblTeikou   ,O  ,Double   ,ねらいρ
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_Top_Btm_TEIKOU(ingot As Integer, dblTeikou As Double)
    Dim udtTmpResPosCal As type_ResPosCal
    Dim udtCof          As type_Coefficient
    Dim dblMenseki      As Double
    Dim dblTopWght      As Double
    Dim intPos          As Integer
    Dim lngWgtCharge    As Long     '偏析計算用パラメータ
    Dim dblWgtTop       As Double   '偏析計算用パラメータ
    Dim dblWgtTopCut    As Double   '偏析計算用パラメータ
    Dim dblDM           As Double   '偏析計算用パラメータ
    Dim sBlkId          As String
    Dim intBlkPos       As Integer
    Dim dblCalcPos      As Double

    lblMsg.Caption = ""
    txtBlkID = ""
    txtBlkP = ""
    txtCryP.text = ""

    udtCof.TOPSMPLPOS = tblSXL.INGOTPOS
    udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT
    udtCof.TOPRES = tblTotal.typ_y013(1, WFRES).MESDATA5
    udtCof.BOTRES = tblTotal.typ_y013(2, WFRES).MESDATA5
    
    ''偏析計算用パラメータ取得 マルチ引上対応 参照関数変更 2008/05/22 SETsw Nakada
    If GetCoeffParams_new(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
'    If GetCoeffParams(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "偏析計算用パラメータの取得に失敗した"
    End If

    dblMenseki = AreaOfCircle(dblDM)
    dblTopWght = dblWgtTop + dblWgtTopCut
    udtCof.DUNMENSEKI = dblMenseki                      ' 断面積
    udtCof.CHARGEWEIGHT = lngWgtCharge                  ' チャージ量
    udtCof.TOPWEIGHT = dblTopWght                       ' トップ重量
    udtCof.TOPSMPLPOS = tblSXL.INGOTPOS                 ' トップ位置
    udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT  ' ボトム位置

    udtTmpResPosCal.COEFFICIENT = CoefficientCalculation(udtCof)
    udtTmpResPosCal.DUNMENSEKI = udtCof.DUNMENSEKI
    udtTmpResPosCal.CHARGEWEIGHT = udtCof.CHARGEWEIGHT
    udtTmpResPosCal.TOPWEIGHT = udtCof.TOPWEIGHT
    udtTmpResPosCal.TOPSMPLPOS = udtCof.TOPSMPLPOS
    udtTmpResPosCal.TOPRES = udtCof.TOPRES
    udtTmpResPosCal.target = val(ingot)

    '' 偏析計算関数の呼び出し
    dblCalcPos = ResCalculation(udtTmpResPosCal)
    If dblCalcPos <= -9999 Or dblCalcPos > 9999 Then
        lblMsg.Caption = GetMsgStr("ECLC2")
        Exit Sub
    End If

    dblTeikou = dblCalcPos
    Debug.Print "計算抵抗値 = " & dblTeikou
End Sub

'*******************************************************************************
'*    関数名        : sub_DispClear
'*
'*    処理概要      : 1.画面の全データをクリアする
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_DispClear()
    lblMsg.Caption = ""
    lblKanren.Visible = False   '08/01/31 ooba
    txtStaffID.text = ""
    txtJfName.text = ""
    txtCryNum.text = ""
    txtKSXLID.text = ""
    txtTopRsltR.text = ""
    txtBotRsltR.text = ""
    txtTarget.text = ""
    txtBlkID.text = ""
    txtBlkP.text = ""
    txtCryP.text = ""
    sprSpec.MaxRows = 0
    sprExamine.MaxRows = 0
    sprWarp.MaxRows = 0
    bSampFlag = False
    cmdF(2).Enabled = False
    cmdF(5).Enabled = False '特採ボタン
    cmdF(5).backColor = &H8000000F
    cmdF(12).Enabled = False

    Call Pic_Disp(0)
'----- サンプルボタン状態初期化　2010/12/08 Marushita
    bSampleBtn = True
    cmdF(10).Enabled = True
End Sub

'*******************************************************************************
'*    関数名        : fnc_DispHinSpec
'*
'*    処理概要      : 1.製品仕様データを表示する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    intsyoki    ,I  ,Integer  ,0:初期表示、1:サンプルボタン
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Function fnc_DispHinSpec(intsyoki As Integer) As Boolean
    Dim udtTmpHin       As tFullHinban
    Dim sHin            As String
    Dim sErrMsg         As String
    Dim vTemp           As Variant
    Dim intNLen         As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim vNukisiFlg      As Variant
    Dim vOldNukisiFlg   As Variant
    Dim intHinLoop      As Integer
    Dim intCnt          As Integer
    Dim sRev            As String
    Dim blSameHin       As Boolean
    Dim s               As Integer

    fnc_DispHinSpec = False
    With sprExamine
        '' 品番チェック
        For i = 1 To sprExamine.MaxRows - 1
            '品番を1列追加したことによる列の変更-------start iida 2003/09/06
'            sprExamine.GetText 29, i, vNukisiFlg
            ''残存酸素検査項目追加による変更　03/12/15 ooba
'            sprExamine.GetText 30, i, vNukisiFlg
            'GD追加による変更　05/02/17 ooba
            'sprExamine.GetText 31, i, vNukisiFlg
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            sprExamine.GetText 37, i, vNukisiFlg
            If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Then    '抜試行だったら
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                sprExamine.GetText 37, i + 1, vOldNukisiFlg
                If (vNukisiFlg = 1) And (vOldNukisiFlg = 1) Then  '２行続けて初期表示抜試行だった時
                    '何もしない
                Else
                    If (vNukisiFlg = 1) Then
                        Call sprExamine.GetText(2, i, vTemp)
                    Else
                        Call sprExamine.GetText(2, i + 1, vTemp)
                    End If
                    vTemp = Trim(vTemp)
                    If Trim$(vTemp) <> "Z" Then
                        Select Case ChkString(vTemp, 8, 8)
                        Case CHK_NG, CHK_NULL
                            lblMsg.Caption = GetMsgStr(EHIN1)
                            Exit Function
                        End Select
                        sHin = vTemp
                        If GetLastHinban(sHin, udtTmpHin) = FUNCTION_RETURN_FAILURE Then
                            lblMsg.Caption = GetMsgStr(EHIN0)
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i

        '' 品番の設定
        m = .MaxRows
        ReDim tblHinbanRs(m)
        If m = 0 Then
            Exit Function
        End If
        j = 0
        For i = 1 To m - 1
            .row = i

            '--- 2006/08/15 Cng エピ先行評価追加対応
            sprExamine.GetText 37, i, vNukisiFlg

            If intsyoki = 0 Then  '初期表示のときの処理
                '抜試行だったら
                If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Then
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    sprExamine.GetText 37, i + 1, vOldNukisiFlg

                    If (vNukisiFlg = 1) Then
                        .col = 2
                        sHin = Trim$(.text)
                        .col = 3
                        sRev = Trim$(.text)
                    Else
                        .row = i + 1
                        .col = 2
                        sHin = Trim$(.text)
                        .col = 3
                        sRev = Trim$(.text)
                    End If
                    If sHin <> "Z" Then
                        .col = 4
                        intNLen = val(.text)
                        If vNukisiFlg = 1 Then      '2003/11/14 SystemBrain 良く解らないけど、無理やりに・・・ ▽
                            .row = i + 1                    'でも、複数ﾌﾞﾛｯｸ1SXLの時、だめでした
                        Else
                            .row = i + 2
                        End If                      '2003/11/14 SystemBrain 良く解らないけど、無理やりに・・・ △
                        If i = 1 Or i <> m Then
                            .col = 4
                            intNLen = val(.text) - intNLen

                            '同一の12桁品番が既に存在するかﾁｪｯｸし、既にあれば作成しない 2003/11/14 SystemBrain ▽
                            blSameHin = False
                            For s = 1 To j
                                If tblHinbanRs(s).HIN.hinban = sHin And tblHinbanRs(s).HIN.mnorevno = val(Mid(sRev, 1, 2)) And _
                                   tblHinbanRs(s).HIN.factory = Mid(sRev, 3, 1) And tblHinbanRs(s).HIN.opecond = Mid(sRev, 4, 1) Then
                                    tblHinbanRs(s).LENGHT = tblHinbanRs(s).LENGHT + intNLen    'たぶん加算 ⇒ vNukisiFlg=0のﾃﾞｰﾀはここを通らないので長さが足りない
                                    blSameHin = True
                                    Exit For
                                End If
                            Next s

                            '同一の12桁品番が既に存在するかﾁｪｯｸし、既にあれば作成しない 2003/11/14 SystemBrain △
                            If Not blSameHin Then
                                j = j + 1

                                With tblHinbanRs(j)                                                 '表示してる品番の仕様なんだって 2003/11/14 ▽
                                    .CRYNUM = tblSXL.CRYNUM
                                    .HIN.hinban = sHin                        ' 品番
                                    .HIN.mnorevno = val(Mid(sRev, 1, 2))      ' 製品番号改訂番号
                                    .HIN.factory = Mid(sRev, 3, 1)            ' 工場
                                    .HIN.opecond = Mid(sRev, 4, 1)            ' 操業条件
                                    .LENGHT = intNLen                            ' 長さ
                                End With                                                            '表示してる品番の仕様なんだって 2003/11/14 △
                            End If
                        End If
                    End If
                End If
            Else
'                '条件を追加(2ブロック1SXL仕様のない品番に振替対応) 初期表示以外
                If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Or (i Mod 2 = 0 And vNukisiFlg = 3 And UBound(tblWafInd()) > 2) Then
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    sprExamine.GetText 37, i + 1, vOldNukisiFlg

                    If (vNukisiFlg = 1) Then
                        .col = 2
                        sHin = Trim$(.text)
                        .col = 3
                        sRev = Trim$(.text)
                    Else
                        .row = i + 1
                        .col = 2
                        sHin = Trim$(.text)
                        .col = 3
                        sRev = Trim$(.text)
                    End If
                    If sHin <> "Z" Then
                        .col = 4
                        intNLen = val(.text)
                        If vNukisiFlg = 1 Then      '2003/11/14 SystemBrain 良く解らないけど、無理やりに・・・ ▽
                            .row = i + 1                    'でも、複数ﾌﾞﾛｯｸ1SXLの時、だめでした
                        Else
                            .row = i + 2
                        End If                      '2003/11/14 SystemBrain 良く解らないけど、無理やりに・・・ △
                        If i = 1 Or i <> m Then
                            .col = 4
                            intNLen = val(.text) - intNLen

                            '同一の12桁品番が既に存在するかﾁｪｯｸし、既にあれば作成しない 2003/11/14 SystemBrain ▽
                            blSameHin = False
                            For s = 1 To j
                                If tblHinbanRs(s).HIN.hinban = sHin And tblHinbanRs(s).HIN.mnorevno = val(Mid(sRev, 1, 2)) And _
                                   tblHinbanRs(s).HIN.factory = Mid(sRev, 3, 1) And tblHinbanRs(s).HIN.opecond = Mid(sRev, 4, 1) Then
                                    tblHinbanRs(s).LENGHT = tblHinbanRs(s).LENGHT + intNLen    'たぶん加算 ⇒ vNukisiFlg=0,3のﾃﾞｰﾀはここを通らないので長さが足りない
                                    blSameHin = True
                                    Exit For
                                End If
                            Next s

                            '同一の12桁品番が既に存在するかﾁｪｯｸし、既にあれば作成しない 2003/11/14 SystemBrain △
                            If Not blSameHin Then
                                j = j + 1
                                With tblHinbanRs(j)                                                 '表示してる品番の仕様なんだって 2003/11/14 ▽
                                    .CRYNUM = tblSXL.CRYNUM
                                    .HIN.hinban = sHin                        ' 品番
                                    .HIN.mnorevno = val(Mid(sRev, 1, 2))      ' 製品番号改訂番号
                                    .HIN.factory = Mid(sRev, 3, 1)            ' 工場
                                    .HIN.opecond = Mid(sRev, 4, 1)            ' 操業条件
                                    .LENGHT = intNLen                            ' 長さ
                                End With                                                            '表示してる品番の仕様なんだって 2003/11/14 △
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        ReDim Preserve tblHinbanRs(j)
    End With

    '' DBからデータを読み込む
    '' 仕様部分のチェック
    If DBDRV_scmzc_fcmlc001d_DispSiyou(tblHinbanRs, tblsiyou, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = sErrMsg
        Exit Function
    End If

    '' 製品仕様データの表示
    m = UBound(tblHinbanRs())
    With sprSpec
        j = 0
        .MaxRows = m
        For i = 1 To m
            If Trim$(tblHinbanRs(i).HIN.hinban) <> "" Then
                j = j + 1
                .row = j
                .col = 1
                .text = tblHinbanRs(i).HIN.hinban   ' 品番
                .col = 2
                .text = Format(tblHinbanRs(j).HIN.mnorevno, "00") & tblHinbanRs(j).HIN.factory & tblHinbanRs(j).HIN.opecond
                .col = 3                            ' 比抵抗
                .text = toRsStr_nl(tblsiyou(i).HWFRMIN, tblsiyou(i).HWFRMAX) '2001/12/26 S.Sano
                .col = 4
                .text = tblsiyou(i).KEIKAKUL        ' 計画長
                .col = 5
                .text = tblHinbanRs(i).LENGHT       ' 推定長
                .col = 6
                .text = tblsiyou(i).HWFRHWYS        ' Rs
                .col = 7
                .text = tblsiyou(i).HWFONHWS        ' Oi
                .col = 8
                .text = tblsiyou(i).HWFBM1HS        ' B1
                .col = 9
                .text = tblsiyou(i).HWFBM2HS        ' B2
                .col = 10
                .text = tblsiyou(i).HWFBM3HS        ' B3
                .col = 11
                .text = tblsiyou(i).HWFOF1HS        ' L1
                .col = 12
                .text = tblsiyou(i).HWFOF2HS        ' L2
                .col = 13
                .text = tblsiyou(i).HWFOF3HS        ' L3
                .col = 14
                '.text = tblsiyou(i).HWFOF4HS        ' L4
                'Change 2010/01/06 Y.Hitomi
                .text = tblsiyou(i).HWFSIRDHS       ' SD(SIRD)
                .col = 15
                .text = tblsiyou(i).HWFDSOHS        ' DS
                .col = 16
                .text = tblsiyou(i).HWFMKHWS        ' DZ
                .col = 17

                '拡散長,Nr濃度追加　06/06/08 ooba START ===================>
                If tblsiyou(i).HWFSPVHS = "H" Or _
                   tblsiyou(i).HWFDLHWS = "H" Or _
                   tblsiyou(i).HWFNRHS = "H" Then
                    .text = "H"
                ElseIf tblsiyou(i).HWFSPVHS = "S" Or _
                       tblsiyou(i).HWFDLHWS = "S" Or _
                       tblsiyou(i).HWFNRHS = "S" Then
                    .text = "S"
                Else
                    .text = tblsiyou(i).HWFSPVHS
                End If

                .col = 18
                .text = tblsiyou(i).HWFOS1HS        ' D1
                .col = 19
                .text = tblsiyou(i).HWFOS2HS        ' D2
                .col = 20
                .text = tblsiyou(i).HWFOS3HS        ' D3
                .col = 21
                .text = tblsiyou(i).HWFZOHWS        ' AO        ''残存酸素追加
                ''残存酸素検査項目追加による変更　03/12/15 ooba

                '仕様ｳｨﾝﾄﾞｩOT1､OT2に指示の有無を表示する
                If tblsiyou(i).HWFOT1 = "1" Then
                    .col = 22
                    .text = "有"
                ElseIf tblsiyou(i).HWFOT1 = "0" Then
                    .col = 22
                    .text = "無"
                End If
                If tblsiyou(i).HWFOT2 = "1" Then
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
                    .col = 30
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                    .text = "有"
                ElseIf tblsiyou(i).HWFOT2 = "0" Then
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
                    .col = 30
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                    .text = "無"
                End If

'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
                .col = 23
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
                If tblsiyou(i).HWFDENHS = "H" Or tblsiyou(i).HWFLDLHS = "H" Or _
                        tblsiyou(i).HWFDVDHS = "H" Then
                    .text = "H"
                ElseIf tblsiyou(i).HWFDENHS = "S" Or tblsiyou(i).HWFLDLHS = "S" Or _
                        tblsiyou(i).HWFDVDHS = "S" Then
                    .text = "S"
                Else
                    .text = " "
                End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                .col = 24:  .text = tblsiyou(i).HEPBM1HS
                .col = 25:  .text = tblsiyou(i).HEPBM2HS
                .col = 26:  .text = tblsiyou(i).HEPBM3HS
                .col = 27:  .text = tblsiyou(i).HEPOF1HS
                .col = 28:  .text = tblsiyou(i).HEPOF2HS
                .col = 29:  .text = tblsiyou(i).HEPOF3HS
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                '>>>>> 中間抜試規格セット追加 2011/07/15 Marushita
                .col = 31:  .text = tblsiyou(i).CHUTAN
                .col = 32:  .text = tblsiyou(i).CHUKYO
                .col = 33:  .text = tblsiyou(i).CHUFLG
                '<<<<< 中間抜試規格セット追加 2011/07/15 Marushita
            End If
        Next i
        .MaxRows = j

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        '' WF仕様を初期表示する
        Call sub_cmbc039_3_ChangeHinSpec(0)
        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "ＷＦ >>"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
    End With

    '仕様がないデータへの振替を行ったときはエラーにしない
    '初期表示行の品番だけを変更したとき
    '初期表示処理のときはチェックをしない
    If intsyoki <> 0 Then
    For i = 1 To UBound(tblsiyou())
        intCnt = 0
        If tblsiyou(i).HWFRHWYS <> "H" And tblsiyou(i).HWFRHWYS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFONHWS <> "H" And tblsiyou(i).HWFONHWS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFBM1HS <> "H" And tblsiyou(i).HWFBM1HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFBM2HS <> "H" And tblsiyou(i).HWFBM2HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFBM3HS <> "H" And tblsiyou(i).HWFBM3HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOF1HS <> "H" And tblsiyou(i).HWFOF1HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOF2HS <> "H" And tblsiyou(i).HWFOF2HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOF3HS <> "H" And tblsiyou(i).HWFOF3HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOF4HS <> "H" And tblsiyou(i).HWFOF4HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFDSOHS <> "H" And tblsiyou(i).HWFDSOHS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFMKHWS <> "H" And tblsiyou(i).HWFMKHWS <> "S" Then
            intCnt = intCnt + 1
        End If
        '拡散長,Nr濃度追加
        If tblsiyou(i).HWFSPVHS <> "H" And tblsiyou(i).HWFSPVHS <> "S" And _
           tblsiyou(i).HWFDLHWS <> "H" And tblsiyou(i).HWFDLHWS <> "S" And _
           tblsiyou(i).HWFNRHS <> "H" And tblsiyou(i).HWFNRHS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOS1HS <> "H" And tblsiyou(i).HWFOS1HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOS2HS <> "H" And tblsiyou(i).HWFOS2HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOS3HS <> "H" And tblsiyou(i).HWFOS3HS <> "S" Then
            intCnt = intCnt + 1
        End If
        ''残存酸素追加
        If tblsiyou(i).HWFZOHWS <> "H" And tblsiyou(i).HWFZOHWS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOT1 = "0" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOT2 = "0" Then
            intCnt = intCnt + 1
        End If
        'GD追加
        If tblsiyou(i).HWFDENHS <> "H" And tblsiyou(i).HWFDENHS <> "S" And _
                tblsiyou(i).HWFLDLHS <> "H" And tblsiyou(i).HWFLDLHS <> "S" And _
                tblsiyou(i).HWFDVDHS <> "H" And tblsiyou(i).HWFDVDHS <> "S" Then
            intCnt = intCnt + 1
        End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        If tblsiyou(i).HEPOF1HS <> "H" And tblsiyou(i).HEPOF1HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HEPOF2HS <> "H" And tblsiyou(i).HEPOF2HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HEPOF3HS <> "H" And tblsiyou(i).HEPOF3HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HEPBM1HS <> "H" And tblsiyou(i).HEPBM1HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HEPBM2HS <> "H" And tblsiyou(i).HEPBM2HS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HEPBM3HS <> "H" And tblsiyou(i).HEPBM3HS <> "S" Then
            intCnt = intCnt + 1
        End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        If intCnt = 25 Then
            If UBound(tblWafInd()) > 2 Then
                lblMsg.Caption = "仕様がありません"
                cmdF(12).Enabled = False
                Exit Function
            End If
        End If
    Next
    End If
    fnc_DispHinSpec = True
End Function

'*******************************************************************************
'*    関数名        : sub_DispExamineData
'*
'*    処理概要      : 1.抜試指示データを表示する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_DispExamineData()
    Dim Blk             As c_cmzcBlk
    Dim sList           As String
    Dim intREBlockPos   As Integer
    Dim intBlockStPos   As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim c0              As Integer
    Dim lngSelectpoint  As Long

    If Trim$(tblsmp(1).SMPLID) = "" Then
        ReDim tblsmp(2)
        tblsmp(1).INGOTPOS = tblSXL.INGOTPOS
        tblsmp(1).hinban = tblSXL.hinban
        tblsmp(2).INGOTPOS = tblSXL.INGOTPOS + tblSXL.COUNT
        tblsmp(2).hinban = tblSXL.hinban
    End If

    '' 抜試指示データの表示
    With sprExamine
        .MaxRows = 2
        For i = 1 To 2
            '' ブロックID
            .row = i
            .col = 1
            If tblsmp(i).INGOTPOS <> 0 Then
                If i = 1 Then
                    intBlockStPos = orgXl.Blks.LowerArea(CStr(tblsmp(i).INGOTPOS))
                Else
                    intBlockStPos = orgXl.Blks.UpperArea(CStr(tblsmp(i).INGOTPOS))
                End If
            Else
                intBlockStPos = 0
            End If
            Set Blk = orgXl.Blks(CStr(intBlockStPos))
            .text = Right(Blk.BLOCKID, 3)
            .backColor = COLOR_OK

            '' ブロックP
            .col = 2
            intREBlockPos = DBData2DispData(tblsmp(i).INGOTPOS - intBlockStPos, "0")
            .text = intREBlockPos

            '' 結晶P
            .col = 3
            .text = tblsmp(i).INGOTPOS

            '' 区分コンボの設定
            If i = 1 Then
                m = UBound(tblPrcList)
                sList = GetGPCodeDspStr(tblPrcList(1).CODE, tblPrcList(1).INFO1)
                For j = 2 To m
                    sList = sList & vbTab & GetGPCodeDspStr(tblPrcList(j).CODE, tblPrcList(j).INFO1)
                Next j
                .col = 4
                .TypeComboBoxList = sList
                .TypeComboBoxCurSel = 0
            End If

            '' 品番
            .col = 5
            If i = 1 Then
                .text = Trim(tblTotal.typ_Param.hinban)
            Else
                '' SXLの下端は内側の品番を表示させる
                .text = tblsmp(1).hinban
            End If
        Next i

        '' WF総合判定にて再抜試の指示がでた場合、行追加
        For i = 1 To 2
            If tblTotal.bOKNG(i) = False And _
               tblTotal.dblScut(i) > tblSXL.INGOTPOS And _
               tblTotal.dblScut(i) < tblSXL.INGOTPOS + tblSXL.COUNT Then
                .MaxRows = .MaxRows + 1
                .row = .ActiveRow
                .Action = ActionInsertRow

                '' ブロックID
                intBlockStPos = orgXl.Blks.UpperPos(tblTotal.dblScut(i))
                Set Blk = orgXl.Blks(CStr(intBlockStPos))

                '' ブロックIDコンボの設定
                .col = 1
                m = UBound(SxlIntoBlock)
                If m = 1 Then
                    .Lock = True
                    .CellType = CellTypeStaticText
                    .text = SxlIntoBlock(1).SORTID
                Else
                    .Lock = False
                    .CellType = CellTypeComboBox
                    sList = SxlIntoBlock(1).SORTID
                    For j = 2 To m
                        sList = sList & vbTab & SxlIntoBlock(j).SORTID
                    Next j
                    .TypeComboBoxList = sList
                    m = .TypeComboBoxCount
                    For j = 0 To m
                        .TypeComboBoxCurSel = 0
                        If .text = Right(Blk.BLOCKID, 3) Then
                            .TypeComboBoxCurSel = j
                            Exit For
                        End If
                    Next j
                End If
                .TypeHAlign = TypeHAlignCenter

                '' ブロックP
                .col = 2
                intREBlockPos = DBData2DispData(tblTotal.dblScut(i) - orgXl.Blks.UpperPos(tblTotal.dblScut(i)), "0")
                .text = intREBlockPos
                .col = 3
                .text = DBData2DispData(tblTotal.dblScut(i), "0")

                '' 区分コンボの設定
                m = UBound(tblPrcList)
                sList = GetGPCodeDspStr(tblPrcList(1).CODE, tblPrcList(1).INFO1)
                For j = 2 To m
                    sList = sList & vbTab & GetGPCodeDspStr(tblPrcList(j).CODE, tblPrcList(j).INFO1)
                Next j
                .col = 4
                .TypeComboBoxList = sList
                .TypeComboBoxCurSel = 0

                '' サンプルIDの設定
                '' 追加行の品番設定(下側のZ品番は追加行に対してZをたてるため、このループ内で設定する)
                .col = 5
                If i = 1 Then
                    .text = tblSXL.hinban
                ElseIf i = 2 Then
                    .text = "Z"
                End If
                InMaxRow = InMaxRow + 1
            End If
        Next i

        '' 結晶位置にてソートをかける
        .col = 1
        .col2 = .MaxCols
        .row = 1
        .row2 = .MaxRows
        .SortBy = SortByRow
        .SortKey(1) = 3
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Action = ActionSort

        '' 品番に廃棄情報設定(上側NGの場合は、1行目をZにする)
        If tblTotal.bOKNG(1) = False And _
           tblTotal.dblScut(1) > tblSXL.INGOTPOS And _
           tblTotal.dblScut(1) < tblSXL.INGOTPOS + tblSXL.COUNT Then
            .row = 1
            .col = 5
            .TypeComboBoxCurSel = 1
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_DispSample
'*
'*    処理概要      : 1.規定値の検査指示サンプルデータを表示する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    blnsflg　　 ,I  ,Boolean  ,データチェック完了フラグ
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Function sub_DispSample(Optional blnsflg As Boolean) As Boolean
    Dim sSampID1        As String
    Dim sSampID2        As String
    Dim m               As Integer
    Dim n               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim blAns           As Boolean
    Dim vTemp           As Variant
    Dim vTemp1          As Variant
    Dim vTemp2          As String
    Dim c0              As Long
    Dim udtFullHinban   As tFullHinban
    Dim udtTmpWafSmp()  As typ_XSDCW
    Dim blNukisiOk      As Boolean
    Dim intLoopCnt      As Integer
    Dim vNukisiFlg      As Variant
    Dim intNukisiRow      As Integer
    Dim vgetlotid       As Variant
    Dim vGetLotId2      As Variant
    Dim vGetBlkP        As Variant
    Dim vGetIngotP      As Variant
    Dim vGetBlkSeq      As Variant
    Dim vGetBlkSeq2     As Variant
    Dim vGetHinban      As Variant
    Dim vSample1        As Variant
    Dim vSample2        As Variant
    Dim intNextBlkP     As Integer
    Dim vNextBlkP       As Variant
    Dim vGetWfNum       As Variant
    Dim vViewSmpId      As Variant
    Dim vViewSmpId2     As Variant
    Dim vAllNum         As Variant
    Dim vBlockId        As Variant
    Dim intCngBlpnt     As Integer
    Dim vWFNumber       As Variant
    Dim vKeturaku       As Variant
    Dim vNull           As Variant
    Dim vZERO           As Variant
    Dim sNextIngotP     As String
    Dim sHinban         As String
    Dim intRowNum       As Integer
    Dim intCol          As Integer
    Dim sNum            As String
    Dim sColor          As String
    Dim sNukishi        As String
    Dim intWfChkLoop    As Integer
    Dim vGetUpWf        As Variant
    Dim vGetDnWf        As Variant
    Dim vGetUpBlk       As Variant
    Dim vGetDnBlk       As Variant
    Dim blKensaLock     As Boolean
    Dim blKirikaeflg    As Boolean
    Dim vGetToHin       As Variant
    Dim vGetToSamp      As Variant
    Dim sMsg            As String
    Dim intModori       As Integer  '戻り値
    Dim sGHin           As String   '品番
    Dim blnhflg         As Boolean  '反映判定フラグ
    Dim intZkbn         As Integer  'Z区分
    Dim udtFHin         As tFullHinban
    Dim intSmpkbn       As Integer 'サンプル区分が代表サンプルか判断する
    Dim sHinRev         As String
    Dim intBcnt         As Integer  '共有(別)サンプル数
    
    Dim flg          As Integer
    Dim now_bid      As String
    Dim old_bid      As String
    Dim iJCnt1      As Integer
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
    Dim sird1stBlockSet As Boolean          '結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END
     

    sub_DispSample = False
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
    sird1stBlockSet = False             '結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END
    
    intBcnt = 1
    ReDim iBetuRow(sprExamine.MaxRows)

    blNukisiOk = False
    '' 再抜試の新規行が存在しなければ処理を抜ける&全数廃棄のZ品番のみ有効とする

    With sprExamine
        For intLoopCnt = 1 To .MaxRows
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            .GetText 37, intLoopCnt, vNukisiFlg
            If (vNukisiFlg <> "1") And (vNukisiFlg <> "3") Then
                blNukisiOk = True
                Exit For
            End If
        Next
        If blNukisiOk = False Then
            For intLoopCnt = 1 To .MaxRows Step 2
                .row = intLoopCnt
                .col = 2
                sHinRev = Trim$(.text)
                .col = 3
                sHinRev = sHinRev & Trim$(.text)
                If sHinRev <> Trim(tblTotal.typ_Param.hinban) + _
                              Trim(Format(tblTotal.typ_Param.REVNUM, "00")) + _
                              Trim(tblTotal.typ_Param.factory) + _
                              Trim(tblTotal.typ_Param.opecond) Then
                    blNukisiOk = True
                    Exit For
                End If
            Next intLoopCnt

            For intLoopCnt = 1 To .MaxRows
                .row = intLoopCnt
                .col = 1
                If .CellType = CellTypeCheckBox Then
                    If .text = "1" Then
                        blNukisiOk = True
                    End If
                End If
            Next intLoopCnt
            If blNukisiOk = False Then
                lblMsg.Caption = GetMsgStr("SET48")
                Exit Function
            End If

        End If
    End With

    '' 抜試指示一覧の入力チェック（WFﾏｯﾌﾟ）
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

'品番のＮＵＬＬチェックはすでにはじいてある。
'ここでＮＵＬＬをはじいているのは、画面レイアウト変更により、ＮＵＬＬ行(品番入力不可)が存在するため。
    '' 品番チェック
    For c0 = 1 To sprExamine.MaxRows - 1
        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg

        If (vNukisiFlg = "1") Or (vNukisiFlg = "2") Then    '抜試行だったら
            If (vNukisiFlg = "1") Then
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                sprExamine.GetText 37, c0 + 1, vNukisiFlg

                If vNukisiFlg = "1" Then
                    '追加抜試の無い行は何もしない
                ElseIf vNukisiFlg = "2" Then
                    sprExamine.GetText 2, c0, vTemp             '各行に品番が書いてあるわけではない、ダミーでは全行に品番を持っているのでそれを見る
                    vTemp = Trim(vTemp)
                    If Trim$(vTemp) <> "Z" Then
                        Select Case ChkString(vTemp, 8, 8)
                        Case CHK_NG, CHK_NULL
                            lblMsg.Caption = GetMsgStr(EHIN1)
                            Exit Function
                        End Select
                        vTemp2 = vTemp
                        If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                            lblMsg.Caption = GetMsgStr(EHIN0)
                            Exit Function
                        ''残存酸素仕様チェック
                        Else
                            iChkAoi = ChkAoiSiyou(udtFullHinban)
                            If iChkAoi < 0 Then
                                lblMsg.Caption = "残存酸素(AOi)仕様エラー"
                                Exit Function
                            End If
                        End If
                    End If
                End If
            Else
                sprExamine.GetText 2, c0 + 1, vTemp
                vTemp = Trim(vTemp)
                If Trim$(vTemp) <> "Z" Then
                    Select Case ChkString(vTemp, 8, 8)
                    Case CHK_NG, CHK_NULL
                        lblMsg.Caption = GetMsgStr(EHIN1)
                        Exit Function
                    End Select
                    vTemp2 = vTemp
                    If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                        lblMsg.Caption = GetMsgStr(EHIN0)
                        Exit Function
                    ''残存酸素仕様チェック
                    Else
                        iChkAoi = ChkAoiSiyou(udtFullHinban)
                        If iChkAoi < 0 Then
                            lblMsg.Caption = "残存酸素(AOi)仕様エラー"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next c0

   '抜試指示用構造体
    ReDim Preserve tblNukishi(m)

    '' 区分エラーチェック
    If fnc_CheckHinbanZ() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EHIN6")
        Exit Function
    End If

    '' エラーチェック
    If fnc_CheckBlockP() = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

    '抜試ﾃﾞｰﾀ不正表示チェック
    If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr(sMsg)
        Exit Function
    End If

    'データチェックが完了したらフラグを立てる
    blnsflg = True

    '' 検査指示確認メッセージ
    If MsgBox(GetMsgStr("PSMP1"), vbOKCancel, "検査指示サンプル処理") = vbCancel Then
        sub_DispSample = False
        Exit Function
    End If

    lblMsg.Caption = GetMsgStr("PWAIT")

     '構造体クリア
    ReDim tblKns(0) '検査項目切替用の構造体をクリアする

    DoEvents

    bSampFlag = False
    bMotoGDcpyFlg(1) = False
    bMotoGDcpyFlg(2) = False

    '' 再抜試指示テーブルの更新
    If fnc_UpdateData() = FUNCTION_RETURN_FAILURE Then
        cmdF(12).Enabled = False
        bSampFlag = False
    Else
        cmdF(12).Enabled = True
        bSampFlag = True
    End If

    '' 抜試指示データの表示
    With sprExamine
        .ReDraw = False
        If .MaxRows = 0 Then
            Exit Function
        End If

       '' 検査項目の内容を設定
        m = .MaxRows
        n = UBound(tblWafInd)
        intNukisiRow = 0
        intCngBlpnt = 1
        For i = 1 To m
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            .GetText 37, i, vNukisiFlg

    ''''''' '初期表示のブロックの境
            If i Mod 2 = 0 And vNukisiFlg = "3" Then
                'サンプルクリア
                .SetText 10, i, vbNullString
                .SetText 10, i + 1, vbNullString
                .SetText 8, i, gsWF_STA_NORMAL
                .SetText 8, i + 1, gsWF_STA_NORMAL
                '検査項目クリア
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                For j = 11 To 35
                    .SetText j, i, vbNullString
                    .row = i
                    .row2 = i
                    .col = j
                    .col2 = j
                    .Lock = True
                    .BlockMode = True
                    .backColor = vbWhite
                    .Lock = False
                    .SetText j, i + 1, vbNullString
                    .row = i + 1
                    .col = j
                    .Lock = True
                    .BlockMode = True
                    .backColor = vbWhite
                    .BlockMode = False
                Next j
            End If

            If (vNukisiFlg <> "1") Then
                .row = i
                If bSampFlag = False Then
                    Exit Function
                End If

                '' 検査指示の設定
                For j = 1 To n
                    With tblWafInd(j)
                        .SMP.CRYINDRS = GetWFSamp(.HINUP, .HINDN, 1)
                        .SMP.CRYINDOI = GetWFSamp(.HINUP, .HINDN, 2)
                        .SMP.CRYINDB1 = GetWFSamp(.HINUP, .HINDN, 3)
                        .SMP.CRYINDB2 = GetWFSamp(.HINUP, .HINDN, 4)
                        .SMP.CRYINDB3 = GetWFSamp(.HINUP, .HINDN, 5)
                        .SMP.CRYINDL1 = GetWFSamp(.HINUP, .HINDN, 6)
                        .SMP.CRYINDL2 = GetWFSamp(.HINUP, .HINDN, 7)
                        .SMP.CRYINDL3 = GetWFSamp(.HINUP, .HINDN, 8)
                        .SMP.CRYINDL4 = GetWFSamp(.HINUP, .HINDN, 9)
                        .SMP.CRYINDDS = GetWFSamp(.HINUP, .HINDN, 10)
                        .SMP.CRYINDDZ = GetWFSamp(.HINUP, .HINDN, 11)
                        .SMP.CRYINDSP = GetWFSamp(.HINUP, .HINDN, 12)
                        .SMP.CRYINDD1 = GetWFSamp(.HINUP, .HINDN, 13)
                        .SMP.CRYINDD2 = GetWFSamp(.HINUP, .HINDN, 14)
                        .SMP.CRYINDD3 = GetWFSamp(.HINUP, .HINDN, 15)
                        .SMP.CRYOTHER1 = GetWFSamp(.HINUP, .HINDN, 16)
                        .SMP.CRYOTHER2 = GetWFSamp(.HINUP, .HINDN, 17)
                        .SMP.CRYINDAO = GetWFSamp(.HINUP, .HINDN, 18)    '残存酸素追加
                        .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)    'GD追加

                        'GDﾗｲﾝﾁｪｯｸ機能追加
                        If Trim(.SMP.CRYINDGD) = "3" Then
                            .SMP.CRYINDGD2 = GetWFSamp(.HINUP, .HINDN, 26)
                        Else
                            .SMP.CRYINDGD2 = ""
                        End If

                        '--- 2006/08/15 Add エピ先行評価追加対応
                        .SMP.EPIINDB1 = GetWFSamp(.HINUP, .HINDN, 20)
                        .SMP.EPIINDB2 = GetWFSamp(.HINUP, .HINDN, 21)
                        .SMP.EPIINDB3 = GetWFSamp(.HINUP, .HINDN, 22)
                        .SMP.EPIINDL1 = GetWFSamp(.HINUP, .HINDN, 23)
                        .SMP.EPIINDL2 = GetWFSamp(.HINUP, .HINDN, 24)
                        .SMP.EPIINDL3 = GetWFSamp(.HINUP, .HINDN, 25)
                    End With
                Next j

                '検査項目の分類表示
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .GetText 37, i, vNukisiFlg

                If CheckGetSampleID(i) = True Then
                    intNukisiRow = intNukisiRow + 1
                    '表示色一度白に
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    For j = 11 To 35
                        .col = j
                        .row = i
                        .Lock = False
                        .backColor = vbWhite
                        .ForeColor = vbWhite
                        .Lock = True
                        .row = i + 1
                        .Lock = False
                        .backColor = vbWhite
                        .ForeColor = vbWhite
                        .Lock = True
                    Next j

                    '反映推定処理追加のため修正
                    If GetSampleID(intNukisiRow, sSampID1, sSampID2, CInt(vNukisiFlg)) = True Then
                    '共有サンプルが存在し、切り替えができる場合
                        blKirikaeflg = True  '2003/05/26 okazaki

                        If Right(sSampID1, 1) = "U" Then
                            vViewSmpId = sSampID1
                            vViewSmpId2 = vbNullString
                        Else
                            vViewSmpId = sSampID2
                            vViewSmpId2 = vbNullString
                        End If

                        intZkbn = 0 'Z区分を初期化

                        '反映推定チェック関数呼び出し
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                        Call sub_Hanei(i, intNukisiRow, intZkbn)
                        Call sub_Hanei(i, intNukisiRow, intZkbn, sird1stBlockSet)           'ﾊﾟﾗﾒｰﾀ追加：結晶内SIRDｻﾝﾌﾟﾙ指示設定有無
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP END

                         intSmpkbn = 0 '初期化
                        .row = i
                        .col = 11 'Rs
                        .backColor = vbWhite
                        .text = IIf(tblWafInd(intNukisiRow).SMP.CRYINDRS = "0" Or tblWafInd(intNukisiRow).SMP.CRYINDRS = "1", "", "1")
                            If .text = "1" Then
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                            ElseIf .text = "" Then
                                tblNukishi(i).WFSMPLIDRSCW = ""
                            End If
                        .row = i + 1
                        .col = 11
                        .backColor = vbWhite

                        '保証方法変更対応
                        If tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW Then
                            .text = IIf(tblWafInd(intNukisiRow).SMP.CRYINDRS = "0" Or tblWafInd(intNukisiRow).SMP.CRYINDRS = "2", "", "2")
                        Else
                            .text = IIf(tblWafInd(intNukisiRow).SMP.CRYINDRS = "0" Or tblWafInd(intNukisiRow).SMP.CRYINDRS = "2", "", "1")
                        End If

                        If .text = "2" Then
                            tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                        ElseIf .text = "1" Then
                            tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                        ElseIf .text = "" Then
                            tblNukishi(i + 1).WFSMPLIDRSCW = ""
                        End If

                        .col = 12
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then                                                                '共通関数で反映ができないとき
                            If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                '仕様があるとき
                                If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then '両方とも実測
                                    tblNukishi(i + 1).WFINDOICW = "2"                                       '下のデータを反映にする
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                    tblNukishi(i + 1).WFRESOICW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                             End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDOICW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDOICW = "1") Then
                            intSmpkbn = 1
                        End If
                        .col = 13
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDB1 = "3" Then
                                If tblNukishi(i).WFINDB1CW = "1" And tblNukishi(i + 1).WFINDB1CW = "1" Then
                                    tblNukishi(i + 1).WFINDB1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                                    tblNukishi(i + 1).WFRESB1CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDB1CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDB1CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 14
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDB2 = "3" Then
                                If tblNukishi(i).WFINDB2CW = "1" And tblNukishi(i + 1).WFINDB2CW = "1" Then
                                    tblNukishi(i + 1).WFINDB2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB2CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDB2CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDB2CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 15
                        .row = i
                        .backColor = vbWhite

                         If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDB3 = "3" Then
                                If tblNukishi(i).WFINDB3CW = "1" And tblNukishi(i + 1).WFINDB3CW = "1" Then
                                    tblNukishi(i + 1).WFINDB3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB3CW
                                    tblNukishi(i + 1).WFRESB3CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                         ElseIf (.text = "2" And tblNukishi(i + 1).WFINDB3CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDB3CW = "1") Then
                            intSmpkbn = 1
                         End If

                        .col = 16
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDL1 = "3" Then
                                If tblNukishi(i).WFINDL1CW = "1" And tblNukishi(i + 1).WFINDL1CW = "1" Then
                                    tblNukishi(i + 1).WFINDL1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                                    tblNukishi(i + 1).WFRESL1CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDL1CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDL1CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 17
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDL2 = "3" Then
                                If tblNukishi(i).WFINDL2CW = "1" And tblNukishi(i + 1).WFINDL2CW = "1" Then
                                    tblNukishi(i + 1).WFINDL2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                                    tblNukishi(i + 1).WFRESL2CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDL2CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDL2CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 18
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDL3 = "3" Then
                                If tblNukishi(i).WFINDL3CW = "1" And tblNukishi(i + 1).WFINDL3CW = "1" Then
                                    tblNukishi(i + 1).WFINDL3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                                    tblNukishi(i + 1).WFRESL3CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDL3CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDL3CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 19
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDL4 = "3" Then
                                If tblNukishi(i).WFINDL4CW = "1" And tblNukishi(i + 1).WFINDL4CW = "1" Then
                                    tblNukishi(i + 1).WFINDL4CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                                    tblNukishi(i + 1).WFRESL4CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDL4CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDL4CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 20
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDDS = "3" Then
                                If tblNukishi(i).WFINDDSCW = "1" And tblNukishi(i + 1).WFINDDSCW = "1" Then
                                    tblNukishi(i + 1).WFINDDSCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                                    tblNukishi(i + 1).WFRESDSCW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDDSCW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDDSCW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 21
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDDZ = "3" Then
                               If tblNukishi(i).WFINDDZCW = "1" And tblNukishi(i + 1).WFINDDZCW = "1" Then
                                    tblNukishi(i + 1).WFINDDZCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                                    tblNukishi(i + 1).WFRESDZCW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                               End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDDZCW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDDZCW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 22
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDSP = "3" Then
                                If tblNukishi(i).WFINDSPCW = "1" And tblNukishi(i + 1).WFINDSPCW = "1" Then
                                    tblNukishi(i + 1).WFINDSPCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                                    tblNukishi(i + 1).WFRESSPCW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDSPCW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDSPCW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 23
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDD1 = "3" Then
                                If tblNukishi(i).WFINDDO1CW = "1" And tblNukishi(i + 1).WFINDDO1CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                                    tblNukishi(i + 1).WFRESDO1CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDDO1CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDDO1CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 24
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDD2 = "3" Then
                                If tblNukishi(i + 1).WFINDDO2CW = "1" And tblNukishi(i).WFINDDO2CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO2CW = tblNukishi(i).WFSMPLIDDO2CW
                                    tblNukishi(i + 1).WFRESDO2CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDDO2CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDDO2CW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 25
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDD3 = "3" Then
                                If tblNukishi(i).WFINDDO3CW = "1" And tblNukishi(i + 1).WFINDDO3CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                                    tblNukishi(i + 1).WFRESDO3CW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDDO3CW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDDO3CW = "1") Then
                            intSmpkbn = 1
                        End If

                        ''残存酸素追加
                        .col = 26
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDAO = "3" Then
                                If tblNukishi(i).WFINDAOICW = "1" And tblNukishi(i + 1).WFINDAOICW = "1" Then
                                    tblNukishi(i + 1).WFINDAOICW = "2"
                                    tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                                    tblNukishi(i + 1).WFRESAOICW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDAOICW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDAOICW = "1") Then
                            intSmpkbn = 1
                        End If

                        .col = 27
                        .row = i
                        .backColor = vbWhite
                        .text = IIf(tblWafInd(intNukisiRow).SMP.CRYOTHER1 = "0", "", "1")

                        If .text = "1" Then
                            tblNukishi(i).WFSMPLIDOT1CW = tblNukishi(i).REPSMPLIDCW
                        ElseIf .text = "" Then
                            tblNukishi(i).WFSMPLIDOT1CW = ""
                        End If

                        .col = 27
                        .row = i + 1
                        .backColor = vbWhite
                        .text = IIf(tblWafInd(intNukisiRow).SMP.CRYOTHER1 = "0", "", "1")

                        If .text = "1" Then
                            tblNukishi(i + 1).WFSMPLIDOT1CW = tblNukishi(i + 1).REPSMPLIDCW
                        ElseIf .text = "" Then
                            tblNukishi(i).WFSMPLIDOT1CW = ""
                        End If

                        'エピ先行評価追加対応
                        .col = 35
                        .row = i
                        .backColor = vbWhite
                        .text = IIf(tblWafInd(intNukisiRow).SMP.CRYOTHER2 = "0", "", "1")

                        If .text = "1" Then
                            tblNukishi(i).WFSMPLIDOT2CW = tblNukishi(i).REPSMPLIDCW
                        ElseIf .text = "" Then
                            tblNukishi(i).WFSMPLIDOT2CW = ""
                        End If

                        .col = 35
                        .row = i + 1
                        .backColor = vbWhite
                        .text = IIf(tblWafInd(intNukisiRow).SMP.CRYOTHER2 = "0", "", "1")

                        If .text = "1" Then
                            tblNukishi(i + 1).WFSMPLIDOT2CW = tblNukishi(i + 1).REPSMPLIDCW
                        ElseIf .text = "" Then
                            tblNukishi(i).WFSMPLIDOT2CW = ""
                        End If

                        ''GD追加
                        .col = 28
                        .row = i
                        .backColor = vbWhite

                        If .text <> "2" Then
                            If tblWafInd(intNukisiRow).SMP.CRYINDGD = "3" Then
                                If tblNukishi(i).WFINDGDCW = "1" And tblNukishi(i + 1).WFINDGDCW = "1" Then
                                    tblNukishi(i + 1).WFINDGDCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDGDCW = tblNukishi(i).WFSMPLIDGDCW
                                    tblNukishi(i + 1).WFRESGDCW = "0"
                                    tblNukishi(i + 1).WFHSGDCW = "0"
                                    .row = i + 1
                                    .text = "2"
                                Else
                                    intSmpkbn = 1
                                End If
                            End If
                        ElseIf (.text = "2" And tblNukishi(i + 1).WFINDGDCW = "2") Or (.text = "2" And tblNukishi(i + 1).WFINDGDCW = "1") Then
                            intSmpkbn = 1
                        End If

                        'エピ先行評価追加対応
                        ' VB上の制約(プロシージャ容量64k制限)のため､エピ分は別関数で処理する
                        Call sub_DispSumple_Hanei_Ep_1(i, intNukisiRow, intSmpkbn)

                        ''比抵抗で必ず実績を立てる処理を追加
                        .col = 11
                        .row = i
                        Dim intCntrs As Integer
                        If .text = "2" Then
                            'エピ先行評価追加対応
                            For intCntrs = 12 To 35
                                If intCntrs <> 27 And intCntrs <> 35 Then
                                    .col = intCntrs
                                    If .text = "1" Then
                                        .col = 11
                                        .text = "1"
                                        tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                                    End If
                                End If
                            Next intCntrs
                        End If
                        .col = 11
                        .row = i + 1
                        If .text = "2" Then
                            'エピ先行評価追加対応
                            For intCntrs = 12 To 35
                                If intCntrs <> 27 And intCntrs <> 35 Then
                                    .col = intCntrs
                                    If .text = "1" Then
                                        .col = 11
                                        .text = "1"
                                        tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                                    End If
                                End If
                            Next intCntrs
                        End If
                        .row = i

                        '実データチェック処理 実データ処理と色を塗る処理は2行まとめて行う
                          Call sub_Jitu(i + 1, blKirikaeflg, blnhflg, intZkbn)

                        '実データチェック処理
                        '色を塗る
                        Call sub_Paint(i + 1)
                   Else
                        '切り替えなしの場合
                        If Right(sSampID1, 1) = "U" Then
                            vViewSmpId = sSampID1
                            vViewSmpId2 = sSampID2
                        Else
                            vViewSmpId = sSampID2
                            vViewSmpId2 = sSampID1
                        End If

                        intZkbn = 0
                        '上品番または下品番がZだった場合の処理をする
                        '上品番がZ
                          If Trim(tblWafInd(intNukisiRow).HINUP.hinban) = "Z" Then
                            intZkbn = 1
                                If tblWafInd(intNukisiRow).HINDN.hinban = tblSXL.hinban Then  '振替えられていないときと振替えた品番を元に戻すとき
                                    intZkbn = 3
                                End If

                                udtFHin.factory = tblSXL.factory
                                udtFHin.hinban = tblSXL.hinban
                                udtFHin.mnorevno = tblSXL.REVNUM
                                udtFHin.opecond = tblSXL.opecond
                            With tblWafInd(intNukisiRow)
                            ''''''''''''''''''''''''''振替前の品番-------------------,振替後の品番
                            .SMP.CRYINDRS = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 1)
                            .SMP.CRYINDOI = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 2)
                            .SMP.CRYINDB1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 3)
                            .SMP.CRYINDB2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 4)
                            .SMP.CRYINDB3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 5)
                            .SMP.CRYINDL1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 6)
                            .SMP.CRYINDL2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 7)
                            .SMP.CRYINDL3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 8)
                            .SMP.CRYINDL4 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 9)
                            .SMP.CRYINDDS = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 10)
                            .SMP.CRYINDDZ = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 11)
                            .SMP.CRYINDSP = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 12)
                            .SMP.CRYINDD1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 13)
                            .SMP.CRYINDD2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 14)
                            .SMP.CRYINDD3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 15)
                            .SMP.CRYOTHER1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 16)  '03/05/26 後藤
                            .SMP.CRYOTHER2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 17)   '03/05/28 後藤
                            .SMP.CRYINDAO = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 18)    '残存酸素追加　03/12/15 ooba
                            .SMP.CRYINDGD = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 19)    'GD追加　05/02/18 ooba

                            'エピ先行評価追加対応
                            .SMP.EPIINDB1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 20)
                            .SMP.EPIINDB2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 21)
                            .SMP.EPIINDB3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 22)
                            .SMP.EPIINDL1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 23)
                            .SMP.EPIINDL2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 24)
                            .SMP.EPIINDL3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 25)
                            End With
                        '下品番がZ
                        ElseIf Trim(tblWafInd(intNukisiRow).HINDN.hinban) = "Z" Then
                            intZkbn = 2
                                If tblWafInd(intNukisiRow).HINUP.hinban = tblSXL.hinban Then '振替えられていないときと振替えた品番を元に戻すとき
                                    intZkbn = 4
                                End If
                                udtFHin.factory = tblSXL.factory
                                udtFHin.hinban = tblSXL.hinban
                                udtFHin.mnorevno = tblSXL.REVNUM
                                udtFHin.opecond = tblSXL.opecond
                            With tblWafInd(intNukisiRow)

                            ''''''''''''''''''''''''''振替後の品番---------------,振替前の品番
                            .SMP.CRYINDRS = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 1)
                            .SMP.CRYINDOI = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 2)
                            .SMP.CRYINDB1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 3)
                            .SMP.CRYINDB2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 4)
                            .SMP.CRYINDB3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 5)
                            .SMP.CRYINDL1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 6)
                            .SMP.CRYINDL2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 7)
                            .SMP.CRYINDL3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 8)
                            .SMP.CRYINDL4 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 9)
                            .SMP.CRYINDDS = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 10)
                            .SMP.CRYINDDZ = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 11)
                            .SMP.CRYINDSP = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 12)
                            .SMP.CRYINDD1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 13)
                            .SMP.CRYINDD2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 14)
                            .SMP.CRYINDD3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 15)
                            .SMP.CRYOTHER1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 16)  '03/05/26 後藤
                            .SMP.CRYOTHER2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 17)   '03/05/28 後藤
                            .SMP.CRYINDAO = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 18)    '残存酸素追加　03/12/15 ooba
                            .SMP.CRYINDGD = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 19)    'GD追加　05/02/18 ooba

                            'エピ先行評価追加対応
                            .SMP.EPIINDB1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 20)
                            .SMP.EPIINDB2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 21)
                            .SMP.EPIINDB3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 22)
                            .SMP.EPIINDL1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 23)
                            .SMP.EPIINDL2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 24)
                            .SMP.EPIINDL3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 25)
                            End With
                        End If

                        '実測ﾁｪｯｸのために共有(別)の抜試行をｾｯﾄする
                        If intZkbn = 0 Then
                            iBetuRow(intBcnt) = i + 1
                            intBcnt = intBcnt + 1
                        End If

                        '反映推定チェック関数呼び出し
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                        Call sub_Hanei(i, intNukisiRow, intZkbn)
                        Call sub_Hanei(i, intNukisiRow, intZkbn, sird1stBlockSet)           'ﾊﾟﾗﾒｰﾀ追加：結晶内SIRDｻﾝﾌﾟﾙ指示設定有無
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP END

                        'エピ先行評価追加対応
                        Dim skensa1(24) As String
                        Dim skensa2(24) As String
                         Call sub_Betu(i, intNukisiRow, skensa1(), skensa2(), intZkbn, blnhflg)

                        'Zのとき
                          '2
                          '1
                          '1
                          '2
                        '共有の判定
                        .row = i
                        .col = 11
                        .backColor = vbWhite

                        '変更---
                        'Rsは隣のデータの反映
                        .text = skensa1(0)
                        If intZkbn = 4 Or intZkbn = 2 Then
                            If skensa2(0) = "2" Or skensa2(0) = "1" Then  '共有可
                                .text = "1" '実測を取る
                                tblNukishi(i).WFINDRSCW = "1"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                                tblNukishi(i).WFRESRS1CW = "0" '実績無
                                'コピー
                                tblNukishi(i + 1).WFINDRSCW = "2"  '実測を立てない
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i).WFSMPLIDRSCW
                                tblNukishi(i + 1).WFRESRS1CW = tblNukishi(i).WFRESRS1CW  '実績フラグが立ってないため0
                                .row = i + 1
                                .text = "2"
                                .row = i
                            ElseIf skensa2(0) = "0" Then
                                If skensa1(0) = "0" Then
                                    .text = ""
                                    tblNukishi(i).WFINDRSCW = "0"
                                    tblNukishi(i).WFSMPLIDRSCW = ""
                                    tblNukishi(i).WFRESRS1CW = "0"
                                Else
                                    .text = "1"
                                    tblNukishi(i).WFINDRSCW = "1"
                                    tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                                    tblNukishi(i).WFRESRS1CW = "0"
                                End If
                                tblNukishi(i + 1).WFINDRSCW = "0"
                                tblNukishi(i + 1).WFSMPLIDRSCW = ""
                                tblNukishi(i + 1).WFRESRS1CW = "0"
                                .row = i + 1
                                .text = ""
                                .row = i
                            End If
                        ElseIf intZkbn = 0 Then 'Zじゃないとき
                            .text = skensa1(0)
                            If .text = "1" Then
                                tblNukishi(i).WFINDRSCW = "1"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                            ElseIf .text = "2" Then
                                tblNukishi(i).WFINDRSCW = "2"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                                tblNukishi(i).WFRESRS1CW = "0" '実績無
                            End If
                        End If

                        .col = 12
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then '共通関数で反映不可
                            If intZkbn = 4 Or intZkbn = 2 Then '
                                If tblNukishi(i + 1).WFINDOICW = "1" Then '2が入っている
                                    tblNukishi(i).WFINDOICW = "2"
                                    tblNukishi(i).WFSMPLIDOICW = ""
                                    tblNukishi(i).WFRESOICW = "1" '実績立てる
                                Else
                                    '共通関数で判定している結果を
                                    'コピー
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                    tblNukishi(i + 1).WFRESOICW = tblNukishi(i).WFRESOICW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                 '仕様があって
                                    If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then  '両方実測の場合
                                          If skensa1(1) = "2" Then                                             '仕様の厳しくないほうが反映
                                            tblNukishi(i).WFINDOICW = "2"
                                            tblNukishi(i).WFSMPLIDOICW = tblNukishi(i + 1).WFSMPLIDOICW
                                            tblNukishi(i).WFRESOICW = "0"
                                            .text = skensa1(1)
                                          End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then  '共通関数で反映可のときのサンプルIDをコピー
                            tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(1) = "0" Then
                            tblNukishi(i + 1).WFINDOICW = "0"
                            tblNukishi(i + 1).WFSMPLIDOICW = ""
                            tblNukishi(i + 1).WFRESOICW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 13
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDB1CW = "1" Then
                                    tblNukishi(i).WFINDB1CW = "2"
                                    tblNukishi(i).WFSMPLIDB1CW = ""
                                    tblNukishi(i).WFRESB1CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                                    tblNukishi(i + 1).WFRESB1CW = tblNukishi(i).WFRESB1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDB1 = "3" Then
                                    If tblNukishi(i).WFINDB1CW = "1" And tblNukishi(i + 1).WFINDB1CW = "1" Then
                                        If skensa1(2) = "2" Then
                                            tblNukishi(i).WFINDB1CW = "2"
                                            tblNukishi(i).WFSMPLIDB1CW = tblNukishi(i + 1).WFSMPLIDB1CW
                                            tblNukishi(i).WFRESB1CW = "0"
                                            .text = skensa1(2)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(2) = "0" Then
                            tblNukishi(i + 1).WFINDB1CW = "0"
                            tblNukishi(i + 1).WFSMPLIDB1CW = ""
                            tblNukishi(i + 1).WFRESB1CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                         .col = 14
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDB2CW = "1" Then
                                    tblNukishi(i).WFINDB2CW = "2"
                                    tblNukishi(i).WFSMPLIDB2CW = ""
                                    tblNukishi(i).WFRESB2CW = "1"
                                Else
                                    'コピー
                                    tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB2CW = tblNukishi(i).WFRESB2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDB2 = "3" Then
                                    If tblNukishi(i).WFINDB2CW = "1" And tblNukishi(i + 1).WFINDB2CW = "1" Then
                                        If skensa1(3) = "2" Then
                                            tblNukishi(i).WFINDB2CW = "2"
                                            tblNukishi(i).WFSMPLIDB2CW = tblNukishi(i + 1).WFSMPLIDB2CW
                                            tblNukishi(i).WFRESB2CW = "0"
                                            .text = skensa1(3)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(3) = "0" Then
                            tblNukishi(i + 1).WFINDB2CW = "0"
                            tblNukishi(i + 1).WFSMPLIDB2CW = ""
                            tblNukishi(i + 1).WFRESB2CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 15
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDB3CW = "1" Then
                                    tblNukishi(i).WFINDB3CW = "2"
                                    tblNukishi(i).WFSMPLIDB3CW = ""
                                    tblNukishi(i).WFRESB3CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB3CW
                                    tblNukishi(i + 1).WFRESB3CW = tblNukishi(i).WFRESB3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDB3 = "3" Then
                                    If tblNukishi(i).WFINDB3CW = "1" And tblNukishi(i + 1).WFINDB3CW = "1" Then
                                        If skensa1(4) = "2" Then
                                            tblNukishi(i).WFINDB3CW = "2"
                                            tblNukishi(i).WFSMPLIDB3CW = tblNukishi(i + 1).WFSMPLIDB3CW
                                            tblNukishi(i).WFRESB3CW = "0"
                                            .text = skensa1(4)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(4) = "0" Then
                            tblNukishi(i + 1).WFINDB3CW = "0"
                            tblNukishi(i + 1).WFSMPLIDB3CW = ""
                            tblNukishi(i + 1).WFRESB3CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 16
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDL1CW = "1" Then
                                    tblNukishi(i).WFINDL1CW = "2"
                                    tblNukishi(i).WFSMPLIDL1CW = ""
                                    tblNukishi(i).WFRESL1CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                                    tblNukishi(i + 1).WFRESL1CW = tblNukishi(i).WFRESL1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDL1 = "3" Then
                                    If tblNukishi(i).WFINDL1CW = "1" And tblNukishi(i + 1).WFINDL1CW = "1" Then
                                        If skensa1(5) = "2" Then
                                            tblNukishi(i).WFINDL1CW = "2"
                                            tblNukishi(i).WFSMPLIDL1CW = tblNukishi(i + 1).WFSMPLIDL1CW
                                            tblNukishi(i).WFRESL1CW = "0"
                                            .text = skensa1(5)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(5) = "0" Then
                            tblNukishi(i + 1).WFINDL1CW = "0"
                            tblNukishi(i + 1).WFSMPLIDL1CW = ""
                            tblNukishi(i + 1).WFRESL1CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 17
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDL2CW = "1" Then
                                    tblNukishi(i).WFINDL2CW = "2"
                                    tblNukishi(i).WFSMPLIDL2CW = ""
                                    tblNukishi(i).WFRESL2CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                                    tblNukishi(i + 1).WFRESL2CW = tblNukishi(i).WFRESL2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDL2 = "3" Then
                                    If tblNukishi(i).WFINDL2CW = "1" And tblNukishi(i + 1).WFINDL2CW = "1" Then
                                        If skensa1(6) = "2" Then
                                            tblNukishi(i).WFINDL2CW = "2"
                                            tblNukishi(i).WFSMPLIDL2CW = tblNukishi(i + 1).WFSMPLIDL2CW
                                            tblNukishi(i).WFRESL2CW = "0"
                                            .text = skensa1(6)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(6) = "0" Then
                            tblNukishi(i + 1).WFINDL2CW = "0"
                            tblNukishi(i + 1).WFSMPLIDL2CW = ""
                            tblNukishi(i + 1).WFRESL2CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 18
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDL3CW = "1" Then
                                    tblNukishi(i).WFINDL3CW = "2"
                                    tblNukishi(i).WFSMPLIDL3CW = ""
                                    tblNukishi(i).WFRESL3CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                                    tblNukishi(i + 1).WFRESL3CW = tblNukishi(i).WFRESL3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDL3 = "3" Then
                                    If tblNukishi(i).WFINDL3CW = "1" And tblNukishi(i + 1).WFINDL3CW = "1" Then
                                        If skensa1(7) = "2" Then
                                            tblNukishi(i).WFINDL3CW = "2"
                                            tblNukishi(i).WFSMPLIDL3CW = tblNukishi(i + 1).WFSMPLIDL3CW
                                            tblNukishi(i).WFRESL3CW = "0"
                                            .text = skensa1(7)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(7) = "0" Then
                            tblNukishi(i + 1).WFINDL3CW = "0"
                            tblNukishi(i + 1).WFSMPLIDL3CW = ""
                            tblNukishi(i + 1).WFRESL3CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 19
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDL4CW = "1" Then
                                    tblNukishi(i).WFINDL4CW = "2"
                                    tblNukishi(i).WFSMPLIDL4CW = ""
                                    tblNukishi(i).WFRESL4CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                                    tblNukishi(i + 1).WFRESL4CW = tblNukishi(i).WFRESL4CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDL4 = "3" Then
                                    If tblNukishi(i).WFINDL4CW = "1" And tblNukishi(i + 1).WFINDL4CW = "1" Then
                                        If skensa1(8) = "2" Then
                                            tblNukishi(i).WFINDL4CW = "2"
                                            tblNukishi(i).WFSMPLIDL4CW = tblNukishi(i + 1).WFSMPLIDL4CW
                                            tblNukishi(i).WFRESL4CW = "0"
                                            .text = skensa1(8)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(8) = "0" Then
                            tblNukishi(i + 1).WFINDL4CW = "0"
                            tblNukishi(i + 1).WFSMPLIDL4CW = ""
                            tblNukishi(i + 1).WFRESL4CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 20
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDDSCW = "1" Then
                                    tblNukishi(i).WFINDDSCW = "2"
                                    tblNukishi(i).WFSMPLIDDSCW = ""
                                    tblNukishi(i).WFRESDSCW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                                    tblNukishi(i + 1).WFRESDSCW = tblNukishi(i).WFRESDSCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDDS = "3" Then
                                    If tblNukishi(i).WFINDDSCW = "1" And tblNukishi(i + 1).WFINDDSCW = "1" Then
                                        If skensa1(9) = "2" Then
                                            tblNukishi(i).WFINDDSCW = "2"
                                            tblNukishi(i).WFSMPLIDDSCW = tblNukishi(i + 1).WFSMPLIDDSCW
                                            tblNukishi(i).WFRESDSCW = "0"
                                            .text = skensa1(9)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(9) = "0" Then
                            tblNukishi(i + 1).WFINDDSCW = "0"
                            tblNukishi(i + 1).WFSMPLIDDSCW = ""
                            tblNukishi(i + 1).WFRESDSCW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 21
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDDZCW = "1" Then
                                    tblNukishi(i).WFINDDZCW = "2"
                                    tblNukishi(i).WFSMPLIDDZCW = tblNukishi(i + 1).WFSMPLIDDZCW
                                    tblNukishi(i).WFRESDZCW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                                    tblNukishi(i + 1).WFRESDZCW = tblNukishi(i).WFRESDZCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDDZ = "3" Then
                                    If tblNukishi(i).WFINDDZCW = "1" And tblNukishi(i + 1).WFINDDZCW = "1" Then
                                        If skensa1(10) = "2" Then
                                            tblNukishi(i).WFINDDZCW = "2"
                                            tblNukishi(i).WFSMPLIDDZCW = tblNukishi(i + 1).WFSMPLIDDZCW
                                            tblNukishi(i).WFRESDZCW = "0"
                                            .text = skensa1(10)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(10) = "0" Then
                            tblNukishi(i + 1).WFINDDZCW = "0"
                            tblNukishi(i + 1).WFSMPLIDDZCW = ""
                            tblNukishi(i + 1).WFRESDZCW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 22
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDSPCW = "1" Then
                                    tblNukishi(i).WFINDSPCW = "2"
                                    tblNukishi(i).WFSMPLIDSPCW = ""
                                    tblNukishi(i).WFRESSPCW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                                    tblNukishi(i + 1).WFRESSPCW = tblNukishi(i).WFRESSPCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zじゃないとき
                                If tblWafInd(intNukisiRow).SMP.CRYINDSP = "3" Then
                                    If tblNukishi(i).WFINDSPCW = "1" And tblNukishi(i + 1).WFINDSPCW = "1" Then
                                        If skensa1(11) = "2" Then
                                            tblNukishi(i).WFINDSPCW = "2"
                                            tblNukishi(i).WFSMPLIDSPCW = tblNukishi(i + 1).WFSMPLIDSPCW
                                            tblNukishi(i).WFRESSPCW = "0"
                                            .text = skensa1(11)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(11) = "0" Then
                            tblNukishi(i + 1).WFINDSPCW = "0"
                            tblNukishi(i + 1).WFSMPLIDSPCW = ""
                            tblNukishi(i + 1).WFRESSPCW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 23
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDDO1CW = "1" Then
                                    tblNukishi(i).WFINDDO1CW = "2"
                                    tblNukishi(i).WFSMPLIDDO1CW = ""
                                    tblNukishi(i).WFRESDO1CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                                    tblNukishi(i + 1).WFRESDO1CW = tblNukishi(i).WFRESDO1CW
                                End If
                            ElseIf intZkbn = 0 Then
                                If tblWafInd(intNukisiRow).SMP.CRYINDD1 = "3" Then
                                    If tblNukishi(i).WFINDOT1CW = "1" And tblNukishi(i + 1).WFINDOT1CW = "1" Then
                                        If skensa1(12) = "2" Then
                                            tblNukishi(i).WFINDDO1CW = "2"
                                            tblNukishi(i).WFSMPLIDDO1CW = tblNukishi(i + 1).WFSMPLIDDO1CW
                                            tblNukishi(i).WFRESDO1CW = "0"
                                            .text = skensa1(12)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(12) = "0" Then
                            tblNukishi(i + 1).WFINDDO1CW = "0"
                            tblNukishi(i + 1).WFSMPLIDDO1CW = ""
                            tblNukishi(i + 1).WFRESDO1CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 24
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDDO2CW = "1" Then
                                    tblNukishi(i).WFINDDO2CW = "2"
                                    tblNukishi(i).WFSMPLIDDO2CW = ""
                                    tblNukishi(i).WFRESDO2CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDDO2CW = tblNukishi(i).WFSMPLIDDO2CW
                                    tblNukishi(i + 1).WFRESDO2CW = tblNukishi(i).WFRESDO2CW
                                End If
                            ElseIf intZkbn = 0 Then
                                If tblWafInd(intNukisiRow).SMP.CRYINDD2 = "3" Then
                                    If tblNukishi(i).WFINDDO2CW = "1" And tblNukishi(i + 1).WFINDDO2CW = "1" Then
                                        If skensa1(13) = "2" Then
                                            tblNukishi(i).WFINDDO2CW = "2"
                                            tblNukishi(i).WFSMPLIDDO2CW = tblNukishi(i + 1).WFSMPLIDDO2CW
                                            tblNukishi(i).WFRESDO2CW = "0"
                                            .text = skensa1(13)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDDO2CW = tblNukishi(i).WFSMPLIDDO2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(13) = "0" Then
                            tblNukishi(i + 1).WFINDDO2CW = "0"
                            tblNukishi(i + 1).WFSMPLIDDO2CW = ""
                            tblNukishi(i + 1).WFRESDO2CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 25
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDDO3CW = "1" Then
                                    tblNukishi(i).WFINDDO3CW = "2"
                                    tblNukishi(i).WFSMPLIDDO3CW = ""
                                    tblNukishi(i).WFRESDO3CW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                                    tblNukishi(i + 1).WFRESDO3CW = tblNukishi(i).WFRESDO3CW
                                End If
                            ElseIf intZkbn = 0 Then
                                If tblWafInd(intNukisiRow).SMP.CRYINDD3 = "3" Then
                                    If tblNukishi(i).WFINDDO3CW = "1" And tblNukishi(i + 1).WFINDDO3CW = "1" Then
                                        If skensa1(14) = "2" Then
                                            tblNukishi(i).WFINDDO3CW = "2"
                                            tblNukishi(i).WFSMPLIDDO3CW = tblNukishi(i + 1).WFSMPLIDDO3CW
                                            tblNukishi(i).WFRESDO3CW = "0"
                                            .text = skensa1(14)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(14) = "0" Then
                            tblNukishi(i + 1).WFINDDO3CW = "0"
                            tblNukishi(i + 1).WFSMPLIDDO3CW = ""
                            tblNukishi(i + 1).WFRESDO3CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        ''残存酸素追加
                        .col = 26
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDAOICW = "1" Then
                                    tblNukishi(i).WFINDAOICW = "2"
                                    tblNukishi(i).WFSMPLIDAOICW = ""
                                    tblNukishi(i).WFRESAOICW = "1"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                                    tblNukishi(i + 1).WFRESAOICW = tblNukishi(i).WFRESAOICW
                                End If
                            ElseIf intZkbn = 0 Then
                                If tblWafInd(intNukisiRow).SMP.CRYINDAO = "3" Then
                                    If tblNukishi(i).WFINDAOICW = "1" And tblNukishi(i + 1).WFINDAOICW = "1" Then
                                        If skensa1(15) = "2" Then
                                            tblNukishi(i).WFINDAOICW = "2"
                                            tblNukishi(i).WFSMPLIDAOICW = tblNukishi(i + 1).WFSMPLIDAOICW
                                            tblNukishi(i).WFRESAOICW = "0"
                                            .text = skensa1(15)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(15) = "0" Then
                            tblNukishi(i + 1).WFINDAOICW = "0"
                            tblNukishi(i + 1).WFSMPLIDAOICW = ""
                            tblNukishi(i + 1).WFRESAOICW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 27 'OT1(黒か白)
                        .backColor = vbWhite

                        '変更後
                        If intZkbn = 4 Or intZkbn = 2 Then
                            .text = skensa1(16)
                            If skensa1(16) = "1" Then '仕様有り
                                tblNukishi(i).WFINDOT1CW = "1"
                                tblNukishi(i).WFSMPLIDOT1CW = tblNukishi(i).REPSMPLIDCW
                                'コピー
                                tblNukishi(i + 1).WFINDOT1CW = "0"  'OTは反映はない
                                tblNukishi(i + 1).WFSMPLIDOT1CW = ""
                                tblNukishi(i + 1).WFRESOT1CW = ""
                            ElseIf .text = "" Or .text = "0" Then
                                tblNukishi(i).WFINDOT1CW = ""
                                tblNukishi(i).WFSMPLIDOT1CW = ""
                                tblNukishi(i + 1).WFINDOT1CW = ""
                                tblNukishi(i + 1).WFSMPLIDOT1CW = ""
                                .text = ""
                            End If
                        ElseIf intZkbn = 0 Then
                            .text = skensa1(16)
                                If .text = "1" Then
                                    tblNukishi(i).WFINDOT1CW = "1"
                                    tblNukishi(i).WFSMPLIDOT1CW = tblNukishi(i).REPSMPLIDCW
                                ElseIf .text = "" Then
                                    tblNukishi(i).WFINDOT1CW = ""
                                    tblNukishi(i).WFSMPLIDOT1CW = ""
                                End If
                        End If

                        '変更後
                        .col = 35
                        .backColor = vbWhite
                        If intZkbn = 4 Or intZkbn = 2 Then
                            .text = skensa1(17)
                            If skensa1(17) = "1" Then '仕様有り
                                tblNukishi(i).WFINDOT2CW = "1"
                                tblNukishi(i).WFSMPLIDOT2CW = tblNukishi(i).REPSMPLIDCW

                                'コピー
                                tblNukishi(i + 1).WFINDOT2CW = "0"
                                tblNukishi(i + 1).WFSMPLIDOT2CW = ""
                                tblNukishi(i + 1).WFRESOT2CW = ""
                            ElseIf .text = "" Or .text = "0" Then
                                tblNukishi(i).WFINDOT2CW = ""
                                tblNukishi(i).WFSMPLIDOT2CW = ""
                                tblNukishi(i + 1).WFINDOT2CW = ""
                                tblNukishi(i + 1).WFSMPLIDOT2CW = ""
                                .text = ""
                            End If
                        ElseIf intZkbn = 0 Then
                            .text = skensa1(17)
                            If .text = "1" Then
                                tblNukishi(i).WFINDOT2CW = "1"
                                tblNukishi(i).WFSMPLIDOT2CW = tblNukishi(i).REPSMPLIDCW
                            ElseIf .text = "" Then
                                tblNukishi(i).WFINDOT2CW = ""
                                tblNukishi(i).WFSMPLIDOT2CW = ""
                            End If
                        End If

                        'エピ先行評価追加対応
                        .col = 28
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 4 Or intZkbn = 2 Then
                                If tblNukishi(i + 1).WFINDGDCW = "1" Then
                                    tblNukishi(i).WFINDGDCW = "2"
                                    tblNukishi(i).WFSMPLIDGDCW = ""
                                    tblNukishi(i).WFRESGDCW = "1"
                                    tblNukishi(i).WFHSGDCW = "0"
                                Else
                                    tblNukishi(i + 1).WFSMPLIDGDCW = tblNukishi(i).WFSMPLIDGDCW
                                    tblNukishi(i + 1).WFRESGDCW = tblNukishi(i).WFRESGDCW
                                    tblNukishi(i + 1).WFHSGDCW = "0"
                                End If
                            ElseIf intZkbn = 0 Then
                                If tblWafInd(intNukisiRow).SMP.CRYINDGD = "3" Then
                                    If tblNukishi(i).WFINDGDCW = "1" And tblNukishi(i + 1).WFINDGDCW = "1" Then
                                        If skensa1(18) = "2" Then
                                            tblNukishi(i).WFINDGDCW = "2"
                                            tblNukishi(i).WFSMPLIDGDCW = tblNukishi(i + 1).WFSMPLIDGDCW
                                            tblNukishi(i).WFRESGDCW = "0"
                                            tblNukishi(i).WFHSGDCW = "0"
                                            .text = skensa1(18)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
                            tblNukishi(i + 1).WFSMPLIDGDCW = tblNukishi(i).WFSMPLIDGDCW
                            tblNukishi(i + 1).WFRESGDCW = tblNukishi(i).WFRESGDCW
                            tblNukishi(i + 1).WFHSGDCW = tblNukishi(i).WFHSGDCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(18) = "0" Then
                            tblNukishi(i + 1).WFINDGDCW = "0"
                            tblNukishi(i + 1).WFSMPLIDGDCW = ""
                            tblNukishi(i + 1).WFRESGDCW = "0"
                            tblNukishi(i + 1).WFHSGDCW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        'エピ先行評価追加対応
                        ' VB上の制約(プロシージャ容量64k制限)のため､エピ分は別関数で処理する
                        Call sub_DispSumple_Hanei_Ep_2(i, intNukisiRow, skensa1(), skensa2(), intZkbn)
                        .row = i + 1
                        .col = 11 'Rs
                        .backColor = vbWhite

                        '変更-------
                        .text = skensa2(0)
                        If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                            If skensa1(0) = "2" Or skensa1(0) = "1" Then  '共有できる
                                .text = "1"
                                tblNukishi(i + 1).WFINDRSCW = "1" '実測を取る
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW '修正
                                tblNukishi(i + 1).WFRESRS1CW = "0"
                                'コピー
                                tblNukishi(i).WFINDRSCW = "2"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i + 1).WFSMPLIDRSCW
                                tblNukishi(i).WFRESRS1CW = tblNukishi(i + 1).WFRESRS1CW
                                .row = i
                                .text = "2"
                                .row = i + 1
                            ElseIf skensa1(0) = "0" Then
                                If skensa2(0) = "0" Then
                                    .text = ""
                                    tblNukishi(i + 1).WFINDRSCW = "0"
                                    tblNukishi(i + 1).WFSMPLIDRSCW = ""
                                    tblNukishi(i + 1).WFRESRS1CW = "0"
                                Else
                                    .text = "1"
                                    tblNukishi(i + 1).WFINDRSCW = "1"
                                    tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                                    tblNukishi(i + 1).WFRESRS1CW = "0"
                                End If
                                tblNukishi(i).WFINDRSCW = "0"
                                tblNukishi(i).WFSMPLIDRSCW = ""
                                tblNukishi(i).WFRESRS1CW = "0"
                                .row = i
                                .text = ""
                                .row = i + 1
                            End If
                        ElseIf intZkbn = 0 Then 'Zではない
                            If .text = "2" Then
                                tblNukishi(i + 1).WFINDRSCW = "2"
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                               '実績フラグは立てない
                            ElseIf .text = "1" Then
                                tblNukishi(i + 1).WFINDRSCW = "1"
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                            End If
                        End If

                        '-----変更
                        .col = 12
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then '反映推定関数で反映OK以外
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDOICW = "1" Then  '1が入ってはいけない(反映のはず)
                                    tblNukishi(i + 1).WFINDOICW = "2"
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                    tblNukishi(i + 1).WFRESOICW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDOICW = tblNukishi(i + 1).WFSMPLIDOICW
                                    tblNukishi(i).WFRESOICW = tblNukishi(i + 1).WFRESOICW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                 '仕様があって
                                    If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then '両方実測の場合
                                        If skensa2(1) = "2" Then                                           '仕様の厳しくないほうが反映
                                            tblNukishi(i + 1).WFINDOICW = "2"
                                            tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                            tblNukishi(i + 1).WFRESOICW = "0"
                                            .text = skensa2(1)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDOICW = tblNukishi(i + 1).WFSMPLIDOICW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(1) = "0" Then
                            tblNukishi(i).WFINDOICW = "0"
                            tblNukishi(i).WFSMPLIDOICW = ""
                            tblNukishi(i).WFRESOICW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 13
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDB1CW = "1" Then
                                    tblNukishi(i + 1).WFINDB1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                                    tblNukishi(i + 1).WFRESB1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB1CW = tblNukishi(i + 1).WFSMPLIDB1CW
                                    tblNukishi(i).WFRESB1CW = tblNukishi(i + 1).WFRESB1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDB1 = "3" Then
                                    If tblNukishi(i).WFINDB1CW = "1" And tblNukishi(i + 1).WFINDB1CW = "1" Then
                                        If skensa2(2) = "2" Then
                                            tblNukishi(i + 1).WFINDB1CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                                            tblNukishi(i + 1).WFRESB1CW = "0"
                                            .text = skensa2(2)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDB1CW = tblNukishi(i + 1).WFSMPLIDB1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(2) = "0" Then
                            tblNukishi(i).WFINDB1CW = "0"
                            tblNukishi(i).WFSMPLIDB1CW = ""
                            tblNukishi(i).WFRESB1CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 14
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDB2CW = "1" Then
                                    tblNukishi(i + 1).WFINDB2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB2CW = tblNukishi(i + 1).WFSMPLIDB2CW
                                    tblNukishi(i).WFRESB2CW = tblNukishi(i + 1).WFRESB2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDB2 = "3" Then
                                    If tblNukishi(i).WFINDB2CW = "1" And tblNukishi(i + 1).WFINDB2CW = "1" Then
                                        If skensa2(3) = "2" Then
                                            tblNukishi(i + 1).WFINDB2CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                            tblNukishi(i + 1).WFRESB2CW = "0"
                                            .text = skensa2(3)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDB2CW = tblNukishi(i + 1).WFSMPLIDB2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(3) = "0" Then
                            tblNukishi(i).WFINDB2CW = "0"
                            tblNukishi(i).WFSMPLIDB2CW = ""
                            tblNukishi(i).WFRESB2CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 15
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDB3CW = "1" Then
                                    tblNukishi(i + 1).WFINDB3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB3CW = tblNukishi(i + 1).WFSMPLIDB3CW
                                    tblNukishi(i).WFRESB3CW = tblNukishi(i + 1).WFRESB3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDB3 = "3" Then
                                    If tblNukishi(i).WFINDB3CW = "1" And tblNukishi(i + 1).WFINDB3CW = "1" Then
                                        If skensa2(4) = "2" Then
                                            tblNukishi(i + 1).WFINDB3CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB3CW
                                            tblNukishi(i + 1).WFRESB3CW = "0"
                                            .text = skensa2(4)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDB3CW = tblNukishi(i + 1).WFSMPLIDB3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(4) = "0" Then
                            tblNukishi(i).WFINDB3CW = "0"
                            tblNukishi(i).WFSMPLIDB3CW = ""
                            tblNukishi(i).WFRESB3CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 16
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDL1CW = "1" Then
                                    tblNukishi(i + 1).WFINDL1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                                    tblNukishi(i + 1).WFRESL1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL1CW = tblNukishi(i + 1).WFSMPLIDL1CW
                                    tblNukishi(i).WFRESL1CW = tblNukishi(i + 1).WFRESL1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDL1 = "3" Then
                                    If tblNukishi(i).WFINDL1CW = "1" And tblNukishi(i + 1).WFINDL1CW = "1" Then
                                        If skensa2(5) = "2" Then
                                            tblNukishi(i + 1).WFINDL1CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                                            tblNukishi(i + 1).WFRESL1CW = "0"
                                            .text = skensa2(5)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                                tblNukishi(i).WFSMPLIDL1CW = tblNukishi(i + 1).WFSMPLIDL1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(5) = "0" Then
                            tblNukishi(i).WFINDL1CW = "0"
                            tblNukishi(i).WFSMPLIDL1CW = ""
                            tblNukishi(i).WFRESL1CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 17
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDL2CW = "1" Then
                                    tblNukishi(i + 1).WFINDL2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                                    tblNukishi(i + 1).WFRESL2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL2CW = tblNukishi(i + 1).WFSMPLIDL2CW
                                    tblNukishi(i).WFRESL2CW = tblNukishi(i + 1).WFRESL2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDL2 = "3" Then
                                    If tblNukishi(i).WFINDL2CW = "1" And tblNukishi(i + 1).WFINDL2CW = "1" Then
                                        If skensa2(6) = "2" Then
                                            tblNukishi(i + 1).WFINDL2CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                                            tblNukishi(i + 1).WFRESL2CW = "0"
                                            .text = skensa2(6)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDL2CW = tblNukishi(i + 1).WFSMPLIDL2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(6) = "0" Then
                            tblNukishi(i).WFINDL2CW = "0"
                            tblNukishi(i).WFSMPLIDL2CW = ""
                            tblNukishi(i).WFRESL2CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 18
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDL3CW = "1" Then
                                    tblNukishi(i + 1).WFINDL3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                                    tblNukishi(i + 1).WFRESL3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL3CW = tblNukishi(i + 1).WFSMPLIDL3CW
                                    tblNukishi(i).WFRESL3CW = tblNukishi(i + 1).WFRESL3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDL3 = "3" Then
                                    If tblNukishi(i).WFINDL3CW = "1" And tblNukishi(i + 1).WFINDL3CW = "1" Then
                                        If skensa2(7) = "2" Then
                                            tblNukishi(i + 1).WFINDL3CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                                            tblNukishi(i + 1).WFRESL3CW = "0"
                                            .text = skensa2(7)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDL3CW = tblNukishi(i + 1).WFSMPLIDL3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(7) = "0" Then
                            tblNukishi(i).WFINDL3CW = "0"
                            tblNukishi(i).WFSMPLIDL3CW = ""
                            tblNukishi(i).WFRESL3CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 19
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDL4CW = "1" Then
                                    tblNukishi(i + 1).WFINDL4CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                                    tblNukishi(i + 1).WFRESL4CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL4CW = tblNukishi(i + 1).WFSMPLIDL4CW
                                    tblNukishi(i).WFRESL4CW = tblNukishi(i + 1).WFRESL4CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDL4 = "3" Then
                                    If tblNukishi(i).WFINDL4CW = "1" And tblNukishi(i + 1).WFINDL4CW = "1" Then
                                        If skensa2(8) = "2" Then
                                            tblNukishi(i + 1).WFINDL4CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                                            tblNukishi(i + 1).WFRESL4CW = "0"
                                            .text = skensa2(8)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDL4CW = tblNukishi(i + 1).WFSMPLIDL4CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(8) = "0" Then
                            tblNukishi(i).WFINDL4CW = "0"
                            tblNukishi(i).WFSMPLIDL4CW = ""
                            tblNukishi(i).WFRESL4CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 20
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDDSCW = "1" Then
                                    tblNukishi(i + 1).WFINDDSCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                                    tblNukishi(i + 1).WFRESDSCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDSCW = tblNukishi(i + 1).WFSMPLIDDSCW
                                    tblNukishi(i).WFRESDSCW = tblNukishi(i + 1).WFRESDSCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDDS = "3" Then
                                    If tblNukishi(i).WFINDDSCW = "1" And tblNukishi(i + 1).WFINDDSCW = "1" Then
                                        If skensa2(9) = "2" Then
                                            tblNukishi(i + 1).WFINDDSCW = "2"
                                            tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                                            tblNukishi(i + 1).WFRESDSCW = "0"
                                            .text = skensa2(9)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDDSCW = tblNukishi(i + 1).WFSMPLIDDSCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(9) = "0" Then
                            tblNukishi(i).WFINDDSCW = "0"
                            tblNukishi(i).WFSMPLIDDSCW = ""
                            tblNukishi(i).WFRESDSCW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 21
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDDZCW = "1" Then
                                    tblNukishi(i + 1).WFINDDZCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                                    tblNukishi(i + 1).WFRESDZCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDZCW = tblNukishi(i + 1).WFSMPLIDDZCW
                                    tblNukishi(i).WFRESDZCW = tblNukishi(i + 1).WFRESDZCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDDZ = "3" Then
                                    If tblNukishi(i).WFINDDZCW = "1" And tblNukishi(i + 1).WFINDDZCW = "1" Then
                                        If skensa2(10) = "2" Then
                                            tblNukishi(i + 1).WFINDDZCW = "2"
                                            tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                                            tblNukishi(i + 1).WFRESDZCW = "0"
                                            .text = skensa2(10)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDDZCW = tblNukishi(i + 1).WFSMPLIDDZCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(10) = "0" Then
                            tblNukishi(i).WFINDDZCW = "0"
                            tblNukishi(i).WFSMPLIDDZCW = ""
                            tblNukishi(i).WFRESDZCW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 22
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDSPCW = "1" Then
                                    tblNukishi(i + 1).WFINDSPCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                                    tblNukishi(i + 1).WFRESSPCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDSPCW = tblNukishi(i + 1).WFSMPLIDSPCW
                                    tblNukishi(i).WFRESSPCW = tblNukishi(i + 1).WFRESSPCW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDSP = "3" Then
                                    If tblNukishi(i).WFINDSPCW = "1" And tblNukishi(i + 1).WFINDSPCW = "1" Then
                                        If skensa2(11) = "2" Then
                                            tblNukishi(i + 1).WFINDSPCW = "2"
                                            tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                                            tblNukishi(i + 1).WFRESSPCW = "0"
                                            .text = skensa2(11)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDSPCW = tblNukishi(i + 1).WFSMPLIDSPCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(11) = "0" Then
                            tblNukishi(i).WFINDSPCW = "0"
                            tblNukishi(i).WFSMPLIDSPCW = ""
                            tblNukishi(i).WFRESSPCW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 23
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDDO1CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                                    tblNukishi(i + 1).WFRESDO1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO1CW = tblNukishi(i + 1).WFSMPLIDDO1CW
                                    tblNukishi(i).WFRESDO1CW = tblNukishi(i + 1).WFRESDO1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDD1 = "3" Then
                                    If tblNukishi(i).WFINDOT1CW = "1" And tblNukishi(i + 1).WFINDOT1CW = "1" Then
                                        If skensa2(12) = "2" Then
                                            tblNukishi(i + 1).WFINDDO1CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                                            tblNukishi(i + 1).WFRESDO1CW = "0"
                                            .text = skensa2(12)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDDO1CW = tblNukishi(i + 1).WFSMPLIDDO1CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(12) = "0" Then
                            tblNukishi(i).WFINDDO1CW = "0"
                            tblNukishi(i).WFSMPLIDDO1CW = ""
                            tblNukishi(i).WFRESDO1CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 24
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDDO2CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDDO2CW
                                    tblNukishi(i + 1).WFRESDO2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO2CW = tblNukishi(i + 1).WFSMPLIDDO2CW
                                    tblNukishi(i).WFRESDO2CW = tblNukishi(i + 1).WFRESDO2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDD2 = "3" Then
                                    If tblNukishi(i).WFINDDO2CW = "1" And tblNukishi(i + 1).WFINDDO2CW = "1" Then
                                        If skensa2(13) = "2" Then
                                            tblNukishi(i + 1).WFINDDO2CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDDO2CW = tblNukishi(i).WFSMPLIDDO2CW
                                            tblNukishi(i + 1).WFRESDO2CW = "0"
                                            .text = skensa2(13)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDDO2CW = tblNukishi(i + 1).WFSMPLIDDO2CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(13) = "0" Then
                            tblNukishi(i).WFINDDO2CW = "0"
                            tblNukishi(i).WFSMPLIDDO2CW = ""
                            tblNukishi(i).WFRESDO2CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 25
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDDO3CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                                    tblNukishi(i + 1).WFRESDO3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO3CW = tblNukishi(i + 1).WFSMPLIDDO3CW
                                    tblNukishi(i).WFRESDO3CW = tblNukishi(i + 1).WFRESDO3CW
                                End If
                            ElseIf intZkbn = 0 Then  'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDD3 = "3" Then
                                    If tblNukishi(i).WFINDDO3CW = "1" And tblNukishi(i + 1).WFINDDO3CW = "1" Then
                                        If skensa2(14) = "2" Then
                                            tblNukishi(i + 1).WFINDDO3CW = "2"
                                            tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                                            tblNukishi(i + 1).WFRESDO3CW = "0"
                                            .text = skensa2(14)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDDO3CW = tblNukishi(i + 1).WFSMPLIDDO3CW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(14) = "0" Then
                            tblNukishi(i).WFINDDO3CW = "0"
                            tblNukishi(i).WFSMPLIDDO3CW = ""
                            tblNukishi(i).WFRESDO3CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        ''残存酸素追加
                        .col = 26
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDAOICW = "1" Then
                                    tblNukishi(i + 1).WFINDAOICW = "2"
                                    tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                                    tblNukishi(i + 1).WFRESAOICW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDAOICW = tblNukishi(i + 1).WFSMPLIDAOICW
                                    tblNukishi(i).WFRESAOICW = tblNukishi(i + 1).WFRESAOICW
                                End If
                            ElseIf intZkbn = 0 Then  'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDAO = "3" Then
                                    If tblNukishi(i).WFINDAOICW = "1" And tblNukishi(i + 1).WFINDAOICW = "1" Then
                                        If skensa2(15) = "2" Then
                                            tblNukishi(i + 1).WFINDAOICW = "2"
                                            tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                                            tblNukishi(i + 1).WFRESAOICW = "0"
                                            .text = skensa2(15)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDAOICW = tblNukishi(i + 1).WFSMPLIDAOICW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(15) = "0" Then
                            tblNukishi(i).WFINDAOICW = "0"
                            tblNukishi(i).WFSMPLIDAOICW = ""
                            tblNukishi(i).WFRESAOICW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        .col = 27 'OT1
                        .backColor = vbWhite
                        If .text <> "2" Then
                            If intZkbn = 3 Or intZkbn = 1 Then
                                .text = skensa2(16)
                                If .text = "1" Then
                                    tblNukishi(i + 1).WFINDOT1CW = "1"
                                    tblNukishi(i + 1).WFSMPLIDOT1CW = tblNukishi(i + 1).REPSMPLIDCW
                                    'コピー
                                    tblNukishi(i).WFINDOT1CW = "0"
                                    tblNukishi(i).WFSMPLIDOT1CW = ""
                                    tblNukishi(i).WFRESOT1CW = ""
                                ElseIf .text = "0" Then
                                    tblNukishi(i + 1).WFINDOT1CW = ""
                                    tblNukishi(i + 1).WFSMPLIDOT1CW = ""
                                End If
                            ElseIf intZkbn = 0 Then
                                .text = skensa2(16)
                                If .text = "2" Then
                                    tblNukishi(i + 1).WFINDOT1CW = "0"
                                    tblNukishi(i + 1).WFSMPLIDOT1CW = ""
                                ElseIf .text = "1" Then
                                    tblNukishi(i + 1).WFINDOT1CW = "1"
                                    tblNukishi(i + 1).WFSMPLIDOT1CW = tblNukishi(i + 1).REPSMPLIDCW
                                End If
                            End If
                        ElseIf skensa1(16) = "" Then
                            .text = ""
                            tblNukishi(i).WFINDOT1CW = "0"
                            tblNukishi(i).WFSMPLIDOT1CW = ""
                        End If

                        'エピ先行評価追加対応
                        .col = 35
                        .backColor = vbWhite
                        If .text <> "2" Then
                            If intZkbn = 3 Or intZkbn = 1 Then
                                .text = skensa2(17)
                                If .text = "1" Then
                                    tblNukishi(i + 1).WFINDOT2CW = "1"
                                    tblNukishi(i + 1).WFSMPLIDOT2CW = tblNukishi(i + 1).REPSMPLIDCW
                                    'コピー
                                    tblNukishi(i).WFINDOT2CW = "0"
                                    tblNukishi(i).WFSMPLIDOT2CW = ""
                                    tblNukishi(i).WFRESOT2CW = ""
                                ElseIf .text = "0" Then
                                    tblNukishi(i + 1).WFINDOT2CW = ""
                                    tblNukishi(i + 1).WFSMPLIDOT2CW = ""
                                End If
                        ElseIf intZkbn = 0 Then
                            .text = skensa2(17)
                                If .text = "2" Then
                                    tblNukishi(i + 1).WFINDOT2CW = "0"
                                    tblNukishi(i + 1).WFSMPLIDOT2CW = ""
                                ElseIf .text = "1" Then
                                    tblNukishi(i + 1).WFINDOT2CW = "1"
                                    tblNukishi(i + 1).WFSMPLIDOT2CW = tblNukishi(i + 1).REPSMPLIDCW
                                End If
                            End If
                        ElseIf skensa2(17) = "" Then
                            .text = ""
                            tblNukishi(i).WFINDOT1CW = "0"
                            tblNukishi(i).WFSMPLIDOT1CW = ""
                        End If

                        'エピ先行評価追加対応
                        .col = 28
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                                If tblNukishi(i).WFINDGDCW = "1" Then
                                    tblNukishi(i + 1).WFINDGDCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDGDCW = tblNukishi(i).WFSMPLIDGDCW
                                    tblNukishi(i + 1).WFRESGDCW = "1"
                                    tblNukishi(i + 1).WFHSGDCW = "0"
                                Else
                                    tblNukishi(i).WFSMPLIDGDCW = tblNukishi(i + 1).WFSMPLIDGDCW
                                    tblNukishi(i).WFRESGDCW = tblNukishi(i + 1).WFRESGDCW
                                    tblNukishi(i).WFHSGDCW = "0"
                                End If
                            ElseIf intZkbn = 0 Then  'Zではない
                                If tblWafInd(intNukisiRow).SMP.CRYINDGD = "3" Then
                                    If tblNukishi(i).WFINDGDCW = "1" And tblNukishi(i + 1).WFINDGDCW = "1" Then
                                        If skensa2(18) = "2" Then
                                            tblNukishi(i + 1).WFINDGDCW = "2"
                                            tblNukishi(i + 1).WFSMPLIDGDCW = tblNukishi(i).WFSMPLIDGDCW
                                            tblNukishi(i + 1).WFRESGDCW = "0"
                                            tblNukishi(i + 1).WFHSGDCW = "0"
                                            .text = skensa2(18)
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                            tblNukishi(i).WFSMPLIDGDCW = tblNukishi(i + 1).WFSMPLIDGDCW
                            tblNukishi(i).WFRESGDCW = tblNukishi(i + 1).WFRESGDCW
                            tblNukishi(i).WFHSGDCW = tblNukishi(i + 1).WFHSGDCW
                        End If

                        '保証方法変更対応
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(18) = "0" Then
                            tblNukishi(i).WFINDGDCW = "0"
                            tblNukishi(i).WFSMPLIDGDCW = ""
                            tblNukishi(i).WFRESGDCW = "0"
                            tblNukishi(i).WFHSGDCW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        'エピ先行評価追加対応
                        ' VB上の制約(プロシージャ容量64k制限)のため､エピ分は別関数で処理する
                        Call sub_DispSumple_Hanei_Ep_3(i, intNukisiRow, skensa1(), skensa2(), intZkbn)

                        ''比抵抗で必ず実績を立てる処理を追加
                        With tblNukishi(i)
                            If .WFINDRSCW = "2" Then
                                'エピ先行評価追加対応
                                If .WFINDOICW = "1" Or .WFINDB1CW = "1" Or .WFINDB2CW = "1" Or _
                                   .WFINDB3CW = "1" Or .WFINDL1CW = "1" Or .WFINDL2CW = "1" Or _
                                   .WFINDL3CW = "1" Or .WFINDL4CW = "1" Or .WFINDDSCW = "1" Or _
                                   .WFINDDZCW = "1" Or .WFINDSPCW = "1" Or .WFINDDO1CW = "1" Or _
                                   .WFINDDO2CW = "1" Or .WFINDDO3CW = "1" Or .WFINDAOICW = "1" Or _
                                   .WFINDGDCW = "1" Or _
                                   .EPINDB1CW = "1" Or .EPINDB2CW = "1" Or .EPINDB3CW = "1" Or _
                                   .EPINDL1CW = "1" Or .EPINDL2CW = "1" Or .EPINDL3CW = "1" Then
                                    .WFINDRSCW = "1"
                                    .WFSMPLIDRSCW = .REPSMPLIDCW
                                    .WFRESRS1CW = "0"
                                    sprExamine.col = 11
                                    sprExamine.row = i
                                    sprExamine.text = "1"
                                End If
                            End If
                        End With
                        With tblNukishi(i + 1)
                            If .WFINDRSCW = "2" Then
                                'エピ先行評価追加対応
                                If .WFINDOICW = "1" Or .WFINDB1CW = "1" Or .WFINDB2CW = "1" Or _
                                   .WFINDB3CW = "1" Or .WFINDL1CW = "1" Or .WFINDL2CW = "1" Or _
                                   .WFINDL3CW = "1" Or .WFINDL4CW = "1" Or .WFINDDSCW = "1" Or _
                                   .WFINDDZCW = "1" Or .WFINDSPCW = "1" Or .WFINDDO1CW = "1" Or _
                                   .WFINDDO2CW = "1" Or .WFINDDO3CW = "1" Or .WFINDAOICW = "1" Or _
                                   .WFINDGDCW = "1" Or _
                                   .EPINDB1CW = "1" Or .EPINDB2CW = "1" Or .EPINDB3CW = "1" Or _
                                   .EPINDL1CW = "1" Or .EPINDL2CW = "1" Or .EPINDL3CW = "1" Then
                                    .WFINDRSCW = "1"
                                    .WFSMPLIDRSCW = .REPSMPLIDCW
                                    .WFRESRS1CW = "0"
                                    sprExamine.col = 11
                                    sprExamine.row = i + 1
                                    sprExamine.text = "1"
                                    sprExamine.row = i
                                End If
                            End If
                        End With

                        '実データのチェック
                         Call sub_Jitu(i + 1, blKirikaeflg, blnhflg, intZkbn)

                        '色を塗る
                        Call sub_Paint(i + 1)
                        '色を塗る
                End If
                    If i Mod 2 = 0 Then
                        ReDim Preserve gtSprWfMap(i + 1)
                        gtSprWfMap(i).ADD_FLG = 2
                        gtSprWfMap(i + 1).ADD_FLG = 0
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        .GetText 39, i, vgetlotid
                        gtSprWfMap(i).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                        .GetText 39, i + 1, vGetLotId2
                        gtSprWfMap(i + 1).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vGetLotId2)

                        .GetText 4, i, vGetBlkP
                        gtSprWfMap(i).blockp = Trim(vGetBlkP)
                        gtSprWfMap(i + 1).blockp = Trim(vGetBlkP)
                        .GetText 4, i + 1, vGetBlkP
                        If vGetBlkP <> "" And vNukisiFlg <> 2 Then    '追加行は処理しない　04/06/07 ooba
                            If vgetlotid <> vGetLotId2 Then
                                gtSprWfMap(i + 1).blockp = Trim(vGetBlkP)
                            End If
                        End If
                        .GetText 4, i, vGetBlkP
                        .GetText 2, i - 1, vGetHinban
                        gtSprWfMap(i).hinban = Trim(vGetHinban)
                        .GetText 2, i + 1, vGetHinban
                        gtSprWfMap(i + 1).hinban = Trim(vGetHinban)
                        gtSprWfMap(i - 1).opecond = tblWafInd(intNukisiRow).HINUP.opecond
                        gtSprWfMap(i - 1).REVNUM = tblWafInd(intNukisiRow).HINUP.mnorevno
                        gtSprWfMap(i - 1).factory = tblWafInd(intNukisiRow).HINUP.factory
                        gtSprWfMap(i).opecond = tblWafInd(intNukisiRow).HINUP.opecond
                        gtSprWfMap(i).REVNUM = tblWafInd(intNukisiRow).HINUP.mnorevno
                        gtSprWfMap(i).factory = tblWafInd(intNukisiRow).HINUP.factory
                        gtSprWfMap(i + 1).opecond = tblWafInd(intNukisiRow).HINDN.opecond
                        gtSprWfMap(i + 1).REVNUM = tblWafInd(intNukisiRow).HINDN.mnorevno
                        gtSprWfMap(i + 1).factory = tblWafInd(intNukisiRow).HINDN.factory
                        'TBCMY011から該当データ取得
                        vSample1 = Trim(Right(sSampID1, 1))
                        vSample2 = Trim(Right(sSampID2, 1))
                        If i <= m - 2 Then
                            .GetText 4, i + 2, vNextBlkP
                            intNextBlkP = CInt(Trim(vNextBlkP))
                        End If

                        'WF枚数、サンプルID取得
                        If DBDRV_GET_WFMAP(gtSprWfMap(i).LOTID, tblSXL.SXLID, gtSprWfMap(i).blockp, _
                                            vGetBlkP, vGetIngotP, sNextIngotP, vGetBlkSeq, vGetBlkSeq2, _
                                            vSample1, vSample2, intNextBlkP, vGetWfNum) = FUNCTION_RETURN_FAILURE Then
                            lblMsg.Caption = GetMsgStr("EWFM1") '03/06/06 後藤
                            Exit Function
                        End If

                        gtSprWfMap(i).KESSYOUP = Trim(vGetIngotP)
                        gtSprWfMap(i + 1).KESSYOUP = Trim(vGetIngotP)
                        gtSprWfMap(i).BLOCKSEQ = Trim(vGetBlkSeq)
                        If IsNull(vGetBlkSeq2) = False And vGetBlkSeq2 <> "" Then
                            gtSprWfMap(i + 1).BLOCKSEQ = Trim(vGetBlkSeq2)
                        Else
                            gtSprWfMap(i + 1).BLOCKSEQ = Trim(vGetBlkSeq)
                        End If
                        If vSample1 <> vbNullString Then
                            gtSprWfMap(i).SAMPLEID = vSample1
                        Else
                            gtSprWfMap(i).SAMPLEID = tblNukishi(i).REPSMPLIDCW
                        End If
                        If vSample2 <> vbNullString Then
                            gtSprWfMap(i + 1).SAMPLEID = vSample2
                        Else
                            gtSprWfMap(i + 1).SAMPLEID = tblNukishi(i + 1).REPSMPLIDCW
                        End If
                        gtSprWfMap(i).wfnum = vGetWfNum
                        '共有、サンプルの選択可能の場合、SampleIDを別隠しカラムに保存
                        '共有の場合でもサンプルIDが代表サンプルIDのときに処理をするように変更　2003/11/11
                        If blKirikaeflg = True And intSmpkbn <> 1 Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .SetText 44, i, vSample1
                            gtSprWfMap(i).SAMPLEID = vSample1
                            .SetText 44, i, vSample1
                            gtSprWfMap(i + 1).SAMPLEID = vSample1
                        End If
                    End If
                Else
                    '区分３で抜試を伴わなかったもの
                    If i Mod 2 = 0 Then
                        ReDim Preserve gtSprWfMap(i + 1)
                        gtSprWfMap(i).ADD_FLG = 3
                        gtSprWfMap(i + 1).ADD_FLG = 3
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        .GetText 39, i, vgetlotid
                        gtSprWfMap(i).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                         .GetText 39, i + 1, vgetlotid
                        gtSprWfMap(i + 1).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                        .GetText 4, i, vGetBlkP
                        gtSprWfMap(i).blockp = Trim(vGetBlkP)

                        .GetText 4, i + 1, vGetBlkP
                        gtSprWfMap(i + 1).blockp = Trim(vGetBlkP)
                        '初期表示の結晶位置,連番は画面から　2003/04/23
                        .GetText 5, i, vGetIngotP
                        gtSprWfMap(i).KESSYOUP = Trim(vGetIngotP)

                        .GetText 5, i + 1, vGetIngotP
                        gtSprWfMap(i + 1).KESSYOUP = Trim(vGetIngotP)
                        .GetText 6, i, vGetBlkSeq
                        If vGetBlkSeq <> "" Then
                            gtSprWfMap(i).BLOCKSEQ = Trim(vGetBlkSeq)
                        End If
                        .GetText 6, i + 1, vGetBlkSeq
                        If vGetBlkSeq <> "" Then
                            gtSprWfMap(i + 1).BLOCKSEQ = Trim(vGetBlkSeq)
                        End If

                        .GetText 2, i - 1, vGetHinban
                        gtSprWfMap(i).hinban = Trim(vGetHinban)
                        .GetText 2, i + 1, vGetHinban
                        gtSprWfMap(i + 1).hinban = Trim(vGetHinban)
                        gtSprWfMap(i - 1).opecond = tblWafInd(intNukisiRow + 1).HINUP.opecond
                        gtSprWfMap(i - 1).REVNUM = tblWafInd(intNukisiRow + 1).HINUP.mnorevno
                        gtSprWfMap(i - 1).factory = tblWafInd(intNukisiRow + 1).HINUP.factory
                        gtSprWfMap(i).opecond = tblWafInd(intNukisiRow + 1).HINUP.opecond
                        gtSprWfMap(i).REVNUM = tblWafInd(intNukisiRow + 1).HINUP.mnorevno
                        gtSprWfMap(i).factory = tblWafInd(intNukisiRow + 1).HINUP.factory
                        gtSprWfMap(i + 1).opecond = tblWafInd(intNukisiRow + 1).HINUP.opecond
                        gtSprWfMap(i + 1).REVNUM = tblWafInd(intNukisiRow + 1).HINUP.mnorevno
                        gtSprWfMap(i + 1).factory = tblWafInd(intNukisiRow + 1).HINUP.factory
                        If i <= m - 2 Then
                            .GetText 4, i + 2, vNextBlkP
                            intNextBlkP = CInt(Trim(vNextBlkP))
                        End If
                    End If
                End If
            Else
                intNukisiRow = intNukisiRow + 1 '1行目＆最終行で、INDEX+1しておく
                ReDim Preserve gtSprWfMap(i)
                If vNukisiFlg = 1 Then
                    gtSprWfMap(i).ADD_FLG = 1
                Else
                    gtSprWfMap(i).ADD_FLG = 3
                End If
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .GetText 39, i, vgetlotid
                gtSprWfMap(i).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                .GetText 4, i, vGetBlkP
                gtSprWfMap(i).blockp = Trim(vGetBlkP)
                .GetText 5, i, vGetIngotP
                gtSprWfMap(i).KESSYOUP = Trim(vGetIngotP)
                .GetText 6, i, vGetBlkSeq
                gtSprWfMap(i).BLOCKSEQ = Trim(vGetBlkSeq)
                If i = 1 Then
                    .GetText 2, 1, vGetHinban
                    gtSprWfMap(i).hinban = Trim(vGetHinban)
                Else
                    .GetText 2, m - 1, vGetHinban
                    gtSprWfMap(i).hinban = Trim(vGetHinban)
                End If
                gtSprWfMap(i).opecond = tblWafInd(intNukisiRow).HINDN.opecond
                gtSprWfMap(i).REVNUM = tblWafInd(intNukisiRow).HINDN.mnorevno
                gtSprWfMap(i).factory = tblWafInd(intNukisiRow).HINDN.factory

                ''全振替時の結晶GD引継ぎ対応
                .row = i
                .col = 28
                'GDの指示がない場合
                If .text <> "1" And .text <> "2" Then
                    With tblWafInd(intNukisiRow)
                        '初期行のTOP側で品番を振替えた場合
                        If i = 1 And Trim(.HINDN.hinban) <> "Z" And _
                                    (.HINDN.hinban <> tblSXL.hinban Or _
                                    .HINDN.mnorevno <> tblSXL.REVNUM Or _
                                    .HINDN.factory <> tblSXL.factory Or _
                                    .HINDN.opecond <> tblSXL.opecond) Then

                            .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)       'GD
                            'GDの指示がある場合
                            If .SMP.CRYINDGD <> "0" And .SMP.CRYINDGD <> "2" And _
                                        IsNumeric(CpyCrySmpl.TsmplidGD) Then

                                .SMP.WFHSGD = "1"
                                sprExamine.text = CpyCrySmpl.TindGD
                                sprExamine.backColor = COLOR_CryJitsu
                                sprExamine.ForeColor = COLOR_CryJitsu
                                bMotoGDcpyFlg(1) = True
                            End If

                        '初期行のBOT側で品番を振替えた場合
                        ElseIf i = m And Trim(.HINUP.hinban) <> "Z" And _
                                    (.HINUP.hinban <> tblSXL.hinban Or _
                                    .HINUP.mnorevno <> tblSXL.REVNUM Or _
                                    .HINUP.factory <> tblSXL.factory Or _
                                    .HINUP.opecond <> tblSXL.opecond) Then

                            .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)       'GD
                            'GDの指示がある場合
                            If .SMP.CRYINDGD <> "0" And .SMP.CRYINDGD <> "1" And _
                                        IsNumeric(CpyCrySmpl.BsmplidGD) Then

                                .SMP.WFHSGD = "1"
                                sprExamine.text = CpyCrySmpl.BindGD
                                sprExamine.backColor = COLOR_CryJitsu
                                sprExamine.ForeColor = COLOR_CryJitsu
                                bMotoGDcpyFlg(2) = True
                            End If

                        End If
                    End With
                End If
            End If
        Next i

        '画面表示
        m = .MaxRows
        intNukisiRow = 0
        For i = 1 To m
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            .GetText 37, i, vNukisiFlg
            If CheckGetSampleID(i) = True Then  '3の上と偶数行
                .SetText 4, i, gtSprWfMap(i).blockp
                .SetText 5, i, gtSprWfMap(i).KESSYOUP
                .SetText 6, i, gtSprWfMap(i).BLOCKSEQ
                If vNukisiFlg = 2 Then
                    .SetText 5, i + 1, gtSprWfMap(i + 1).KESSYOUP
                    .SetText 4, i + 1, gtSprWfMap(i + 1).blockp
                    .SetText 6, i + 1, gtSprWfMap(i + 1).BLOCKSEQ
                    .SetText 7, i + 1, gtSprWfMap(i).wfnum
                End If
                If gtSprWfMap(i).SAMPLEID = vbNullString Then
                    .SetText 10, i, gsWF_SMPL_JOINT
                    .SetText 8, i, gsWF_STA_NORMAL
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .SetText 44, i, gtSprWfMap(i).SAMPLEID
                    If blKirikaeflg = True And intSmpkbn <> 1 Then  '共通のときだけ
                        If i Mod 2 = 1 Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .SetText 38, i, gtSprWfMap(i - 1).SAMPLEID
                        Else
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .SetText 38, i, gtSprWfMap(i + 1).SAMPLEID
                        End If
                    End If
                Else
                    .SetText 10, i, Right(gtSprWfMap(i).LOTID, 3) & "-" & Right(gtSprWfMap(i).SAMPLEID, 4)
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .SetText 38, i, gtSprWfMap(i).SAMPLEID
                    .SetText 8, i, gsWF_STA_SIJI
                    .SetText 44, i, gtSprWfMap(i).SAMPLEID
                End If

                blKensaLock = False '実測をチェック
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                For intLoopCnt = 11 To 35
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    If intLoopCnt <> 27 And intLoopCnt <> 35 Then
                        .col = intLoopCnt
                        .row = i + 1
                        If .text = "1" Then '実測がある
                            blKensaLock = True
                        End If
                    End If
                Next
                If blKirikaeflg = True And intSmpkbn <> 1 Then       '共通
                    .SetText 10, i + 1, gsWF_SMPL_JOINT
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                Else
                    If blKensaLock = False Then  '実測無し
                        .SetText 10, i + 1, Right(gtSprWfMap(i + 1).LOTID, 3) & "-" & Right(gtSprWfMap(i + 1).SAMPLEID, 4)
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 44, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 8, i + 1, gsWF_STA_SIJI
                    Else                            '実測有り
                        .SetText 10, i + 1, Right(gtSprWfMap(i + 1).LOTID, 3) & "-" & Right(gtSprWfMap(i + 1).SAMPLEID, 4)
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 44, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 8, i + 1, gsWF_STA_SIJI
                    End If
                End If
            End If
        Next

        'WF枚数計算
        For i = 1 To .MaxRows - 2
            If i Mod 2 = 0 Then

                '枚数0枚はブロック境界 またはZ品番
                If CInt(Trim$(gtSprWfMap(i).wfnum)) = 0 Then
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .GetText 43, i - 1, vAllNum
                    .GetText 7, intCngBlpnt, vWFNumber
                    If vWFNumber <> "" Then
                        If Not (CInt(vWFNumber) = 0) Then
                            .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                        End If
                    End If
                    intCngBlpnt = i + 1
                ElseIf Trim$(gtSprWfMap(i).hinban) = "Z" Then
                    .SetText 7, i - 1, "0"
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .GetText 43, i - 1, vAllNum
                    If i < .MaxRows Then
                        .GetText 1, i + 1, vBlockId
                        If vBlockId = "" And vAllNum <> "" Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .SetText 43, i + 1, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                            If (i = .MaxRows - 2) Then
                                .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                                intCngBlpnt = i + 1
                            End If

                        Else
                            .GetText 7, intCngBlpnt, vWFNumber
                            If Not (CInt(vWFNumber) = 0) Then
                                .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                                intCngBlpnt = i + 1
                            End If
                        End If
                    Else
                        If Not (CInt(vWFNumber) = 0) Then
                            .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                            intCngBlpnt = i + 1
                        End If
                   End If
                Else
                    If CInt(gtSprWfMap(i).wfnum) <> 0 Then  'ブロック境界への対処
                        .SetText 7, i + 1, Trim$(gtSprWfMap(i).wfnum)
                    End If
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .GetText 43, i - 1, vAllNum
                    .GetText 1, intCngBlpnt, vBlockId
                    If vBlockId = Mid(gtSprWfMap(i).LOTID, 10, 3) Then
                        If (gtSprWfMap(i).LOTID) = (gtSprWfMap(i + 1).LOTID) Then 'ブロック境界への対処    'upd 2003/04/28 hitec)matsumoto WF状態0の時は、ブロック境だけでない
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .SetText 43, i + 1, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                        End If

                        If (i = .MaxRows - 2) Then
                            .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                        End If

                    Else
                        .GetText 6, intCngBlpnt, vWFNumber
                        If Not (CInt(vWFNumber) = 0) Then
                            .SetText 7, intCngBlpnt, CInt(vAllNum) - CInt(Trim$(gtSprWfMap(i).wfnum))
                            intCngBlpnt = i + 1
                        End If
                    End If
                End If
            Else
                If 0 = i Mod 2 Then
                    If i < .MaxRows Then
                        .GetText 1, i + 1, vBlockId
                        If vBlockId <> "" Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .GetText 43, i - 1, vAllNum
                            .SetText 7, intCngBlpnt, vAllNum
                            intCngBlpnt = i + 1
                        End If

                        '次のブロックの結晶位置は後のデータ展開のため同じ位置にしておく（4/23現在）
                        .GetText 5, i, vGetIngotP
                        .SetText 5, i + 1, vGetIngotP
                    End If
                End If
            End If
        Next i

        m = .MaxRows
        For i = 1 To m
                If i = 1 Then
                .GetText 2, i, vGetHinban
                If vGetHinban = "Z" Then
                    .col = 8
                    .row = i + 1
                    .text = "欠落"
                End If
            ElseIf i = .MaxRows Then
                .GetText 2, i - 1, vGetHinban
                If vGetHinban = "Z" Then
                    .col = 8
                    .row = i - 1
                    .text = "欠落"
                End If
            Else
                If i Mod 2 = 1 And i <> .MaxRows - 1 And i <> .MaxRows Then
                    .GetText 2, i, vGetHinban
                    If vGetHinban = "Z" Then
                        .col = 8
                        .row = i
                        .text = "欠落"
                        .row = i + 1
                        .text = "欠落"
                    End If
                End If
            End If
        Next i

        vKeturaku = "欠落"
        vNull = ""
        vZERO = "0"
        For i = 1 To .MaxRows Step 2
            .GetText 2, i, vGetHinban
            If vGetHinban = "Z" Then
                .SetText 7, i, vZERO
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .GetText 37, i, vNukisiFlg
                If vNukisiFlg <> "1" Then
                    .SetText 10, i, vNull
                    For intLoopCnt = i To 1 Step -1  'add 2003/06/11 hitec)matsumoto z品番時のサンプルIDバックアップ
                        If intLoopCnt = 1 Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
                            .SetText 38, i, vGetToSamp

                            Exit For
                        End If
                        If intLoopCnt Mod 2 = 0 Then
                            .GetText 2, intLoopCnt - 1, vGetToHin
                        Else
                            .GetText 2, intLoopCnt, vGetToHin
                        End If
                        If vGetToHin <> "Z" Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
                            .SetText 38, i, vGetToSamp
                            Exit For
                        End If
                    Next
                End If

                ' 2006/08/15 Cng エピ先行評価追加対応 SAMPO)kondoh
                .GetText 37, i + 1, vNukisiFlg

                If vNukisiFlg <> "1" Then
                    .SetText 10, i + 1, vNull
                    For intLoopCnt = i + 1 To .MaxRows Step 1  'add 2003/06/11 hitec)matsumoto z品番時のサンプルIDバックアップ
                        If intLoopCnt = .MaxRows Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
'                            .SetText 38, i, vGetToSamp
                            .SetText 38, i + 1, vGetToSamp      '08/08/25 ooba
                            Exit For
                        End If
                        If intLoopCnt Mod 2 = 0 Then
                            .GetText 2, intLoopCnt - 1, vGetToHin
                        Else
                            .GetText 2, intLoopCnt, vGetToHin
                        End If
                        If vGetToHin <> "Z" Then
                            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
                            .SetText 38, i + 1, vGetToSamp
                            Exit For
                        End If
                    Next
                End If

                '初期表示行（1,最終行）のサンプルIDは消去しない
                .col = 10
                .row = i
                .ForeColor = &H8080FF
                .col = 10
                .row = i + 1
                .ForeColor = &H8080FF
                ''Z品番を赤表示にする
                .row = i + 1

                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .col = 37
                sNukishi = .text

                If sNukishi <> "3" Then
                    .BlockMode = True
                    .row = i
                    .row2 = i + 1
                    .col = 4
                    .col2 = .MaxCols
                    .backColor = &H8080FF
                    .BlockMode = False
                    .col = 4
                    .row = i + 1
                    .backColor = vbWhite

                Else
                    .row = i
                    .row2 = i + 1
                    .col = 4
                    .col2 = .MaxCols
                    .BlockMode = True
                    .backColor = &H8080FF
                    .BlockMode = False

                End If
               ''検査項目の文字を赤にする
                .row = i
                .row2 = i + 1
                .col = 11
                .col2 = .MaxCols
                .BlockMode = True
                .ForeColor = &H8080FF
                .backColor = &H8080FF
                .BlockMode = False
                If i = .MaxRows - 1 Then
                    .row = i
                    .row2 = i + 1
                    .col = 4
                    .col2 = 4
                    .BlockMode = True
                    .backColor = &H8080FF
                    .BlockMode = False
                End If
            Else
                .row = i
                .col = 3
                .backColor = &H80FF80

                'ブロックＰの列
                .row = i + 1
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .col = 37
                sNukishi = .text

                If sNukishi <> "3" Then
                    .BlockMode = True
                    .row = i
                    .row2 = i + 1
                    .col = 4
                    .col2 = 4
                    .backColor = &H80FF80
                    .BlockMode = False
                    .col = 4
                    .row = i + 1
                    .backColor = vbWhite

                Else
                    .row = i
                    .row2 = i + 1
                    .col = 4
                    .col2 = 4
                    .BlockMode = True
                    .backColor = &H80FF80
                    .BlockMode = False
                End If

                .row = i
                .row2 = i + 1
                .col = 5
                .col2 = 10
                .BlockMode = True
                .backColor = &H80FF80
                .ForeColor = vbBlack
                .BlockMode = False
            End If
        Next i

        i = 1
        intRowNum = .MaxRows
        .row = intRowNum - 1  '' 04/24 後藤
        .col = 2
        sHinban = .text
        If (intRowNum = .MaxRows) And (sHinban = "Z") Then
            .row = intRowNum
            .col = 4
            .col2 = 10
            .BlockMode = True
            .backColor = &H8080FF
'            .ForeColor = vbBlack   '05/28
            .BlockMode = False
        ElseIf (intRowNum = .MaxRows) Then
            .row = intRowNum
            .col = 4
            .col2 = 10
            .BlockMode = True
            .backColor = &H80FF80
            .ForeColor = vbBlack
            .BlockMode = False
        End If

'↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
        'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
        For i = 1 To .MaxRows Step 2
            '品番取得
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .col = 10
                'WF状態取得(TOP)
                .GetText 8, i, vTemp
                'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
                If Trim(vTemp) = gsWF_STA_SIJI_KEKKA Then
                    .row = i
                    .backColor = f_cmbc039_3.Label12.backColor
                End If
                'WF状態取得(BOTTOM)
                .GetText 8, i + 1, vTemp
                'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
                If Trim(vTemp) = gsWF_STA_SIJI_KEKKA Then
                    .row = i + 1
                    .backColor = f_cmbc039_3.Label12.backColor
                End If

            End If
        Next i
'↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------

        'add start 2003/04/28 hitec)matsumoto　検索してきたWFが重複していないかチェックし、重複している場合はエラーメッセージを表示させる----------------------
        For i = 2 To .MaxRows Step 2    '偶数行のみループする
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            .GetText 37, i, vNukisiFlg
            If vNukisiFlg = "2" Then                '追加抜試行の場合
                .GetText 6, i, vGetUpWf    'マップ位置取得
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .GetText 39, i, vGetUpBlk  'ブロックID
                .GetText 8, i - 1, vGetWfNum   '枚数
                If vGetWfNum = 0 Then
                    If vGetHinban <> "Z" Then
                        cmdF(12).Enabled = False
                        bSampFlag = False
                        lblMsg.Caption = GetMsgStr("EWFM2") '03/06/06 後藤
                        Exit Function
                    End If
                End If

                For intWfChkLoop = i + 2 To .MaxRows Step 2   '重複チェックのため、次の偶数行を検索
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    .GetText 37, intWfChkLoop, vNukisiFlg

                    If vNukisiFlg = "2" Then                '追加抜試行の場合
                        .GetText 6, intWfChkLoop, vGetDnWf    'マップ位置取得
                        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                        .GetText 39, intWfChkLoop, vGetDnBlk  'ブロックID
                        If Trim(vGetUpWf) = Trim(vGetDnWf) Then     'マップ位置が、最初に取得した値と同じだったら、
                            If vGetUpBlk = vGetDnBlk Then
                                cmdF(12).Enabled = False
                                bSampFlag = False
                                lblMsg.Caption = GetMsgStr("EWFM3") '03/06/06
                                Exit Function
                            End If
                        End If
                    End If
                Next
            End If
        Next

        .ReDraw = True

        ''Warp判定対応
        'WFﾏｯﾌﾟ上の品番情報取得
        ReDim tMapHin(0)
        m = 0
        For i = 1 To .MaxRows Step 2
            '品番取得
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                If GetLastHinban(CStr(vGetHinban), udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                    lblMsg.Caption = "品番入力エラー"
                    Exit Function
                End If

                m = m + 1
                ReDim Preserve tMapHin(m)

                tMapHin(m).HIN = udtFullHinban

                'ﾌﾞﾛｯｸID取得
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                .GetText 39, i, vBlockId
                tMapHin(m).BLOCKID = left(txtCryNum.text, 9) & CStr(vBlockId)
                'ﾌﾞﾛｯｸSEQ取得
                .GetText 6, i, vTemp
                .GetText 6, i + 1, vTemp1
                tMapHin(m).BLKSEQ_S = CInt(vTemp)
                tMapHin(m).BLKSEQ_E = CInt(vTemp1)
                '振替ﾁｪｯｸﾌﾗｸﾞ
                tMapHin(m).WARPFLG = False
                tMapHin(m).KAKUFLG = False

                'Add Start 2011/04/25 SMPK Miyata
                tMapHin(m).XTALCS = txtCryNum.text      '結晶番号
                .GetText 5, i, vTemp
                .GetText 5, i + 1, vTemp1
                tMapHin(m).INPOSCS_S = CInt(vTemp)      '結晶内位置(Start)
                tMapHin(m).INPOSCS_E = CInt(vTemp1)     '結晶内位置(End)
                'Add End   2011/04/25 SMPK Miyata

            End If
        Next i
    End With

    '' 製品仕様の表示
    cmdF(12).Enabled = False
    bSampFlag = False

    '構造体をとっておく
    ReDim tblksn(UBound(tblNukishi))
    tblKns() = tblNukishi()

    ReDim tWafk(UBound(gtSprWfMap))
    tWafk() = gtSprWfMap()

    ReDim udtTmpWafSmp(UBound(tblWafInd))
    For i = 1 To UBound(tblWafInd)
        udtTmpWafSmp(i).INPOSCW = tblWafInd(i).INGOTPOS
        udtTmpWafSmp(i).HINBCW = tblWafInd(i).HINDN.hinban
        udtTmpWafSmp(i).REVNUMCW = tblWafInd(i).HINDN.mnorevno
        udtTmpWafSmp(i).FACTORYCW = tblWafInd(i).HINDN.factory
        udtTmpWafSmp(i).OPECW = tblWafInd(i).HINDN.opecond
    Next

    ReDim tWarpMeasG(0)
    ReDim tKakuMeasG(0)
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応

    '振替チェック-------start iida 2003/09/29 →移動 cmdF_Click
    If fnc_Furikae_Check = False Then
         bSampFlag = False

        '振替ﾁｪｯｸ未実施品番のWarp/合成角度判定　06/01/12 ooba START ========================>
        For i = 1 To UBound(tMapHin)
            '振替ﾁｪｯｸ実施の確認
            tMapHinG = tMapHin(i)
            For j = 1 To 2
                If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                    m = funChkFurikaeShiyou("CW763", txtKSXLID.text, tMapHinG.HIN, _
                                            tMapHinG.HIN, intModori, sMsg, _
                                            typ_b, typ_CType, 0)

                    tMapHin(i).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                    tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                End If
            Next j
        Next i

        'Warp/合成角度情報表示
        Call WarpKakuDisp(Me)
        Exit Function
    End If

    '>>>>> MOD 2012/09/07 SETsw Marushita 10枚未満流動対応
    ' ブロック保障チェック処理(プロシージャサイズエラーのため分割)
    If fnc_BlockHCheck() < 0 Then
        Exit Function
    End If

    'Warp/合成角度情報表示　06/01/12 ooba
    Call WarpKakuDisp(Me)
    
    '振替チェック-------end iida 2003/09/29
    cmdF(12).Enabled = True
    bSampFlag = True

    sub_DispSample = True
End Function

'*******************************************************************************
'*    関数名        : sub_BlockHCheck
'*
'*    処理概要      : ブロック保障チェック
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*
'*    戻り値        : Integer
'*
'*******************************************************************************
Private Function fnc_BlockHCheck() As Integer
    Dim i            As Integer
    Dim m            As Integer
    Dim flg          As Integer
    Dim now_bid      As String
    Dim old_bid      As String
    Dim vGetHinban   As Variant
    Dim vGetWfNum    As Variant
    Dim iJCnt1       As Integer
    Dim sC_Flg       As String   'チェックフラグ取得用　　　2012/09/07 SETsw Marushita
    Dim sC_Mai       As String   'チェックWafer枚数取得用 　2012/09/07 SETsw Marushita
    Dim iRtn         As Integer  '処理確認用　2012/09/07 SETsw Marushita

    fnc_BlockHCheck = -1

'** ﾌﾞﾛｯｸ保障ﾁｪｯｸ *********************** 2008.03.20 aoyagi *************
    With sprExamine

    flg = 0
    old_bid = ""

    '全行数ループ-------
    For i = 1 To .MaxRows
        .row = i
        .col = 39
        now_bid = .text
        If now_bid <> old_bid Then  'ブロックID変り目---
            old_bid = now_bid
            '有効なSXL数--------
            iJCnt1 = sub_DispSample_SCnt02(i)
            If iJCnt1 <= 1 Then   '1ﾌﾞﾛｯｸ=1SXLなのでﾌﾞﾛｯｸ保障ﾁｪｯｸしない
                '1ﾌﾞﾛｯｸの行数--------
                iJCnt1 = sub_DispSample_SCnt03(i)
                i = i + iJCnt1 - 1  '次に進める
            Else
                'ﾌﾞﾛｯｸ保障ﾁｪｯｸ-------
                iJCnt1 = sub_DispSample_SCnt01(i)
                If iJCnt1 <= 0 Then  'ｻﾝﾌﾟﾙなし
                    flg = 1          'NG
                    Exit For
                End If
            End If
        Else
            'ﾌﾞﾛｯｸ保障ﾁｪｯｸ-------
            iJCnt1 = sub_DispSample_SCnt01(i)
            If iJCnt1 <= 0 Then  'ｻﾝﾌﾟﾙなし
                flg = 1          'NG
                Exit For
            End If
        End If
    Next i
    
    If flg = 1 Then
        lblMsg.Caption = "WFサンプル無しは分割できません。"
        cmdF(12).Enabled = False
    
        Exit Function
    End If
    
    '>>>>> ADD 2012/09/07 SETsw Marushita 10枚未満流動対応
    'チェックフラグの取得(1:エラーチェック、2:アラームチェック)
    sC_Flg = GetCodeA9Field("X", "19", "NUKIMAI", "KCODE01A9")
    'チェックWafer枚数の取得
    sC_Mai = GetCodeA9Field("X", "19", "NUKIMAI", "CTR01A9")
    '<<<<< ADD 2012/09/07 SETsw Marushita 10枚未満流動対応

'Cng Start 2012/02/23 Y.Hitomi
    '仕掛品Wafer10枚判定
    m = .MaxRows
    For i = 1 To m
        .row = i
        .col = 7

        If i Mod 2 <> 0 Then
            .GetText 2, i, vGetHinban
            .GetText 7, i, vGetWfNum
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .backColor = &H80FF80
                '>>>>> Mod 2012/09/07 SETsw Marushita 10枚以下ロット流動対応
                'Wafer枚数のチェック
                If CInt(vGetWfNum) < CInt(sC_Mai) Then
                    If Trim(sC_Flg) = "1" Then
                        .backColor = &H8080FF
                        lblMsg.Caption = "Wafer枚数が" & sC_Mai & "枚未満の為、実行できません。"
                        Exit Function
                    Else
                        If Trim(sC_Flg) = "2" Then
                            iRtn = MsgBox("Wafer枚数が" & sC_Mai & "枚未満です。実行してもよろしいですか？", vbQuestion + vbOKCancel, "確認")
                            'キャンセルの場合は処理を抜ける
                            If iRtn = vbCancel Then
                                Exit Function
                            End If
                        End If
                    End If
                End If
                'If CInt(vGetWfNum) < 11 Then
                '    .backColor = &H8080FF
                '    lblMsg.Caption = "Wafer枚数が10枚以下の為、実行できません。"
                '    Exit Function
                'End If
                '<<<<< Mod 2012/09/07 SETsw Marushita 10枚以下ロット流動対応
            End If
        End If
    Next
'''Add Start SPK 2009/09/14
'    '仕掛品Wafer0枚判定
'    m = .MaxRows
'    For i = 1 To m
'        .row = i
'        .col = 7
'
'        If i Mod 2 <> 0 Then
'            .GetText 2, i, vGetHinban
'            .GetText 7, i, vGetWfNum
'            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
'                .backColor = &H80FF80
'                If CInt(vGetWfNum) = 0 Then
'                    .backColor = &H8080FF
'
'                    lblMsg.Caption = "Wafer枚数が0枚です！"
'                    Exit Function
'                End If
'            End If
'        End If
'    Next
''Add End SPK 2009/09/14
'Cng End 2012/02/23 Y.Hitomi
    
    End With

    fnc_BlockHCheck = 0

End Function

'*******************************************************************************
'*    関数名        : sub_Hanei
'*
'*    処理概要      : 1.反映推定チェック
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    i　　       ,I  ,Integer  ,Row
'*                    j　　       ,I  ,Integer  ,tblWafind
'*                    i　　       ,I  ,Integer  ,Z区分
'*                    pSird1stBlockSet　,IO  ,Boolean  ,結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
'*
'*    戻り値        : なし
'*
'*******************************************************************************
'◆--- 2010/01/20 SIRD対応 SPK habuki REP START  (ﾊﾟﾗﾒｰﾀ追加：結晶内SIRDｻﾝﾌﾟﾙ指示設定有無)
'''Private Sub sub_Hanei(i As Integer, j As Integer, intZkbn As Integer)
Private Sub sub_Hanei(i As Integer, j As Integer, intZkbn As Integer, ByRef pSird1stBlockSet As Boolean)
'◆--- 2010/01/20 SIRD対応 SPK habuki REP END
    'i：Row、j:tblWafind
    Dim k               As Integer '検査
    Dim sKensa()        As String
    Dim sSampID()       As String
    Dim sJflg()         As String
    Dim s               As Integer
    Dim sTB             As String
    Dim udtFullHinban   As tFullHinban
    Dim sGetSmpllid1    As String
    Dim sGetSmpllid2    As String '反映なので使用しない
    Dim intHanSuiKBN    As Integer
    Dim intSmpPos       As Integer
    Dim intModori       As Integer
    Dim sKRs            As String
    Dim sKensac()       As String
    Dim sGetHS          As String        '保証ﾌﾗｸﾞ　05/02/18 ooba
    Dim sGDChk          As String
    Dim sHanFlg         As String
    
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
    Dim SirdYN          As Boolean          'SIRDｻﾝﾌﾟﾙ指示有無[True:情報あり、False:情報無し]
    Dim sirdSmpID       As String           '検査済みｻﾝﾌﾟﾙID（SIRD用）<TBCMJ022>
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END


    sGetHS = "0"        '05/02/24 ooba

    'GDﾗｲﾝﾁｪｯｸ機能追加
    sHanFlg = "0"

    CrySampleID = CpyCrySmpl        '結晶実績引継ぎﾃﾞｰﾀのｺﾋﾟｰ　05/06/13 ooba

    For s = i To i + 1 'Row
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
        ReDim sKensa(21)
        ReDim sKensac(21)
        ReDim sSampID(21)
        ReDim sJflg(21)

        With tblWafInd(j)
            sKensac(0) = .SMP.CRYINDOI
            sKensac(1) = .SMP.CRYINDB1
            sKensac(2) = .SMP.CRYINDB2
            sKensac(3) = .SMP.CRYINDB3
            sKensac(4) = .SMP.CRYINDL1
            sKensac(5) = .SMP.CRYINDL2
            sKensac(6) = .SMP.CRYINDL3
            sKensac(7) = .SMP.CRYINDL4
            sKensac(8) = .SMP.CRYINDDS
            sKensac(9) = .SMP.CRYINDDZ
            sKensac(10) = .SMP.CRYINDSP
            sKensac(11) = .SMP.CRYINDD1
            sKensac(12) = .SMP.CRYINDD2
            sKensac(13) = .SMP.CRYINDD3
            sKensac(14) = .SMP.CRYINDAO       '残存酸素追加　03/12/15 ooba
            sKensac(15) = .SMP.CRYINDGD       'GD追加　05/02/18 ooba

            '--- 2006/08/15 Add エピ先行評価追加対応
            sKensac(16) = .SMP.EPIINDB1
            sKensac(17) = .SMP.EPIINDB2
            sKensac(18) = .SMP.EPIINDB3
            sKensac(19) = .SMP.EPIINDL1
            sKensac(20) = .SMP.EPIINDL2
            sKensac(21) = .SMP.EPIINDL3

            'GDﾗｲﾝﾁｪｯｸ機能追加 2007/06/25 M.Kaga START   ---
            sGDChk = .SMP.CRYINDGD2
        End With

        If s = i Then
        '品番がZのとき
            If Trim(tblWafInd(j).HINUP.hinban) = "Z" Then
                sTB = "B"
                udtFullHinban.hinban = tblWafInd(j).HINDN.hinban
                udtFullHinban.factory = tblWafInd(j).HINDN.factory
                udtFullHinban.mnorevno = tblWafInd(j).HINDN.mnorevno
                udtFullHinban.opecond = tblWafInd(j).HINDN.opecond
                sprExamine.col = 5
                sprExamine.row = i - 1
                intSmpPos = val(sprExamine.text)
            Else
                sTB = "B"
                udtFullHinban.factory = tblWafInd(j).HINUP.factory '品番
                udtFullHinban.hinban = tblWafInd(j).HINUP.hinban
                udtFullHinban.mnorevno = tblWafInd(j).HINUP.mnorevno
                udtFullHinban.opecond = tblWafInd(j).HINUP.opecond
'               sprExamine.GetText 5, i, intSmpPos
                sprExamine.col = 5
                sprExamine.row = i
                intSmpPos = val(sprExamine.text)
            End If
        Else
            If Trim(tblWafInd(j).HINDN.hinban) = "Z" Then
                sTB = "T"
                udtFullHinban.hinban = tblWafInd(j).HINUP.hinban
                udtFullHinban.factory = tblWafInd(j).HINUP.factory
                udtFullHinban.mnorevno = tblWafInd(j).HINUP.mnorevno
                udtFullHinban.opecond = tblWafInd(j).HINUP.opecond
                sprExamine.col = 5
                sprExamine.row = i
                intSmpPos = val(sprExamine.text)
            Else
                sTB = "T"  'TB区分
                udtFullHinban.factory = tblWafInd(j).HINDN.factory
                udtFullHinban.hinban = tblWafInd(j).HINDN.hinban
                udtFullHinban.mnorevno = tblWafInd(j).HINDN.mnorevno
                udtFullHinban.opecond = tblWafInd(j).HINDN.opecond
'               sprExamine.GetText 5, i + 1, intSmpPos
                sprExamine.col = 5
                sprExamine.row = i
                intSmpPos = val(sprExamine.text)
            End If
        End If
        '1：D、2：U、3UD 2行分の判定をする

        '--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
        For k = 0 To 21
            Select Case sKensac(k)
                Case 1
                    If s = i Then
                        sKensa(k) = "0"
                    Else
                        sKensa(k) = "1"
                    End If
                Case 2
                    If s = i Then
                        sKensa(k) = "1"
                    Else
                        sKensa(k) = "0"
                    End If
                Case 3, 4
                    If s = i Then
                        sKensa(k) = "1"
                    Else
                        sKensa(k) = "1"
                    End If
                Case 0, ""
                    If s = i Then
                        sKensa(k) = ""
                    Else
                        sKensa(k) = ""
                    End If
            End Select

            If sKensac(k) <> "0" Then
                If (intZkbn = 3 And s = i) Or (intZkbn = 1 And s = i) Or (intZkbn = 4 And s = i + 1) Or (intZkbn = 2 And s = i + 1) Then   'Zが上のときと下のときは反映:2を入れる共通関数呼ばない
                   sKensa(k) = "2"
                   sJflg(k) = "1"
                   sGetHS = "0"     '05/02/18 ooba
                ElseIf sKensa(k) <> "0" Then
                    'GD追加による変更　05/02/17 ooba
                    If k >= 14 Then k = k + 2
                    '反映推定共通関数の呼び出し
                    intModori = funChkWfHanSui(tblSXL.SXLID, sTB, tblSXL.CRYNUM, udtFullHinban, intSmpPos, k + 2, SIngotP, EIngotP, intHanSuiKBN, sGetSmpllid1, sGetSmpllid2, sGetHS)
                    '戻り値(0:正常終了(反映/推定OK)、1：正常終了(反映/推定NG)、-1：入力引数値エラー、-2：それ以外のエラー)
                    If k >= 16 Then k = k - 2

                    If intModori = 1 Then  '反映NG
                        sKensa(k) = "1"
                        sSampID(k) = tblNukishi(s).REPSMPLIDCW '検査項目に代表サンプルIDを入れる
                        sGetHS = "0"    '05/02/18 ooba
                    ElseIf intModori = 0 Then
                        If intHanSuiKBN = 0 Then '反映OK
                            sKensa(k) = "2"
                            sSampID(k) = sGetSmpllid1
                            sJflg(k) = "1"
                        End If
                    ElseIf intModori <= -1 Then
                        'err
    '                   Exit For
                    End If
                End If
            End If

            With sprExamine
                .row = s
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
                If k < 15 Then .col = 12 + k Else .col = 13 + k
                'GDﾗｲﾝﾁｪｯｸの場合
                If .col = 28 Then
                    If sHanFlg = "0" Then
                        '上品番に実測
                        If Trim(sGDChk) = "2" Then
                            sKensa(k) = "1"
                        ElseIf Trim(sGDChk) = "1" Then
                            sKensa(k) = "2"
                        End If
                        .text = sKensa(k) '結果をスプレッドにセット
                        sHanFlg = "1"
                    ElseIf sHanFlg = "1" Then
                        '下品番に実測
                        If Trim(sGDChk) = "2" Then
                            sKensa(k) = "2"
                        ElseIf Trim(sGDChk) = "1" Then
                            sKensa(k) = "1"
                        End If
                        .text = sKensa(k) '結果をスプレッドにセット
                        sHanFlg = "0"
                    End If
                Else
                    '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                    .text = sKensa(k) '結果をスプレッドにセット
                    
                    If .col = 19 Then
                        '<< SIRD >>
                        If (sKensa(k) = "1") Or (sKensa(k) = "2") Or (sKensa(k) = "3") Or (sKensa(k) = "4") Then
                            '<TBCMJ022検索>
                            Call fncGetSirdSample(Trim(txtCryNum.text), SirdYN, sirdSmpID)
                            If SirdYN Then
                                '<結晶内 1st block 以降>
                                .text = "2"          '共有
                                pSird1stBlockSet = True             '結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
                            Else
                                If Not pSird1stBlockSet Then
                                    '<結晶内 1st block>
                                    .text = "1"      '取得
                                    pSird1stBlockSet = True         '結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
                                Else
                                    '<結晶内 1st block 以降>
                                    .text = "2"      '共有
                                    pSird1stBlockSet = True         '結晶内SIRDｻﾝﾌﾟﾙ指示設定有無[True:設定済み、False:未設定]
                                End If
                            End If
                        Else
'''                            .text = sKensa(k) '結果をスプレッドにセット
                            .text = ""
                        End If
                    Else
                        .text = sKensa(k) '結果をスプレッドにセット
                    End If
                    '◆--- 2010/01/20 SIRD対応 SPK habuki REP END
                End If
            End With

            With tblNukishi(s) '構造体にセット
                Select Case k
                    Case 0                                      '--------Oi
                    .WFINDOICW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDOICW = sSampID(k)
                    .WFRESOICW = sJflg(k)
                    Case 1
                    .WFINDB1CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDB1CW = sSampID(k)
                    .WFRESB1CW = sJflg(k)
                    Case 2
                    .WFINDB2CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDB2CW = sSampID(k)
                    .WFRESB2CW = sJflg(k)
                    Case 3
                    .WFINDB3CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDB3CW = sSampID(k)
                    .WFRESB3CW = sJflg(k)
                    Case 4
                    .WFINDL1CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDL1CW = sSampID(k)
                    .WFRESL1CW = sJflg(k)
                    Case 5
                    .WFINDL2CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDL2CW = sSampID(k)
                    .WFRESL2CW = sJflg(k)
                    Case 6
                    .WFINDL3CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDL3CW = sSampID(k)
                    .WFRESL3CW = sJflg(k)
                    Case 7
                    .WFINDL4CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDL4CW = sSampID(k)
                    .WFRESL4CW = sJflg(k)
                    Case 8
                    .WFINDDSCW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDDSCW = sSampID(k)
                    .WFRESDSCW = sJflg(k)
                    Case 9
                    .WFINDDZCW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDDZCW = sSampID(k)
                    .WFRESDZCW = sJflg(k)
                    Case 10
                    .WFINDSPCW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDSPCW = sSampID(k)
                    .WFRESSPCW = sJflg(k)
                    Case 11
                    .WFINDDO1CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDDO1CW = sSampID(k)
                    .WFRESDO1CW = sJflg(k)
                    Case 12
                    .WFINDDO2CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDDO2CW = sSampID(k)
                    .WFRESDO2CW = sJflg(k)
                    Case 13                                         '-------DO3
                    .WFINDDO3CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDDO3CW = sSampID(k)
                    .WFRESDO3CW = sJflg(k)
                    Case 14     ''残存酸素追加　03/12/15 ooba
                    .WFINDAOICW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDAOICW = sSampID(k)
                    .WFRESAOICW = sJflg(k)
                    Case 15     'GD追加　05/02/18 ooba
                    .WFINDGDCW = IIf(sKensa(k) = "", "0", sKensa(k))    '状態FLG(GD)
'                    .WFSMPLIDGDCW = sSampID(k)                          'ｻﾝﾌﾟﾙID(GD)
                    '09/07/22 Y.Hitomi GDｻﾝﾌﾟﾙ指示不具合対応→GDのみ,WFｻﾝﾌﾟﾙ指示有り時,結晶ｻﾝﾌﾟﾙ反映NG
                        If .WFINDGDCW <> "1" Then
                            .WFSMPLIDGDCW = sSampID(k)                 'ｻﾝﾌﾟﾙID(GD)
                        ElseIf .WFINDGDCW = "1" Then
                            .WFSMPLIDGDCW = tblNukishi(s).REPSMPLIDCW '検査項目に代表サンプルIDを入れる
                        End If
                    
                    .WFRESGDCW = sJflg(k)                               '実績FLG(GD)
                    .WFHSGDCW = sGetHS                                  '保証FLG(GD)

                    'エピ先行評価追加対応
                    Case 16
                    .EPINDB1CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDB1CW = sSampID(k)
                    .EPRESB1CW = sJflg(k)
                    Case 17
                    .EPINDB2CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDB2CW = sSampID(k)
                    .EPRESB2CW = sJflg(k)
                    Case 18
                    .EPINDB3CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDB3CW = sSampID(k)
                    .EPRESB3CW = sJflg(k)
                    Case 19
                    .EPINDL1CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDL1CW = sSampID(k)
                    .EPRESL1CW = sJflg(k)
                    Case 20
                    .EPINDL2CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDL2CW = sSampID(k)
                    .EPRESL2CW = sJflg(k)
                    Case 21
                    .EPINDL3CW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .EPSMPLIDL3CW = sSampID(k)
                    .EPRESL3CW = sJflg(k)
                End Select
            End With
        Next k
    Next s
End Sub

'*******************************************************************************
'*    関数名        : sub_Betu
'*
'*    処理概要      : 1.共有別の場合の設定する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    i　　       ,I  ,Integer  ,Row
'*                    j　　       ,I  ,Integer  ,tblWafind
'*                    sKensa1　 ,I  ,String   ,検査用
'*                    sKensa2　 ,I  ,String　 ,検査用
'*                    intZkbn　　 ,I  ,Integer  ,Z区分
'*                    blnhflg　　 ,I  ,Boolean  ,反映判定フラグ
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_Betu(i As Integer, j As Integer, skensa1() As String, skensa2() As String, intZkbn As Integer, blnhflg As Boolean)
    Dim k           As Integer '検査
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
    Dim sKensa(24)  As String
    Dim k1          As Integer
    Dim k2          As Integer

    With tblWafInd(j)
        sKensa(0) = .SMP.CRYINDRS
        sKensa(1) = .SMP.CRYINDOI
        sKensa(2) = .SMP.CRYINDB1
        sKensa(3) = .SMP.CRYINDB2
        sKensa(4) = .SMP.CRYINDB3
        sKensa(5) = .SMP.CRYINDL1
        sKensa(6) = .SMP.CRYINDL2
        sKensa(7) = .SMP.CRYINDL3
        sKensa(8) = .SMP.CRYINDL4
        sKensa(9) = .SMP.CRYINDDS
        sKensa(10) = .SMP.CRYINDDZ
        sKensa(11) = .SMP.CRYINDSP
        sKensa(12) = .SMP.CRYINDD1
        sKensa(13) = .SMP.CRYINDD2
        sKensa(14) = .SMP.CRYINDD3
        sKensa(15) = .SMP.CRYINDAO        '残存酸素追加　03/12/15 ooba
        sKensa(16) = .SMP.CRYOTHER1
        sKensa(17) = .SMP.CRYOTHER2
        sKensa(18) = .SMP.CRYINDGD        'GD追加　05/02/18 ooba

        'エピ先行評価追加対応
        sKensa(19) = .SMP.EPIINDB1        'BMD1E
        sKensa(20) = .SMP.EPIINDB2        'BMD2E
        sKensa(21) = .SMP.EPIINDB3        'BMD3E
        sKensa(22) = .SMP.EPIINDL1        'OSF1E
        sKensa(23) = .SMP.EPIINDL2        'OSF2E
        sKensa(24) = .SMP.EPIINDL3        'OSF3E
    End With

'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
    For k = 0 To 24
        Select Case sKensa(k)
            Case 1 'Top
               If intZkbn = 2 Or intZkbn = 4 Then
                   '下品番Zなら検査なし　04/04/14 ooba
                   skensa1(k) = "0"
                   skensa2(k) = "0"
               Else
                   skensa1(k) = "0"
                   skensa2(k) = "1"
                   k2 = k2 + 1
               End If
            Case 2 'Tail
               If intZkbn = 1 Or intZkbn = 3 Then
                   '上品番Zなら検査なし　04/04/14 ooba
                   skensa1(k) = "0"
                   skensa2(k) = "0"
               Else
                   skensa1(k) = "1"
                   skensa2(k) = "0"
                   k1 = k1 + 1
               End If
            Case 3 '共通
               If intZkbn = 1 Then
                   skensa1(k) = "2"
                   skensa2(k) = "1"
               ElseIf intZkbn = 2 Then
                   skensa1(k) = "1"
                   skensa2(k) = "2"
               ElseIf intZkbn = 3 Then
                   skensa1(k) = "2"
                   skensa2(k) = "1"
               ElseIf intZkbn = 4 Then
                   skensa1(k) = "1"
                   skensa2(k) = "2"
               Else
                   skensa1(k) = "1"
                   skensa2(k) = "1"
               End If

            '↓追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            '2.1.2 共有チェック追加
            '抵抗のcase 4 の時は無条件で実績を立てる
            Case 4 '両方検査(別)
'Cng Start 2011/11/02 Y.Hitomi
                ''抵抗の時だけ処理を行う
                If k = 0 Then
                    If intZkbn = 1 Then
                        skensa1(k) = "2"
                        skensa2(k) = "1"
                    ElseIf intZkbn = 2 Then
                        skensa1(k) = "1"
                        skensa2(k) = "2"
                    ElseIf intZkbn = 3 Then
                        skensa1(k) = "2"
                        skensa2(k) = "1"
                    ElseIf intZkbn = 4 Then
                        skensa1(k) = "1"
                        skensa2(k) = "2"
                    Else
                    skensa1(k) = "1"
                    skensa2(k) = "1"
                End If
                End If
'                If k = 0 Then
'                    skensa1(k) = "1"
'                    skensa2(k) = "1"
'                End If
'Cng End 2011/11/02 Y.Hitomi
            '↑追加 熱処理判断処理追加 2006/02/15 SMP石川 ---------------
            Case 0  '無し
               skensa1(k) = ""
               skensa2(k) = ""
        End Select
    Next k

    If intZkbn = 0 Then 'Zではないとき
        blnhflg = False

        'エピ先行評価追加対応 SMP)kondoh
        For k = 0 To 24
            If skensa1(k) = "1" And skensa2(k) = "1" Then '共有のときだけ2にする
                If k1 < k2 Then '検査項目が少ないほうが反映となる
                    skensa1(k) = IIf(skensa1(k) = "1", "2", skensa1(k))
                    blnhflg = False '上共有

                Else
                    skensa2(k) = IIf(skensa2(k) = "1", "2", skensa2(k))
                    blnhflg = True  '下共有
                End If
            End If
        Next k
    End If
End Sub

'*******************************************************************************
'*    関数名        : sub_Paint
'*
'*    処理概要      : 1.検査項目のスプレッドの状態により色分けをする
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    intRow        ,I  ,Integer  ,Row
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_Paint(intRow As Integer)
    Dim sTval       As String
    Dim sKval       As String
    Dim intKensa    As Integer
    Dim i           As Integer '列
    Dim k           As Integer '行---2行を色塗る

    'スプレッドの検査項目(""：白、1：黒、2：黄色)
    intKensa = 0
        With sprExamine
        For k = intRow - 1 To intRow
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            For i = 11 To 35
                If i <> 28 Then
                    .col = i
                    .row = k
                    sKval = .text
                    If sKval = "" Then
                        .backColor = vbWhite
                        .ForeColor = vbWhite
                        .Lock = True
                    ElseIf sKval = "1" Then
                        .backColor = vbBlack
                        .ForeColor = vbBlack
                        .Lock = True
                    ElseIf sKval = "2" Then
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                        .backColor = vbYellow
'''                        .ForeColor = vbYellow
'''                        .Lock = False
                        
                        If i = 19 Then
                            '<SIRD>・・・ｸﾞﾚｲ
                            .backColor = COLOR_CryJitsu
                            .ForeColor = COLOR_CryJitsu
                            .Lock = False
                        Else
                            '<SIRD以外>・・・黄色
                            .backColor = vbYellow
                            .ForeColor = vbYellow
                            .Lock = False
                        End If
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
                    End If
                End If
            Next i

            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            .col = 28
            .row = k
            sKval = .text
            '指示無の場合は白表示
            If sKval = "" Then
                .backColor = vbWhite
                .ForeColor = vbWhite
                .Lock = True
            '実測の場合は黒表示
            ElseIf sKval = "1" Then
                .backColor = vbBlack
                .ForeColor = vbBlack
                .Lock = True
            '反映の場合は黄色表示
            ElseIf sKval = "2" And tblNukishi(k).WFHSGDCW <> "1" Then
                .backColor = vbYellow
                .ForeColor = vbYellow
                .Lock = False
            '結晶実績の場合はｸﾞﾚｰ表示
            ElseIf sKval = "2" And tblNukishi(k).WFHSGDCW = "1" Then
                .backColor = COLOR_CryJitsu
                .ForeColor = COLOR_CryJitsu
                .Lock = False
            End If
        Next k
    End With
End Sub

'*******************************************************************************************
'*    関数名        : sub_Jitu
'*
'*    処理概要      : 1.検査を行って反映となった行で実データが1つもない行を探す(共有は除く)
'*                      (実データが存在するかのチェックをする)
'*                    2.実データ無しの場合その行の抵抗(Rs)に実データを立てる
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    k           ,I  ,Integer  ,Row
'*                    blKirikaeflg ,I  ,Boolean  ,切替えフラグ
'*                    blnhflg　　 ,I  ,Boolean  ,反映判定フラグ
'*                    intZkbn　　 ,I  ,Integer  ,Z区分
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_Jitu(k As Integer, blKirikaeflg As Boolean, blnhflg As Boolean, intZkbn As Integer)
    Dim sTval       As String
    Dim sKval       As String
    Dim intKensa    As Integer
    Dim j           As Integer
    Dim i           As Integer
    Dim blTwoFlg    As Boolean     '2004/01/29 ooba

    '①検査を行って反映となった行で実データが1つもない行を探す(共有は除く)
    '②実データ無しの場合その行の抵抗(Rs)に実データを立てる
    '実データのチェック-------start iida 2003/09/12

    For j = k - 1 To k
        blTwoFlg = False     '2004/01/29 ooba

        If blKirikaeflg = True Then
            If j = k Then
                Exit For
            End If
        'Zのとき
        ElseIf intZkbn = 3 Or intZkbn = 4 Or intZkbn = 1 Or intZkbn = 2 Then
            If (intZkbn = 3 And j = k - 1) Or (intZkbn = 1 And j = k - 1) Then
                j = j + 1  '上に実測立てない
            ElseIf (intZkbn = 4 And j = k) Or (intZkbn = 4 And j = k) Then
                Exit For   '下に実測立てない
            End If
        '共有別のとき
        Else
            intKensa = 0    '2004/01/28 ooba
            blTwoFlg = True  '2004/01/29 ooba
        End If

        With sprExamine
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                For i = 11 To 35
                    .col = i
                    .row = j
                    sKval = .text

                    If sKval = "1" Then   '検査無し(実測)
                        intKensa = intKensa + 1
                    End If
                    '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                    If i = 35 Then
                        If intKensa = 0 Then
                            .col = 11
                            sKval = .text             '2004/01/28 ooba
                            If sKval = "2" Then       '2004/01/28 ooba
                                .text = "1"
                                .Lock = True
                                'サンプルIDを代表サンプルIDにする
'                                tblNukishi(k).WFSMPLIDRS1CW = tblNukishi(k).REPSMPLIDCW
                                If blTwoFlg = False Then
                                    If j = k Then
                                        tblNukishi(j).WFSMPLIDRSCW = tblNukishi(j).REPSMPLIDCW
                                        tblNukishi(k - 1).WFSMPLIDRSCW = tblNukishi(j).REPSMPLIDCW
                                    Else
                                        tblNukishi(j).WFSMPLIDRSCW = tblNukishi(j).REPSMPLIDCW
                                        tblNukishi(k).WFSMPLIDRSCW = tblNukishi(j).REPSMPLIDCW
                                    End If
                                Else
                                    '2行ﾁｪｯｸする場合は該当行のみ変更　2004/01/29 ooba
                                    tblNukishi(j).WFSMPLIDRSCW = tblNukishi(j).REPSMPLIDCW
                                End If
                            End If
                        End If
                    End If
                Next i
'            End If
        End With
    Next j
End Sub

'*******************************************************************************
'*    関数名        : sub_DrawImage
'*
'*    処理概要      : 1.結晶イメージウィンドウを表示する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_DrawImage()
    Dim Xl              As c_cmzcXl
    Dim m               As Integer
    Dim i               As Integer
    Dim blAns           As Boolean
    Dim vTemp           As Variant
    Dim vTemp1          As Variant
    Dim vTemp2          As String
    Dim c0              As Long
    Dim udtFullHinban   As tFullHinban
    Dim vNukisiFlg      As Variant
    Dim intCnt          As Integer
    Dim intRow          As Integer
    Dim intRowNum       As Integer
    Dim sHinban         As String
    Dim sNukishi        As String
    Dim intCol          As Integer
    Dim sNum            As String
    Dim vGetHinban      As Variant

    lblMsg.Caption = ""

'>>>>> 頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota ---------------------
'    If Mid(tblSXL.CRYNUM, 1, 1) = "8" Then
'        '' エラーメッセージを表示する
'        lblMsg.Caption = GetMsgStr("EKDE1")
''        Screen.MousePointer = 1
'        Exit Sub
'    End If
'<<<<< 頭8を購入単結晶扱いしない 2007/10/10 SETsw kubota ---------------------

'2001/08/30 S.Sano Start
    '' 品番チェック
    For c0 = 1 To sprExamine.MaxRows
        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg

        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        sprExamine.GetText 40, c0, vTemp
        vTemp = Trim(vTemp)
        If Trim$(vTemp) <> "Z" Then
            Select Case ChkString(vTemp, 8, 8)
            Case CHK_NG, CHK_NULL
                lblMsg.Caption = GetMsgStr(EHIN1)
                Exit Sub
            End Select
            vTemp2 = vTemp
            If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                lblMsg.Caption = GetMsgStr(EHIN0)
                Exit Sub
            End If
        End If
    Next c0

    '' 抜試指示一覧の入力チェック（WFﾏｯﾌﾟ）
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

'2003/03/18 hitec)matsumoto 品番のＮＵＬＬチェックはすでにはじいてある。ここでＮＵＬＬをはじいているのは、画面レイアウト変更により、ＮＵＬＬ行(品番入力不可)が存在するため。
'2001/08/30 S.Sano Start
    '' 品番チェック
    For c0 = 1 To sprExamine.MaxRows - 1
        '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg
        If (vNukisiFlg = "1") Or (vNukisiFlg = "2") Then    '抜試行だったら
            If (vNukisiFlg = "1") Then
                '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
                sprExamine.GetText 37, c0 + 1, vNukisiFlg
                If vNukisiFlg = "1" Then
                    '追加抜試の無い行は何もしない
                ElseIf vNukisiFlg = "2" Then
                    sprExamine.GetText 2, c0, vTemp             '各行に品番が書いてあるわけではない、ダミーでは全行に品番を持っているのでそれを見る
                    vTemp = Trim(vTemp)
                    If Trim$(vTemp) <> "Z" Then
                        Select Case ChkString(vTemp, 8, 8)
                        Case CHK_NG, CHK_NULL
                            lblMsg.Caption = GetMsgStr(EHIN1)
                            Exit Sub
                        End Select
                        vTemp2 = vTemp
                        If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                            lblMsg.Caption = GetMsgStr(EHIN0)
                            Exit Sub
                        End If
                    End If
                End If
            Else
                sprExamine.GetText 2, c0 + 1, vTemp
                vTemp = Trim(vTemp)
                If Trim$(vTemp) <> "Z" Then
                    Select Case ChkString(vTemp, 8, 8)
                    Case CHK_NG, CHK_NULL
                        lblMsg.Caption = GetMsgStr(EHIN1)
                        Exit Sub
                    End Select
                    vTemp2 = vTemp
                    If GetLastHinban(vTemp2, udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                        lblMsg.Caption = GetMsgStr(EHIN0)
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next c0

    '' 結晶クラスからブロック情報を取得するためにGetxl関数呼び出し
    Set Xl = GetXl(tblSXL.CRYNUM, "f_cmbc039_3")

    With Xl
        Call sub_MakeTBCME042
        '' 再抜試位置
        m = UBound(tblWafInd) - 1
        For i = 1 To m
            '--- 2006/08/15 Cng エピ先行評価追加対応 SAMPO)sekine
            sprExamine.GetText 37, i, vNukisiFlg
            If i = 1 Then
                blAns = sprExamine.GetText(2, i, vTemp)
                Xl.WfSmps(CStr(tblSXL.WFSMP(1).INPOSCW) & tblSXL.WFSMP(1).SMPKBNCW).hinban = Trim(vTemp)
            Else
                .AddWfSample tblWafInd(i).INGOTPOS, tblWafInd(i).HINDN.hinban
            End If
        Next i
        .GenerateSxl SIngotP, EIngotP
    End With

    '' 結晶図を描画する
    f_cmzc003a.Draw Xl  '' 結晶の情報を描画する
    f_cmzc003a.Show     '' 結晶図ウィンドウを表示する
End Sub

'*******************************************************************************
'*    関数名        : sub_InitData
'*
'*    処理概要      : 1.データ初期設定
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_InitData()

    Dim Blk     As c_cmzcBlk
    Dim blFlag  As Boolean
    Dim intSP   As Integer
    Dim intEP   As Integer
    Dim intSBP  As Integer
    Dim intEBP  As Integer
    Dim m       As Integer
    Dim i       As Integer

    '' 前画面からのデータ受け取り
    tblTotal = typ_CType
    tblSXL = typ_CType.typ_Param

    '' 対象ブロックID表示用の構造体を取得する
    '' 行挿入した際のブロックIDの選択コンボボックスに使用する
    Set orgXl = GetXl(tblSXL.CRYNUM, Me.Name)

    '' 結晶クラスからブロック内容を取得し、SXLの対象ブロックかを判別
    m = orgXl.Blks.COUNT
    ReDim SxlIntoBlock(m)
    i = 0
    For Each Blk In orgXl.Blks
        intSP = tblSXL.INGOTPOS
        intEP = intSP + tblSXL.COUNT
        intSBP = Blk.INGOTPOS
        intEBP = intSBP + Blk.LENGTH
        blFlag = False

        '' ブロックがSXLの中に完全に含まれている場合
        If intSP <= intSBP And intEP >= intEBP Then
            blFlag = True
        '' ブロックがSXLの開始位置より上にあり、かつ終端位置よりも長い場合
        ElseIf intSP >= intSBP And intEP <= intEBP Then
            blFlag = True
        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが上側。ただしブロックの終端とSXLの開始位置が一致しないこと)
        ElseIf intSP > intSBP And intSP < intEBP And intSP <> intEBP Then
            blFlag = True
        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが下側。ただしSXLの終端とブロックの開始位置が一致しないこと)
        ElseIf intSP < intSBP And intEP > intSBP And intEP <> intSBP Then
            blFlag = True
        End If
        If blFlag = True Then
            i = i + 1
            SxlIntoBlock(i).SORTID = Right(Blk.BLOCKID, 3)
            SxlIntoBlock(i).FULLID = Blk.BLOCKID
        End If
    Next
    ReDim Preserve SxlIntoBlock(i)

    '' ブロック管理テーブルを構造体に設定
    m = orgXl.Blks.COUNT
    ReDim tblBlkInf(m)
    i = 1
    For Each Blk In orgXl.Blks
        With tblBlkInf(i)
            .COF.TOPSMPLPOS = Blk.INGOTPOS
            .LENGTH = Blk.LENGTH
            .REALLEN = Blk.REALLEN
            .BLOCKID = Blk.BLOCKID
            .KRPROCCD = Blk.KRPROCCD
            .NOWPROC = Blk.NOWPROC
            .LPKRPROCCD = Blk.LPKRPROCCD
            .LASTPASS = Blk.LASTPASS
            .RSTATCLS = Blk.RSTATCLS
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .SAMPFLAG = False
        End With
        i = i + 1
    Next

    '' 区分コードの取得
    Call GetCodeListSC18("SC", "18", "WF", tblPrcList)
End Sub

'*******************************************************************************
'*    関数名        : sub_LoadAndDisp
'*
'*    処理概要      : 1.データのロードと表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_LoadAndDisp()

    Dim m       As Integer
    Dim n       As Integer
    Dim i       As Integer
    Dim sMsg    As String
    Dim vt      As Variant
'>>>>> add start 2011/06/30 Marushita
    Dim iMinMidCnt      As Integer       '中間抜試の必要数
    Dim iRstMidCnt      As Integer       '中間抜試の件数
'<<<<< add start 2011/06/30 Marushita

    '' 前画面から引き渡された値を表示する
    txtStaffID.text = tblTotal.StrStaffId                                   ' 担当者コード
    txtJfName.text = tblTotal.strStaffName                                  ' 担当者名
    txtKSXLID.text = tblSXL.SXLID                                           ' 仮SXLID
    txtCryNum.text = tblSXL.CRYNUM                                          ' インゴットID
    If sKanrenFlg = "1" Then lblKanren.Visible = True       '関連ﾌﾞﾛｯｸ表示　08/01/31 ooba
    txtTopRsltR.text = toRsStr(CDbl(tblTotal.typ_y013(1, WFRES).MESDATA5))  ' T側実績ρ
    txtBotRsltR.text = toRsStr(CDbl(tblTotal.typ_y013(2, WFRES).MESDATA5))  ' B側実績ρ

    If fnc_GetMukesaki_XSDCB(Trim(txtKSXLID.text)) = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' 処理時間セット
    SetPresentTime lblTime

    InMaxRow = 2

    '' 結晶に含まれる品番すべてを取得
    If GetXlHinban(tblSXL.CRYNUM, tblHinNum) = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' 品番の設定
    ReDim tblHinbanRs(1)
    With tblHinbanRs(1)
        .CRYNUM = tblSXL.CRYNUM
        .HIN.hinban = tblSXL.hinban
        .HIN.mnorevno = tblSXL.REVNUM
        .HIN.factory = tblSXL.factory
        .HIN.opecond = tblSXL.opecond
    End With

    '' DBから関連データを読み込む
    If fnc_LoadData = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' サンプルのないSXLだった場合、画面の処理はなにも行えないものとする
    If Trim$(tblSXL.WFSMP(1).XTALCW) <> "" And _
       Trim$(tblSXL.WFSMP(2).XTALCW) <> "" Then
        '' 抜試指示の表示
        If fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE Then
            Exit Sub
        End If

        '' 製品仕様の表示
        If fnc_DispHinSpec(0) = False Then
            Exit Sub
        End If

        '品番を1列追加したことによる列の変更
        With sprExamine
            m = .MaxRows
            n = .MaxCols

            '' 対象SXLの結晶位置を保持
            .row = 1
            SIngotP = tblsmp(1).INGOTPOS    'トップ位置

            'エピ先行評価追加対応
            .SetText 42, 1, tblsmp(1).INGOTPOS
'            '' 対象SXLの結晶位置を保持
            .row = m
            EIngotP = tblsmp(2).INGOTPOS    'ボトム位置

            'エピ先行評価追加対応
            .SetText 42, m, tblsmp(2).INGOTPOS
        End With

        '抜試ﾃﾞｰﾀ不正表示チェック
        If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr(sMsg)
            Exit Sub
        End If
'>>>>> add start 2011/06/30 Marushita
        ' 中間抜試品か？
        With sprSpec
            lblNukishi.Caption = ""
            For i = 1 To .MaxRows
                .GetText 33, i, vt
                If CStr(vt) = "1" Then
                    lblNukishi.Caption = "中間抜試"
                End If
            Next i
        End With
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'            lblNukishi.Visible = True
'            lblNukishi.Caption = "中間抜試"
'            With sprSpec
'                .ColWidth(31) = 3.88      ' 表示
'                .ColWidth(32) = 3.88      ' 表示
'            End With
'            '中間抜試単位(中間抜試許容値(枚数)/(mm))
'            sprSpec.SetText 32, 1, CInt(typ_CType.typ_si.MSMPCONSTMAI)
'            '中間抜試必要数
'            lblMSMP_SUU.Visible = True
'            '中間抜試の必要数 = (SXLのWF枚数 - 中間抜試許容値(枚数)) / 中間抜試単位(枚数)
'            iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
'            'マイナスの場合、０とする
'            If iMinMidCnt < 0 Then iMinMidCnt = 0
'            '中間抜試の件数
'            iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
'            lblMSMP_SUU.Caption = "抜試数/必要数： " & _
'            CInt(iRstMidCnt) & "/" & CInt(iMinMidCnt) & " 枚"
'            '中間抜試単位(枚数)
'            sprSpec.SetText 31, 1, CInt(typ_CType.typ_si.MSMPTANIMAI)
'        Else
'            lblNukishi.Visible = False
'            lblNukishi.Caption = ""
'            lblMSMP_SUU.Visible = False
'            With sprSpec
'                .ColWidth(31) = 0      ' 非表示
'                .ColWidth(32) = 0      ' 非表示
'            End With
'        End If
'<<<<< add end 2011/06/30 Marushita
    
    Else
        CutCntFlg = 1
        bSampFlag = True
        cmdF(12).Enabled = True
        lblMsg.Caption = GetMsgStr("SET47")
    End If
End Sub

'*******************************************************************************
'*    関数名        : fnc_LoadData
'*
'*    処理概要      : 1.テーブルから必要レコードを取得
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_LoadData() As FUNCTION_RETURN

    Dim udtTmpLackWaf() As typ_LackWaf
    Dim udtTmpBlkInf()  As typ_BlkInf3
    Dim sErrMsg         As String
    Dim m               As Integer
    Dim i               As Integer

    '' DBからデータを読み込む
    If DBDRV_scmzc_fcmlc001d_DispSiyou(tblHinbanRs, tblsiyou, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージ表示
        lblMsg.Caption = sErrMsg
        fnc_LoadData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '' 再抜試指示データの取得
    If DBDRV_scmzc_fcmlc001d_DispSmp(tblSXL.SXLID, tblsmp, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージ表示
        lblMsg.Caption = sErrMsg
        fnc_LoadData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '' 欠落情報の取得
    If DBDRV_scmzc_fcmlc001d_LostInfo(tblSXL.CRYNUM, udtTmpLackWaf) = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージ表示
        lblMsg.Caption = GetMsgStr("EAPLY") & "Y006"
        fnc_LoadData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    m = UBound(tblBlkInf)
    ReDim udtTmpBlkInf(m)
    For i = 1 To m
        DoEvents
        udtTmpBlkInf(i).BLOCKID = tblBlkInf(i).BLOCKID
        udtTmpBlkInf(i).REALLEN = tblBlkInf(i).REALLEN
    Next i

    '' 欠落ウェハーテーブルの作成
    If LackMapMake(udtTmpBlkInf, udtTmpLackWaf, 1, m) = FUNCTION_RETURN_FAILURE Then
    End If

    fnc_LoadData = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    関数名        : sub_SaveData
'*
'*    処理概要      : 1.テーブルへの追加、更新を行う
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_SaveData()

    Dim RET                     As FUNCTION_RETURN
    Dim sErrMsg                 As String
    Dim m                       As Integer
    Dim i                       As Integer
    Dim j                       As Integer
    Dim k                       As Integer
    Dim udtTmpWafSmp()          As typ_XSDCW
    Dim udtTmpEpMesInd()        As typ_TBCMY020
    Dim vNukisiFlg              As Variant
    Dim vGetHinban              As Variant
    Dim intLoopCnt              As Integer
    Dim intRowCnt               As Integer
    Dim sMsg                    As String
    Dim intCnt                  As Integer
    Dim udtNewData              As typImgData       '07/10/05 miyatake START ================>
    Dim udtNewData_Detail(0)    As typImgData_Detail

    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
    Dim SirdYN                  As Boolean      'SIRDｻﾝﾌﾟﾙ指示有無[True:情報あり、False:情報無し]
    Dim sirdSmpID               As String       '検査済みｻﾝﾌﾟﾙID（SIRD用）<TBCMJ022>
    Dim sirdSmpID_Inp           As String       '抜試指示ｻﾝﾌﾟﾙID（SIRD用）
    Dim ix                      As Integer      'SIRDｻﾝﾌﾟﾙ検索用INDEX
    Dim iy                      As Integer      'SIRDｻﾝﾌﾟﾙ検索用INDEX
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END

    udtNewData.detail = udtNewData_Detail     '07/10/05 miyatake END ==================>

    lblMsg.Caption = ""

    '' サンプルボタンが押されているか
    If bSampFlag = False Then
        '' エラーメッセージを表示して処理を抜ける
        lblMsg.Caption = GetMsgStr("ESAMP")
        Exit Sub
    End If

    '' 抜試指示一覧の入力チェック（WFﾏｯﾌﾟ）
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' 区分エラーチェック
    If fnc_CheckHinbanZ() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EHIN6")
        Exit Sub
    End If

    '向先転用チェック
    If fnc_ChkMukesaki = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = "向先が変更されています。"
        Exit Sub
    End If

    '抜試ﾃﾞｰﾀ不正表示チェック
    If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr(sMsg)
        Exit Sub
    End If

    '' 流動監視チェック add 09/03/17 SETkimizuka
    If CheckXODY4(WATCH_PROCCD_NUKISI, "", txtKSXLID.text) = False Then
        lblMsg.Caption = Y4_STOP_ERR
        Exit Sub
    End If
        
    '払出規制チェック    *2010/02/15 Kameda
    If F_HaraiKisei = False Then
        lblMsg.Caption = GetMsgStr("EREG1")
        Exit Sub
    End If
    
    If MsgBox(GetMsgStr("PIN01"), vbOKCancel, "再抜試") = vbCancel Then
        cmdF(12).Enabled = True
        Exit Sub
    End If

    ' 承認機能追加による修正  2007/10/05 miyatake ===================> START
    '' コメント入力
    If Me.chk_Png = 1 Then
        If f_comment.GetComment(sComment) <> vbOK Then
            Exit Sub
        End If
        Call SetForceForegroundWindow(Me.hwnd)
    End If
    DoEvents
    ' 承認機能追加による修正  2007/10/05 miyatake ===================> START

    '' 再抜試テーブルの更新
    If fnc_UpdateData() = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージを表示して処理を抜ける
        '実測ﾁｪｯｸｴﾗｰのﾒｯｾｰｼﾞ表示　2004/01/28 ooba
        If bJituChkFlg = True Then
            lblMsg.Caption = "サンプル位置に実測がありません。"
        Else
            lblMsg.Caption = GetMsgStr("SET49")
        End If
        Exit Sub
    End If

    '' U/Dを上下に分割
    Call SeparateUD

    'サンプル管理とSXL管理の順序を変えた
    '' SXL管理テーブル構造体の設定
    Call sub_MakeTBCME042

    '' WFサンプル管理テーブル構造体の設定
    '新サンプル管理(XSDCW)に変更
    Call sub_MakeTBCME044

    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
    '検査済みのｻﾝﾌﾟﾙ情報を検索<TBCMJ022>
    If Not fncGetSirdSample(Trim(txtCryNum.text), SirdYN, sirdSmpID) Then
        MsgBox "検査済みのサンプル情報の検索に失敗しました" & vbCrLf & "( TBCMJ022 )", vbInformation + vbOKOnly
        Exit Sub
    End If

    '-------------------------------------------------------------------------
    'SIRDの抜試指示が有る場合、そこから下の情報は今回指示したｻﾝﾌﾟﾙを共有する
    'SIRDの抜試指示が無い場合、既に検査済みのｻﾝﾌﾟﾙIDを共有する
    '-------------------------------------------------------------------------
    For ix = 1 To UBound(tblWfSample)
        'SIRDの「ｻﾝﾌﾟﾙ共有」について、共有するｻﾝﾌﾟﾙIDを検索し置き換える（位置が手前で最も近いｻﾝﾌﾟﾙ）
        If tblWfSample(ix).WFSMP.WFINDL4CW = "2" Then
            
            '今回登録情報内にSIRDの抜試指示があるかﾁｪｯｸする
            sirdSmpID_Inp = ""
            For iy = 1 To ix - 1
            
                '今回登録情報内にSIRDの抜試指示がある場合、位置情報を確認
                If tblWfSample(iy).WFSMP.WFINDL4CW = "1" Then
                
                    '抜試指示ｻﾝﾌﾟﾙの位置が自分より手前であればｻﾝﾌﾟﾙIDをKeep
                    If tblWfSample(iy).WFSMP.INPOSCW < tblWfSample(ix).WFSMP.INPOSCW Then
                        sirdSmpID_Inp = tblWfSample(iy).WFSMP.REPSMPLIDCW           '代表サンプルID
                    End If
                End If
            Next iy
            
            If sirdSmpID_Inp = "" Then
                '今回登録情報内にSIRDの抜試指示が共有より手前に無い場合、<TBCMJ022>の検査済みｻﾝﾌﾟﾙIDをｾｯﾄ
                tblWfSample(ix).WFSMP.WFSMPLIDL4CW = sirdSmpID
            Else
                '今回登録情報内にSIRDの抜試指示が共有より手前に有る場合、検索したｻﾝﾌﾟﾙIDをｾｯﾄ
                tblWfSample(ix).WFSMP.WFSMPLIDL4CW = sirdSmpID_Inp
            End If
            
        End If
    Next ix
    '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END
    
    '' SXL確定指示テーブル構造体の設定
    Call sub_MakeTBCMY007

    '' WF総合判定実績テーブル構造体の設定
    Call sub_MakeTBCMW005

    '' 振替廃棄実績テーブル構造体の設定
    Call sub_MakeTBCMW006

'' 測定評価方法指示テーブル構造体の設定 Start
    '' TBCME041構造体に品番設定(測定評価方法指示テーブル作成用)
    m = UBound(tblHinbanRs)
    ReDim tblHinMng(m)
    For i = 1 To m
        tblHinMng(i).hinban = tblHinbanRs(i).HIN.hinban
        tblHinMng(i).REVNUM = tblHinbanRs(i).HIN.mnorevno
        tblHinMng(i).factory = tblHinbanRs(i).HIN.factory
        tblHinMng(i).opecond = tblHinbanRs(i).HIN.opecond
    Next i

    For intLoopCnt = 1 To UBound(tblWfSample())
        With f_cmbc039_3.sprExamine
            For intRowCnt = 1 To .MaxRows
                .row = intRowCnt
                .col = 10
                If left((.text), 3) = Mid(tblWfSample(intLoopCnt).WFSMP.REPSMPLIDCW, 10, 3) And _
                   Right((.text), 4) = Right(tblWfSample(intLoopCnt).WFSMP.REPSMPLIDCW, 4) Then
                    ''残存酸素検査項目追加による変更　03/12/15 ooba
                    .col = 27
                    If .backColor = vbWhite Then
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT1CW = "0"
                    Else
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT1CW = IIf(.text = "", "0", .text)
                    End If

                    'エピ先行評価追加対応
                    .col = 35
                    If .backColor = vbWhite Then
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT2CW = "0"
                    Else
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT2CW = IIf(.text = "", "0", .text)
                    End If
                    Exit For
                End If
            Next
        End With
    Next

    '' TBCME044構造体に設定(測定評価方法指示テーブル作成用)
    m = UBound(tblWfSample)
    ReDim tblWafSmp(m)
    j = 0
    For i = 1 To m
        If Trim(tblWfSample(i).BLOCKID) <> "" Then
            j = j + 1
            tblWafSmp(j).XTALCW = tblWfSample(i).WFSMP.XTALCW
            tblWafSmp(j).INPOSCW = tblWfSample(i).WFSMP.INPOSCW
            tblWafSmp(j).SMPKBNCW = tblWfSample(i).WFSMP.SMPKBNCW
            tblWafSmp(j).REPSMPLIDCW = tblWfSample(i).WFSMP.REPSMPLIDCW
            tblWafSmp(j).HINBCW = tblWfSample(i).WFSMP.HINBCW
            tblWafSmp(j).REVNUMCW = tblWfSample(i).WFSMP.REVNUMCW
            tblWafSmp(j).FACTORYCW = tblWfSample(i).WFSMP.FACTORYCW
            tblWafSmp(j).OPECW = tblWfSample(i).WFSMP.OPECW
            tblWafSmp(j).WFINDRSCW = tblWfSample(i).WFSMP.WFINDRSCW
            tblWafSmp(j).WFINDOICW = tblWfSample(i).WFSMP.WFINDOICW
            tblWafSmp(j).WFINDB1CW = tblWfSample(i).WFSMP.WFINDB1CW
            tblWafSmp(j).WFINDB2CW = tblWfSample(i).WFSMP.WFINDB2CW
            tblWafSmp(j).WFINDB3CW = tblWfSample(i).WFSMP.WFINDB3CW
            tblWafSmp(j).WFINDL1CW = tblWfSample(i).WFSMP.WFINDL1CW
            tblWafSmp(j).WFINDL2CW = tblWfSample(i).WFSMP.WFINDL2CW
            tblWafSmp(j).WFINDL3CW = tblWfSample(i).WFSMP.WFINDL3CW
            tblWafSmp(j).WFINDL4CW = tblWfSample(i).WFSMP.WFINDL4CW
            tblWafSmp(j).WFINDDSCW = tblWfSample(i).WFSMP.WFINDDSCW
            tblWafSmp(j).WFINDDZCW = tblWfSample(i).WFSMP.WFINDDZCW
            tblWafSmp(j).WFINDSPCW = tblWfSample(i).WFSMP.WFINDSPCW
            tblWafSmp(j).WFINDDO1CW = tblWfSample(i).WFSMP.WFINDDO1CW
            tblWafSmp(j).WFINDDO2CW = tblWfSample(i).WFSMP.WFINDDO2CW
            tblWafSmp(j).WFINDDO3CW = tblWfSample(i).WFSMP.WFINDDO3CW
            tblWafSmp(j).WFINDOT1CW = tblWfSample(i).WFSMP.WFINDOT1CW       ' WF検査指示（OT1)
            tblWafSmp(j).WFINDOT2CW = tblWfSample(i).WFSMP.WFINDOT2CW       ' WF検査指示（OT2)
            tblWafSmp(j).WFINDAOICW = tblWfSample(i).WFSMP.WFINDAOICW       ' 残存酸素追加
            tblWafSmp(j).WFINDGDCW = tblWfSample(i).WFSMP.WFINDGDCW         ' GD追加
            tblWafSmp(j).WFHSGDCW = tblWfSample(i).WFSMP.WFHSGDCW           ' 保証FLG(GD)

            'エピ先行評価追加対応
            tblWafSmp(j).EPINDB1CW = tblWfSample(i).WFSMP.EPINDB1CW
            tblWafSmp(j).EPINDB2CW = tblWfSample(i).WFSMP.EPINDB2CW
            tblWafSmp(j).EPINDB3CW = tblWfSample(i).WFSMP.EPINDB3CW
            tblWafSmp(j).EPINDL1CW = tblWfSample(i).WFSMP.EPINDL1CW
            tblWafSmp(j).EPINDL2CW = tblWfSample(i).WFSMP.EPINDL2CW
            tblWafSmp(j).EPINDL3CW = tblWfSample(i).WFSMP.EPINDL3CW
        End If
    Next i
    ReDim Preserve tblWafSmp(j)
    ReDim tblSokuSizi(0)

    'エピ先行評価追加対応
    ReDim udtTmpEpMesInd(0)

    If UBound(tblWafSmp) > 0 Then
        'エピ先行評価追加対応
        If MakeMesIndTbl(tblWfSxlMng, tblWafSmp, tblSokuSizi, udtTmpEpMesInd()) = FUNCTION_RETURN_FAILURE Then

            lblMsg.Caption = GetMsgStr("EGET2", "Y003")
            Exit Sub
        End If
    End If

    m = UBound(tblSokuSizi)
    For i = 1 To m
        tblSokuSizi(i).SENDFLAG = "0"
        tblSokuSizi(i).MUKESAKI = sCmbMukesaki
    Next i

    m = UBound(udtTmpEpMesInd)
    For i = 1 To m
        udtTmpEpMesInd(i).SENDFLAG = "0"
        udtTmpEpMesInd(i).MUKESAKI = sCmbMukesaki
    Next i


'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
For i = 1 To UBound(tblWafInd)
    Debug.Print "tblWafInd(" & i & ").BLOCKID=" & tblWafInd(i).BLOCKID & " : " & "tblWafInd(" & i & ").SAMPLEID=" & tblWafInd(i).SAMPLEID & " : " & "tblWafInd(" & i & ").SMP.CRYINDL4=" & tblWafInd(i).SMP.CRYINDL4
Next i
For i = 1 To UBound(tblSokuSizi)
    Debug.Print "tblSokuSizi(" & i & ").SAMPLEID=" & tblSokuSizi(i).SAMPLEID & " : " & "tblSokuSizi(" & i & ").OSITEM=" & tblSokuSizi(i).OSITEM
Next i
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END

    '' DBに登録
    OraDB.BeginTrans

    If DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージを表示して処理を抜ける
        lblMsg.Caption = GetMsgStr("EWFM4") '03/06/06 後藤
        OraDB.Rollback
        Exit Sub
    End If

    '' WF_GD実績(TBCMJ015)更新処理
    If UBound(typ_J015_WFGDUpd) > 0 Then
        'ﾃﾞｰﾀ数分UPDATE
        For intCnt = 1 To UBound(typ_J015_WFGDUpd)
            If DBDRV_scmzc_fcmlc001c_UpdGDdata(typ_J015_WFGDUpd(intCnt), txtStaffID.text) _
                                        <> FUNCTION_RETURN_SUCCESS Then
                lblMsg.Caption = GetMsgStr("EAPLY") & "J015"
                OraDB.Rollback
            Exit Sub
            End If
        Next
    End If

    'エピ先行評価追加対応
    RET = DBDRV_scmzc_fcmlc001d_Exec(tblWfSample, tblWfSxlMng, tblWfHantei, tblHuriHai, tblSokuSizi, tblSxlKSiji, udtTmpEpMesInd(), sErrMsg)

    If RET = FUNCTION_RETURN_FAILURE Then
        '' エラーメッセージを表示する
        lblMsg.Caption = sErrMsg
        OraDB.Rollback
        Exit Sub
    End If

    '### 基本処理 ###
    Debug.Print "新DB書込み処理開始"
    If MakeParameter(SAINUKISI_FORM) <> FUNCTION_RETURN_SUCCESS Then
        'エラーメッセージはすでに出してある
        OraDB.Rollback
        Debug.Print "新DB書込み処理異常終了"
        Call clearType  '構造体初期化
        EndProcess '' プロセス終了
        Exit Sub
    End If

    '履歴管理DB登録
    '品番の振替を行った場合履歴管理DBにデータを登録する
     If fnc_RirekiKanriDB_Touroku(sErrMsg) = False Then
     lblMsg.Caption = sErrMsg
        OraDB.Rollback
        Exit Sub
     End If

     Call clearType  '構造体初期化

    ' 承認機能追加による修正  07/10/05 miyatake ===================> START
    ''PNGファイル作成
    If Me.chk_Png = 1 Then
        udtNewData.xtal = txtCryNum
        udtNewData.STAFFID = txtStaffID
        udtNewData.SXLID = txtKSXLID
        udtNewData.memo = sComment
        '工程をCC710からCW760に変更 2010/04/30 SETsw kubota
        'If FileCreate_PNG(PROCD_NUKISI_SIJI, udtNewData, Me, sErrMsg, Nothing, pic_Png) = False Then ' upd 09/02/04 SETmiyatake
        If FileCreate_PNG(PROCD_WFC_SAINUKISI, udtNewData, Me, sErrMsg, Nothing, pic_Png) = False Then
            OraDB.Rollback
            lblMsg.Caption = sErrMsg
            Exit Sub
        End If
    End If
    ' 承認機能追加による修正  07/10/05 miyatake ===================> END

    'OraDB.Rollback   'test用コメントを戻す
    OraDB.CommitTrans
    Debug.Print "新DB書込み処理正常終了"

    ' 承認機能追加による修正  07/10/05 miyatake ===================> START
    If Me.chk_Png = 1 Then
        ''PNGファイル送信
        '工程をCC710からCW760に変更 2010/04/30 SETsw kubota
        'Call FileReSend_PNG(PROCD_NUKISI_SIJI)
        Call FileReSend_PNG(PROCD_WFC_SAINUKISI)
    End If
    ' 承認機能追加による修正  07/10/05 miyatake ===================> END

    '' 終了メッセージを表示する
    lblMsg.Caption = GetMsgStr("PPROK")

    ' 終了メッセージを表示する
    lblMsg.Caption = GetMsgStr("PPROK")

    sprExamine.Enabled = False
    cmdF(2).Enabled = True
    cmdF(3).Enabled = False
    cmdF(6).Enabled = False
    cmdF(7).Enabled = False
    cmdF(8).Enabled = False
    cmdF(9).Enabled = False
    cmdF(10).Enabled = False
    cmdF(11).Enabled = False
    cmdF(12).Enabled = False
End Sub

'*******************************************************************************
'*    関数名        : fnc_CheckBlockP
'*
'*    処理概要      : 1.入力されたブロック位置が正しいか判別
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　　　　　　　　　なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_CheckBlockP() As FUNCTION_RETURN

    Dim sTmpBlockID     As String
    Dim sTmpHinban      As String
    Dim sNowHinban      As String
    Dim intTmpBlockP    As Integer
    Dim intRetIngotPos  As Integer
    Dim intCmdIndex     As Integer
    Dim intBlockPErrFlg As Integer
    Dim intRow          As Integer
    Dim intPos          As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim vBlkNullChk     As Variant
    Dim vNukisiFlg      As Variant
    Dim vNextNukisiFlg  As Variant
    Dim sBlockId        As String
    Dim sNextIngotP     As String
    Dim vGetBlkSeq      As Variant
    Dim vGetBlkSeq2     As Variant
    Dim vSample1        As Variant
    Dim vSample2        As Variant
    Dim sSample1        As String
    Dim sSample2        As String
    Dim intNextBlkP     As Integer
    Dim vGetWfNum       As Variant
    Dim vGetBlkP        As Variant
    Dim vGetIngotP      As Variant
    Dim intBlockp       As Integer
    Dim iRtn            As FUNCTION_RETURN

    fnc_CheckBlockP = FUNCTION_RETURN_SUCCESS

    '' ブロック位置、結晶位置の確定
    '品番を1列追加したことによる列の変更
    With sprExamine
        m = .MaxRows
        If m = 0 Then
            Exit Function
        End If

        For i = 1 To m
            If i <> 1 And i <> m Then
                'エピ先行評価追加対応
                .GetText 37, i, vNukisiFlg
                If (vNukisiFlg = "2") Then
                    .row = i
                    .col = 1
                    intCmdIndex = .TypeComboBoxCurSel

                    'エピ先行評価追加対応
                    .col = 39
                    sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                    .col = 4
                    '' エラーチェック(入力されたか)
                    If Trim$(.text) = "" Then
                        .backColor = COLOR_NG
                        Exit Function
                    End If
                    intTmpBlockP = .text
                    .col = 3
                    intRetIngotPos = orgXl.Blks.GetPosByID(sTmpBlockID)
                    '' エラーチェック
                    If fnc_CheckBlockP2(intRetIngotPos, sTmpBlockID, intTmpBlockP, i, intBlockPErrFlg) = FUNCTION_RETURN_FAILURE Then
                        '' エラーでなければ結晶Pを設定する
                        .row = i
                        .col = 4
                        .backColor = COLOR_NG
                    Else
                        .row = i

                        'エピ先行評価追加対応
                        .col = 39
                        sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                        .col = 4
                        intBlockp = CInt(.text)
                        .row = i + 2
                        .col = 4
                        intNextBlkP = CInt(.text)

                        .row = i
                        vSample1 = "U"
                        vSample2 = " " 'DのサンプルIDも作成する

                        If DBDRV_GET_WFMAP(sTmpBlockID, tblSXL.SXLID, intBlockp, vGetBlkP, vGetIngotP, sNextIngotP, vGetBlkSeq, vGetBlkSeq2, vSample1, vSample2, intNextBlkP, vGetWfNum) = FUNCTION_RETURN_FAILURE Then
                            .col = 4
                            .col2 = 4
                            .row = i
                            .row2 = i
                            .BlockMode = True
                            .backColor = &H8080FF
                            .BlockMode = False
                            lblMsg.Caption = GetMsgStr("EWFM1")
                            fnc_CheckBlockP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        Else
                            .col = 4
                            .col2 = 4
                            .row = i
                            .row2 = i
                            .BlockMode = True
                            .backColor = &H80FF80
                            .BlockMode = False
                        End If
                        ReDim Preserve tblNukishi(i + 1)
                        tblNukishi(i).REPSMPLIDCW = vSample1 '代表サンプルIDを作成
                        tblNukishi(i + 1).REPSMPLIDCW = vSample2

                        'ここで作られた結晶Pは表示させない -----------
                        .col = 5
                        .text = CStr(vGetIngotP)
                        .col = 4
                            .backColor = COLOR_OK
                    End If
                'vNukisiFlg = "3"のときの処理を追加(サンプルID切替のため)
                ElseIf vNukisiFlg = "3" And i Mod 2 = 0 Then
                    .row = i

                    'エピ先行評価追加対応
                    .col = 39
                    sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                    .col = 4
                    intBlockp = CInt(.text)
                    .row = i + 2
                    .col = 4
                    intNextBlkP = CInt(.text)

                    .row = i
                    vSample1 = "U"
                    vSample2 = " " 'DのサンプルIDも作成する

                    If DBDRV_GET_WFMAP(sTmpBlockID, tblSXL.SXLID, intBlockp, vGetBlkP, vGetIngotP, sNextIngotP, vGetBlkSeq, vGetBlkSeq2, vSample1, vSample2, intNextBlkP, vGetWfNum) = FUNCTION_RETURN_SUCCESS Then
                        ReDim Preserve tblNukishi(i + 1)
                        tblNukishi(i).REPSMPLIDCW = vSample1 '代表サンプルIDを作成
                        tblNukishi(i + 1).REPSMPLIDCW = vSample2
                    End If
                End If
            End If
        Next i

        '' エラーだった場合、処理を抜ける
        If intBlockPErrFlg <> 0 Then
            '' エラーメッセージを表示して処理を抜ける
            Select Case intBlockPErrFlg
                Case 1      ' SXL範囲外
                    lblMsg.Caption = GetMsgStr("EBLK2")
                Case 2      ' 位置重複
                    lblMsg.Caption = GetMsgStr("EBLK3")
                Case 3      ' 品番エラー
                    lblMsg.Caption = GetMsgStr("EHIN3")
                Case 4      ' 欠落
                    lblMsg.Caption = GetMsgStr("EBLK4")
            End Select

            fnc_CheckBlockP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    End With
End Function

'*******************************************************************************
'*    関数名        : fnc_CheckBlockP2
'*
'*    処理概要      : 1.ブロックPエラーチェック（サブ）
'*　　　　　　　　　　  (ブロックPが範囲外入力されていないかチェック)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*          　　      intIngotpos   ,I  ,Integer　,ブロック開始位置
'*　　      　　      InBlockId     ,I  ,String 　,ブロックID
'*　　      　　      intBlockp     ,I  ,String 　,ブロック位置
'*　　      　　      intRowCount   ,I  ,String 　,SPREADの位置
'*　　      　　      intErrP       ,O  ,String 　,エラー内容フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_CheckBlockP2(intIngotpos As Integer, sBlockId As String, intBlockp As Integer, intRowCount As Integer, ByRef intErrP As Integer) As FUNCTION_RETURN

    Dim sBlk            As String
    Dim sTmpBlockID     As String
    Dim intTmpBlockPos  As Integer
    Dim intTmpBlockLen  As Integer
    Dim intTopIngotPos  As Integer
    Dim intTailIngotPos As Integer
    Dim blFlag          As Boolean
    Dim intPos          As Integer
    Dim m               As Integer
    Dim n               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim vNukisiFlg      As Variant

    fnc_CheckBlockP2 = FUNCTION_RETURN_SUCCESS

    '' エラーチェック（SXL範囲内かどうか）
    If intIngotpos + intBlockp > EIngotP Or intIngotpos + intBlockp < SIngotP Then
        fnc_CheckBlockP2 = FUNCTION_RETURN_FAILURE
        intErrP = 1
        Exit Function
    End If

    '' エラーチェック（ブロックの範囲内か）
    If intIngotpos = 0 Then
        intTmpBlockLen = orgXl.Blks.LowerPos(1)
    Else
        intTopIngotPos = orgXl.Blks.LowerPos(intIngotpos + intBlockp)
        intTailIngotPos = orgXl.Blks.UpperArea(intIngotpos + intBlockp)
        intTmpBlockLen = intTopIngotPos - intTailIngotPos
    End If
    If intTmpBlockLen < intBlockp Then
        fnc_CheckBlockP2 = FUNCTION_RETURN_FAILURE
        intErrP = 1
        Exit Function
    End If

    '品番を1列追加したことによる列の変更
    With sprExamine
        m = .MaxRows

        '' ブロックＰの欠落チェック
        n = UBound(tblLackMap)
        For i = 1 To m
            'エピ先行評価追加対応
            .GetText 37, i, vNukisiFlg
            If (vNukisiFlg = "2") Then  '追加抜試行のみ処理を行う
                .row = i

                'エピ先行評価追加対応
                .col = 39
                sBlk = left(txtCryNum.text, 9) & .text
                .col = 4

                intPos = val(.text)
                For j = 1 To n
                    If tblLackMap(j).BLOCKID = sBlk And _
                       tblLackMap(j).LACKPOSS <= intPos And _
                       tblLackMap(j).LACKPOSE >= intPos Then
                        .backColor = COLOR_NG
                        blFlag = True
                        Exit For
                    End If
                Next j
            End If
        Next i

        If blFlag = True Then
            fnc_CheckBlockP2 = FUNCTION_RETURN_FAILURE
            intErrP = 4
            Exit Function
        End If
    End With
End Function

'*******************************************************************************
'*    関数名        : fnc_CheckHinbanZ
'*
'*    処理概要      : 1.Z品番が設定されているときに区分0ならエラーとする
'*　　　　　　　　　　  (品番エラーチェック)
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_CheckHinbanZ() As FUNCTION_RETURN

    Dim sNowHinban      As String
    Dim sNowKubun       As String
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim vHinChk         As Variant
    Dim vNukisiFlg      As Variant
    Dim vNextNukisiFlg  As Variant

    fnc_CheckHinbanZ = FUNCTION_RETURN_SUCCESS

    '品番を1列追加したことによる列の変更
    With sprExamine
        m = .MaxRows
        For i = 1 To m
            If i <> m Then
                'エピ先行評価追加対応
                .GetText 37, i, vNukisiFlg  '現在行
                .GetText 37, i + 1, vNextNukisiFlg  '次行

                If i Mod 2 <> 0 Then    '奇数行のみ処理を行う
                     .GetText 2, i, vHinChk
                    sNowHinban = vHinChk
                    .row = i
                    .col = 9
                    j = .TypeComboBoxCurSel + 1
                    sNowKubun = Trim$(tblPrcList(j).CODE)
                    If sNowHinban = "Z" And sNowKubun = "0" Then
                        .backColor = COLOR_NG  '色変えしない
                        fnc_CheckHinbanZ = FUNCTION_RETURN_FAILURE
                    ElseIf (sNowHinban = "Z" And i = 1) Or i <> 1 Then
                         .backColor = &H80FF80     '2003/04/16
                    End If
                End If
            End If
        Next i
    End With
End Function

'*******************************************************************************
'*    関数名        : fnc_UpdateData
'*
'*    処理概要      : 1.抜試指示一覧の入力内容により全データを更新する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_UpdateData() As FUNCTION_RETURN

    Dim udtTmpHin   As tFullHinban
    Dim udtTmpUPHin As tFullHinban
    Dim udtTmpDNHin As tFullHinban
    Dim sHin        As String
    Dim blFlag      As Boolean
    Dim m           As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim intSmpRow   As Integer
    Dim vNukisiFlg  As Variant
    Dim vGetSample1 As Variant
    Dim vGetSample2 As Variant
    Dim vSmpleID    As Variant
    Dim lngRow      As Long
    Dim intRow      As Integer

    lblMsg.Caption = ""

    '' 再抜試指示データの取得
    '品番を1列追加したことによる列の変更

    bJituChkFlg = False

    With sprExamine
        blFlag = False
        m = .MaxRows
        If m = 0 Then
            fnc_UpdateData = FUNCTION_RETURN_SUCCESS
            Exit Function
        End If
        intSmpRow = 0

        For i = 1 To m
            ReDim Preserve tblNukishi(m)
            tblNukishi(i).INPOSCW = -1

            '' WFサンプル指示データを更新する
            If i = 1 Or i = m Or CheckGetSampleID(i) = True Or CheckGetSampleID(i) = False Then
                'エピ先行評価追加対応
                .GetText 37, i, vNukisiFlg

                If (vNukisiFlg = 1) Then   '初期表示抜試行の処理を行う
                    intSmpRow = intSmpRow + 1
                    ReDim Preserve tblWafInd(intSmpRow)     '抜試行のみ構造体を作成する
                    .row = i

                    'エピ先行評価追加対応
                    .col = 39
                    'サンプルIDも入れる UPDATEの条件に使用
                    tblWafInd(intSmpRow).BLOCKID = left(tblSXL.SXLID, 9) & .text

                    'サンプルIDが変わらない（初期表示の１・最終行）では開始・終了位置は変えない
                    .col = 4
                    tblWafInd(intSmpRow).BlockPos = val(.text)

                    .col = 5
                    If Not .text = "" Then
                        tblWafInd(intSmpRow).INGOTPOS = .text
                    End If

                    If i = 1 Then   '１行目の結晶位置は既存位置になる
                        tblWafInd(intSmpRow).INGOTPOS = SIngotP
                        tblNukishi(i).INPOSCW = SIngotP
                    ElseIf i = m Then
                        tblWafInd(intSmpRow).INGOTPOS = EIngotP   '最終行も既存位置
                        tblNukishi(i).INPOSCW = EIngotP
                    End If

                    .col = 2

                    If i = 1 Then
                        lngRow = i
                    Else
                        lngRow = i - 1
                    End If
                    .row = lngRow
                    sHin = Trim$(.text)
                    .row = i

                    If sHin <> "Z" And intSmpRow <> m Then
                        '' 品番設定されているならフル桁品番を取得
                        If sHin <> "" Then
                            If GetLastHinban(sHin, udtTmpHin) = FUNCTION_RETURN_FAILURE Then
                                '' エラー処理
                                tblWafInd(intSmpRow).ERRDNFLG = True
                                blFlag = True
                            End If
                        End If
                    End If

                    '最後の行の下品番には初期表示の品番を設定
                    If i = m Then
                        sHin = ""
                    End If

                    With tblWafInd(intSmpRow)
                        If sHin = "Z" Then
                            .HINDN.hinban = sHin                    ' 品番
                            .HINDN.mnorevno = 0                     ' 製品番号改訂番号
                            .HINDN.factory = ""                     ' 工場
                            .HINDN.opecond = ""                     ' 操業条件
                        ElseIf sHin = "" Then
                            .HINDN.hinban = tblSXL.hinban           ' 品番
                            .HINDN.mnorevno = tblSXL.REVNUM         ' 製品番号改訂番号
                            .HINDN.factory = tblSXL.factory         ' 工場
                            .HINDN.opecond = tblSXL.opecond         ' 操業条件
                        Else
                            .HINDN.hinban = sHin                    ' 品番
                            .HINDN.mnorevno = udtTmpHin.mnorevno    ' 製品番号改訂番号
                            .HINDN.factory = udtTmpHin.factory      ' 工場
                            .HINDN.opecond = udtTmpHin.opecond      ' 操業条件
                        End If
                    End With

                    Call sub_GetWafHinban01(intSmpRow)

                    If i = 1 Then
                        tblNukishi(i).HINBCW = tblWafInd(intSmpRow).HINDN.hinban
                        tblNukishi(i).FACTORYCW = tblWafInd(intSmpRow).HINDN.factory
                        tblNukishi(i).REVNUMCW = tblWafInd(intSmpRow).HINDN.mnorevno
                        tblNukishi(i).OPECW = tblWafInd(intSmpRow).HINDN.opecond
                    ElseIf i = m Then
                        tblNukishi(i).HINBCW = tblWafInd(intSmpRow).HINUP.hinban
                        tblNukishi(i).FACTORYCW = tblWafInd(intSmpRow).HINUP.factory
                        tblNukishi(i).REVNUMCW = tblWafInd(intSmpRow).HINUP.mnorevno
                        tblNukishi(i).OPECW = tblWafInd(intSmpRow).HINUP.opecond
                    End If

                    .col = 11
                    tblWafInd(intSmpRow).SMP.CRYINDRS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDRSCW = IIf(.text = "", "0", .text)

                    .col = 12
                    tblWafInd(intSmpRow).SMP.CRYINDOI = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOICW = IIf(.text = "", "0", .text)

                    .col = 13
                    tblWafInd(intSmpRow).SMP.CRYINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB1CW = IIf(.text = "", "0", .text)

                    .col = 14
                    tblWafInd(intSmpRow).SMP.CRYINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB2CW = IIf(.text = "", "0", .text)

                    .col = 15
                    tblWafInd(intSmpRow).SMP.CRYINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB3CW = IIf(.text = "", "0", .text)

                    .col = 16
                    tblWafInd(intSmpRow).SMP.CRYINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL1CW = IIf(.text = "", "0", .text)

                    .col = 17
                    tblWafInd(intSmpRow).SMP.CRYINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL2CW = IIf(.text = "", "0", .text)

                    .col = 18
                    tblWafInd(intSmpRow).SMP.CRYINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL3CW = IIf(.text = "", "0", .text)

                    .col = 19
                    tblWafInd(intSmpRow).SMP.CRYINDL4 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL4CW = IIf(.text = "", "0", .text)

                    .col = 20
                    tblWafInd(intSmpRow).SMP.CRYINDDS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDSCW = IIf(.text = "", "0", .text)

                    .col = 21
                    tblWafInd(intSmpRow).SMP.CRYINDDZ = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDZCW = IIf(.text = "", "0", .text)

                    .col = 22
                    tblWafInd(intSmpRow).SMP.CRYINDSP = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDSPCW = IIf(.text = "", "0", .text)

                    .col = 23
                    tblWafInd(intSmpRow).SMP.CRYINDD1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO1CW = IIf(.text = "", "0", .text)

                    .col = 24
                    tblWafInd(intSmpRow).SMP.CRYINDD2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO2CW = IIf(.text = "", "0", .text)

                    .col = 25
                    tblWafInd(intSmpRow).SMP.CRYINDD3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO3CW = IIf(.text = "", "0", .text)

                    ''残存酸素追加
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = 0
                    If .text <> vbNullString Then   '検査項目ONの部分のみ検査項目取得
                        If .backColor = vbBlack Then
                            tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                            tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)
                        End If
                    End If

                    'エピ先行評価追加対応
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = 0
                    If .text <> vbNullString Then   '検査項目ONの部分のみ検査項目取得
                        If .backColor = vbBlack Then
                            tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                            tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)
                        End If
                    End If

                    'GD追加
                    'エピ先行評価追加対応
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)     '状態ﾌﾗｸﾞ(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)               '状態ﾌﾗｸﾞ(GD)

                    'ｻﾝﾌﾟﾙ未処理またはZ品番でない場合
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg) '保証ﾌﾗｸﾞ(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD        '保証ﾌﾗｸﾞ(GD)
                    'ｻﾝﾌﾟﾙ処理済かつZ品番の場合
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW        '保証ﾌﾗｸﾞ(GD)
                    End If

                    'エピ先行評価追加対応
                    .col = 29
                    tblWafInd(intSmpRow).SMP.EPIINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB1CW = IIf(.text = "", "0", .text)

                    .col = 30
                    tblWafInd(intSmpRow).SMP.EPIINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB2CW = IIf(.text = "", "0", .text)

                    .col = 31
                    tblWafInd(intSmpRow).SMP.EPIINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB3CW = IIf(.text = "", "0", .text)

                    .col = 32
                    tblWafInd(intSmpRow).SMP.EPIINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL1CW = IIf(.text = "", "0", .text)

                    .col = 33
                    tblWafInd(intSmpRow).SMP.EPIINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL2CW = IIf(.text = "", "0", .text)

                    .col = 34
                    tblWafInd(intSmpRow).SMP.EPIINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL3CW = IIf(.text = "", "0", .text)

                    .col = 10

                    '初期表示の先頭、最後はサSMPLKBN2ンプルUPDATEのため区分は入れない
                    tblWafInd(intSmpRow).SMPLKBN2 = ""
                    tblWafInd(intSmpRow).SMPLKBN1 = ""

                    'エピ先行評価追加対応
                    .col = 38
                    tblWafInd(intSmpRow).SAMPLEID = .text

                    If i = 1 Then
                        tblNukishi(i).SMPKBNCW = tKensa(0).SMPKBNCW
                    Else
                        tblNukishi(i).SMPKBNCW = tKensa(1).SMPKBNCW
                    End If
                    tblNukishi(i).TBKBNCW = IIf(tblNukishi(i).SMPKBNCW = "B" Or tblNukishi(i).SMPKBNCW = "U", "B", "T") 'TB区分
                ElseIf vNukisiFlg = 2 Or (vNukisiFlg = 3 And CheckGetSampleID(i) = True) Then  '追加抜試行の処理を行う
                    intSmpRow = intSmpRow + 1
                    ReDim Preserve tblWafInd(intSmpRow)   'add 2003/04/14 hitec)matsumoto 抜試行のみ構造体を作成する

                    If vNukisiFlg = 2 Then
                        .row = i
                        'エピ先行評価追加対応
                        .col = 39
                        tblWafInd(intSmpRow).BLOCKID = left(tblSXL.SXLID, 9) & .text

                        .col = 4
                        tblWafInd(intSmpRow).BlockPos = val(.text)

                        .col = 5
                        tblWafInd(intSmpRow).INGOTPOS = .text
                        tblNukishi(i).INPOSCW = .text
                    ElseIf vNukisiFlg = 3 Then
                        .row = i + 1
                        'エピ先行評価追加対応
                        .col = 39
                        tblWafInd(intSmpRow).BLOCKID = left(tblSXL.SXLID, 9) & .text

                        .col = 4
                        tblWafInd(intSmpRow).BlockPos = val(.text)

                        .col = 5
                        tblWafInd(intSmpRow).INGOTPOS = .text
                        tblNukishi(i).INPOSCW = .text
                    End If

                    .row = i + 1
                    .col = 2
                    sHin = Trim$(.text)

                    If sHin <> "Z" And intSmpRow <> m Then
                        '' 品番設定されているならフル桁品番を取得
                        If sHin <> "" Then
                            If GetLastHinban(.text, udtTmpHin) = FUNCTION_RETURN_FAILURE Then
                                '' エラー処理
                                tblWafInd(intSmpRow).ERRDNFLG = True
                                blFlag = True
                            End If
                        End If
                    End If

                    With tblWafInd(intSmpRow)
                        If sHin = "Z" Then
                            .HINDN.hinban = sHin                    ' 品番
                            .HINDN.mnorevno = 0                     ' 製品番号改訂番号
                            .HINDN.factory = ""                     ' 工場
                            .HINDN.opecond = ""                     ' 操業条件
                        ElseIf sHin = "" Then
                            .HINDN.hinban = tblSXL.hinban           ' 品番
                            .HINDN.mnorevno = tblSXL.REVNUM         ' 製品番号改訂番号
                            .HINDN.factory = tblSXL.factory         ' 工場
                            .HINDN.opecond = tblSXL.opecond         ' 操業条件
                        Else
                            .HINDN.hinban = sHin                    ' 品番
                            .HINDN.mnorevno = udtTmpHin.mnorevno    ' 製品番号改訂番号
                            .HINDN.factory = udtTmpHin.factory      ' 工場
                            .HINDN.opecond = udtTmpHin.opecond      ' 操業条件
                        End If
                    End With

                    tblWafInd(intSmpRow).HINUP.hinban = tblWafInd(intSmpRow - 1).HINDN.hinban
                    tblWafInd(intSmpRow).HINUP.mnorevno = tblWafInd(intSmpRow - 1).HINDN.mnorevno
                    tblWafInd(intSmpRow).HINUP.factory = tblWafInd(intSmpRow - 1).HINDN.factory
                    tblWafInd(intSmpRow).HINUP.opecond = tblWafInd(intSmpRow - 1).HINDN.opecond

                    tblNukishi(i).HINBCW = tblWafInd(intSmpRow).HINUP.hinban
                    tblNukishi(i).FACTORYCW = tblWafInd(intSmpRow).HINUP.factory
                    tblNukishi(i).REVNUMCW = tblWafInd(intSmpRow).HINUP.mnorevno
                    tblNukishi(i).OPECW = tblWafInd(intSmpRow).HINUP.opecond

                    .GetText 10, i, vGetSample1
                    .GetText 10, i + 1, vGetSample2

                    If Trim(vGetSample1) = gsWF_SMPL_JOINT Then
                        tblWafInd(intSmpRow).SMPLKBN1 = Right(vGetSample2, 1)
                        tblWafInd(intSmpRow).SMPLKBN2 = ""

                        'エピ先行評価追加対応
                        .GetText 38, i + 1, vGetSample2
                        .GetText 44, i, vGetSample1

                        tblWafInd(intSmpRow).SAMPLEID = CStr(vGetSample2)
                        tblWafInd(intSmpRow).SAMPLEID2 = ""
                    ElseIf Trim(vGetSample2) = gsWF_SMPL_JOINT Then
                        tblWafInd(intSmpRow).SMPLKBN1 = Right(vGetSample1, 1)
                        tblWafInd(intSmpRow).SMPLKBN2 = ""

                        'エピ先行評価追加対応
                        .GetText 38, i, vGetSample1
                        .GetText 44, i + 1, vGetSample2

                        tblWafInd(intSmpRow).SAMPLEID = CStr(vGetSample1)
                        tblWafInd(intSmpRow).SAMPLEID2 = ""
                    Else
                        tblWafInd(intSmpRow).SMPLKBN1 = Right(vGetSample1, 1)
                        tblWafInd(intSmpRow).SMPLKBN2 = Right(vGetSample2, 1)

                        'エピ先行評価追加対応
                        .GetText 38, i, vGetSample1
                        .GetText 44, i + 1, vGetSample2

                        tblWafInd(intSmpRow).SAMPLEID = CStr(vGetSample1)
                        tblWafInd(intSmpRow).SAMPLEID2 = CStr(vGetSample2)
                    End If

                    .row = i
                    .col = 11
                    tblWafInd(intSmpRow).SMP.CRYINDRS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDRSCW = IIf(.text = "", "0", .text)

                    .col = 12
                    tblWafInd(intSmpRow).SMP.CRYINDOI = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOICW = IIf(.text = "", "0", .text)

                    .col = 13
                    tblWafInd(intSmpRow).SMP.CRYINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB1CW = IIf(.text = "", "0", .text)

                    .col = 14
                    tblWafInd(intSmpRow).SMP.CRYINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB2CW = IIf(.text = "", "0", .text)

                    .col = 15
                    tblWafInd(intSmpRow).SMP.CRYINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB3CW = IIf(.text = "", "0", .text)

                    .col = 16
                    tblWafInd(intSmpRow).SMP.CRYINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL1CW = IIf(.text = "", "0", .text)

                    .col = 17
                    tblWafInd(intSmpRow).SMP.CRYINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL2CW = IIf(.text = "", "0", .text)

                    .col = 18
                    tblWafInd(intSmpRow).SMP.CRYINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL3CW = IIf(.text = "", "0", .text)

                    .col = 19
                    tblWafInd(intSmpRow).SMP.CRYINDL4 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL4CW = IIf(.text = "", "0", .text)

                    .col = 20
                    tblWafInd(intSmpRow).SMP.CRYINDDS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDSCW = IIf(.text = "", "0", .text)

                    .col = 21
                    tblWafInd(intSmpRow).SMP.CRYINDDZ = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDZCW = IIf(.text = "", "0", .text)

                    .col = 22
                    tblWafInd(intSmpRow).SMP.CRYINDSP = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDSPCW = IIf(.text = "", "0", .text)

                    .col = 23
                    tblWafInd(intSmpRow).SMP.CRYINDD1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO1CW = IIf(.text = "", "0", .text)

                    .col = 24
                    tblWafInd(intSmpRow).SMP.CRYINDD2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO2CW = IIf(.text = "", "0", .text)

                    .col = 25
                    tblWafInd(intSmpRow).SMP.CRYINDD3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO3CW = IIf(.text = "", "0", .text)

                    ''残存酸素追加
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)

                    'エピ先行評価追加対応
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)

                    'GD追加　05/02/18 ooba
                    'エピ先行評価追加対応
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)     '状態ﾌﾗｸﾞ(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)               '状態ﾌﾗｸﾞ(GD)

                    'ｻﾝﾌﾟﾙ未処理またはZ品番でない場合
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg) '保証ﾌﾗｸﾞ(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD        '保証ﾌﾗｸﾞ(GD)
                    'ｻﾝﾌﾟﾙ処理済かつZ品番の場合
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW        '保証ﾌﾗｸﾞ(GD)
                    End If

                    'エピ先行評価追加対応
                    .col = 29
                    tblWafInd(intSmpRow).SMP.EPIINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB1CW = IIf(.text = "", "0", .text)

                    .col = 30
                    tblWafInd(intSmpRow).SMP.EPIINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB2CW = IIf(.text = "", "0", .text)

                    .col = 31
                    tblWafInd(intSmpRow).SMP.EPIINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB3CW = IIf(.text = "", "0", .text)

                    .col = 32
                    tblWafInd(intSmpRow).SMP.EPIINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL1CW = IIf(.text = "", "0", .text)

                    .col = 33
                    tblWafInd(intSmpRow).SMP.EPIINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL2CW = IIf(.text = "", "0", .text)

                    .col = 34
                    tblWafInd(intSmpRow).SMP.EPIINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL3CW = IIf(.text = "", "0", .text)

                    i = i + 1
                     .row = i      '抜試行下処理追加

                    tblNukishi(i).HINBCW = tblWafInd(intSmpRow).HINDN.hinban
                    tblNukishi(i).FACTORYCW = tblWafInd(intSmpRow).HINDN.factory
                    tblNukishi(i).REVNUMCW = tblWafInd(intSmpRow).HINDN.mnorevno
                    tblNukishi(i).OPECW = tblWafInd(intSmpRow).HINDN.opecond
                    tblNukishi(i).INPOSCW = tblNukishi(i - 1).INPOSCW
                    .col = 11
                    tblWafInd(intSmpRow).SMP.CRYINDRS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDRSCW = IIf(.text = "", "0", .text)

                    .col = 12
                    tblWafInd(intSmpRow).SMP.CRYINDOI = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOICW = IIf(.text = "", "0", .text)

                    .col = 13
                    tblWafInd(intSmpRow).SMP.CRYINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB1CW = IIf(.text = "", "0", .text)

                    .col = 14
                    tblWafInd(intSmpRow).SMP.CRYINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB2CW = IIf(.text = "", "0", .text)

                    .col = 15
                    tblWafInd(intSmpRow).SMP.CRYINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDB3CW = IIf(.text = "", "0", .text)

                    .col = 16
                    tblWafInd(intSmpRow).SMP.CRYINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL1CW = IIf(.text = "", "0", .text)

                    .col = 17
                    tblWafInd(intSmpRow).SMP.CRYINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL2CW = IIf(.text = "", "0", .text)

                    .col = 18
                    tblWafInd(intSmpRow).SMP.CRYINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL3CW = IIf(.text = "", "0", .text)

                    .col = 19
                    tblWafInd(intSmpRow).SMP.CRYINDL4 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDL4CW = IIf(.text = "", "0", .text)

                    .col = 20
                    tblWafInd(intSmpRow).SMP.CRYINDDS = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDSCW = IIf(.text = "", "0", .text)

                    .col = 21
                    tblWafInd(intSmpRow).SMP.CRYINDDZ = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDZCW = IIf(.text = "", "0", .text)

                    .col = 22
                    tblWafInd(intSmpRow).SMP.CRYINDSP = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDSPCW = IIf(.text = "", "0", .text)

                    .col = 23
                    tblWafInd(intSmpRow).SMP.CRYINDD1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO1CW = IIf(.text = "", "0", .text)

                    .col = 24
                    tblWafInd(intSmpRow).SMP.CRYINDD2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO2CW = IIf(.text = "", "0", .text)

                    .col = 25
                    tblWafInd(intSmpRow).SMP.CRYINDD3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDDO3CW = IIf(.text = "", "0", .text)

                    ''残存酸素追加
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)

                    'エピ先行評価追加対応
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)

                    'エピ先行評価追加対応
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)   '状態ﾌﾗｸﾞ(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)           '状態ﾌﾗｸﾞ(GD)
                    'ｻﾝﾌﾟﾙ未処理またはZ品番でない場合
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg)   '保証ﾌﾗｸﾞ(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD      '保証ﾌﾗｸﾞ(GD)
                    'ｻﾝﾌﾟﾙ処理済かつZ品番の場合
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW      '保証ﾌﾗｸﾞ(GD)
                    End If

                    .col = 29
                    tblWafInd(intSmpRow).SMP.EPIINDB1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB1CW = IIf(.text = "", "0", .text)

                    .col = 30
                    tblWafInd(intSmpRow).SMP.EPIINDB2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB2CW = IIf(.text = "", "0", .text)

                    .col = 31
                    tblWafInd(intSmpRow).SMP.EPIINDB3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDB3CW = IIf(.text = "", "0", .text)

                    .col = 32
                    tblWafInd(intSmpRow).SMP.EPIINDL1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL1CW = IIf(.text = "", "0", .text)

                    .col = 33
                    tblWafInd(intSmpRow).SMP.EPIINDL2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL2CW = IIf(.text = "", "0", .text)

                    .col = 34
                    tblWafInd(intSmpRow).SMP.EPIINDL3 = IIf(.text = "", "0", .text)
                    tblNukishi(i).EPINDL3CW = IIf(.text = "", "0", .text)

                    .GetText 10, i, vGetSample1
                    .GetText 10, i + 1, vGetSample2

                    '実測ﾁｪｯｸ処理追加
                    If giFKeyFlg = 2 Then
                        If fnc_CheckJituData(i) = FUNCTION_RETURN_FAILURE Then
                            fnc_UpdateData = FUNCTION_RETURN_FAILURE
                            bJituChkFlg = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i

        '' 品番に誤りがあるなら処理を抜ける
        If blFlag = True Then
            fnc_UpdateData = FUNCTION_RETURN_FAILURE
            lblMsg.Caption = GetMsgStr("EHIN3")
            Exit Function
        End If
    End With

    '' WF対象品番構造体の設定
    sHin = ""
    j = 0
    m = UBound(tblWafInd) - 1
    ReDim tblHinbanRs(m)
    For i = 1 To m
        With tblWafInd(i)
            If Trim$(.HINDN.hinban) <> sHin And _
               Trim$(.HINDN.hinban) <> "Z" And _
               Trim$(.HINDN.hinban) <> "" Then
                j = j + 1
                tblHinbanRs(j).CRYNUM = tblSXL.CRYNUM
                tblHinbanRs(j).HIN.hinban = .HINDN.hinban
                tblHinbanRs(j).HIN.mnorevno = .HINDN.mnorevno
                tblHinbanRs(j).HIN.factory = .HINDN.factory
                tblHinbanRs(j).HIN.opecond = .HINDN.opecond
            End If
            sHin = Trim$(.HINDN.hinban)
        End With
    Next i
    ReDim Preserve tblHinbanRs(j)

    '' 長さの更新
    m = UBound(tblWafInd)
    For i = 1 To m
        If i = m Then
            tblWafInd(i).LENGTH = 0
        Else

            tblWafInd(i).LENGTH = tblWafInd(i + 1).INGOTPOS - tblWafInd(i).INGOTPOS
        End If
    Next i

    fnc_UpdateData = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    関数名        : fnc_Get_CellHsFlg
'*
'*    処理概要      : 1.ｾﾙ色から保証ﾌﾗｸﾞを取得する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      vNukisi　　　 ,I  ,Variant  ,抜試フラグ
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_Get_CellHsFlg(vNukisi As Variant) As String

    Dim intRow      As Integer            '処理行
    Dim intCol      As Integer            '処理列
    Dim sHin        As String             '8桁品番
    Dim udtFullHin  As tFullHinban        '12桁品番
    Dim sBlkflg     As String             'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ

    '保証ﾌﾗｸﾞ｢0｣：WF実績
    fnc_Get_CellHsFlg = "0"

    With sprExamine
        intRow = .row
        intCol = .col

        '抜試位置の品番取得
        If intRow Mod 2 = 1 Then .row = intRow Else .row = intRow - 1
        .col = 2
        sHin = Trim$(.text)
        If GetLastHinban(sHin, udtFullHin) = FUNCTION_RETURN_FAILURE Then Exit Function

        'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞ取得
        If chkBlkTanFlg(udtFullHin, sBlkflg) = FUNCTION_RETURN_FAILURE Then Exit Function

        .row = intRow
        .col = intCol

        'ﾌﾞﾛｯｸ単位保証ﾌﾗｸﾞのﾁｪｯｸ除外
        If .backColor = COLOR_CryJitsu Then fnc_Get_CellHsFlg = "1"
    End With
End Function

'*******************************************************************************
'*    関数名        : fnc_CheckJituData
'*
'*    処理概要      : 1.実行処理時、追加抜試行に対し実測ﾃﾞｰﾀﾁｪｯｸをする
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      intRow　　   　 ,I  ,Integer  ,Spreadの行
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_CheckJituData(intRow As Integer) As FUNCTION_RETURN
    Dim intChkRow       As Integer      '実績ﾁｪｯｸ抜試行
    Dim intChkCol       As Integer      '実績ﾁｪｯｸ検査項目
    Dim intJituChkCnt   As Integer      '実績あり検査項目数
    Dim blBetuFlg       As Boolean      '共有(別)ｻﾝﾌﾟﾙﾌﾗｸﾞ
    Dim intCnt          As Integer

    fnc_CheckJituData = FUNCTION_RETURN_FAILURE
    intJituChkCnt = 0
    intCnt = 1
    blBetuFlg = False

    With sprExamine

        '実績ﾁｪｯｸ行が共有(別)ｻﾝﾌﾟﾙかﾁｪｯｸ
        Do While iBetuRow(intCnt) > 0
            If intRow = iBetuRow(intCnt) Then
                blBetuFlg = True
            End If
            intCnt = intCnt + 1
            If intCnt > .MaxRows Then Exit Do
        Loop

        '共有(別)ｻﾝﾌﾟﾙの場合は1行ずつ実績ﾁｪｯｸする
        If blBetuFlg = True Then
            For intChkRow = intRow - 1 To intRow
                .row = intChkRow

                'エピ先行評価追加対応
                For intChkCol = 11 To 35
                    .col = intChkCol
                    If .text = "1" Then
                        intJituChkCnt = intJituChkCnt + 1
                    End If
                Next

                '実測ﾃﾞｰﾀが1つでもある場合はﾁｪｯｸOK。ない場合はﾁｪｯｸNGで実行処理を中止する。
                If intJituChkCnt > 0 Then
                    fnc_CheckJituData = FUNCTION_RETURN_SUCCESS
                    intJituChkCnt = 0
                Else
                    fnc_CheckJituData = FUNCTION_RETURN_FAILURE
                    Exit For
                End If
            Next
        Else
            For intChkRow = intRow - 1 To intRow
                .row = intChkRow

                'エピ先行評価追加対応
                For intChkCol = 11 To 35
                    .col = intChkCol
                    If .text = "1" Then
                        intJituChkCnt = intJituChkCnt + 1
                    End If
                Next
            Next

            '実測ﾃﾞｰﾀが1つでもある場合はﾁｪｯｸOK。ない場合はﾁｪｯｸNGで実行処理を中止する。
            If intJituChkCnt > 0 Then
                fnc_CheckJituData = FUNCTION_RETURN_SUCCESS
            End If
        End If
    End With
End Function

'*******************************************************************************
'*    関数名        : sub_GetWafHinban
'*
'*    処理概要      : 1.テーブルの上下品番を取得する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_GetWafHinban()

    Dim sHin    As String
    Dim intPos  As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim i       As Integer
    Dim j       As Integer

    '' 抜試指示データの更新
    m = UBound(tblWafInd)
    n = UBound(tblHinNum)
    For i = 1 To m
        With tblWafInd(i)
            If i = 1 Then
                '' 上品番の取得
                intPos = orgXl.Hins.UpperArea(.INGOTPOS)
                If intPos < 0 Then
                    .HINUP.hinban = ""
                Else
                    .HINUP.hinban = orgXl.Hins(CStr(intPos)).hinban
                End If

                sHin = Trim$(.HINUP.hinban)

                If sHin = "" Or sHin = "G" Or sHin = "Z" Then
                    .HINUP.mnorevno = 0
                    .HINUP.factory = ""
                    .HINUP.opecond = ""
                Else
                    For j = 0 To n
                        If .HINUP.hinban = tblHinNum(j).hinban Then
                            .HINUP.mnorevno = tblHinNum(j).mnorevno
                            .HINUP.factory = tblHinNum(j).factory
                            .HINUP.opecond = tblHinNum(j).opecond
                            Exit For
                        End If
                    Next j
                End If
            Else
                If i = m Then
                    '' 上品番の取得
                    .HINUP.hinban = tblWafInd(i - 1).HINDN.hinban
                    .HINUP.mnorevno = tblWafInd(i - 1).HINDN.mnorevno
                    .HINUP.factory = tblWafInd(i - 1).HINDN.factory
                    .HINUP.opecond = tblWafInd(i - 1).HINDN.opecond

                    ' 下品番の取得
                    intPos = orgXl.Hins.LowerArea(.INGOTPOS)
                    If (intPos = 9999) Then
                        .HINDN.hinban = ""
                    Else
                        .HINDN.hinban = orgXl.Hins(CStr(intPos)).hinban
                    End If
                    sHin = Trim$(.HINUP.hinban)
                    If sHin = "" Or sHin = "G" Or sHin = "Z" Then
                        .HINDN.mnorevno = 0
                        .HINDN.factory = ""
                        .HINDN.opecond = ""
                    Else
                        For j = 0 To n
                            If .HINDN.hinban = tblHinNum(j).hinban Then
                                .HINDN.mnorevno = tblHinNum(j).mnorevno
                                .HINDN.factory = tblHinNum(j).factory
                                .HINDN.opecond = tblHinNum(j).opecond
                                Exit For
                            End If
                        Next j
                    End If
                Else
                    .HINUP.hinban = tblWafInd(i - 1).HINDN.hinban
                    .HINUP.mnorevno = tblWafInd(i - 1).HINDN.mnorevno
                    .HINUP.factory = tblWafInd(i - 1).HINDN.factory
                    .HINUP.opecond = tblWafInd(i - 1).HINDN.opecond
                End If
            End If
        End With
    Next i
End Sub

'*******************************************************************************
'*    関数名        : sub_GetWafHinban01
'*
'*    処理概要      : 1.テーブルの上下品番を取得する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_GetWafHinban01(i As Integer)

    Dim sHin    As String
    Dim intPos  As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim j       As Integer

    '' 抜試指示データの更新
    m = UBound(tblWafInd)
    n = UBound(tblHinNum)
    With tblWafInd(i)
        If i = 1 Then
            '' 上品番の取得
            intPos = orgXl.Hins.UpperArea(.INGOTPOS)
            If intPos < 0 Then
                .HINUP.hinban = ""
            Else
                .HINUP.hinban = orgXl.Hins(CStr(intPos)).hinban
            End If
            sHin = Trim$(.HINUP.hinban)
            If sHin = "" Or sHin = "G" Or sHin = "Z" Then
                .HINUP.mnorevno = 0
                .HINUP.factory = ""
                .HINUP.opecond = ""
            Else
                For j = 0 To n
                    If .HINUP.hinban = tblHinNum(j).hinban Then
                        .HINUP.mnorevno = tblHinNum(j).mnorevno
                        .HINUP.factory = tblHinNum(j).factory
                        .HINUP.opecond = tblHinNum(j).opecond
                        Exit For
                    End If
                Next j
            End If
        Else
            If i = m Then
                '' 上品番の取得
                .HINUP.hinban = tblWafInd(i - 1).HINDN.hinban
                .HINUP.mnorevno = tblWafInd(i - 1).HINDN.mnorevno
                .HINUP.factory = tblWafInd(i - 1).HINDN.factory
                .HINUP.opecond = tblWafInd(i - 1).HINDN.opecond

                ' 下品番の取得
                intPos = orgXl.Hins.LowerArea(.INGOTPOS)
                If (intPos = 9999) Then
                    .HINDN.hinban = ""
                Else
                    .HINDN.hinban = orgXl.Hins(CStr(intPos)).hinban
                End If

                sHin = Trim$(.HINUP.hinban)

                If sHin = "" Or sHin = "G" Or sHin = "Z" Then
                    .HINDN.mnorevno = 0
                    .HINDN.factory = ""
                    .HINDN.opecond = ""
                Else
                    For j = 0 To n
                        If .HINDN.hinban = tblHinNum(j).hinban Then
                            .HINDN.mnorevno = tblHinNum(j).mnorevno
                            .HINDN.factory = tblHinNum(j).factory
                            .HINDN.opecond = tblHinNum(j).opecond
                            Exit For
                        End If
                    Next j
                End If
            End If
        End If
    End With
End Sub

'*******************************************************************************
'*    関数名        : fnc_BsmpID
'*
'*    処理概要      : 1.結晶サンプルとWFサンプル紐付け関数呼び出しをする
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      udtWfSample　 ,I  ,typ_WfSampleGr ,WFサンプル管理
'*　　      　　      j          　 ,I  ,Integer        ,tblWfSample構造体配列用ｶｳﾝﾀ
'*
'*    戻り値        : 結晶サンプルのブロックID
'*
'*******************************************************************************
Private Function fnc_BsmpID(udtWfSample() As typ_WfSampleGr, j As Integer) As String
    Dim udtWfsmp    As typ_Wf_Smpl   '紐付け共通関数用(XSDCW) 2003/09/29 追加
    Dim udtCrsmp    As typ_Cry_Smpl  '紐付け共通関数用(XSDCS戻り値)

    With udtWfsmp
        .SXLIDCW = ""
        .TBKBNCW = udtWfSample(j).WFSMP.TBKBNCW
        .XTALCW = udtWfSample(j).WFSMP.XTALCW
        .INPOSCW = udtWfSample(j).WFSMP.INPOSCW
        .HINBCW = udtWfSample(j).WFSMP.HINBCW
        .REVNUMCW = udtWfSample(j).WFSMP.REVNUMCW
        .FACTORYCW = udtWfSample(j).WFSMP.FACTORYCW
        .OPECW = udtWfSample(j).WFSMP.OPECW
    End With

    If funConSxl_Wf_Sampl(udtWfsmp, udtCrsmp) < 0 Then
        'エラーメッセージ
        fnc_BsmpID = ""
    End If

    fnc_BsmpID = udtCrsmp.CRYNUMCS
End Function

'*******************************************************************************
'*    関数名        : sub_Set_SMP_TB
'*
'*    処理概要      : 1.共有行がﾌﾞﾛｯｸの境界でSXL分割した行の場合、UDをTBに変更する。(各検査項目のｻﾝﾌﾟﾙID)
'*                      (UDをTBに変更(各検査項目のｻﾝﾌﾟﾙID))
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      udtTmpWafSmp  ,I  ,typ_WfSampleGr ,WFサンプル管理
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_Set_SMP_TB(udtTmpWafSmp() As typ_WfSampleGr)

    Dim intCnt As Integer      'ﾙｰﾌﾟｶｳﾝﾀ

    'udtTmpWafSmpの数分ループする (列)
    For intCnt = 1 To UBound(udtTmpWafSmp)
        With udtTmpWafSmp(intCnt)
            '各検査項目のWFINDが、0以外ならTB変換処理を行なう
            If .WFSMP.WFINDRSCW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDRSCW)         'Rs
            If .WFSMP.WFINDOICW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDOICW)         'Oi
            If .WFSMP.WFINDB1CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDB1CW)         'BMD1
            If .WFSMP.WFINDB2CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDB2CW)         'BMD2
            If .WFSMP.WFINDB3CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDB3CW)         'BMD3
            If .WFSMP.WFINDL1CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDL1CW)         'OSF1
            If .WFSMP.WFINDL2CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDL2CW)         'OSF2
            If .WFSMP.WFINDL3CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDL3CW)         'OSF3
            If .WFSMP.WFINDL4CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDL4CW)         'OSF4
            If .WFSMP.WFINDDSCW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDDSCW)         'DS
            If .WFSMP.WFINDDZCW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDDZCW)         'DZ
            If .WFSMP.WFINDSPCW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDSPCW)         'SPVE
            If .WFSMP.WFINDDO1CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDDO1CW)       'DOi1
            If .WFSMP.WFINDDO2CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDDO2CW)       'DOi2
            If .WFSMP.WFINDDO3CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDDO3CW)       'DOi3
            If .WFSMP.WFINDOT1CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDOT1CW)       'OT1
            If .WFSMP.WFINDOT2CW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDOT2CW)       'OT2
            If .WFSMP.WFINDAOICW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDAOICW)       'AOI
            If .WFSMP.WFINDGDCW <> "0" Then Call Cnv_GetSample(.WFSMP.WFSMPLIDGDCW)         'GD     '05/02/21 ooba

            'エピ先行評価追加対応
            If .WFSMP.EPINDB1CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDB1CW)         'BMD1E
            If .WFSMP.EPINDB2CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDB2CW)         'BMD2E
            If .WFSMP.EPINDB3CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDB3CW)         'BMD3E
            If .WFSMP.EPINDL1CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDL1CW)         'OSF1E
            If .WFSMP.EPINDL2CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDL2CW)         'OSF2E
            If .WFSMP.EPINDL3CW <> "0" Then Call Cnv_GetSample(.WFSMP.EPSMPLIDL3CW)         'OSF3E
        End With
    Next intCnt
End Sub

'*******************************************************************************
'*    関数名        : sub_MakeTBCME044
'*
'*    処理概要      : 1.テーブルへ登録させるために構造体構築
'*                      (WFサンプル管理テーブルの作成)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_MakeTBCME044()
    Dim m           As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim ii          As Integer
    Dim intCnt      As Integer      'SXL管理用ｶｳﾝﾀ
    Dim k           As Integer      'tblWfSample構造体配列用ｶｳﾝﾀ

    Dim intDx1      As Integer
    Dim intDx2      As Integer
    Dim blBlkFlg    As Boolean      'SXLﾁｪｯｸﾎﾞｯｸｽﾌﾗｸﾞ
    Dim vGetIngotP1 As Variant
    Dim vGetIngotP2 As Variant

    '' WFサンプル管理テーブル構造体の設定
    m = UBound(tblNukishi)
    ReDim tblWfSample(0)
    ReDim CngSmpID_UD(0)

    i = 0
    k = 0

    For j = 1 To m
        If j <> 1 And j <> m Then
            If tblNukishi(j).INPOSCW <> -1 Then
                k = k + 1
                ReDim Preserve tblWfSample(k)

                i = i + 1
                With tblWafInd(i)
                    tblWfSample(k).BLOCKID = .BLOCKID                   ' ブロックID
                    tblWfSample(k).blockp = .BlockPos                   ' ブロック位置
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(i - 1 * ((j + 1) Mod 2)).SXLID  'SXLID

                    tblWfSample(k).WFSMP.SMPKBNCW = Right(tblNukishi(j).REPSMPLIDCW, 1) ' サンプル区分
                    tblWfSample(k).WFSMP.TBKBNCW = IIf(tblWfSample(k).WFSMP.SMPKBNCW = "U" Or tblWfSample(k).WFSMP.SMPKBNCW = "B", "B", "T") 'TB区分
                    tblWfSample(k).WFSMP.REPSMPLIDCW = tblNukishi(j).REPSMPLIDCW '代表サンプルID
                    tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM            ' 結晶番号
                    tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                ' 結晶内位置

                    'サンプルIDがマップのIDと異なることへの対応 2003/04/23
                    If tblWfSample(k).WFSMP.TBKBNCW = "B" Then
                        tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban         ' 品番
                        tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno     ' 製品番号改訂番号
                        tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory     ' 工場
                        tblWfSample(k).WFSMP.OPECW = .HINUP.opecond         ' 操業条件
                    Else
                        tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban         ' 品番
                        tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno     ' 製品番号改訂番号
                        tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory     ' 工場
                        tblWfSample(k).WFSMP.OPECW = .HINDN.opecond         ' 操業条件
                    End If

                    tblWfSample(k).WFSMP.KTKBNCW = "0"                      ' 確定区分
                    tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k)              'サンプルブロックID←紐付け関数
                    tblWfSample(k).WFSMP.WFSMPLIDRSCW = IIf(tblNukishi(j).WFINDRSCW = "0", "", tblNukishi(j).WFSMPLIDRSCW) 'サンプルID
                    tblWfSample(k).WFSMP.WFSMPLIDOICW = IIf(tblNukishi(j).WFINDOICW = "0", "", tblNukishi(j).WFSMPLIDOICW)
                    tblWfSample(k).WFSMP.WFSMPLIDB1CW = IIf(tblNukishi(j).WFINDB1CW = "0", "", tblNukishi(j).WFSMPLIDB1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDB2CW = IIf(tblNukishi(j).WFINDB2CW = "0", "", tblNukishi(j).WFSMPLIDB2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDB3CW = IIf(tblNukishi(j).WFINDB3CW = "0", "", tblNukishi(j).WFSMPLIDB3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL1CW = IIf(tblNukishi(j).WFINDL1CW = "0", "", tblNukishi(j).WFSMPLIDL1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL2CW = IIf(tblNukishi(j).WFINDL2CW = "0", "", tblNukishi(j).WFSMPLIDL2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL3CW = IIf(tblNukishi(j).WFINDL3CW = "0", "", tblNukishi(j).WFSMPLIDL3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL4CW = IIf(tblNukishi(j).WFINDL4CW = "0", "", tblNukishi(j).WFSMPLIDL4CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDSCW = IIf(tblNukishi(j).WFINDDSCW = "0", "", tblNukishi(j).WFSMPLIDDSCW)
                    tblWfSample(k).WFSMP.WFSMPLIDDZCW = IIf(tblNukishi(j).WFINDDZCW = "0", "", tblNukishi(j).WFSMPLIDDZCW)
                    tblWfSample(k).WFSMP.WFSMPLIDSPCW = IIf(tblNukishi(j).WFINDSPCW = "0", "", tblNukishi(j).WFSMPLIDSPCW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO1CW = IIf(tblNukishi(j).WFINDDO1CW = "0", "", tblNukishi(j).WFSMPLIDDO1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO2CW = IIf(tblNukishi(j).WFINDDO2CW = "0", "", tblNukishi(j).WFSMPLIDDO2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO3CW = IIf(tblNukishi(j).WFINDDO3CW = "0", "", tblNukishi(j).WFSMPLIDDO3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDOT1CW = IIf(tblNukishi(j).WFINDOT1CW = "0", "", tblNukishi(j).WFSMPLIDOT1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDOT2CW = IIf(tblNukishi(j).WFINDOT2CW = "0", "", tblNukishi(j).WFSMPLIDOT2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDAOICW = IIf(tblNukishi(j).WFINDAOICW = "0", "", tblNukishi(j).WFSMPLIDAOICW)   '残存酸素追加　03/12/15 ooba
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = IIf(tblNukishi(j).WFINDGDCW = "0", "", tblNukishi(j).WFSMPLIDGDCW)      'GD追加　05/02/21 ooba

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPSMPLIDB1CW = IIf(tblNukishi(j).EPINDB1CW = "0", "", tblNukishi(j).EPSMPLIDB1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB2CW = IIf(tblNukishi(j).EPINDB2CW = "0", "", tblNukishi(j).EPSMPLIDB2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB3CW = IIf(tblNukishi(j).EPINDB3CW = "0", "", tblNukishi(j).EPSMPLIDB3CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL1CW = IIf(tblNukishi(j).EPINDL1CW = "0", "", tblNukishi(j).EPSMPLIDL1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL2CW = IIf(tblNukishi(j).EPINDL2CW = "0", "", tblNukishi(j).EPSMPLIDL2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL3CW = IIf(tblNukishi(j).EPINDL3CW = "0", "", tblNukishi(j).EPSMPLIDL3CW)
                    tblWfSample(k).WFSMP.WFINDRSCW = tblNukishi(j).WFINDRSCW        ' WF検査指示（Rs)
                    tblWfSample(k).WFSMP.WFINDOICW = tblNukishi(j).WFINDOICW        ' WF検査指示（Oi)
                    tblWfSample(k).WFSMP.WFINDB1CW = tblNukishi(j).WFINDB1CW        ' WF検査指示（B1)
                    tblWfSample(k).WFSMP.WFINDB2CW = tblNukishi(j).WFINDB2CW        ' WF検査指示（B2)
                    tblWfSample(k).WFSMP.WFINDB3CW = tblNukishi(j).WFINDB3CW        ' WF検査指示（B3)
                    tblWfSample(k).WFSMP.WFINDL1CW = tblNukishi(j).WFINDL1CW        ' WF検査指示（L1)
                    tblWfSample(k).WFSMP.WFINDL2CW = tblNukishi(j).WFINDL2CW        ' WF検査指示（L2)
                    tblWfSample(k).WFSMP.WFINDL3CW = tblNukishi(j).WFINDL3CW        ' WF検査指示（L3)
                    tblWfSample(k).WFSMP.WFINDL4CW = tblNukishi(j).WFINDL4CW        ' WF検査指示（L4)
                    tblWfSample(k).WFSMP.WFINDDSCW = tblNukishi(j).WFINDDSCW        ' WF検査指示（DS)
                    tblWfSample(k).WFSMP.WFINDDZCW = tblNukishi(j).WFINDDZCW        ' WF検査指示（DZ)
                    tblWfSample(k).WFSMP.WFINDSPCW = tblNukishi(j).WFINDSPCW        ' WF検査指示（SP)
                    tblWfSample(k).WFSMP.WFINDDO1CW = tblNukishi(j).WFINDDO1CW      ' WF検査指示（D1)
                    tblWfSample(k).WFSMP.WFINDDO2CW = tblNukishi(j).WFINDDO2CW      ' WF検査指示（D2)
                    tblWfSample(k).WFSMP.WFINDDO3CW = tblNukishi(j).WFINDDO3CW      ' WF検査指示（D3)
                    tblWfSample(k).WFSMP.WFINDAOICW = tblNukishi(j).WFINDAOICW      ' WF検査指示 (AOi)    '残存酸素追加　03/12/15 ooba
                    tblWfSample(k).WFSMP.WFINDGDCW = tblNukishi(j).WFINDGDCW        ' WF検査指示 (GD)     'GD追加　05/02/21 ooba

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPINDB1CW = tblNukishi(j).EPINDB1CW        ' WF検査指示（B1E)
                    tblWfSample(k).WFSMP.EPINDB2CW = tblNukishi(j).EPINDB2CW        ' WF検査指示（B2E)
                    tblWfSample(k).WFSMP.EPINDB3CW = tblNukishi(j).EPINDB3CW        ' WF検査指示（B3E)
                    tblWfSample(k).WFSMP.EPINDL1CW = tblNukishi(j).EPINDL1CW        ' WF検査指示（L1E)
                    tblWfSample(k).WFSMP.EPINDL2CW = tblNukishi(j).EPINDL2CW        ' WF検査指示（L2E)
                    tblWfSample(k).WFSMP.EPINDL3CW = tblNukishi(j).EPINDL3CW        ' WF検査指示（L3E)
                    tblWfSample(k).WFSMP.WFINDOT1CW = tblNukishi(j).WFINDOT1CW
                    tblWfSample(k).WFSMP.WFINDOT2CW = tblNukishi(j).WFINDOT2CW
                    tblWfSample(k).WFSMP.WFRESRS1CW = IIf(tblNukishi(j).WFRESRS1CW = "1", "1", "0")                ' WF検査実績（Rs)
                    tblWfSample(k).WFSMP.WFRESOICW = IIf(tblNukishi(j).WFRESOICW = "1", "1", "0")                  ' WF検査実績（Oi)
                    tblWfSample(k).WFSMP.WFRESB1CW = IIf(tblNukishi(j).WFRESB1CW = "1", "1", "0")                  ' WF検査実績（B1)
                    tblWfSample(k).WFSMP.WFRESB2CW = IIf(tblNukishi(j).WFRESB2CW = "1", "1", "0")                  ' WF検査実績（B2）
                    tblWfSample(k).WFSMP.WFRESB3CW = IIf(tblNukishi(j).WFRESB3CW = "1", "1", "0")                  ' WF検査実績（B3)
                    tblWfSample(k).WFSMP.WFRESL1CW = IIf(tblNukishi(j).WFRESL1CW = "1", "1", "0")                  ' WF検査実績（L1)
                    tblWfSample(k).WFSMP.WFRESL2CW = IIf(tblNukishi(j).WFRESL2CW = "1", "1", "0")                  ' WF検査実績（L2)
                    tblWfSample(k).WFSMP.WFRESL3CW = IIf(tblNukishi(j).WFRESL3CW = "1", "1", "0")                  ' WF検査実績（L3)
                    tblWfSample(k).WFSMP.WFRESL4CW = IIf(tblNukishi(j).WFRESL4CW = "1", "1", "0")                  ' WF検査実績（DS)
                    tblWfSample(k).WFSMP.WFRESDSCW = IIf(tblNukishi(j).WFRESDSCW = "1", "1", "0")                  ' WF検査実績（DZ)
                    tblWfSample(k).WFSMP.WFRESDZCW = IIf(tblNukishi(j).WFRESDZCW = "1", "1", "0")                  ' WF検査実績（DZ)
                    tblWfSample(k).WFSMP.WFRESSPCW = IIf(tblNukishi(j).WFRESSPCW = "1", "1", "0")                  ' WF検査実績（SP)
                    tblWfSample(k).WFSMP.WFRESDO1CW = IIf(tblNukishi(j).WFRESDO1CW = "1", "1", "0")                ' WF検査実績（DO1)
                    tblWfSample(k).WFSMP.WFRESDO2CW = IIf(tblNukishi(j).WFRESDO2CW = "1", "1", "0")                ' WF検査実績（DO2)
                    tblWfSample(k).WFSMP.WFRESDO3CW = IIf(tblNukishi(j).WFRESDO3CW = "1", "1", "0")                ' WF検査実績（DO3)
                    tblWfSample(k).WFSMP.WFRESOT1CW = IIf(tblNukishi(j).WFRESOT1CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESOT2CW = IIf(tblNukishi(j).WFRESOT2CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESAOICW = IIf(tblNukishi(j).WFRESAOICW = "1", "1", "0")                ' WF検査実績（AOi)  '残存酸素追加
                    tblWfSample(k).WFSMP.WFRESGDCW = IIf(tblNukishi(j).WFRESGDCW = "1", "1", "0")                  ' WF検査実績（GD)   'GD追加

                    tblWfSample(k).WFSMP.WFHSGDCW = IIf(tblNukishi(j).WFHSGDCW = "1", "1", "0")                    ' 保証FLG（GD)

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPRESB1CW = IIf(tblNukishi(j).EPRESB1CW = "1", "1", "0")                  ' WF検査実績（B1E)
                    tblWfSample(k).WFSMP.EPRESB2CW = IIf(tblNukishi(j).EPRESB2CW = "1", "1", "0")                  ' WF検査実績（B2E）
                    tblWfSample(k).WFSMP.EPRESB3CW = IIf(tblNukishi(j).EPRESB3CW = "1", "1", "0")                  ' WF検査実績（B3E)
                    tblWfSample(k).WFSMP.EPRESL1CW = IIf(tblNukishi(j).EPRESL1CW = "1", "1", "0")                  ' WF検査実績（L1E)
                    tblWfSample(k).WFSMP.EPRESL2CW = IIf(tblNukishi(j).EPRESL2CW = "1", "1", "0")                  ' WF検査実績（L2E)
                    tblWfSample(k).WFSMP.EPRESL3CW = IIf(tblNukishi(j).EPRESL3CW = "1", "1", "0")                  ' WF検査実績（L3E)
                    tblWfSample(k).WFSMP.TSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.LIVKCW = "0"

                    j = j + 1
                    k = k + 1
                    ReDim Preserve tblWfSample(k)

                    tblWfSample(k).BLOCKID = .BLOCKID                   ' ブロックID
                    tblWfSample(k).blockp = .BlockPos                   ' ブロック位置

                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(i - 1 * ((j + 1) Mod 2)).SXLID   'SXLID
                    tblWfSample(k).WFSMP.SMPKBNCW = Right(tblNukishi(j).REPSMPLIDCW, 1)         ' サンプル区分
                    tblWfSample(k).WFSMP.TBKBNCW = IIf(tblWfSample(k).WFSMP.SMPKBNCW = "U" Or tblWfSample(k).WFSMP.SMPKBNCW = "B", "B", "T") 'TB区分
                    tblWfSample(k).WFSMP.REPSMPLIDCW = tblNukishi(j).REPSMPLIDCW                '代表サンプルID
                    tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM                                 ' 結晶番号
                    tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                                    ' 結晶内位置

                    If tblNukishi(j).TBKBNCW = "B" Then
                        tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban     ' 品番
                        tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno ' 製品番号改訂番号
                        tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory ' 工場
                        tblWfSample(k).WFSMP.OPECW = .HINUP.opecond     ' 操業条件
                    Else
                        tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban     ' 品番
                        tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno ' 製品番号改訂番号
                        tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory ' 工場
                        tblWfSample(k).WFSMP.OPECW = .HINDN.opecond     ' 操業条件
                    End If
                    tblWfSample(k).WFSMP.KTKBNCW = "0"                  ' 確定区分
                    tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k)    ' サンプルブロックID←紐付け関数

                    tblWfSample(k).WFSMP.WFSMPLIDRSCW = IIf(tblNukishi(j).WFINDRSCW = "0", "", tblNukishi(j).WFSMPLIDRSCW) 'サンプルID
                    tblWfSample(k).WFSMP.WFSMPLIDOICW = IIf(tblNukishi(j).WFINDOICW = "0", "", tblNukishi(j).WFSMPLIDOICW)
                    tblWfSample(k).WFSMP.WFSMPLIDB1CW = IIf(tblNukishi(j).WFINDB1CW = "0", "", tblNukishi(j).WFSMPLIDB1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDB2CW = IIf(tblNukishi(j).WFINDB2CW = "0", "", tblNukishi(j).WFSMPLIDB2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDB3CW = IIf(tblNukishi(j).WFINDB3CW = "0", "", tblNukishi(j).WFSMPLIDB3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL1CW = IIf(tblNukishi(j).WFINDL1CW = "0", "", tblNukishi(j).WFSMPLIDL1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL2CW = IIf(tblNukishi(j).WFINDL2CW = "0", "", tblNukishi(j).WFSMPLIDL2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL3CW = IIf(tblNukishi(j).WFINDL3CW = "0", "", tblNukishi(j).WFSMPLIDL3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDL4CW = IIf(tblNukishi(j).WFINDL4CW = "0", "", tblNukishi(j).WFSMPLIDL4CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDSCW = IIf(tblNukishi(j).WFINDDSCW = "0", "", tblNukishi(j).WFSMPLIDDSCW)
                    tblWfSample(k).WFSMP.WFSMPLIDDZCW = IIf(tblNukishi(j).WFINDDZCW = "0", "", tblNukishi(j).WFSMPLIDDZCW)
                    tblWfSample(k).WFSMP.WFSMPLIDSPCW = IIf(tblNukishi(j).WFINDSPCW = "0", "", tblNukishi(j).WFSMPLIDSPCW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO1CW = IIf(tblNukishi(j).WFINDDO1CW = "0", "", tblNukishi(j).WFSMPLIDDO1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO2CW = IIf(tblNukishi(j).WFINDDO2CW = "0", "", tblNukishi(j).WFSMPLIDDO2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDDO3CW = IIf(tblNukishi(j).WFINDDO3CW = "0", "", tblNukishi(j).WFSMPLIDDO3CW)
                    tblWfSample(k).WFSMP.WFSMPLIDOT1CW = IIf(tblNukishi(j).WFINDOT1CW = "0", "", tblNukishi(j).WFSMPLIDOT1CW)
                    tblWfSample(k).WFSMP.WFSMPLIDOT2CW = IIf(tblNukishi(j).WFINDOT2CW = "0", "", tblNukishi(j).WFSMPLIDOT2CW)
                    tblWfSample(k).WFSMP.WFSMPLIDAOICW = IIf(tblNukishi(j).WFINDAOICW = "0", "", tblNukishi(j).WFSMPLIDAOICW)   '残存酸素追加　03/12/15 ooba
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = IIf(tblNukishi(j).WFINDGDCW = "0", "", tblNukishi(j).WFSMPLIDGDCW)      'GD追加　05/02/21 ooba

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPSMPLIDB1CW = IIf(tblNukishi(j).EPINDB1CW = "0", "", tblNukishi(j).EPSMPLIDB1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB2CW = IIf(tblNukishi(j).EPINDB2CW = "0", "", tblNukishi(j).EPSMPLIDB2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB3CW = IIf(tblNukishi(j).EPINDB3CW = "0", "", tblNukishi(j).EPSMPLIDB3CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL1CW = IIf(tblNukishi(j).EPINDL1CW = "0", "", tblNukishi(j).EPSMPLIDL1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL2CW = IIf(tblNukishi(j).EPINDL2CW = "0", "", tblNukishi(j).EPSMPLIDL2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL3CW = IIf(tblNukishi(j).EPINDL3CW = "0", "", tblNukishi(j).EPSMPLIDL3CW)

                    tblWfSample(k).WFSMP.WFINDRSCW = tblNukishi(j).WFINDRSCW        ' WF検査指示（Rs)
                    tblWfSample(k).WFSMP.WFINDOICW = tblNukishi(j).WFINDOICW        ' WF検査指示（Oi)
                    tblWfSample(k).WFSMP.WFINDB1CW = tblNukishi(j).WFINDB1CW        ' WF検査指示（B1)
                    tblWfSample(k).WFSMP.WFINDB2CW = tblNukishi(j).WFINDB2CW        ' WF検査指示（B2)
                    tblWfSample(k).WFSMP.WFINDB3CW = tblNukishi(j).WFINDB3CW        ' WF検査指示（B3)
                    tblWfSample(k).WFSMP.WFINDL1CW = tblNukishi(j).WFINDL1CW        ' WF検査指示（L1)
                    tblWfSample(k).WFSMP.WFINDL2CW = tblNukishi(j).WFINDL2CW        ' WF検査指示（L2)
                    tblWfSample(k).WFSMP.WFINDL3CW = tblNukishi(j).WFINDL3CW        ' WF検査指示（L3)
                    tblWfSample(k).WFSMP.WFINDL4CW = tblNukishi(j).WFINDL4CW        ' WF検査指示（L4)
                    tblWfSample(k).WFSMP.WFINDDSCW = tblNukishi(j).WFINDDSCW        ' WF検査指示（DS)
                    tblWfSample(k).WFSMP.WFINDDZCW = tblNukishi(j).WFINDDZCW        ' WF検査指示（DZ)
                    tblWfSample(k).WFSMP.WFINDSPCW = tblNukishi(j).WFINDSPCW        ' WF検査指示（SP)
                    tblWfSample(k).WFSMP.WFINDDO1CW = tblNukishi(j).WFINDDO1CW      ' WF検査指示（D1)
                    tblWfSample(k).WFSMP.WFINDDO2CW = tblNukishi(j).WFINDDO2CW      ' WF検査指示（D2)
                    tblWfSample(k).WFSMP.WFINDDO3CW = tblNukishi(j).WFINDDO3CW      ' WF検査指示（D3)
                    tblWfSample(k).WFSMP.WFINDOT1CW = tblNukishi(j).WFINDOT1CW
                    tblWfSample(k).WFSMP.WFINDOT2CW = tblNukishi(j).WFINDOT2CW
                    tblWfSample(k).WFSMP.WFINDAOICW = tblNukishi(j).WFINDAOICW      ' WF検査指示 (AOi)     '残存酸素追加　03/12/15 ooba
                    tblWfSample(k).WFSMP.WFINDGDCW = tblNukishi(j).WFINDGDCW        ' WF検査指示 (GD)      'GD追加　05/02/21 ooba

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPINDB1CW = tblNukishi(j).EPINDB1CW        ' WF検査指示（B1)
                    tblWfSample(k).WFSMP.EPINDB2CW = tblNukishi(j).EPINDB2CW        ' WF検査指示（B2)
                    tblWfSample(k).WFSMP.EPINDB3CW = tblNukishi(j).EPINDB3CW        ' WF検査指示（B3)
                    tblWfSample(k).WFSMP.EPINDL1CW = tblNukishi(j).EPINDL1CW        ' WF検査指示（L1)
                    tblWfSample(k).WFSMP.EPINDL2CW = tblNukishi(j).EPINDL2CW        ' WF検査指示（L2)
                    tblWfSample(k).WFSMP.EPINDL3CW = tblNukishi(j).EPINDL3CW        ' WF検査指示（L3)
                    tblWfSample(k).WFSMP.WFRESRS1CW = IIf(tblNukishi(j).WFRESRS1CW = "1", "1", "0")               ' WF検査実績（Rs)
                    tblWfSample(k).WFSMP.WFRESOICW = IIf(tblNukishi(j).WFRESOICW = "1", "1", "0")                  ' WF検査実績（Oi)
                    tblWfSample(k).WFSMP.WFRESB1CW = IIf(tblNukishi(j).WFRESB1CW = "1", "1", "0")                  ' WF検査実績（B1)
                    tblWfSample(k).WFSMP.WFRESB2CW = IIf(tblNukishi(j).WFRESB2CW = "1", "1", "0")                  ' WF検査実績（B2）
                    tblWfSample(k).WFSMP.WFRESB3CW = IIf(tblNukishi(j).WFRESB3CW = "1", "1", "0")                  ' WF検査実績（B3)
                    tblWfSample(k).WFSMP.WFRESL1CW = IIf(tblNukishi(j).WFRESL1CW = "1", "1", "0")                  ' WF検査実績（L1)
                    tblWfSample(k).WFSMP.WFRESL2CW = IIf(tblNukishi(j).WFRESL2CW = "1", "1", "0")                  ' WF検査実績（L2)
                    tblWfSample(k).WFSMP.WFRESL3CW = IIf(tblNukishi(j).WFRESL3CW = "1", "1", "0")                  ' WF検査実績（L3)
                    tblWfSample(k).WFSMP.WFRESL4CW = IIf(tblNukishi(j).WFRESL4CW = "1", "1", "0")                  ' WF検査実績（DS)
                    tblWfSample(k).WFSMP.WFRESDSCW = IIf(tblNukishi(j).WFRESDSCW = "1", "1", "0")                  ' WF検査実績（DZ)
                    tblWfSample(k).WFSMP.WFRESDZCW = IIf(tblNukishi(j).WFRESDZCW = "1", "1", "0")                  ' WF検査実績（DZ)
                    tblWfSample(k).WFSMP.WFRESSPCW = IIf(tblNukishi(j).WFRESSPCW = "1", "1", "0")                  ' WF検査実績（SP)
                    tblWfSample(k).WFSMP.WFRESDO1CW = IIf(tblNukishi(j).WFRESDO1CW = "1", "1", "0")                 ' WF検査実績（DO1)
                    tblWfSample(k).WFSMP.WFRESDO2CW = IIf(tblNukishi(j).WFRESDO2CW = "1", "1", "0")                ' WF検査実績（DO2)
                    tblWfSample(k).WFSMP.WFRESDO3CW = IIf(tblNukishi(j).WFRESDO3CW = "1", "1", "0")                 ' WF検査実績（DO3)
                    tblWfSample(k).WFSMP.WFRESOT1CW = IIf(tblNukishi(j).WFRESOT1CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESOT2CW = IIf(tblNukishi(j).WFRESOT2CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESAOICW = IIf(tblNukishi(j).WFRESAOICW = "1", "1", "0")                 ' WF検査実績 (AOi)  '残存酸素追加
                    tblWfSample(k).WFSMP.WFRESGDCW = IIf(tblNukishi(j).WFRESGDCW = "1", "1", "0")                   ' WF検査実績 (GD)   'GD追加
                    tblWfSample(k).WFSMP.WFHSGDCW = IIf(tblNukishi(j).WFHSGDCW = "1", "1", "0")                     ' 保証FLG（GD)

                    'エピ先行評価追加対応
                    tblWfSample(k).WFSMP.EPRESB1CW = IIf(tblNukishi(j).EPRESB1CW = "1", "1", "0")                  ' WF検査実績（B1E)
                    tblWfSample(k).WFSMP.EPRESB2CW = IIf(tblNukishi(j).EPRESB2CW = "1", "1", "0")                  ' WF検査実績（B2E）
                    tblWfSample(k).WFSMP.EPRESB3CW = IIf(tblNukishi(j).EPRESB3CW = "1", "1", "0")                  ' WF検査実績（B3E)
                    tblWfSample(k).WFSMP.EPRESL1CW = IIf(tblNukishi(j).EPRESL1CW = "1", "1", "0")                  ' WF検査実績（L1E)
                    tblWfSample(k).WFSMP.EPRESL2CW = IIf(tblNukishi(j).EPRESL2CW = "1", "1", "0")                  ' WF検査実績（L2E)
                    tblWfSample(k).WFSMP.EPRESL3CW = IIf(tblNukishi(j).EPRESL3CW = "1", "1", "0")                  ' WF検査実績（L3E)
                    tblWfSample(k).WFSMP.TSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.LIVKCW = "0"

                    'ﾌﾞﾛｯｸの境界でSXL分割した場合UDをTBに変更する
                    intDx1 = k - 1:   intDx2 = k

                    blBlkFlg = False
                    With sprExamine
                        .col = 1
                        For ii = intDx1 To .MaxRows Step 2
                            .row = ii
                            If .CellType = CellTypeCheckBox Then
                                If .Value = "1" Then
                                    .GetText 5, ii, vGetIngotP1
                                    .GetText 5, ii + 1, vGetIngotP2
                                    If tblWfSample(intDx1).WFSMP.INPOSCW = CInt(vGetIngotP1) Or _
                                                tblWfSample(intDx1).WFSMP.INPOSCW = CInt(vGetIngotP2) Then
                                        blBlkFlg = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End With

                    If tblWfSample(intDx1).WFSMP.SMCRYNUMCW <> tblWfSample(intDx2).WFSMP.SMCRYNUMCW _
                       Or blBlkFlg = True Then
                        ' 変更対象に追加
                        ReDim Preserve CngSmpID_UD(UBound(CngSmpID_UD) + 2)
                        CngSmpID_UD(UBound(CngSmpID_UD) - 1) = tblWfSample(intDx1).WFSMP.REPSMPLIDCW
                        CngSmpID_UD(UBound(CngSmpID_UD)) = tblWfSample(intDx2).WFSMP.REPSMPLIDCW

                        tblWfSample(intDx1).WFSMP.SMPKBNCW = Cnv_Smp_KB(tblWfSample(intDx1).WFSMP.SMPKBNCW)
                        tblWfSample(intDx2).WFSMP.SMPKBNCW = Cnv_Smp_KB(tblWfSample(intDx2).WFSMP.SMPKBNCW)
                        tblWfSample(intDx1).WFSMP.REPSMPLIDCW = left(tblWfSample(intDx1).WFSMP.REPSMPLIDCW, Len(tblWfSample(intDx1).WFSMP.REPSMPLIDCW) - 1) + tblWfSample(intDx1).WFSMP.SMPKBNCW
                        tblWfSample(intDx2).WFSMP.REPSMPLIDCW = left(tblWfSample(intDx2).WFSMP.REPSMPLIDCW, Len(tblWfSample(intDx2).WFSMP.REPSMPLIDCW) - 1) + tblWfSample(intDx2).WFSMP.SMPKBNCW
                    End If
                End With
            End If
        Else
            i = i + 1
            With tblWafInd(i)
                k = k + 1
                ReDim Preserve tblWfSample(k)

                tblWfSample(k).BLOCKID = ""                                 ' ブロックID
                tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM                 ' 結晶番号
                tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                    ' 結晶内位置
                If i = 1 Then
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(1).SXLID     'SXLID
                    tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban             ' 品番
                    tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno         ' 製品番号改訂番号
                    tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory         ' 工場
                    tblWfSample(k).WFSMP.OPECW = .HINDN.opecond             ' 操業条件
                Else
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(UBound(tblWfSxlMng)).SXLID  'SXLID 最後の
                    tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban             ' 品番
                    tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno         ' 製品番号改訂番号
                    tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory         ' 工場
                    tblWfSample(k).WFSMP.OPECW = .HINUP.opecond             ' 操業条件
                End If
                tblWfSample(k).WFSMP.SMPKBNCW = tblNukishi(j).SMPKBNCW
                tblWfSample(k).WFSMP.TBKBNCW = tblNukishi(j).TBKBNCW        'TB区分
                tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k) 'サンプルブロックID←紐付け関数
                tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text             '更新社員ID

                ''全振替時の結晶GD引継ぎ対応
                If i = 1 And bMotoGDcpyFlg(1) Then
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = CpyCrySmpl.TsmplidGD
                    tblWfSample(k).WFSMP.WFINDGDCW = CpyCrySmpl.TindGD
                    tblWfSample(k).WFSMP.WFRESGDCW = "1"
                    tblWfSample(k).WFSMP.WFHSGDCW = .SMP.WFHSGD
                ElseIf i <> 1 And bMotoGDcpyFlg(2) Then
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = CpyCrySmpl.BsmplidGD
                    tblWfSample(k).WFSMP.WFINDGDCW = CpyCrySmpl.BindGD
                    tblWfSample(k).WFSMP.WFRESGDCW = "1"
                    tblWfSample(k).WFSMP.WFHSGDCW = .SMP.WFHSGD
                End If
            End With
        End If
    Next j

    'ﾌﾞﾛｯｸの境界でSXL分割した場合UDをTBに変更する
    If UBound(CngSmpID_UD) > 0 Then
        Call sub_Set_SMP_TB(tblWfSample())
    End If
End Sub

'*******************************************************************************
'*    関数名        : sub_MakeTBCME042
'*
'*    処理概要      : 1.SXL管理テーブルへ登録させるために構造体構築
'*                      (SXL管理テーブルの作成)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_MakeTBCME042()

    Dim udtTmpHin       As tFullHinban  ' フル桁品番取得用構造体
    Dim sNowHinban      As String       ' 現在行の品番
    Dim intNowIngotPos  As Integer      ' 現在行のインゴット位置
    Dim intOldIngotPos  As Integer
    Dim sTmpSXLID       As String       ' SXLID
    Dim intNowSxlLen    As Integer      ' SXLの長さ
    Dim intSPoint       As Integer      ' 行位置保存
    Dim intFlg          As Integer      ' 判別フラグ
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim k               As Integer
    Dim vSampl          As Variant
    Dim intCounter      As Integer
    Dim sTemp           As String
    Dim vFlg            As Variant

    ReDim tblWfSxlMng(0)

    m = UBound(tblWafInd)
    k = 0
    For i = 1 To m - 1
        k = k + 1
        ReDim Preserve tblWfSxlMng(k)
        With tblWfSxlMng(k)
            If i = 1 Then
                .SXLID = tblSXL.SXLID     'SXLを位置から作成しなおすと、SXLIDがずれてしまう為、元SXLIDをそのまま使うよう修正
                .CRYNUM = ""                        ' 結晶番号
            Else
                .SXLID = Mid(tblWafInd(i).BLOCKID, 1, 10) & GetWafPos(tblWafInd(i).INGOTPOS)
                .CRYNUM = tblSXL.CRYNUM              ' 結晶番号
            End If

            '先頭SXLの開始位置は前工程の開始位置となる
            If i = 1 Then
                .INGOTPOS = SIngotP
            Else
                .INGOTPOS = tblWafInd(i).INGOTPOS               ' 結晶内開始位置
            End If
            .LENGTH = tblWafInd(i).LENGTH                       ' 長さ
            .hinban = tblWafInd(i).HINDN.hinban                 ' 品番
            .REVNUM = tblWafInd(i).HINDN.mnorevno               ' 製品番号改訂番号
            .factory = tblWafInd(i).HINDN.factory               ' 工場
            .opecond = tblWafInd(i).HINDN.opecond               ' 操業条件
            .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI                 ' 管理工程
            .NOWPROC = PROCD_WFC_SOUGOUHANTEI                   ' 現在工程
            .LPKRPROCCD = MGPRCD_WFC_SOUGOUHANTEI               ' 最終通過管理工程
            .LASTPASS = PROCD_WFC_SOUGOUHANTEI                  ' 最終通過工程

            '' Ｚ品番なら削除区分と最終状態区分を変える
            If Trim(.hinban) = "Z" Then
                .DELCLS = "1"                                   ' 削除区分
                .LSTATCLS = "H"                                 ' 最終状態区分
            Else
                .DELCLS = "0"                                   ' 削除区分
                .LSTATCLS = "T"                                 ' 最終状態区分
            End If
            .HOLDCLS = "0"                                      ' ホールド区分

            '品番を1列追加したことによる列の変更
            sprExamine.row = i
            sprExamine.col = 9
            j = sprExamine.TypeComboBoxCurSel + 1
            .BDCAUS = Trim$(tblPrcList(j).CODE)                 ' 不良理由
            .COUNT = "0"                                        ' 枚数
        End With
    Next i
End Sub

'*******************************************************************************
'*    関数名        : sub_MakeTBCMY007
'*
'*    処理概要      : 1.SXL確定指示テーブルへ登録させるために構造体構築
'*                      (SXL確定指示テーブルの作成)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_MakeTBCMY007()
    Dim sNowBlockID     As String
    Dim intNowIngotPos  As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim j               As Integer
    Dim vChgSmpId       As Variant
    Dim vFlg            As Variant
    Dim vHinban         As Variant
    Dim intFuryoCmb     As Integer
    Dim sSmpId(2)       As String       '共有ｻﾝﾌﾟﾙID
    Dim k               As Integer
    Dim intCntSXL       As Integer      'SXLﾚｺｰﾄﾞ数
    Dim udtRsHIN        As tFullHinban  '比抵抗(Rs)仕様取得品番
    Dim sRsData(10)     As String       '比抵抗(Rs)ﾃﾞｰﾀ
    Dim sRsPtn          As String       '比抵抗ﾃﾞｰﾀ取得ﾊﾟﾀｰﾝ
    Dim intSxlPtn       As Integer      '登録SXLﾊﾟﾀｰﾝ
    Dim sHosho          As String       '保証方法
'Add Start 2012/03/08 Y.Hitomi
    Dim iBlockHoshoCnt  As String       'ブロック確定（廃棄）回数
'Add End   2012/03/08 Y.Hitomi

    
'                                       ●1 : 全廃棄SXL
'                                       ●2 : 上追込みSXL
'                                       ●3 : 下追込みSXL
'                                       ●4 : SXLの間をZ
'                                       ●0 : 取得ﾃﾞｰﾀなし

    sSmpId(1) = ""       'ｻﾝﾌﾟﾙID(From)初期化
    sSmpId(2) = ""       'ｻﾝﾌﾟﾙID(To)初期化
    intCntSXL = UBound(tblWfSample)
'Add Start 2012/03/08 Y.Hitomi
    iBlockHoshoCnt = 0
'Add End   2012/03/08 Y.Hitomi

    '品番を1列追加したことによる列の変更
    With sprExamine
        m = .MaxRows
        ReDim tblSxlKSiji(m)
        j = 0
        For i = 1 To m - 1 Step 2
            '' Z品番ならテーブル作成
            .row = i

            'エピ先行評価追加対応
            .GetText 37, i, vFlg
            .GetText 2, i, vHinban

            If vHinban = "Z" Then
                j = j + 1

                'エピ先行評価追加対応
                .col = 39
                sNowBlockID = Mid(tblSXL.CRYNUM, 1, 9) & .text

                If i = 1 Then   '既存位置からSXLIDは作らない
                    intNowIngotPos = SIngotP
                    tblSxlKSiji(j).SXL_ID = tblSXL.SXLID
                Else
                    .col = 5
                    intNowIngotPos = .text
                    '' SXL-ID
                    tblSxlKSiji(j).SXL_ID = GetSXLID(sNowBlockID, intNowIngotPos)
                End If

'                If CutCntFlg = 1 Then
                'Change 2009/10/20 Y.Hitomi
                If ChkHosho(tblSXL.SXLID, sHosho) = FUNCTION_RETURN_SUCCESS Then
                    If sHosho = 1 Then ''ﾌﾞﾛｯｸ保証
                    
                        'Add Start 2012/03/08 Y.Hitomi　ブロック確定（廃棄）は、1回のみとする。
                        If iBlockHoshoCnt = 0 Then
                            iBlockHoshoCnt = iBlockHoshoCnt + 1
                        Else
                            j = j - 1
                            Exit For
                        End If
                        'Add End   2012/03/08 Y.Hitomi
                        
                        '' ブロックID
                        tblSxlKSiji(j).BLOCKID = sNowBlockID
                        '' サンプルID(From)
                        tblSxlKSiji(j).SAMPLE_FROM = ""
                        '' サンプルID(To)
                        tblSxlKSiji(j).SAMPLE_TO = ""
                    ElseIf sHosho = 2 Then 'WF保証
                        '' ブロックID
                        tblSxlKSiji(j).BLOCKID = ""
                        '' サンプルID(From)
                        If i = 1 Then
                            tblSxlKSiji(j).SAMPLE_FROM = tblSXL.WFSMP(1).REPSMPLIDCW
                        Else
                            'エピ先行評価追加対応
                            .GetText 38, i, vChgSmpId
                            tblSxlKSiji(j).SAMPLE_FROM = vChgSmpId  'サンプルIDを作り直す必要はない
                        End If
    
                        '関連ﾌﾞﾛｯｸ対応 08/08/25 ooba
                        .row = i + 1
                        .col = 1
                        If .CellType = CellTypeCheckBox Then
                            If .Value = "0" And i < m - 1 Then
                                .GetText 2, i + 2, vHinban
                                If vHinban = "Z" Then
                                    i = i + 2
                                End If
                            End If
                        End If
                        
                        '' サンプルID(To)
                        If i = m - 1 Then
                            tblSxlKSiji(j).SAMPLE_TO = tblSXL.WFSMP(2).REPSMPLIDCW
                        Else
                            'エピ先行評価追加対応
                            .GetText 38, i + 1, vChgSmpId
                            tblSxlKSiji(j).SAMPLE_TO = vChgSmpId
                        End If
    
                        ''廃棄ｻﾝﾌﾟﾙID共有対応
                        sSmpId(1) = tblSxlKSiji(j).SAMPLE_FROM
                        sSmpId(2) = tblSxlKSiji(j).SAMPLE_TO
    
                        '登録ｼﾝｸﾞﾙの各検査項目に対するｻﾝﾌﾟﾙ有無をﾁｪｯｸする
                        If chkComSAMPL(tblSXL.SXLID, sSmpId(1), sSmpId(1)) = FUNCTION_RETURN_SUCCESS Then
                            '実測がひとつもない場合、共有ｻﾝﾌﾟﾙIDを登録する
                            If sSmpId(1) <> tblSxlKSiji(j).SAMPLE_FROM Then
                                tblSxlKSiji(j).SAMPLE_FROM = sSmpId(1)
                            End If
                        End If
    
                        '登録ｼﾝｸﾞﾙの各検査項目に対するｻﾝﾌﾟﾙ有無をﾁｪｯｸする
                        If chkComSAMPL(tblSXL.SXLID, sSmpId(2), sSmpId(2)) = FUNCTION_RETURN_SUCCESS Then
                            '実測がひとつもない場合、共有ｻﾝﾌﾟﾙIDを登録する
                            If sSmpId(2) <> tblSxlKSiji(j).SAMPLE_TO Then
                                tblSxlKSiji(j).SAMPLE_TO = sSmpId(2)
                            End If
                        End If
                        ''廃棄ｻﾝﾌﾟﾙID共有対応
                    End If
                End If

                '' 確定品番
                tblSxlKSiji(j).hinban = tblSXL.hinban & Format(tblSXL.REVNUM, "00")

                '' 区分コード
                .row = i
                .col = 9
                intFuryoCmb = .TypeComboBoxCurSel + 1
                tblSxlKSiji(j).KUBUN = Trim$(tblPrcList(intFuryoCmb).CODE)       ' 不良理由

                '' トランザクションID
                tblSxlKSiji(j).TXID = "TX853I"

                ''比抵抗ﾃﾞｰﾀ取得
                tblSxlKSiji(j).MESDATA1TOP = " "
                tblSxlKSiji(j).MESDATA2TOP = " "
                tblSxlKSiji(j).MESDATA3TOP = " "
                tblSxlKSiji(j).MESDATA4TOP = " "
                tblSxlKSiji(j).MESDATA5TOP = " "
                tblSxlKSiji(j).MESDATA1BOT = " "
                tblSxlKSiji(j).MESDATA2BOT = " "
                tblSxlKSiji(j).MESDATA3BOT = " "
                tblSxlKSiji(j).MESDATA4BOT = " "
                tblSxlKSiji(j).MESDATA5BOT = " "
            End If
        Next i
    ReDim Preserve tblSxlKSiji(j)
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_MakeTBCMW005
'*
'*    処理概要      : 1.テーブル登録用にWF総合判定実績テーブル構造体の設定
'*                      (WF総合判定実績テーブル構造体設定)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_MakeTBCMW005()

    Dim sCode As String

    With sprExamine
        .row = 1
        .col = 2
        sCode = IIf(Trim$(.text) = "Z", "2", "1")
    End With

    With tblWfHantei
        .CRYNUM = tblSXL.CRYNUM             ' 結晶番号
        .INGOTPOS = tblSXL.INGOTPOS         ' インゴット位置
        .CRYLEN = tblSXL.COUNT              ' 長さ
        .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI ' 管理工程コード
        .PROCCODE = PROCD_WFC_SOUGOUHANTEI  ' 工程コード
        .SXLID = tblSXL.SXLID               ' SXLID
        .CODE = sCode                       ' 区分コード
        .TSTAFFID = txtStaffID.text         ' 登録社員ID
        .KSTAFFID = ""                      ' 更新社員ID
    End With
End Sub

'*******************************************************************************
'*    関数名        : sub_MakeTBCMW006
'*
'*    処理概要      : 1.テーブル登録用に振替廃棄実績テーブル構造体の設定
'*                      (振替廃棄実績テーブル構造体の設定)
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_MakeTBCMW006()
    Dim udtTmpHin       As tFullHinban
    Dim sHin            As String
    Dim intNowIngotPos  As Integer
    Dim intDwnIngotPos  As Integer
    Dim m               As Integer
    Dim i               As Integer
    Dim vNukisiFlg      As Variant
    Dim intRowCnt       As Integer

    With sprExamine
        m = .MaxRows
        ReDim tblHuriHai(0) '構造体は該当行の時のみ増やす
        intRowCnt = 0

        For i = 1 To m
            .row = i

            'エピ先行評価追加対応
            sprExamine.GetText 37, i, vNukisiFlg

            If (vNukisiFlg = 1) Or CheckGetSampleID(i) = True Then    '抜試行だったら
                If i = m Then
                    Exit For
                End If
                intRowCnt = intRowCnt + 1

                ReDim Preserve tblHuriHai(intRowCnt) '構造体は該当行の時のみ増やす
                .col = 5

                '先頭SXLの開始位置は前工程の開始位置となる
                If i = 1 Then
                    tblHuriHai(intRowCnt).INGOTPOS = SIngotP
                    intNowIngotPos = SIngotP
                Else
                    tblHuriHai(intRowCnt).INGOTPOS = .text
                    intNowIngotPos = .text
                End If

                '最終SXLの終了位置は前工程の終了位置となる
                If i = m - 2 Then
                    intDwnIngotPos = EIngotP
                Else
                    .row = i + 1
                    intDwnIngotPos = .text
                    .row = i
                End If

                .col = 2
                If .row Mod 2 = 0 Then
                    Call GetLastHinban("", udtTmpHin)           ' フル品番
                    sHin = ""
                Else
                    Call GetLastHinban(.text, udtTmpHin)        ' フル品番
                    sHin = Trim$(.text)
                End If

                With tblHuriHai(intRowCnt)
                    .CRYNUM = tblSXL.CRYNUM                     ' 結晶番号
                    If UBound(tblWafInd()) = 2 Then             '2行の時だけ、長さ取得処理変更
                        .CRYLEN = EIngotP - SIngotP
                    Else
                        .CRYLEN = intDwnIngotPos - intNowIngotPos     ' 長さ
                    End If
                    .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI         ' 管理工程コード
                    .PROCCODE = PROCD_WFC_SOUGOUHANTEI          ' 工程コード
                    .TRANCLS = IIf(sHin = "Z", "1", "0")        ' 処理区分
                    .DUOGNUM = tblSXL.hinban                    ' 転用元品番
                    .DUOGREV = tblSXL.REVNUM                    ' 転用元品番 製品番号改訂番号
                    .DUOGFACT = tblSXL.factory                  ' 転用元品番 工場
                    .DUOGOPCD = tblSXL.opecond                  ' 転用元品番 操業条件
                    .DUNWNUM = sHin                             ' 転用先品番
                    .DUNWREV = udtTmpHin.mnorevno               ' 転用先品番 製品番号改訂番号
                    .DUNWFACT = udtTmpHin.factory               ' 転用先品番 工場
                    .DUNWOPCD = udtTmpHin.opecond               ' 転用先品番 操業条件
                    .TSTAFFID = txtStaffID.text                 ' 登録社員ID
                    .KSTAFFID = ""                              ' 更新社員ID
                    .MUKESAKI = sCmbMukesaki
                End With
            End If
        Next i
    End With
End Sub

'*******************************************************************************
'*    関数名        : fnc_Nukisi_LOAD_DISP
'*
'*    処理概要      : 1.SXLIDからブロックＩＤ（はみ出た部分）取得
'*                    2.ブロック、SXLID毎のブロックＳＥＱ（ブロック内連番）の
'*                      最大、最小値とその他情報取得
'*                    3.検査項目取得
'*                    4.振替ﾁｪｯｸ(初期表示)
'*                    5.Warp/合成角度情報表示
'*                    6.チェックボックスを初期表示
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_Nukisi_LOAD_DISP() As FUNCTION_RETURN
    Dim intRtn          As Integer
    Dim sBlockId        As String
    Dim intIngotpos()   As Integer
    Dim intWfNum        As Integer
    Dim i, j            As Integer
    Dim intErrCode      As Integer
    Dim intErrMsg       As String

    fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE

    '前画面からのSXLIDを取得
     ReDim tSXLID(0)
     tSXLID(0).SXLID = tblSXL.SXLID

    ''SXLIDからブロックＩＤ（はみ出た部分）取得
    If DBDRV_BLOCKIDGET() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("ESXL2")
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''ブロック、SXLID毎のブロックＳＥＱ（ブロック内連番）の最大、最小値とその他情報取得
    If DBDRV_MIN_MAX_SEQGET(intWfNum) = FUNCTION_RETURN_FAILURE Then
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '検査項目取得
    If DVDRV_KENSA_KOUMOKU(tKensa()) = FUNCTION_RETURN_FAILURE Then
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '’画面表示
    Call sub_ExamineDisp(tKensa(), intIngotpos(), intWfNum)

    '振替ﾁｪｯｸ(初期表示)
    lblMsg.Caption = ""
    ReDim tWarpMeasG(0)
    ReDim tKakuMeasG(0)
    'Add Start 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura 結晶面傾きチェック追加対応

    For i = 1 To UBound(tMapHin)
        tMapHinG = tMapHin(i)
        For j = 1 To 2
            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                intRtn = funChkFurikaeShiyou("CW763", txtKSXLID.text, _
                                            tMapHinG.HIN, tMapHinG.HIN, _
                                            intErrCode, intErrMsg, typ_b, typ_CType, 0)

                tMapHin(i).WARPFLG = tMapHinG.WARPFLG       'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG       '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ

                '判定NG
                If intRtn = 1 Then
                    If Not tMapHinG.KAKUFLG Then
                        lblMsg.Caption = "Warp判定エラー　品番振替を行ってください。"
                    Else
                        lblMsg.Caption = "合成角度判定エラー　品番振替を行ってください。"
                    End If

                '振替ﾁｪｯｸｴﾗｰ
                ElseIf intRtn < 0 Then
                    fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
                    lblMsg.Caption = intErrMsg
                    Exit Function
                End If
            End If
        Next j
    Next i

    If bMapWarpFlg Then lblMsg.Caption = "WFマップとWarp実績の不一致エラー"

    'Warp/合成角度情報表示
    Call WarpKakuDisp(Me)

    'チェックボックスを初期表示
    Call sub_LOADDISP_ADD_CHECKBOX

    '振替チェック追加による修正
    Call sub_FurikaeMotoDataSet

    Erase intIngotpos

    fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_SUCCESS

End Function

'*******************************************************************************
'*    関数名        : sub_ExamineDisp
'*
'*    処理概要      : 1.sprExamineに表示する
'*
'*    パラメータ    : 変数名        ,IO ,型             ,説明
'*　　      　　      udtKensa        ,I  ,typ_XSDCW      ,新サンプル管理(SXL)
'*　　      　　      intIngotpos()   ,I  ,Integer        ,
'*　　      　　      intWfNum        ,I  ,Integer        ,WF枚数
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_ExamineDisp(ByRef udtKensa() As typ_XSDCW, ByRef intIngotpos() As Integer, ByRef intWfNum As Integer)
    Dim intUCount       As Integer
    Dim i, j, k         As Integer
    Dim intRtn          As Integer
    Dim sSmplID         As String
    Dim vInsertData     As Variant
    Dim intRowNo        As Integer
    Dim intUCount2      As Integer
    Dim dblSmplTopPos   As Double
    Dim intRowMod       As Integer
    Dim intKensaLoop    As Integer
    Dim m               As Integer
    Dim sList           As String
    Dim sBkSMPLID       As String
    Dim dblteikoku      As Double
    Dim intTingot       As Integer
    Dim intBingot       As Integer
    Dim vGetIngot       As Variant
    Dim blSampleChk     As Boolean
    Dim udtChkHin       As tFullHinban  'ﾁｪｯｸ用12桁品番
    Dim sBlkSmplID      As String       '結晶ｻﾝﾌﾟﾙ実績
    Dim sCryindFlg      As String       '結晶状態FLG
    Dim vGetHinban      As Variant      '品番
    Dim vGetBlockID     As Variant      'ﾌﾞﾛｯｸID
    Dim vGetBlockSEQ_S  As Variant      'ﾌﾞﾛｯｸSEQ(Start)
    Dim vGetBlockSEQ_E  As Variant      'ﾌﾞﾛｯｸSEQ(End)
    Dim vlvGetWFstatus  As Variant      'WF状態
    Dim sBuf            As String
    Dim sHinban         As String

    intUCount = UBound(tExamine)

    '品番を1列追加することによる列の変更
    With sprExamine
        ''ループ開始
        .MaxRows = 0
        .MaxRows = 1

        '結晶実績引継ぎﾃﾞｰﾀ初期化
        CpyCrySmpl.TsmplidGD = ""           'TOP_ｻﾝﾌﾟﾙID(GD)
        CpyCrySmpl.TindGD = ""              'TOP_状態FLG(GD)
        CpyCrySmpl.BsmplidGD = ""           'BOT_ｻﾝﾌﾟﾙID(GD)
        CpyCrySmpl.BindGD = ""              'BOT_状態FLG(GD)

        For i = 0 To intUCount
            blSampleChk = False
            .BlockMode = True
            .row = i + 1

            '' ブロックID
            If i = 0 Then   'i = 0処理と、elseif時、と同じ処理をする
                vInsertData = Right(tExamine(i).LOTID, 3)
                .SetText 1, i + 1, vInsertData
                .BlockMode = True
                .backColor = COLOR_DISABLE
                .BlockMode = False
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
            ElseIf tExamine(i).LOTID <> tExamine(i - 1).LOTID Then
                .col = 1
                vInsertData = Right(tExamine(i).LOTID, 3)
                .SetText 1, i + 1, vInsertData
                .BlockMode = True
                .backColor = COLOR_DISABLE
                .BlockMode = False
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
            End If

            ''　品番
            If (i = 0) Then
                .col = 2
                .SetText 2, i + 1, tExamine(i).hinban
                sHinban = tExamine(i).hinban

                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
                .Lock = False

                '' ＷＦ枚数
                .col = 7
                .SetText 7, i + 1, tExamine(i).CURRWPCS

                'エピ先行評価追加対応
                .SetText 43, i + 1, tExamine(i).CURRWPCS
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16

                '品番2追加
                .col = 3
                .SetText 3, i + 1, Format(tExamine(i).REVNUM, "00") & tExamine(i).factory & tExamine(i).opecond
                .backColor = COLOR_DISABLE
                .BlockMode = False
            ElseIf (tExamine(i).hinban <> tExamine(i - 1).hinban) Or (tExamine(i).LOTID <> tExamine(i - 1).LOTID) Then
                .col = 2
                .SetText 2, i + 1, tExamine(i).hinban
                sHinban = tExamine(i).hinban

                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
                .Lock = False

                '' ＷＦ枚数
                .col = 7
                .SetText 7, i + 1, tExamine(i).CURRWPCS

                'エピ先行評価追加対応
                .SetText 43, i + 1, tExamine(i).CURRWPCS
                .CellBorderType = 4     '罫線上
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16

                '品番2追加
                .col = 3
                .SetText 3, i + 1, Format(tExamine(i).REVNUM, "00") & tExamine(i).factory & tExamine(i).opecond
                .backColor = COLOR_DISABLE
                .BlockMode = False
            Else
                '’ブロックＰ
                sBuf = GetMukesaki(sHinban)
                .SetText 2, i + 1, sBuf
                sBaseMukesaki = sCmbMukeName

                .BlockMode = True
                .backColor = COLOR_DISABLE
                .BlockMode = False
            End If

            '’ブロックＰ
            .SetText 4, i + 1, CStr(tExamine(i).RTOP_POS)
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            ''　結晶Ｐ
            .SetText 5, i + 1, CStr(tExamine(i).RITOP_POS)
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            '’マップ位置
            .SetText 6, i + 1, Trim(CStr(tExamine(i).BLOCKSEQ))
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            '’ＷＦ状態
            'WF状態判定
            .col = 8
            .row = i + 1
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            Select Case CStr(tExamine(i).WFSTA)
                Case gsWF_STA_0   '通常
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                    'サンプルフラグ判定
                    Select Case CStr(tExamine(i).SHAFLAG)
                        Case gsWF_SMPL_0
                            .SetText 8, i + 1, gsWF_STA_NORMAL
                        Case gsWF_SMPL_1
                            .SetText 8, i + 1, gsWF_STA_SIJI
                        Case gsWF_SMPL_2
                            .SetText 8, i + 1, gsWF_STA_SIJI_OK
                        Case gsWF_SMPL_3
                            .SetText 8, i + 1, gsWF_STA_SIJI_NG
                        Case gsWF_SMPL_4
                            .SetText 8, i + 1, gsWF_STA_SIJI_KEKKA
                    End Select
                Case gsWF_STA_1   '共有
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                Case gsWF_STA_4   '欠落
                    'サンプルフラグ判定
                    Select Case CStr(tExamine(i).SHAFLAG)
                        Case gsWF_SMPL_0
                            .SetText 8, i + 1, gsWF_STA_NORMAL
                        Case gsWF_SMPL_1
                            .SetText 8, i + 1, gsWF_STA_SIJI
                        Case gsWF_SMPL_2
                            .SetText 8, i + 1, gsWF_STA_SIJI_OK
                        Case gsWF_SMPL_3
                            .SetText 8, i + 1, gsWF_STA_SIJI_NG
                        Case gsWF_SMPL_4
                            .SetText 8, i + 1, gsWF_STA_SIJI_KEKKA
                    End Select
            End Select

            '’不良区分
            .col = 9
            intRowMod = (i + 1) Mod 2
            If intRowMod = 0 Then
                .row = i + 1
                .CellType = CellTypeEdit
                .TypeHAlign = TypeHAlignLeft
                .TypeEditCharCase = TypeEditCharCaseSetNone
                .TypeEditMultiLine = False
                .TypeMaxEditLen = 60
                .col = 9
                .CellBorderType = 8     '罫線下
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
            End If

            '' 区分コンボの設定
            If i Mod 2 = 0 Then
                m = UBound(tblPrcList)
                sList = GetGPCodeDspStr(tblPrcList(1).CODE, tblPrcList(1).INFO1)

                For j = 2 To m
                    sList = sList & vbTab & GetGPCodeDspStr(tblPrcList(j).CODE, tblPrcList(j).INFO1)
                Next j

                .TypeComboBoxList = sList
                .TypeComboBoxCurSel = 0
                .Lock = False
            End If

            sSmplID = Mid(tExamine(i).SMPLEID, 10, 3) & "-" & Right(tExamine(i).SMPLEID, 4)
            sBkSMPLID = tExamine(i).SMPLEID

            Select Case tExamine(i).WFSTA
                Case "1" '共通
                    .SetText 10, i + 1, gsWF_SMPL_JOINT

                    'エピ先行評価追加対応
                    .SetText 38, i + 1, sBkSMPLID
                Case "0" '通常
                    If (tExamine(i).SHAFLAG <> "0") Then
                        .SetText 10, i + 1, sSmplID

                        'エピ先行評価追加対応
                        .SetText 38, i + 1, sBkSMPLID
                    End If
                Case "4" '欠落
                    If (tExamine(i).SHAFLAG <> "0") Then
                        .SetText 10, i + 1, sSmplID
                        'エピ先行評価追加対応
                        .SetText 38, i + 1, sBkSMPLID
                    End If
                Case Else
            End Select

            .col = 10
            .CellBorderType = 16    '罫線外枠
            .CellBorderStyle = 1
            .CellBorderColor = vbBlack
            .Action = 16

            '検査項目表示
            .col = 11    '３～９列目

            'エピ先行評価追加対応
            .col2 = 35
            .row = i + 1
            .row2 = i + 1
            .BlockMode = True
            .backColor = vbWhite
            .BlockMode = False

            j = i
            blSampleChk = True

            If blSampleChk = True And tExamine(i).SHAFLAG <> "0" Then
                .row = i + 1
                .row2 = i + 1
                If IsNumeric(udtKensa(j).WFINDRSCW) = True Then
                        .col = 11
                        .col2 = 11
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDRSCW = "0", "", udtKensa(j).WFINDRSCW)
                        .BlockMode = False
                    If .text = "1" Then    '0:検査なし ,1:通常,2:反映,3:推定
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDOICW) = True Then
                        .col = 12
                        .col2 = 12
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDOICW = "0", "", udtKensa(j).WFINDOICW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDB1CW) = True Then
                        .col = 13
                        .col2 = 13
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDB1CW = "0", "", udtKensa(j).WFINDB1CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDB2CW) = True Then
                        .col = 14
                        .col2 = 14
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDB2CW = "0", "", udtKensa(j).WFINDB2CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDB3CW) = True Then
                        .col = 15
                        .col2 = 15
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDB3CW = "0", "", udtKensa(j).WFINDB3CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDL1CW) = True Then
                        .col = 16
                        .col2 = 16
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDL1CW = "0", "", udtKensa(j).WFINDL1CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDL2CW) = True Then
                        .col = 17
                        .col2 = 17
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDL2CW = "0", "", udtKensa(j).WFINDL2CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDL3CW) = True Then
                        .col = 18
                        .col2 = 18
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDL3CW = "0", "", udtKensa(j).WFINDL3CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDL4CW) = True Then
                        .col = 19
                        .col2 = 19
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDL4CW = "0", "", udtKensa(j).WFINDL4CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                        .backColor = vbYellow
'''                        .ForeColor = vbYellow
                        
                        ' ｸﾞﾚｰ
                        .backColor = COLOR_CryJitsu
                        .ForeColor = COLOR_CryJitsu
                        '◆--- 2010/01/20 SIRD対応 SPK habuki REP  END
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDDSCW) = True Then
                        .col = 20
                        .col2 = 20
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDDSCW = "0", "", udtKensa(j).WFINDDSCW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDDZCW) = True Then
                        .col = 21
                        .col2 = 21
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDDZCW = "0", "", udtKensa(j).WFINDDZCW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDSPCW) = True Then
                        .col = 22
                        .col2 = 22
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDSPCW = "0", "", udtKensa(j).WFINDSPCW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDDO1CW) = True Then
                        .col = 23
                        .col2 = 23
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDDO1CW = "0", "", udtKensa(j).WFINDDO1CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDDO2CW) = True Then
                        .col = 24
                        .col2 = 24
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDDO2CW = "0", "", udtKensa(j).WFINDDO2CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDDO3CW) = True Then
                        .col = 25
                        .col2 = 25
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDDO3CW = "0", "", udtKensa(j).WFINDDO3CW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                ''残存酸素追加
                If IsNumeric(udtKensa(j).WFINDAOICW) = True Then
                        .col = 26
                        .col2 = 26
                        .BlockMode = True
                        .text = IIf(udtKensa(j).WFINDAOICW = "0", "", udtKensa(j).WFINDAOICW)
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                If IsNumeric(udtKensa(j).WFINDOT1CW) = True Then
                        .col = 27
                        .col2 = 27
                        .BlockMode = True
                       .text = IIf(udtKensa(j).WFINDOT1CW = "0", "", udtKensa(j).WFINDOT1CW)
                       .ForeColor = vbWhite
                        .BlockMode = False
                    If .text = "1" Then
                        .backColor = vbBlack
                        .ForeColor = vbBlack
                    ElseIf .text = "0" Then
                    End If
                End If
                If IsNumeric(udtKensa(j).WFINDOT2CW) = True Then
                    'エピ先行評価追加対応
                    .col = 35
                    .col2 = 35
                    .BlockMode = True
                    .text = IIf(udtKensa(j).WFINDOT2CW = "0", "", udtKensa(j).WFINDOT2CW)
                    .ForeColor = vbWhite
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                        .ForeColor = vbBlack
                    ElseIf .text = "0" Then
                    End If
                End If

                'GD追加
                If IsNumeric(udtKensa(j).WFINDGDCW) = True Then
                    'エピ先行評価追加対応
                    .col = 28
                    .col2 = 28

                    .BlockMode = True
                    .text = IIf(udtKensa(j).WFINDGDCW = "0", "", udtKensa(j).WFINDGDCW)
                    .BlockMode = False

                    '実測
                    If .text = "1" And udtKensa(j).WFHSGDCW <> "1" Then
                        .backColor = vbBlack
                    '反映
                    ElseIf .text = "2" And udtKensa(j).WFHSGDCW <> "1" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    '結晶実績
                    ElseIf (.text = "1" Or .text = "2") And udtKensa(j).WFHSGDCW = "1" Then
                        .backColor = COLOR_CryJitsu
                        .ForeColor = COLOR_CryJitsu
                    End If
                End If

                'エピ先行評価追加対応
                'BMD1E
                If IsNumeric(udtKensa(j).EPINDB1CW) = True Then
                    .col = 29
                    .col2 = 29
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDB1CW = "0", "", udtKensa(j).EPINDB1CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                'BMD2E
                If IsNumeric(udtKensa(j).EPINDB2CW) = True Then
                    .col = 30
                    .col2 = 30
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDB2CW = "0", "", udtKensa(j).EPINDB2CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                'BMD3E
                If IsNumeric(udtKensa(j).EPINDB3CW) = True Then
                    .col = 31
                    .col2 = 31
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDB3CW = "0", "", udtKensa(j).EPINDB3CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                'OSF1E
                If IsNumeric(udtKensa(j).EPINDL1CW) = True Then
                    .col = 32
                    .col2 = 32
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDL1CW = "0", "", udtKensa(j).EPINDL1CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                'OSF2E
                If IsNumeric(udtKensa(j).EPINDL2CW) = True Then
                    .col = 33
                    .col2 = 33
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDL2CW = "0", "", udtKensa(j).EPINDL2CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If

                'OSF3E
                If IsNumeric(udtKensa(j).EPINDL3CW) = True Then
                    .col = 34
                    .col2 = 34
                    .BlockMode = True
                    .text = IIf(udtKensa(j).EPINDL3CW = "0", "", udtKensa(j).EPINDL3CW)
                    .BlockMode = False

                    If .text = "1" Then
                        .backColor = vbBlack
                    ElseIf .text = "2" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    End If
                End If
            End If

            '' 結晶GD反映対応
            'SXLのTOP位置とBOT位置の結晶実績を取得する
            If (i = 0) Or (i = intUCount) Then
                udtChkHin.hinban = ""
                udtChkHin.mnorevno = 0
                udtChkHin.factory = ""
                udtChkHin.opecond = ""

                '引継ぎﾃﾞｰﾀが存在すればその結晶ｻﾝﾌﾟﾙID／結晶状態FLGをｾｯﾄ
                '--GD
                If funBlkSmpDataGet(udtKensa(j).SMCRYNUMCW, udtKensa(j).TBKBNCW, udtKensa(j).INPOSCW, _
                                        udtChkHin, 1, sBlkSmplID, sCryindFlg) = FUNCTION_RETURN_SUCCESS Then
                    'TOP
                    If udtKensa(j).TBKBNCW = "T" Then
                        CpyCrySmpl.TsmplidGD = sBlkSmplID       'ｻﾝﾌﾟﾙID(GD)
                        CpyCrySmpl.TindGD = sCryindFlg          '状態FLG(GD)
                    'BOT
                    ElseIf udtKensa(j).TBKBNCW = "B" Then
                        CpyCrySmpl.BsmplidGD = sBlkSmplID       'ｻﾝﾌﾟﾙID(GD)
                        CpyCrySmpl.BindGD = sCryindFlg          '状態FLG(GD)
                    End If
                End If
            End If

            '削除ボタンの消去禁止フラグ
            If (i = 0) Or (i = intUCount) Then
                'エピ先行評価追加対応
                .SetText 37, i + 1, "1"
            Else
                'エピ先行評価追加対応
                .SetText 37, i + 1, "3"
            End If

            'エピ先行評価追加対応
            .SetText 39, i + 1, Right(tExamine(i).LOTID, 3)

            '予備フィールドに品番保存
            'エピ先行評価追加対応
            .SetText 40, i + 1, tExamine(i).hinban

            .BlockMode = False
            .MaxRows = .MaxRows + 1
        Next i

        .MaxRows = .MaxRows - 1

        '色をつける
        .col = 1    '１列目
        .col2 = 1
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = COLOR_DISABLE
        .BlockMode = False

        .col = 2    '２列目
        .col2 = 2
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = vbWhite
        .BlockMode = False

        .col = 3    '3列目
        .col2 = 3
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = COLOR_DISABLE
        .BlockMode = False

        .col = 4    '３～９列目
        .col2 = 10
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = COLOR_DISABLE
        .BlockMode = False

        .col = 1
        .col2 = 1
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .CellBorderType = 16
        .CellBorderStyle = CellBorderStyleSolid
        .CellBorderColor = vbBlack
        .Action = ActionSetCellBorder
        .BlockMode = False

        'WFﾏｯﾌﾟ上の品番情報取得
        ReDim tMapHin(0)
        m = 0
        For i = 1 To .MaxRows Step 2
            '品番取得
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                For j = 0 To UBound(tExamine)
                    If tExamine(j).hinban = CStr(vGetHinban) Then
                        m = m + 1
                        ReDim Preserve tMapHin(m)

                        tMapHin(m).HIN.hinban = tExamine(j).hinban
                        tMapHin(m).HIN.mnorevno = tExamine(j).REVNUM
                        tMapHin(m).HIN.factory = tExamine(j).factory
                        tMapHin(m).HIN.opecond = tExamine(j).opecond

                        'ﾌﾞﾛｯｸID取得
                        .GetText 39, i, vGetBlockID
                        tMapHin(m).BLOCKID = left(txtCryNum.text, 9) & CStr(vGetBlockID)

                        'ﾌﾞﾛｯｸSEQ取得
                        .GetText 6, i, vGetBlockSEQ_S
                        .GetText 6, i + 1, vGetBlockSEQ_E
                        tMapHin(m).BLKSEQ_S = CInt(vGetBlockSEQ_S)
                        tMapHin(m).BLKSEQ_E = CInt(vGetBlockSEQ_E)

                        '振替ﾁｪｯｸﾌﾗｸﾞ
                        tMapHin(m).WARPFLG = False
                        tMapHin(m).KAKUFLG = False

                        'Add Start 2011/04/25 SMPK Miyata
                        tMapHin(m).XTALCS = txtCryNum.text      '結晶番号
                        .GetText 5, i, vGetIngot
                        tMapHin(m).INPOSCS_S = CInt(vGetIngot)  '結晶内位置(Start)
                        .GetText 5, i + 1, vGetIngot
                        tMapHin(m).INPOSCS_E = CInt(vGetIngot)  '結晶内位置(End)
                        'Add End   2011/04/25 SMPK Miyata

                        Exit For
                    End If
                Next j
            End If
        Next i

        '熱処理判断処理追加
        'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
        For i = 1 To .MaxRows Step 2
            '品番取得
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .col = 10
                'WF状態取得(TOP)
                .GetText 8, i, vlvGetWFstatus
                'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
                If Trim(vlvGetWFstatus) = gsWF_STA_SIJI_KEKKA Then
                    .row = i
                    .backColor = f_cmbc039_3.Label12.backColor
                End If
                'WF状態取得(BOTTOM)
                .GetText 8, i + 1, vlvGetWFstatus
                'WF状態が"結果"の場合、サンプルIDの背景色を水色にする
                If Trim(vlvGetWFstatus) = gsWF_STA_SIJI_KEKKA Then
                    .row = i + 1
                    .backColor = f_cmbc039_3.Label12.backColor
                End If

            End If
        Next i
    End With

    'ねらい抵抗の上限下限を画面範囲に変更して表示する
    With sprExamine
        .GetText 5, 1, vGetIngot
        intTingot = CInt(vGetIngot)
        .GetText 5, .MaxRows, vGetIngot
        intBingot = CInt(vGetIngot)
    End With

    'トップ側抵抗値
    Call sub_Top_Btm_TEIKOU(intTingot, dblteikoku)
    txtTopRsltR.text = CStr(Format(dblteikoku, "0.0000"))

    'ボトム側抵抗値
    Call sub_Top_Btm_TEIKOU(intBingot, dblteikoku)
    txtBotRsltR.text = CStr(Format(dblteikoku, "0.0000"))

    End Sub

'*******************************************************************************************
'*    関数名        : sprExamine_ButtonClicked
'*
'*    処理概要      : 1.サンプルフラグオフにする
'*
'*    パラメータ    : 変数名      ,IO ,型      ,説明
'*　　      　　      col         ,I  ,Long    ,ボタンがクリックされたセルの列番号
'*　　      　　      Row()       ,I  ,Long    ,ボタンがクリックされたセルの行番号
'*　　      　　      ButtonDown  ,I  ,Integer ,保持型ボタンの状態（チェックボックスの状態）
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sprExamine_ButtonClicked(ByVal col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    With sprExamine
        If (col <> 1) Then
            Exit Sub
        End If

        '' サンプルフラグオフ
        bSampFlag = False
    End With
End Sub

'*******************************************************************************************
'*    関数名        : sprExamine_Click
'*
'*    処理概要      : 1.sprExamine_Click処理
'*                      (検査項目の有無を変更)
'*
'*    パラメータ    : 変数名      ,IO ,型      ,説明
'*　　      　　      col         ,I  ,Long    ,列
'*　　      　　      Row         ,I  ,Long    ,行
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sprExamine_Click(ByVal col As Long, ByVal row As Long)
    Dim vSMPLID         As Variant
    Dim intRtn          As Integer
    Dim intNukisiFlg    As Integer
    Dim vNukisiFlg      As Variant
    Dim intLoopCnt      As Integer
    Dim vKensa          As Variant
    Dim lngRow          As Long
    Dim vSamp           As Variant
    Dim sGetKensa       As String
    Dim vGetHinUp       As Variant
    Dim vGetHinDn       As Variant
    Dim udtHinUp        As tFullHinban
    Dim udtHinDn        As tFullHinban
    Dim udtHinban       As tFullHinban
    Dim vGetSample      As Variant
    Dim intNo           As Integer
    Dim vGetHin         As Variant
    Dim sOT1            As String
    Dim sOT2            As String
    Dim sKensaFlg       As String
    Dim sKensa          As String
    Dim i               As Integer
    Dim intk            As Integer
    Dim inth            As Integer
    Dim sHin            As String
    Dim intR1           As Integer
    Dim sMAI1           As String
    Dim sMAI2           As String

    '品番を1列追加したことによる列の変更
    With sprExamine
        If (col < 10) Then
            Exit Sub
        End If

        '品番がZのときは処理をしない
        If row Mod 2 = 0 Then
            .col = 2
            .row = row - 1
            sHin = Trim(.text)
        Else
            .col = 2
            .row = row
            sHin = Trim(.text)
        End If

        If sHin = "Z" Then
            Exit Sub
        End If

        'エピ先行評価追加対応
        If col = 27 Or col = 35 Then
            .GetText 7, row, vSamp

            If vSamp = "欠落" Then
                Exit Sub
            End If

            'エピ先行評価追加対応
            .GetText 37, row, vNukisiFlg

            If (vNukisiFlg <> "9") And (vNukisiFlg <> "3") And (vNukisiFlg <> "1") Then
                .col = col
                .row = row

                If .backColor = vbWhite Then
                    'add start 2003/05/26 hitec)matsumoto ------------------------
                    If row Mod 2 = 0 Then
                        .GetText 2, row - 1, vGetHin
                    Else
                        .GetText 2, row, vGetHin
                    End If
                    udtHinban.hinban = vbNullString   '初期化
                    udtHinban.factory = vbNullString
                    udtHinban.mnorevno = 0
                    udtHinban.opecond = vbNullString
                    For intLoopCnt = 1 To UBound(tblsiyou)  '検査仕様構造体と比較し、該当のフル桁品番取得
                        If tblHinbanRs(intLoopCnt).HIN.hinban = Trim(vGetHin) Then
                            udtHinban.hinban = tblHinbanRs(intLoopCnt).HIN.hinban
                            udtHinban.factory = tblHinbanRs(intLoopCnt).HIN.factory
                            udtHinban.mnorevno = tblHinbanRs(intLoopCnt).HIN.mnorevno
                            udtHinban.opecond = tblHinbanRs(intLoopCnt).HIN.opecond
                            Exit For
                        End If
                    Next
                    If udtHinban.hinban = vbNullString Then   '該当の品番が構造体になかった場合何もしない
                        Exit Sub
                    End If
                    If scmzc_getE036(udtHinban, sOT1, sOT2, sMAI1, sMAI2) = FUNCTION_RETURN_FAILURE Then
                        Exit Sub    'エラーの場合何もしない
                    End If

                    ''残存酸素検査項目追加による変更
                    If col = 27 Then
                        sKensaFlg = sOT1
                    Else
                        sKensaFlg = sOT2
                    End If
                    If sKensaFlg = "1" Then
                        .backColor = vbBlack
                        .ForeColor = vbBlack
                        .SetText col, row, "1"
                    End If
                Else
                    ' 必ずE036テーブルをチェックするように変更したので処理追加
                    ' 黒（実測）→白（無し）
                    .backColor = vbWhite
                    .ForeColor = vbWhite
                    .SetText col, row, ""
                End If
            End If

            Exit Sub
        End If

        .GetText 37, row, vNukisiFlg

        If vNukisiFlg <> "1" Then
            .col = col
            .row = row

            If .Lock = False And .backColor <> vbWhite Then
                If col <> 28 Then
                    .GetText col, row, vKensa
                    '◆--- 2010/01/20 SIRD対応 SPK habuki REP START
'''                    If vKensa = "2" Then
'''                        vKensa = "1"
'''                        Call sub_Cksmp(row, col, .text)
'''                        .ForeColor = vbBlack
'''                        .backColor = vbBlack
'''                    ElseIf vKensa = "1" Then
'''                        vKensa = "2"
'''                        Call sub_Cksmp(row, col, .text)
'''                        .ForeColor = vbYellow
'''                        .backColor = vbYellow
'''                    End If
                    
                    If col = 19 Then
                        If vKensa = "2" Then
                            vKensa = "1"
                            Call sub_Cksmp(row, col, .text)     ' ｸﾞﾚｰ→黒
                            .ForeColor = vbBlack
                            .backColor = vbBlack
                        ElseIf vKensa = "1" Then
                            vKensa = "2"
                            Call sub_Cksmp(row, col, .text)     ' 黒→ｸﾞﾚｰ
                            .ForeColor = COLOR_CryJitsu
                            .backColor = COLOR_CryJitsu
                        End If
                    Else
                        If vKensa = "2" Then
                            vKensa = "1"
                            Call sub_Cksmp(row, col, .text)
                            .ForeColor = vbBlack
                            .backColor = vbBlack
                        ElseIf vKensa = "1" Then
                            vKensa = "2"
                            Call sub_Cksmp(row, col, .text)
                            .ForeColor = vbYellow
                            .backColor = vbYellow
                        End If
                    End If
                    
                    '◆--- 2010/01/20 SIRD対応 SPK habuki REP  END
                    .SetText col, row, vKensa
                Else
                    'ｻﾝﾌﾟﾙ未処理の場合は処理を抜ける
                    If bSampFlag = False Then Exit Sub
                    .GetText col, row, vKensa
                    '反映／結晶実績　→　実測
                    If vKensa = "2" Then
                        vKensa = "1"
                        Call sub_Cksmp(row, col, .text)
                        .ForeColor = vbBlack
                        .backColor = vbBlack
                    ElseIf vKensa = "1" Then
                        vKensa = "2"
                        Call sub_Cksmp(row, col, .text)
                        '実測　→　反映
                        If tblNukishi(row).WFHSGDCW <> "1" Then
                            .ForeColor = vbYellow
                            .backColor = vbYellow
                        '実測　→　結晶実績
                        Else
                            .ForeColor = COLOR_CryJitsu
                            .backColor = COLOR_CryJitsu
                        End If
                    End If
                    .SetText col, row, vKensa
                End If

                If row Mod 2 = 0 Then
                    intR1 = row + 1
                Else
                    intR1 = row - 1
                End If

                If tWafk(row).SAMPLEID = tWafk(intR1).SAMPLEID Then
                    .GetText 10, row, vSMPLID

                    If vSMPLID = gsWF_SMPL_JOINT Then
                        .GetText 44, row, vSMPLID   'UD別保存のサンプルID
                        .SetText 38, row, vSMPLID

                        .SetText 10, row, Trim$(Mid(tblNukishi(row).REPSMPLIDCW, 10, 3) & "-" & Right(tblNukishi(row).REPSMPLIDCW, 4))
                        .SetText 8, row, gsWF_STA_SIJI
                    Else
                        For i = 0 To 24
                            If i <> 16 And i <> 24 Then
                                .col = i + 11
                                .row = row
                                sKensa = .text
                                If sKensa <> "" And sKensa <> "0" Then
                                    intk = intk + 1

                                    If sKensa = "2" Or .backColor = COLOR_CryJitsu Then   '05/02/24 ooba
                                        inth = inth + 1
                                    End If
                                End If
                            End If
                        Next i
                        If intk = inth Then '検査項目が全て黄色(反映)
                            .SetText 10, row, gsWF_SMPL_JOINT
                            .SetText 8, row, gsWF_STA_NORMAL

                            '別の構造体からコピー
                            vSMPLID = tWafk(row).SAMPLEID
                            .SetText 38, row, vSMPLID
                        Else '実データをとる場合
                            .GetText 44, row, vSMPLID
                            .SetText 38, row, vSMPLID
                        End If
                    End If
                End If
            End If
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START(TOP,BOTは検査済みのため変更不可)
'''        Else
'''            '<vNukisiFlg="1">・・・col=37
'''            .col = col: .row = row
'''            If .Lock = False Then
'''                If col = 19 Then
'''                    If .text = "1" Then
'''                        .SetText col, row, "2"  ' 黒→ｸﾞﾚｰ
'''                        .backColor = COLOR_CryJitsu
'''                        .ForeColor = COLOR_CryJitsu
'''                    Else
'''                        .SetText col, row, "1"  ' ｸﾞﾚｰ→黒
'''                        .backColor = vbBlack
'''                        .ForeColor = vbBlack
'''                    End If
'''                End If
'''            End If
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD  END
        End If
    End With
End Sub

'*******************************************************************************************
'*    関数名        : sub_Cksmp
'*
'*    処理概要      : 1.検査項目がクリックされたときにサンプルIDを変更する
'*
'*
'*    パラメータ    : 変数名      ,IO ,型      ,説明
'*　　      　　      R           ,I  ,Long    ,ボタンがクリックされたセルの列番号
'*　　      　　      C           ,I  ,Long    ,ボタンがクリックされたセルの行番号
'*　　      　　      sHkbn        ,I  ,String  ,RCで選択されたTextの値
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_Cksmp(R As Long, C As Long, sHkbn As String)
    If sHkbn = "2" Then '黄色→黒
        Select Case C
            Case 11
                tblNukishi(R).WFSMPLIDRSCW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESRS1CW = "0"
            Case 12
                tblNukishi(R).WFSMPLIDOICW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESOICW = "0"
            Case 13
                tblNukishi(R).WFSMPLIDB1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESB1CW = "0"
            Case 14
                tblNukishi(R).WFSMPLIDB2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESB2CW = "0"
            Case 15
                tblNukishi(R).WFSMPLIDB3CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESB3CW = "0"
            Case 16
                tblNukishi(R).WFSMPLIDL1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESL1CW = "0"
            Case 17
                tblNukishi(R).WFSMPLIDL2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESL2CW = "0"
            Case 18
                tblNukishi(R).WFSMPLIDL3CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESL3CW = "0"
            Case 19
                tblNukishi(R).WFSMPLIDL4CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESL4CW = "0"
            Case 20
                tblNukishi(R).WFSMPLIDDSCW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESDSCW = "0"
            Case 21
                tblNukishi(R).WFSMPLIDDZCW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESDZCW = "0"
            Case 22
                tblNukishi(R).WFSMPLIDSPCW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESSPCW = "0"
            Case 23
                tblNukishi(R).WFSMPLIDDO1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESDO1CW = "0"
            Case 24
                tblNukishi(R).WFSMPLIDDO2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESDO2CW = "0"
            Case 25
                tblNukishi(R).WFSMPLIDDO3CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESDO3CW = "0"
            ''残存酸素追加
            Case 26
                tblNukishi(R).WFSMPLIDAOICW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESAOICW = "0"
            Case 27
                tblNukishi(R).WFSMPLIDOT1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESOT1CW = "0"
            ' エピ先行評価追加
            Case 35
                tblNukishi(R).WFSMPLIDOT2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESOT2CW = "0"
            'GD追加
            Case 28
                tblNukishi(R).WFSMPLIDGDCW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESGDCW = "0"
                tblNukishi(R).WFHSGDCW = "0"
            Case 29
                tblNukishi(R).EPSMPLIDB1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESB1CW = "0"
            Case 30
                tblNukishi(R).EPSMPLIDB2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESB2CW = "0"
            Case 31
                tblNukishi(R).EPSMPLIDB3CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESB3CW = "0"
            Case 32
                tblNukishi(R).EPSMPLIDL1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESL1CW = "0"
            Case 33
                tblNukishi(R).EPSMPLIDL2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESL2CW = "0"
            Case 34
                tblNukishi(R).EPSMPLIDL3CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).EPRESL3CW = "0"
        End Select

        'Zかどうかの判断をしてZのときはサンプルIDをコピーし実績フラグを0にする
        Call sub_Zsample(R, C, sHkbn)
    ElseIf sHkbn = "1" Then '→黒→黄色
        Select Case C
            Case 11
                tblNukishi(R).WFSMPLIDRSCW = tblKns(R).WFSMPLIDRSCW
                tblNukishi(R).WFRESRS1CW = tblKns(R).WFRESRS1CW
            Case 12
                tblNukishi(R).WFSMPLIDOICW = tblKns(R).WFSMPLIDOICW
                tblNukishi(R).WFRESOICW = tblKns(R).WFRESOICW
            Case 13
                tblNukishi(R).WFSMPLIDB1CW = tblKns(R).WFSMPLIDB1CW
                tblNukishi(R).WFRESB1CW = tblKns(R).WFRESB1CW
            Case 14
                tblNukishi(R).WFSMPLIDB2CW = tblKns(R).WFSMPLIDB2CW
                tblNukishi(R).WFRESB2CW = tblKns(R).WFRESB2CW
            Case 15
                tblNukishi(R).WFSMPLIDB3CW = tblKns(R).WFSMPLIDB3CW
                tblNukishi(R).WFRESB3CW = tblKns(R).WFRESB3CW
            Case 16
                tblNukishi(R).WFSMPLIDL1CW = tblKns(R).WFSMPLIDL1CW
                tblNukishi(R).WFRESL1CW = tblKns(R).WFRESL1CW
            Case 17
                tblNukishi(R).WFSMPLIDL2CW = tblKns(R).WFSMPLIDL2CW
                tblNukishi(R).WFRESL2CW = tblKns(R).WFRESL2CW
            Case 18
                tblNukishi(R).WFSMPLIDL3CW = tblKns(R).WFSMPLIDL3CW
                tblNukishi(R).WFRESL3CW = tblKns(R).WFRESL3CW
            Case 19
                tblNukishi(R).WFSMPLIDL4CW = tblKns(R).WFSMPLIDL4CW
                tblNukishi(R).WFRESL4CW = tblKns(R).WFRESL4CW
            Case 20
                tblNukishi(R).WFSMPLIDDSCW = tblKns(R).WFSMPLIDDSCW
                tblNukishi(R).WFRESDSCW = tblKns(R).WFRESDSCW
            Case 21
                tblNukishi(R).WFSMPLIDDZCW = tblKns(R).WFSMPLIDDZCW
                tblNukishi(R).WFRESDZCW = tblKns(R).WFRESDZCW
            Case 22
                tblNukishi(R).WFSMPLIDSPCW = tblKns(R).WFSMPLIDSPCW
                tblNukishi(R).WFRESSPCW = tblKns(R).WFRESSPCW
            Case 23
                tblNukishi(R).WFSMPLIDDO1CW = tblKns(R).WFSMPLIDDO1CW
                tblNukishi(R).WFRESDO1CW = tblKns(R).WFRESDO1CW
            Case 24
                tblNukishi(R).WFSMPLIDDO2CW = tblKns(R).WFSMPLIDDO2CW
                tblNukishi(R).WFRESDO2CW = tblKns(R).WFRESDO2CW
            Case 25
                tblNukishi(R).WFSMPLIDDO3CW = tblKns(R).WFSMPLIDDO3CW
                tblNukishi(R).WFRESDO3CW = tblKns(R).WFRESDO3CW
            ''残存酸素追加
            Case 26
                tblNukishi(R).WFSMPLIDAOICW = tblKns(R).WFSMPLIDAOICW
                tblNukishi(R).WFRESAOICW = tblKns(R).WFRESAOICW
            Case 27
                tblNukishi(R).WFSMPLIDOT1CW = tblKns(R).WFSMPLIDOT1CW
                tblNukishi(R).WFRESOT1CW = tblKns(R).WFRESOT1CW
            ' エピ先行評価追加
            Case 35
                tblNukishi(R).WFSMPLIDOT2CW = tblKns(R).WFSMPLIDOT2CW
                tblNukishi(R).WFRESOT2CW = tblKns(R).WFRESOT2CW
            'GD追加
            Case 28
                tblNukishi(R).WFSMPLIDGDCW = tblKns(R).WFSMPLIDGDCW
                tblNukishi(R).WFRESGDCW = tblKns(R).WFRESGDCW
                tblNukishi(R).WFHSGDCW = tblKns(R).WFHSGDCW
            Case 29
                tblNukishi(R).EPSMPLIDB1CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESB1CW = tblKns(R).EPRESB1CW
            Case 30
                tblNukishi(R).EPSMPLIDB2CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESB2CW = tblKns(R).EPRESB2CW
            Case 31
                tblNukishi(R).EPSMPLIDB3CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESB3CW = tblKns(R).EPRESB3CW
            Case 32
                tblNukishi(R).EPSMPLIDL1CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESL1CW = tblKns(R).EPRESL1CW
            Case 33
                tblNukishi(R).EPSMPLIDL2CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESL2CW = tblKns(R).EPRESL2CW
            Case 34
                tblNukishi(R).EPSMPLIDL3CW = tblKns(R).REPSMPLIDCW
                tblNukishi(R).EPRESL3CW = tblKns(R).EPRESL3CW
        End Select

        Call sub_Zsample(R, C, sHkbn)
    End If
End Sub

'*******************************************************************************************
'*    関数名        : sub_Zsample
'*
'*    処理概要      : 1.上または下がZ品番だったら
'*　　　　　　　　　　　Zに選択されたデータのサンプルIDをコピーし実績フラグを0にする
'*　　　　　　　　　　　構造体にデータをセットする
'*
'*    パラメータ    : 変数名      ,IO ,型      ,説明
'*　　      　　      R           ,I  ,Long    ,ボタンがクリックされたセルの列番号
'*　　      　　      C           ,I  ,Long    ,ボタンがクリックされたセルの行番号
'*　　      　　      sHkbn        ,I  ,String  ,RCで選択されたTextの値
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_Zsample(R As Long, C As Long, sHkbn As String)
    'クリックした品番の上品番または下品番がZだった場合
    'Zに選択されたデータのサンプルIDをコピーし実績フラグを0にする
    '構造体にデータをセットする
    '3のとき(上品番がZ)
    If R <= 2 Then
        Exit Sub
    End If

    If sHkbn = "2" Then '黄色を黒に
        If R Mod 2 = 0 Then
        '4のとき(下品番Z)
            If Trim(tblNukishi(R + 1).HINBCW) = "Z" Then
                Select Case C
                    Case 11
                        tblNukishi(R + 1).WFSMPLIDRSCW = tblNukishi(R).WFSMPLIDRSCW
                        tblNukishi(R + 1).WFRESRS1CW = "0"
                    Case 12
                        tblNukishi(R + 1).WFSMPLIDOICW = tblNukishi(R).WFSMPLIDOICW
                        tblNukishi(R + 1).WFRESOICW = "0"
                    Case 13
                        tblNukishi(R + 1).WFSMPLIDB1CW = tblNukishi(R).WFSMPLIDB1CW
                        tblNukishi(R + 1).WFRESB1CW = "0"
                    Case 14
                        tblNukishi(R + 1).WFSMPLIDB2CW = tblNukishi(R).WFSMPLIDB2CW
                        tblNukishi(R + 1).WFRESB2CW = "0"
                    Case 15
                        tblNukishi(R + 1).WFSMPLIDB3CW = tblNukishi(R).WFSMPLIDB3CW
                        tblNukishi(R + 1).WFRESB3CW = "0"
                    Case 16
                        tblNukishi(R + 1).WFSMPLIDL1CW = tblNukishi(R).WFSMPLIDL1CW
                        tblNukishi(R + 1).WFRESL1CW = "0"
                    Case 17
                        tblNukishi(R + 1).WFSMPLIDL2CW = tblNukishi(R).WFSMPLIDL2CW
                        tblNukishi(R + 1).WFRESL2CW = "0"
                    Case 18
                        tblNukishi(R + 1).WFSMPLIDL3CW = tblNukishi(R).WFSMPLIDL3CW
                        tblNukishi(R + 1).WFRESL3CW = "0"
                    Case 19
                        tblNukishi(R + 1).WFSMPLIDL4CW = tblNukishi(R).WFSMPLIDL4CW
                        tblNukishi(R + 1).WFRESL4CW = "0"
                    Case 20
                        tblNukishi(R + 1).WFSMPLIDDSCW = tblNukishi(R).WFSMPLIDDSCW
                        tblNukishi(R + 1).WFRESDSCW = "0"
                    Case 21
                        tblNukishi(R + 1).WFSMPLIDDZCW = tblNukishi(R).WFSMPLIDDZCW
                        tblNukishi(R + 1).WFRESDZCW = "0"
                    Case 22
                        tblNukishi(R + 1).WFSMPLIDSPCW = tblNukishi(R).WFSMPLIDSPCW
                        tblNukishi(R + 1).WFRESSPCW = "0"
                    Case 23
                        tblNukishi(R + 1).WFSMPLIDDO1CW = tblNukishi(R).WFSMPLIDDO1CW
                        tblNukishi(R + 1).WFRESDO1CW = "0"
                    Case 24
                        tblNukishi(R + 1).WFSMPLIDDO2CW = tblNukishi(R).WFSMPLIDDO2CW
                        tblNukishi(R + 1).WFRESDO2CW = "0"
                    Case 25
                        tblNukishi(R + 1).WFSMPLIDDO3CW = tblNukishi(R).WFSMPLIDDO3CW
                        tblNukishi(R + 1).WFRESDO3CW = "0"
                    ''残存酸素
                    Case 26
                        tblNukishi(R + 1).WFSMPLIDAOICW = tblNukishi(R).WFSMPLIDAOICW
                        tblNukishi(R + 1).WFRESAOICW = "0"
                    Case 27
                        tblNukishi(R + 1).WFSMPLIDOT1CW = ""
                        tblNukishi(R + 1).WFRESOT1CW = "0"
                    ' エピ先行評価追加
                    Case 35
                        tblNukishi(R + 1).WFSMPLIDOT2CW = ""
                        tblNukishi(R + 1).WFRESOT2CW = "0"
                    'GD追加
                    Case 28
                        tblNukishi(R + 1).WFSMPLIDGDCW = tblNukishi(R).WFSMPLIDGDCW
                        tblNukishi(R + 1).WFRESGDCW = "0"
                        tblNukishi(R + 1).WFHSGDCW = tblNukishi(R).WFHSGDCW
                    Case 29
                        tblNukishi(R + 1).EPSMPLIDB1CW = tblNukishi(R).EPSMPLIDB1CW
                        tblNukishi(R + 1).EPRESB1CW = "0"
                    Case 30
                        tblNukishi(R + 1).EPSMPLIDB2CW = tblNukishi(R).EPSMPLIDB2CW
                        tblNukishi(R + 1).EPRESB2CW = "0"
                    Case 31
                        tblNukishi(R + 1).EPSMPLIDB3CW = tblNukishi(R).EPSMPLIDB3CW
                        tblNukishi(R + 1).EPRESB3CW = "0"
                    Case 32
                        tblNukishi(R + 1).EPSMPLIDL1CW = tblNukishi(R).EPSMPLIDL1CW
                        tblNukishi(R + 1).EPRESL1CW = "0"
                    Case 33
                        tblNukishi(R + 1).EPSMPLIDL2CW = tblNukishi(R).EPSMPLIDL2CW
                        tblNukishi(R + 1).EPRESL2CW = "0"
                    Case 34
                        tblNukishi(R + 1).EPSMPLIDL3CW = tblNukishi(R).EPSMPLIDL3CW
                        tblNukishi(R + 1).EPRESL3CW = "0"
                End Select
            End If
        Else
            '3のとき(上品番Z)
            If Trim(tblNukishi(R - 1).HINBCW) = "Z" Then
                Select Case C
                    Case 11
                        tblNukishi(R - 1).WFSMPLIDRSCW = tblNukishi(R).WFSMPLIDRSCW
                        tblNukishi(R - 1).WFRESRS1CW = "0"
                    Case 12
                        tblNukishi(R - 1).WFSMPLIDOICW = tblNukishi(R).WFSMPLIDOICW
                        tblNukishi(R - 1).WFRESOICW = "0"
                    Case 13
                        tblNukishi(R - 1).WFSMPLIDB1CW = tblNukishi(R).WFSMPLIDB1CW
                        tblNukishi(R - 1).WFRESB1CW = "0"
                    Case 14
                        tblNukishi(R - 1).WFSMPLIDB2CW = tblNukishi(R).WFSMPLIDB2CW
                        tblNukishi(R - 1).WFRESB2CW = "0"
                    Case 15
                        tblNukishi(R - 1).WFSMPLIDB3CW = tblNukishi(R).WFSMPLIDB3CW
                        tblNukishi(R - 1).WFRESB3CW = "0"
                    Case 16
                        tblNukishi(R - 1).WFSMPLIDL1CW = tblNukishi(R).WFSMPLIDL1CW
                        tblNukishi(R - 1).WFRESL1CW = "0"
                    Case 17
                        tblNukishi(R - 1).WFSMPLIDL2CW = tblNukishi(R).WFSMPLIDL2CW
                        tblNukishi(R - 1).WFRESL2CW = "0"
                    Case 18
                        tblNukishi(R - 1).WFSMPLIDL3CW = tblNukishi(R).WFSMPLIDL3CW
                        tblNukishi(R - 1).WFRESL3CW = "0"
                    Case 19
                        tblNukishi(R - 1).WFSMPLIDL4CW = tblNukishi(R).WFSMPLIDL4CW
                        tblNukishi(R - 1).WFRESL4CW = "0"
                    Case 20
                        tblNukishi(R - 1).WFSMPLIDDSCW = tblNukishi(R).WFSMPLIDDSCW
                        tblNukishi(R - 1).WFRESDSCW = "0"
                    Case 21
                        tblNukishi(R - 1).WFSMPLIDDZCW = tblNukishi(R).WFSMPLIDDZCW
                        tblNukishi(R - 1).WFRESDZCW = "0"
                    Case 22
                        tblNukishi(R - 1).WFSMPLIDSPCW = tblNukishi(R).WFSMPLIDSPCW
                        tblNukishi(R - 1).WFRESSPCW = "0"
                    Case 23
                        tblNukishi(R - 1).WFSMPLIDDO1CW = tblNukishi(R).WFSMPLIDDO1CW
                        tblNukishi(R - 1).WFRESDO1CW = "0"
                    Case 24
                        tblNukishi(R - 1).WFSMPLIDDO2CW = tblNukishi(R).WFSMPLIDDO2CW
                        tblNukishi(R - 1).WFRESDO2CW = "0"
                    Case 25
                        tblNukishi(R - 1).WFSMPLIDDO3CW = tblNukishi(R).WFSMPLIDDO3CW
                        tblNukishi(R - 1).WFRESDO3CW = "0"
                    ''残存酸素追加
                    Case 26
                        tblNukishi(R - 1).WFSMPLIDAOICW = tblNukishi(R).WFSMPLIDAOICW
                        tblNukishi(R - 1).WFRESAOICW = "0"
                    Case 27
                        tblNukishi(R - 1).WFSMPLIDOT1CW = ""
                        tblNukishi(R - 1).WFRESOT1CW = "0"
                    ' エピ先行評価追加対応
                    Case 35
                        tblNukishi(R - 1).WFSMPLIDOT2CW = ""
                        tblNukishi(R - 1).WFRESOT2CW = "0"
                    'GD追加
                    Case 28
                        tblNukishi(R - 1).WFSMPLIDGDCW = tblNukishi(R).WFSMPLIDGDCW
                        tblNukishi(R - 1).WFRESGDCW = "0"
                        tblNukishi(R - 1).WFHSGDCW = tblNukishi(R).WFHSGDCW
                    Case 29
                        tblNukishi(R - 1).EPSMPLIDB1CW = tblNukishi(R).EPSMPLIDB1CW
                        tblNukishi(R - 1).EPRESB1CW = "0"
                    Case 30
                        tblNukishi(R - 1).EPSMPLIDB2CW = tblNukishi(R).EPSMPLIDB2CW
                        tblNukishi(R - 1).EPRESB2CW = "0"
                    Case 31
                        tblNukishi(R - 1).EPSMPLIDB3CW = tblNukishi(R).EPSMPLIDB3CW
                        tblNukishi(R - 1).EPRESB3CW = "0"
                    Case 32
                        tblNukishi(R - 1).EPSMPLIDL1CW = tblNukishi(R).EPSMPLIDL1CW
                        tblNukishi(R - 1).EPRESL1CW = "0"
                    Case 33
                        tblNukishi(R - 1).EPSMPLIDL2CW = tblNukishi(R).EPSMPLIDL2CW
                        tblNukishi(R - 1).EPRESL2CW = "0"
                    Case 34
                        tblNukishi(R - 1).EPSMPLIDL3CW = tblNukishi(R).EPSMPLIDL3CW
                        tblNukishi(R - 1).EPRESL3CW = "0"
                End Select
            End If
        End If 'Mod
    ElseIf sHkbn = "1" Then  '→黒→黄色
        If R Mod 2 = 0 Then
        '4のとき(下品番Z)
            If Trim(tblNukishi(R + 1).HINBCW) = "Z" Then
                Select Case C
                    Case 11
                        tblNukishi(R + 1).WFSMPLIDRSCW = tblKns(R + 1).WFSMPLIDRSCW
                        tblNukishi(R + 1).WFRESRS1CW = tblKns(R + 1).WFRESRS1CW
                    Case 12
                        tblNukishi(R + 1).WFSMPLIDOICW = tblKns(R + 1).WFSMPLIDOICW
                        tblNukishi(R + 1).WFRESOICW = tblKns(R + 1).WFRESOICW
                    Case 13
                        tblNukishi(R + 1).WFSMPLIDB1CW = tblKns(R + 1).WFSMPLIDB1CW
                        tblNukishi(R + 1).WFRESB1CW = tblKns(R + 1).WFRESB1CW
                    Case 14
                        tblNukishi(R + 1).WFSMPLIDB2CW = tblKns(R + 1).WFSMPLIDB2CW
                        tblNukishi(R + 1).WFRESB2CW = tblKns(R + 1).WFRESB2CW
                    Case 15
                        tblNukishi(R + 1).WFSMPLIDB3CW = tblKns(R + 1).WFSMPLIDB3CW
                        tblNukishi(R + 1).WFRESB3CW = tblKns(R + 1).WFRESB3CW
                    Case 16
                        tblNukishi(R + 1).WFSMPLIDL1CW = tblKns(R + 1).WFSMPLIDL1CW
                        tblNukishi(R + 1).WFRESL1CW = tblKns(R + 1).WFRESL1CW
                    Case 17
                        tblNukishi(R + 1).WFSMPLIDL2CW = tblKns(R + 1).WFSMPLIDL2CW
                        tblNukishi(R + 1).WFRESL2CW = tblKns(R + 1).WFRESL2CW
                    Case 18
                        tblNukishi(R + 1).WFSMPLIDL3CW = tblKns(R + 1).WFSMPLIDL3CW
                        tblNukishi(R + 1).WFRESL3CW = tblKns(R + 1).WFRESL3CW
                    Case 19
                        tblNukishi(R + 1).WFSMPLIDL4CW = tblKns(R + 1).WFSMPLIDL4CW
                        tblNukishi(R + 1).WFRESL4CW = tblKns(R + 1).WFRESL4CW
                    Case 20
                        tblNukishi(R + 1).WFSMPLIDDSCW = tblKns(R + 1).WFSMPLIDDSCW
                        tblNukishi(R + 1).WFRESDSCW = tblKns(R + 1).WFRESDSCW
                    Case 21
                        tblNukishi(R + 1).WFSMPLIDDZCW = tblKns(R + 1).WFSMPLIDDZCW
                        tblNukishi(R + 1).WFRESDZCW = tblKns(R + 1).WFRESDZCW
                    Case 22
                        tblNukishi(R + 1).WFSMPLIDSPCW = tblKns(R + 1).WFSMPLIDSPCW
                        tblNukishi(R + 1).WFRESSPCW = tblKns(R + 1).WFRESSPCW
                    Case 23
                        tblNukishi(R + 1).WFSMPLIDDO1CW = tblKns(R + 1).WFSMPLIDDO1CW
                        tblNukishi(R + 1).WFRESDO1CW = tblKns(R + 1).WFRESDO1CW
                    Case 24
                        tblNukishi(R + 1).WFSMPLIDDO2CW = tblKns(R + 1).WFSMPLIDDO2CW
                        tblNukishi(R + 1).WFRESDO2CW = tblKns(R + 1).WFRESDO2CW
                    Case 25
                        tblNukishi(R + 1).WFSMPLIDDO3CW = tblKns(R + 1).WFSMPLIDDO3CW
                        tblNukishi(R + 1).WFRESDO3CW = tblKns(R + 1).WFRESDO3CW
                    ''残存酸素追加
                    Case 26
                        tblNukishi(R + 1).WFSMPLIDAOICW = tblKns(R + 1).WFSMPLIDAOICW
                        tblNukishi(R + 1).WFRESAOICW = tblKns(R + 1).WFRESAOICW
                    Case 27
                        tblNukishi(R + 1).WFSMPLIDOT1CW = tblKns(R + 1).WFSMPLIDOT1CW
                        tblNukishi(R + 1).WFRESOT1CW = tblKns(R + 1).WFRESOT1CW
                    ' エピ先行評価追加
                    Case 35
                        tblNukishi(R + 1).WFSMPLIDOT2CW = tblKns(R + 1).WFSMPLIDOT2CW
                        tblNukishi(R + 1).WFRESOT2CW = tblKns(R + 1).WFRESOT2CW
                    'GD追加
                    Case 28
                        tblNukishi(R + 1).WFSMPLIDGDCW = tblKns(R + 1).WFSMPLIDGDCW
                        tblNukishi(R + 1).WFRESGDCW = tblKns(R + 1).WFRESGDCW
                        tblNukishi(R + 1).WFHSGDCW = tblKns(R + 1).WFHSGDCW
                    Case 29
                        tblNukishi(R + 1).EPSMPLIDB1CW = tblKns(R + 1).EPSMPLIDB1CW
                        tblNukishi(R + 1).EPRESB1CW = tblKns(R + 1).EPRESB1CW
                    Case 30
                        tblNukishi(R + 1).EPSMPLIDB2CW = tblKns(R + 1).EPSMPLIDB2CW
                        tblNukishi(R + 1).EPRESB2CW = tblKns(R + 1).EPRESB2CW
                    Case 31
                        tblNukishi(R + 1).EPSMPLIDB3CW = tblKns(R + 1).EPSMPLIDB3CW
                        tblNukishi(R + 1).EPRESB3CW = tblKns(R + 1).EPRESB3CW
                    Case 32
                        tblNukishi(R + 1).EPSMPLIDL1CW = tblKns(R + 1).EPSMPLIDL1CW
                        tblNukishi(R + 1).EPRESL1CW = tblKns(R + 1).EPRESL1CW
                    Case 33
                        tblNukishi(R + 1).EPSMPLIDL2CW = tblKns(R + 1).EPSMPLIDL2CW
                        tblNukishi(R + 1).EPRESL2CW = tblKns(R + 1).EPRESL2CW
                    Case 34
                        tblNukishi(R + 1).EPSMPLIDL3CW = tblKns(R + 1).EPSMPLIDL3CW
                        tblNukishi(R + 1).EPRESL3CW = tblKns(R + 1).EPRESL3CW
                End Select
            End If
        Else
            If Trim(tblNukishi(R - 1).HINBCW) = "Z" Then
                Select Case C
                    Case 11
                        tblNukishi(R - 1).WFSMPLIDRSCW = tblKns(R - 1).WFSMPLIDRSCW
                        tblNukishi(R - 1).WFRESRS1CW = tblKns(R - 1).WFRESRS1CW
                    Case 12
                        tblNukishi(R - 1).WFSMPLIDOICW = tblKns(R - 1).WFSMPLIDOICW
                        tblNukishi(R - 1).WFRESOICW = tblKns(R - 1).WFRESOICW
                    Case 13
                        tblNukishi(R - 1).WFSMPLIDB1CW = tblKns(R - 1).WFSMPLIDB1CW
                        tblNukishi(R - 1).WFRESB1CW = tblKns(R - 1).WFRESB1CW
                    Case 14
                        tblNukishi(R - 1).WFSMPLIDB2CW = tblKns(R - 1).WFSMPLIDB2CW
                        tblNukishi(R - 1).WFRESB2CW = tblKns(R - 1).WFRESB2CW
                    Case 15
                        tblNukishi(R - 1).WFSMPLIDB3CW = tblKns(R - 1).WFSMPLIDB3CW
                        tblNukishi(R - 1).WFRESB3CW = tblKns(R - 1).WFRESB3CW
                    Case 16
                        tblNukishi(R - 1).WFSMPLIDL1CW = tblKns(R - 1).WFSMPLIDL1CW
                        tblNukishi(R - 1).WFRESL1CW = tblKns(R - 1).WFRESL1CW
                    Case 17
                        tblNukishi(R - 1).WFSMPLIDL2CW = tblKns(R - 1).WFSMPLIDL2CW
                        tblNukishi(R - 1).WFRESL2CW = tblKns(R - 1).WFRESL2CW
                    Case 18
                        tblNukishi(R - 1).WFSMPLIDL3CW = tblKns(R - 1).WFSMPLIDL3CW
                        tblNukishi(R - 1).WFRESL3CW = tblKns(R - 1).WFRESL3CW
                    Case 19
                        tblNukishi(R - 1).WFSMPLIDL4CW = tblKns(R - 1).WFSMPLIDL4CW
                        tblNukishi(R - 1).WFRESL4CW = tblKns(R - 1).WFRESL4CW
                    Case 20
                        tblNukishi(R - 1).WFSMPLIDDSCW = tblKns(R - 1).WFSMPLIDDSCW
                        tblNukishi(R - 1).WFRESDSCW = tblKns(R - 1).WFRESDSCW
                    Case 21
                        tblNukishi(R - 1).WFSMPLIDDZCW = tblKns(R - 1).WFSMPLIDDZCW
                        tblNukishi(R - 1).WFRESDZCW = tblKns(R - 1).WFRESDZCW
                    Case 22
                        tblNukishi(R - 1).WFSMPLIDSPCW = tblKns(R - 1).WFSMPLIDSPCW
                        tblNukishi(R - 1).WFRESSPCW = tblKns(R - 1).WFRESSPCW
                    Case 23
                        tblNukishi(R - 1).WFSMPLIDDO1CW = tblKns(R - 1).WFSMPLIDDO1CW
                        tblNukishi(R - 1).WFRESDO1CW = tblKns(R - 1).WFRESDO1CW
                    Case 24
                        tblNukishi(R - 1).WFSMPLIDDO2CW = tblKns(R - 1).WFSMPLIDDO2CW
                        tblNukishi(R - 1).WFRESDO2CW = tblKns(R - 1).WFRESDO2CW
                    Case 25
                        tblNukishi(R - 1).WFSMPLIDDO3CW = tblKns(R - 1).WFSMPLIDDO3CW
                        tblNukishi(R - 1).WFRESDO3CW = tblKns(R - 1).WFRESDO3CW
                    ''残存酸素
                    Case 26
                        tblNukishi(R - 1).WFSMPLIDAOICW = tblKns(R - 1).WFSMPLIDAOICW
                        tblNukishi(R - 1).WFRESAOICW = tblKns(R - 1).WFRESAOICW
                    Case 27
                        tblNukishi(R - 1).WFSMPLIDOT1CW = tblKns(R - 1).WFSMPLIDOT1CW
                        tblNukishi(R - 1).WFRESOT1CW = tblKns(R - 1).WFRESOT1CW
                    ' エピ先行評価追加対応
                    Case 35
                        tblNukishi(R - 1).WFSMPLIDOT2CW = tblKns(R - 1).WFSMPLIDOT2CW
                        tblNukishi(R - 1).WFRESOT2CW = tblKns(R - 1).WFRESOT2CW
                    'GD追加
                    Case 28
                        tblNukishi(R - 1).WFSMPLIDGDCW = tblKns(R - 1).WFSMPLIDGDCW
                        tblNukishi(R - 1).WFRESGDCW = tblKns(R - 1).WFRESGDCW
                        tblNukishi(R - 1).WFHSGDCW = tblKns(R - 1).WFHSGDCW
                    Case 29
                        tblNukishi(R - 1).EPSMPLIDB1CW = tblKns(R - 1).EPSMPLIDB1CW
                        tblNukishi(R - 1).EPRESB1CW = tblKns(R - 1).EPRESB1CW
                    Case 30
                        tblNukishi(R - 1).EPSMPLIDB2CW = tblKns(R - 1).EPSMPLIDB2CW
                        tblNukishi(R - 1).EPRESB2CW = tblKns(R - 1).EPRESB2CW
                    Case 31
                        tblNukishi(R - 1).EPSMPLIDB3CW = tblKns(R - 1).EPSMPLIDB3CW
                        tblNukishi(R - 1).EPRESB3CW = tblKns(R - 1).EPRESB3CW
                    Case 32
                        tblNukishi(R - 1).EPSMPLIDL1CW = tblKns(R - 1).EPSMPLIDL1CW
                        tblNukishi(R - 1).EPRESL1CW = tblKns(R - 1).EPRESL1CW
                    Case 33
                        tblNukishi(R - 1).EPSMPLIDL2CW = tblKns(R - 1).EPSMPLIDL2CW
                        tblNukishi(R - 1).EPRESL2CW = tblKns(R - 1).EPRESL2CW
                    Case 34
                        tblNukishi(R - 1).EPSMPLIDL3CW = tblKns(R - 1).EPSMPLIDL3CW
                        tblNukishi(R - 1).EPRESL3CW = tblKns(R - 1).EPRESL3CW
                End Select
            End If
        End If
    End If
End Sub

'Del Start 2011/03/11 SMPK Miyata
''*******************************************************************************************
''*    関数名        : Combo1_Change
''*
''*    処理概要      : 1.使用していない気がする？
''*
''*
''*    パラメータ    : 変数名      ,IO ,型      ,説明
''*　　      　　      なし
''*
''*    戻り値        : なし
''*
''*******************************************************************************************
'Private Sub Combo1_Change()
'    Dim blRtn       As Boolean
'    Dim sErrMsg     As String
'    Dim sBlkId      As String
'    Dim sSXLID      As String
'
'    Dim intLoopCnt  As Integer
'    Dim intSprSta   As Integer
'    Dim sSprSta     As String
'    Dim vSprSta     As Variant
'
'    With sprExamine
'        .ReDraw = False
'        For intLoopCnt = 1 To .MaxRows
'            Select Case cmbSprChg.ListIndex
'                Case intConSprChg_0  '全件指定
'                    .row = intLoopCnt
'                    .RowHidden = False
'                Case intConSprChg_1  '良品指定
'                    'エピ先行評価追加対応
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_1 Then  '良品以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        .row = intLoopCnt
'                        .RowHidden = False
'                    End If
'                Case intConSprChg_2  'サンプル指定
'                    'エピ先行評価追加対応
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_2 Then  'サンプル以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        .row = intLoopCnt
'                        .RowHidden = False
'                    End If
'                Case intConSprChg_3  '不良指定
'                    'エピ先行評価追加対応
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_3 Then  '不良以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        .row = intLoopCnt
'                        .RowHidden = False
'                    End If
'            End Select
'        Next
'        .ReDraw = True
'    End With
'End Sub
'Del End   2011/03/11 SMPK Miyata

'*******************************************************************************************
'*    関数名        : sub_F_InsertRow
'*
'*    処理概要      : 1.１行挿入を行う
'*
'*    パラメータ    : 変数名      ,IO ,型      ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_F_InsertRow()
    Dim intResult       As Integer
    Dim lngNewRow       As Long
    Dim lngNewRow2      As Long
    Dim vGetHinban      As Variant
    Dim vFirstFlg       As Variant
    Dim vActiveBLOCK    As Variant
    Dim vOldHinban      As Variant
    Dim i               As Integer

    '最大行数追加
    sprExamine.MaxRows = sprExamine.MaxRows + 2

    intResult = sprExamine.ActiveRow Mod 2
    If intResult = 0 Then     '偶数行なら上2行挿入
        lngNewRow = sprExamine.ActiveRow
        lngNewRow2 = sprExamine.ActiveRow + 1
    Else                      '奇数なら下２行挿入
        lngNewRow = sprExamine.ActiveRow + 1
        lngNewRow2 = sprExamine.ActiveRow + 2
    End If

    With sprExamine
        'エピ先行評価追加対応
        .GetText 39, .ActiveRow, vActiveBLOCK   'ブロックIDはcol=30より取得
        .GetText 40, .ActiveRow, vOldHinban

        .row = lngNewRow
        .row2 = lngNewRow2
        .col = (-1)
        .BlockMode = True
        .Action = ActionInsertRow
        .Protect = True

        '色、罫線、ロック設定
        .backColor = vbWhite

        'エピ先行評価追加対応
        .SetText 40, lngNewRow, vOldHinban
        .SetText 40, lngNewRow2, vOldHinban
        .SetText 39, lngNewRow, vActiveBLOCK
        .SetText 39, lngNewRow2, vActiveBLOCK

        .col = 1                '一番左の列
        .col2 = 1
        .CellBorderType = 2
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16
        .backColor = &H80FF80

        .col = 4                '左から3番目以降
        .col2 = 10
        .backColor = &H80FF80
        .CellBorderType = 15
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 7                '左から6番目
        .col2 = 7
        .CellBorderType = 4 Or 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 4                '左から３番目からの罫線
        .col2 = 10
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 4                '左から3番目上ブロックPを編集可能に
        .col2 = 4
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .backColor = vbWhite

        .col = 2                '左から2番目新品番を編集可能に(下ROWのみ)
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 3                '左から3番目
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = True
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 7                'WF枚数
        .col2 = 7
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 7                'WF枚数
        .col2 = 7
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16
        vFirstFlg = 0

        'エピ先行評価追加対応
        .SetText 37, lngNewRow, vFirstFlg
        .SetText 37, lngNewRow2, vFirstFlg
        .col = 9                '不良区分
        .col2 = 9
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 9                '不良区分
        .col2 = 9
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 4                '左から3番目ブロックPを編集不可に(下ROWのみ)
        .row = lngNewRow2
        .Lock = True

        .col = 11                '検査項目の入力不可
        .row = lngNewRow

        'エピ先行評価追加対応
        .col2 = 35
        .row2 = lngNewRow2
        .Lock = True

        .col = 9                '追加行に区分を作る（区分がある初期表示行からのコピー）
        .row = 1
        .col2 = 9
        .row2 = 1
        .DestCol = 9
        .DestRow = lngNewRow2
        .Action = ActionCopyRange

        .Lock = False
        .col = 9
        .col2 = 9
        .row = lngNewRow2
        .row2 = lngNewRow2
        .BlockMode = True
        .backColor = &H80FF80
        .BlockMode = False

        .col = 9                '左から８番目（上行）をテキストに
        .col2 = 9
        .row = lngNewRow
        .row2 = lngNewRow
        .CellType = CellTypeEdit

        .BlockMode = False
    End With

    If intResult = 0 Then     '偶数行なら上2行挿入
    Else                      '奇数なら下２行挿入
        sprExamine.GetText 2, sprExamine.ActiveRow, vGetHinban
        sprExamine.SetText 2, sprExamine.ActiveRow, vbNullString
        sprExamine.SetText 2, sprExamine.ActiveRow + 2, vGetHinban
        sprExamine.GetText 3, sprExamine.ActiveRow, vGetHinban      '品番2の処理を追加
        sprExamine.SetText 3, sprExamine.ActiveRow, vbNullString
        sprExamine.SetText 3, sprExamine.ActiveRow + 2, vGetHinban
    End If
End Sub

'*******************************************************************************************
'*    関数名        : sub_SelWFmap
'*
'*    処理概要      : 1.WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙ（TBCMY011）からﾃﾞｰﾀを取得
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      records()   ,O  ,typ_TBCME037 ,抽出レコード
'*　　      　　      sqlWhere    ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'*　　      　　      sqlOrder    ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function sub_SelWFmap(ByVal sBlkId As String, ByVal sSXLID As String, ByRef sErrMsg As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim bBlkOn      As Boolean 'ﾌﾞﾛｯｸが存在していたら、SXLIDの条件にANDを入れる
    Dim rs          As OraDynaset    'RecordSet
    Dim intDataCnt  As Integer
    Dim sSXLID9     As String

    On Error GoTo proc_err

    sub_SelWFmap = FUNCTION_RETURN_FAILURE
    bBlkOn = False
    intDataCnt = 0

    sSql = vbNullString
    sSql = sSql & "SELECT * FROM TBCMY011"
    sSql = sSql & " WHERE"

    If giFKeyFlg <> 2 Then
        If sSXLID <> vbNullString Then
            sSql = sSql & " MSXLID = '" & sSXLID & "'"
        End If
    Else    '実行ボタン押下後は複数シングルを表示する（初期取得の範囲指定でSELECT）
        sSXLID9 = Mid(sSXLID, 1, 9)
        sSql = sSql & " SUBSTR(MSXLID,1,9) = '" & sSXLID9 & "'"
        sSql = sSql & " AND RITOP_POS > " & SIngotP
        sSql = sSql & " AND RITOP_POS <= " & EIngotP
    End If
    sSql = sSql & " ORDER BY RITOP_POS"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        sub_SelWFmap = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("SET46")
        rs.Close
        Exit Function
    End If

    Do While Not rs.EOF
        ReDim Preserve gtWFmap(intDataCnt) As typeWFmap
        With gtWFmap(intDataCnt)
            .LOTID = CStr(rs.Fields("LOTID"))
            If IsNull(rs.Fields("BLOCKSEQ")) = True Then
                .BLOCKSEQ = 0
            Else
                .BLOCKSEQ = CInt(rs.Fields("BLOCKSEQ"))
            End If
            If IsNull(rs.Fields("INDTM")) = True Then
                .INDTM = vbNullString
            Else
                .INDTM = CStr(rs.Fields("INDTM"))
            End If
            If IsNull(rs.Fields("BASKETID")) = True Then
                .BASKETID = vbNullString
            Else
                .BASKETID = CStr(rs.Fields("BASKETID"))
            End If
            If IsNull(rs.Fields("SLOTNO")) = True Then
                .SLOTNO = 0
            Else
                .SLOTNO = CInt(rs.Fields("SLOTNO"))
            End If
            If IsNull(rs.Fields("CURRWPCS")) = True Then
                .CURRWPCS = 0
            Else
                .CURRWPCS = CInt(rs.Fields("CURRWPCS"))
            End If
            If IsNull(rs.Fields("EXISTFLG")) = True Then
                .EXISTFLG = vbNullString
            Else
                .EXISTFLG = CStr(rs.Fields("EXISTFLG"))
            End If
            If IsNull(rs.Fields("TOP_POS")) = True Then
                .TOP_POS = 0
            Else
                .TOP_POS = rs.Fields("TOP_POS")
            End If
            If IsNull(rs.Fields("REJCAT")) = True Then
                .REJCAT = vbNullString
            Else
                .REJCAT = CStr(rs.Fields("REJCAT"))
            End If
            If IsNull(rs.Fields("TXID")) = True Then
                .TXID = vbNullString
            Else
                .TXID = CStr(rs.Fields("TXID"))
            End If
            If IsNull(rs.Fields("REGDATE")) = True Then
                .REGDATE = vbNullString
            Else
                .REGDATE = CStr(rs.Fields("REGDATE"))
            End If
            If IsNull(rs.Fields("SUMMITSENDFLAG")) = True Then
                .SUMMITSENDFLG = vbNullString
            Else
                .SUMMITSENDFLG = CStr(rs.Fields("SUMMITSENDFLAG"))
            End If
            If IsNull(rs.Fields("SENDFLAG")) = True Then
                .SENDFLG = vbNullString
            Else
                .SENDFLG = CStr(rs.Fields("SENDFLAG"))
            End If
            If IsNull(rs.Fields("SENDDATE")) = True Then
                .SENDDATE = vbNullString
            Else
                .SENDDATE = CStr(rs.Fields("SENDDATE"))
            End If
            If IsNull(rs.Fields("WFSTA")) = True Then
                .WFSTA = vbNullString
            Else
                .WFSTA = CStr(rs.Fields("WFSTA"))
            End If
            If IsNull(rs.Fields("HREJCODE")) = True Then
                .HREJCODE = vbNullString
            Else
                .HREJCODE = CStr(rs.Fields("HREJCODE"))
            End If
            If IsNull(rs.Fields("UPDPROC")) = True Then
                .UPDPROC = vbNullString
            Else
                .UPDPROC = CStr(rs.Fields("UPDPROC"))
            End If
            If IsNull(rs.Fields("UPDDATE")) = True Then
                .UPDDATE = vbNullString
            Else
                .UPDDATE = CStr(rs.Fields("UPDDATE"))
            End If
            If IsNull(rs.Fields("MSXLID")) = True Then
                .SXLID = vbNullString
            Else
                .SXLID = CStr(rs.Fields("MSXLID"))
            End If
            If IsNull(rs.Fields("MHINBAN")) = True Then
                .hinban = vbNullString
            Else
                .hinban = CStr(rs.Fields("MHINBAN"))
            End If
            If IsNull(rs.Fields("MREVNUM")) = True Then
                .REVNUM = 0
            Else
                .REVNUM = CInt(rs.Fields("MREVNUM"))
            End If
            If IsNull(rs.Fields("MFACTORY")) = True Then
                .factory = vbNullString
            Else
                .factory = CStr(rs.Fields("MFACTORY"))
            End If
            If IsNull(rs.Fields("MOPECOND")) = True Then
                .opecond = vbNullString
            Else
                .opecond = CStr(rs.Fields("MOPECOND"))
            End If
            If IsNull(rs.Fields("KANKBN")) = True Then
                .KANKBN = vbNullString
            Else
                .KANKBN = CStr(rs.Fields("KANKBN"))
            End If
            If IsNull(rs.Fields("MSMPLEID")) = True Then
                .SMPLEID = vbNullString
            Else
                .SMPLEID = CStr(rs.Fields("MSMPLEID"))
            End If
            If IsNull(rs.Fields("NREJCODE")) = True Then
                .NREJCODE = vbNullString
            Else
                .NREJCODE = CStr(rs.Fields("NREJCODE"))
            End If
            If IsNull(rs.Fields("SHAFLAG")) = True Then
                .SMPLEFLG = vbNullString
            Else
                .SMPLEFLG = CStr(rs.Fields("SHAFLAG"))
            End If
            If IsNull(rs.Fields("RTOP_POS")) = True Then
                .RTOP_POS = 0
            Else
                .RTOP_POS = rs.Fields("RTOP_POS")
            End If
            If IsNull(rs.Fields("RITOP_POS")) = True Then
                .RITOP_POS = 0
            Else
                .RITOP_POS = rs.Fields("RITOP_POS")
            End If
        End With
        intDataCnt = intDataCnt + 1
        rs.MoveNext
    Loop
    If intDataCnt = 0 Then
        ReDim records(0)
        sub_SelWFmap = FUNCTION_RETURN_FAILURE
        sErrMsg = GetMsgStr("SET46")
        rs.Close
        Exit Function
    End If

    rs.Close
    sub_SelWFmap = FUNCTION_RETURN_SUCCESS
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sub_SelWFmap = FUNCTION_RETURN_FAILURE
    sErrMsg = GetMsgStr("SET47")
    rs.Close
End Function

'Del Start 2011/03/11 SMPK Miyata
''*******************************************************************************************
''*    関数名        : fnc_SetWFmapData
''*
''*    処理概要      : 1.データ表示処理
''*
''*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
''*　　      　　      なし
''*
''*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
''*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
''*
''*******************************************************************************************
'Private Function fnc_SetWFmapData() As FUNCTION_RETURN
'    Dim intLoopCnt    As Integer
'    Dim dblTopPos     As Double
'    Dim dblRTopPos    As Double
'    Dim dblRITopPos   As Double
'    Dim intWarpPoint  As Integer
'
'    With sprWfmapView
'        .MaxRows = 0
'        intWarpPoint = 1
'        For intLoopCnt = 0 To UBound(gtWFmap)
'            .MaxRows = .MaxRows + 1
'            .SetText 3, intLoopCnt + 1, gtWFmap(intLoopCnt).LOTID
'            .SetText 5, intLoopCnt + 1, gtWFmap(intLoopCnt).BLOCKSEQ
'            .SetText 17, intLoopCnt + 1, gtWFmap(intLoopCnt).INDTM
'            .SetText 18, intLoopCnt + 1, gtWFmap(intLoopCnt).BASKETID
'            .SetText 19, intLoopCnt + 1, gtWFmap(intLoopCnt).SLOTNO
'            .SetText 9, intLoopCnt + 1, gtWFmap(intLoopCnt).CURRWPCS
'            .SetText 10, intLoopCnt + 1, gtWFmap(intLoopCnt).EXISTFLG
'            If gtWFmap(intLoopCnt).TOP_POS / 10 = 0 Then
'                .SetText 6, intLoopCnt + 1, 0
'            Else
'                dblTopPos = gtWFmap(intLoopCnt).TOP_POS / 10
'                dblTopPos = dblTopPos
'                .SetText 6, intLoopCnt + 1, dblTopPos
'            End If
'            .SetText 11, intLoopCnt + 1, gtWFmap(intLoopCnt).REJCAT
'            .SetText 26, intLoopCnt + 1, gtWFmap(intLoopCnt).TXID
'            .SetText 25, intLoopCnt + 1, Format(CVar(gtWFmap(intLoopCnt).REGDATE), "yyyy/mm/dd")
'            .SetText 27, intLoopCnt + 1, gtWFmap(intLoopCnt).SUMMITSENDFLG
'            .SetText 28, intLoopCnt + 1, gtWFmap(intLoopCnt).SENDFLG
'            .SetText 29, intLoopCnt + 1, Format(CVar(gtWFmap(intLoopCnt).SENDDATE), "yyyy/mm/dd")
'
'            'WF状態判定
'            Select Case gtWFmap(intLoopCnt).WFSTA
'                Case gsWF_STA_0   '通常
'                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
'                        Case gsWF_SMPL_1
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case gsWF_SMPL_2
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_OK & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case gsWF_SMPL_3
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_NG & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case Else
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_1
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = &H80FF80
'                            .BlockMode = False
'                    End Select
'                Case gsWF_STA_1   '共有
'                    .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                    .SetText 30, intLoopCnt + 1, intConSprChg_1
'                    .col = 1
'                    .col2 = 32          'Warp判定対応
'                    .row = intLoopCnt + 1
'                    .row2 = intLoopCnt + 1
'                    .BlockMode = True
'                    .backColor = &H80FF80
'                    .BlockMode = False
'                Case gsWF_STA_4   '欠落
'                    'サンプルフラグ判定
'                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
'                        Case gsWF_SMPL_4    'upd 2003/05/19 hitec)matsumoto サンプルの結果以外はすべて欠落と判断する
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_KEKKA & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case Else
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_STA_K & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF状態フラグの表示を追加
'                            .SetText 30, intLoopCnt + 1, intConSprChg_3
'                            .col = 1
'                            .col2 = 32          'Warp判定対応
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbRed
'                            .BlockMode = False
'                    End Select
'            End Select
'
'            .SetText 14, intLoopCnt + 1, gtWFmap(intLoopCnt).SMPLEID
'            .SetText 12, intLoopCnt + 1, gtWFmap(intLoopCnt).HREJCODE
'            .SetText 23, intLoopCnt + 1, gtWFmap(intLoopCnt).UPDPROC
'            .SetText 24, intLoopCnt + 1, gtWFmap(intLoopCnt).UPDDATE
'            .SetText 2, intLoopCnt + 1, gtWFmap(intLoopCnt).SXLID
'            .SetText 4, intLoopCnt + 1, gtWFmap(intLoopCnt).hinban
'            .SetText 20, intLoopCnt + 1, gtWFmap(intLoopCnt).REVNUM
'            .SetText 21, intLoopCnt + 1, gtWFmap(intLoopCnt).factory
'            .SetText 22, intLoopCnt + 1, gtWFmap(intLoopCnt).opecond
'            .SetText 13, intLoopCnt + 1, gtWFmap(intLoopCnt).KANKBN
'            .SetText 15, intLoopCnt + 1, gtWFmap(intLoopCnt).NREJCODE
'            .SetText 16, intLoopCnt + 1, gtWFmap(intLoopCnt).SMPLEFLG
'
'            If gtWFmap(intLoopCnt).RTOP_POS = 0 Then
'                .SetText 7, intLoopCnt + 1, 0
'            Else
'                dblRTopPos = gtWFmap(intLoopCnt).RTOP_POS
'                dblRTopPos = Format(dblRTopPos, "0000.0")
'                .SetText 7, intLoopCnt + 1, dblRTopPos
'            End If
'
'            If gtWFmap(intLoopCnt).RITOP_POS = 0 Then
'                .SetText 8, intLoopCnt + 1, 0
'            Else
'                dblRITopPos = gtWFmap(intLoopCnt).RITOP_POS
'                dblRITopPos = Format(dblRITopPos, "0000.0")
'                .SetText 8, intLoopCnt + 1, dblRITopPos
'            End If
'
'            'Warp情報表示追加
'            If UBound(tWarpMeasG) >= intWarpPoint Then
'                'ﾌﾞﾛｯｸIDとﾌﾞﾛｯｸ内連番で紐付け
'                If tWarpMeasG(intWarpPoint).BLOCKID = gtWFmap(intLoopCnt).LOTID And _
'                   tWarpMeasG(intWarpPoint).WAFID = gtWFmap(intLoopCnt).BLOCKSEQ Then
'                    '実ﾃﾞｰﾀが無い場合は表示しない
'                    If tWarpMeasG(intWarpPoint).EXISTFLG > 0 Then
'                        'Warp値
'                        .SetText 31, intLoopCnt + 1, CStr(DBData2DispData_nl(tWarpMeasG(intWarpPoint).MEASDATA))
'                        '判定
'                        .SetText 32, intLoopCnt + 1, IIf(tWarpMeasG(intWarpPoint).Judg, "OK", "NG")
'                    End If
'                    intWarpPoint = intWarpPoint + 1
'                End If
'            End If
'        Next
'    End With
'End Function
'Del End   2011/03/11 SMPK Miyata

'Del Start 2011/03/11 SMPK Miyata
''*******************************************************************************************
''*    関数名        : cmbSprChg_Click
''*
''*    処理概要      : 1.抽出条件により、WFﾏｯﾌﾟ一覧を種別毎の表示の切り替えを行う
''*
''*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
''*　　      　　      なし
''*
''*    戻り値        : なし
''*
''*******************************************************************************************
'Private Sub cmbSprChg_Click()
'    Dim intLoopCnt  As Integer
'    Dim intSprSta   As Integer
'    Dim sSprSta     As String
'    Dim vSprSta     As Variant
'    Dim intRowNo    As Integer
'
'    intRowNo = 0  '表示されている行だけ番号をふりなおす
'     With sprWfmapView
'        .ReDraw = False
'        For intLoopCnt = 1 To .MaxRows
'            Select Case cmbSprChg.ListIndex
'                Case intConSprChg_0  '全件指定
'                    .row = intLoopCnt
'                    .RowHidden = False
'                    intRowNo = intRowNo + 1
'                    .row = intLoopCnt
'                    .RowHidden = False
'                    .col = 0
'                    .row = intLoopCnt
'                    .text = intRowNo
'                Case intConSprChg_1  '良品指定
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_1 Then  '良品以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        intRowNo = intRowNo + 1
'                        .row = intLoopCnt
'                        .RowHidden = False
'                        .col = 0
'                        .row = intLoopCnt
'                        .text = intRowNo
'                    End If
'                Case intConSprChg_2  'サンプル指定
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_2 Then  'サンプル以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        intRowNo = intRowNo + 1
'                        .row = intLoopCnt
'                        .RowHidden = False
'                        .col = 0
'                        .row = intLoopCnt
'                        .text = intRowNo
'                    End If
'                Case intConSprChg_3  '不良指定
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_3 Then  '不良以外だったら、非表示
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        intRowNo = intRowNo + 1
'                        .row = intLoopCnt
'                        .RowHidden = False
'                        .col = 0
'                        .row = intLoopCnt
'                        .text = intRowNo
'                    End If
'            End Select
'        Next
'
'        If .MaxRows > 0 Then
'            .col = 1
'            .row = 1
'            .Action = ActionActiveCell
'        End If
'        .ReDraw = True
'    End With
'End Sub
'Del End   2011/03/11 SMPK Miyata

'*******************************************************************************************
'*    関数名        : fnc_CheckDataWfmap
'*
'*    処理概要      : 1.抜試指示一覧の入力内容をチェックする
'*                      (抜試指示一覧のデータチェック(WFﾏｯﾌﾟ対応）)
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Private Function fnc_CheckDataWfmap() As FUNCTION_RETURN
    Dim intLoopCnt      As Integer
    Dim vFlg            As Variant
    Dim vBlockP         As Variant
    Dim intNukisiFlg    As Integer
    Dim intBlkChkLoop   As Integer
    Dim vChkBlk         As Variant
    Dim intLoopCntAdd   As Integer
    Dim intHinChkLoop   As Integer
    Dim vHinChk         As Variant
    Dim vDummyBlk       As Variant
    Dim vDummyHin       As Variant
    Dim intDummyBlkSet  As Integer
    Dim sNowBlk         As String
    Dim vBlckID         As Variant
    Dim sBlckID         As Variant
    Dim vUBlckP         As Variant
    Dim vMaxBlckP       As Variant
    Dim vUpPos          As Variant
    Dim vSmpId1         As Variant
    Dim vSmpId2         As Variant
    Dim intBlkPChk      As Integer
    Dim vStrB           As Variant
    Dim vStrSam         As Variant

    '品番の未入力エラー
    With sprExamine
        For intHinChkLoop = 1 To .MaxRows Step 2
            .GetText 2, intHinChkLoop, vHinChk
            If vHinChk = vbNullString Then
                fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                lblMsg.Caption = GetMsgStr("EHINA")
                Exit Function
            End If
        Next
    End With

    'ブロックの入力チェック
    intLoopCntAdd = 2
    With sprExamine
        For intLoopCnt = 1 To .MaxRows
            For intBlkPChk = 1 To .MaxRows
                If intBlkPChk Mod 2 = 0 Then
                    .GetText 4, intBlkPChk, vBlockP
                    If vBlockP = vbNullString Then
                        fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                        lblMsg.Caption = GetMsgStr("EBLK8")
                        .col = 4
                        .col2 = 4
                        .row = intBlkPChk
                        .row2 = intBlkPChk
                        .BlockMode = True
                        .backColor = &H8080FF
                        .BlockMode = False
                        Exit Function
                    End If
                    If intLoopCnt Mod 2 = 0 Then
                        If intLoopCnt >= 2 Then
                            .GetText 4, intLoopCnt, vBlockP
                            .GetText 4, intLoopCnt - 1, vUBlckP
                            If vBlockP <> vbNullString And vUBlckP <> vbNullString Then
                                If CInt(vUBlckP) >= CInt(vBlockP) Then
                                    fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                                    lblMsg.Caption = GetMsgStr("EBLKA")
                                    .col = 4
                                    .col2 = 4
                                    .row = intLoopCnt
                                    .row2 = intLoopCnt
                                    .BlockMode = True
                                    .backColor = COLOR_NG
                                    .BlockMode = False
                                    Exit Function
                                End If
                            End If
                        End If
                    End If
                End If
            Next

            If intLoopCnt <> .MaxRows Then
                'エピ先行評価追加対応
                .GetText 37, intLoopCnt, vFlg
                If (vFlg = "1") Or (vFlg = "3") Then    '初期表示の抜試行

                ElseIf intLoopCnt Mod 2 = 0 Then        '挿入された行で、偶数行が入力の行になる
                    .GetText 4, intLoopCnt, vBlockP

                    '同一ブロックでの数値範囲チェック
                    '現在ブロックID
                    'エピ先行評価追加対応
                    .GetText 39, intLoopCnt, vBlckID
                    sBlckID = CStr(vBlckID)

                    '上の値より大きいこと
                    If intLoopCnt > 2 Then
                        'エピ先行評価追加対応
                        .GetText 37, intLoopCnt - 1, vUpPos
                        If (vUpPos = "1") Or (vUpPos = "3") Or (vUpPos = "9") Then
                            'エピ先行評価追加対応
                            .GetText 39, intLoopCnt - 1, vBlckID
                        Else
                            'エピ先行評価追加対応
                            .GetText 39, intLoopCnt - 2, vBlckID
                        End If
                        If sBlckID = CStr(vBlckID) Then
                            'エピ先行評価追加対応
                            .GetText 37, intLoopCnt - 1, vUpPos
                            If (vUpPos = "1") Or (vUpPos = "3") Or (vUpPos = "9") Then
                                .GetText 4, intLoopCnt - 1, vUBlckP
                            Else
                                .GetText 4, intLoopCnt - 2, vUBlckP
                            End If

                            If CInt(vBlockP) <= CInt(vUBlckP) Then
                                fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                                .col = 4
                                .col2 = 4
                                .row = intLoopCnt
                                .row2 = intLoopCnt
                                .BlockMode = True
                                .backColor = COLOR_NG
                                .BlockMode = False
                                lblMsg.Caption = GetMsgStr("EBLKB")
                                Exit Function
                            End If
                        End If
                    End If

                    '下の値よりも小さいこと
                    If intLoopCnt < .MaxRows Then
                        'エピ先行評価追加対応
                        .GetText 39, intLoopCnt + 2, vBlckID
                        If sBlckID = CStr(vBlckID) Then
                            .GetText 4, intLoopCnt + 2, vUBlckP
                            If CInt(vUBlckP) <= CInt(vBlockP) Then
                                fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                                lblMsg.Caption = GetMsgStr("EBLKC")
                                .col = 4
                                .col2 = 4
                                .row = intLoopCnt
                                .row2 = intLoopCnt
                                .BlockMode = True
                                .backColor = COLOR_NG
                                .BlockMode = False
                                Exit Function
                            End If
                        End If
                    End If

                    'ダミーのブロックID・品番に該当ﾃﾞｰﾀを入れておく
                    '現在位置のブロックIDがNULLだったら上の行に探しにいく
                    .GetText 1, intLoopCnt, vDummyBlk
                    If vDummyBlk = vbNullString Then
                        For intDummyBlkSet = intLoopCnt - 1 To 1 Step -1
                            .GetText 1, intDummyBlkSet, vDummyBlk
                            If vDummyBlk <> vbNullString Then
                                Exit For
                            End If
                        Next
                    End If

                    'エピ先行評価追加対応
                    .SetText 39, intLoopCnt, vDummyBlk
                    '品番
                    .GetText 2, intLoopCnt + 1, vDummyHin

                    'エピ先行評価追加対応
                    .SetText 40, intLoopCnt, vDummyHin

                    'エピ先行評価追加対応
                    .SetText 37, intLoopCnt, "2"
                    intLoopCnt = intLoopCnt + 1  '2行下の行を見に行くために＋１しておく
                End If
            End If
        Next

        'サンプルが上下共有の場合のエラーメッセージ---------------
        For intLoopCnt = 2 To .MaxRows Step 2
            .GetText 10, intLoopCnt, vSmpId1
            .GetText 10, intLoopCnt + 1, vSmpId2
            If (vSmpId1 = vSmpId2) And (vSmpId1 <> vbNullString) And (vSmpId2 <> vbNullString) Then    'サンプルIDが、上下共有
                fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                lblMsg.Caption = GetMsgStr("ENSP3")

                'エピ先行評価追加対応
                 .GetText 39, intLoopCnt, vBlckID
                 .GetText 38, intLoopCnt, vStrSam
                 vStrB = Right(vBlckID, 3) & "-" & Right(vStrSam, 4)
                 .SetText 10, intLoopCnt, vStrB ''サンプルID表示
                 .SetText 10, intLoopCnt + 1, "共有"
                Exit Function
            End If
        Next
    End With
    fnc_CheckDataWfmap = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************************
'*    関数名        : fnc_DispChangeDataWfmap
'*
'*    処理概要      : 1.WFマップの表示データに画面での変更を反映する（DB登録前でも）
'*                      (WFマップ表示データ変更(WFﾏｯﾌﾟ対応）)
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*
'*  Chg 2011/03/11 SMPK Miyata  この関数内のWFﾏｯﾌﾟを表示先を変更する。
'*                              f_cmbc039_4.sprWfmapViewをf_cmbc039_4.sprWfmapViewに変更
'*                              プログラムが煩雑になるので修正履歴は残していません。
'*******************************************************************************************
Private Function fnc_DispChangeDataWfmap() As FUNCTION_RETURN
    Dim lngRoopCnt      As Long
    Dim lngMapRoopCnt   As Long
    Dim lngMapMaxRows   As Long
    Dim lngSmplMaxRows  As Long
    Dim vFromMap        As Variant
    Dim vToMap          As Variant
    Dim vNowMap         As Variant
    Dim sBlockId        As String
    Dim vBlockId        As Variant
    Dim vHinban         As Variant
    Dim vMapBlockID     As Variant
    Dim vWFStatus1      As Variant
    Dim vWFStatus2      As Variant
    Dim vStatusWfmap    As Variant
    Dim sWFStatusWfmap  As String
    Dim vGetSampId      As Variant

    fnc_DispChangeDataWfmap = FUNCTION_RETURN_SUCCESS

    lngSmplMaxRows = f_cmbc039_3.sprExamine.MaxRows
    lngMapMaxRows = f_cmbc039_4.sprWfmapView.MaxRows

    'サンプル画面の変更点→マップ画面に
    With f_cmbc039_3.sprExamine

        For lngRoopCnt = 1 To lngSmplMaxRows Step 2
            .GetText 1, lngRoopCnt, vBlockId
            If (vBlockId <> "") Then
                'サンプル_ブロックID保存
                sBlockId = CStr(vBlockId)
            End If
            'サンプル画面データ取得
            .GetText 6, lngRoopCnt, vFromMap        'サンプル_マップ位置（開始）
            .GetText 6, lngRoopCnt + 1, vToMap      'サンプル_マップ位置（終了）
            .GetText 2, lngRoopCnt, vHinban         'サンプル_品番

            'マップ画面検索＆変更
            For lngMapRoopCnt = 1 To lngMapMaxRows
                'マップ_ブロックID取得
                f_cmbc039_4.sprWfmapView.GetText 3, lngMapRoopCnt, vMapBlockID
                If sBlockId = Mid(CStr(vMapBlockID), 10, 3) Then
                    f_cmbc039_4.sprWfmapView.GetText 5, lngMapRoopCnt, vNowMap
                    '抜試行
                    If CInt(vNowMap) = CInt(vFromMap) Or CInt(vNowMap) = CInt(vToMap) Then
                        If vHinban = "Z" Then
                            vStatusWfmap = "欠落" & "(4)"
                            f_cmbc039_4.sprWfmapView.SetText 1, lngMapRoopCnt, vStatusWfmap
                            f_cmbc039_4.sprWfmapView.col = 1
                            f_cmbc039_4.sprWfmapView.col2 = 40
                            f_cmbc039_4.sprWfmapView.row = lngMapRoopCnt
                            f_cmbc039_4.sprWfmapView.row2 = lngMapRoopCnt
                            f_cmbc039_4.sprWfmapView.BlockMode = True
                            f_cmbc039_4.sprWfmapView.backColor = &H8080FF
                            f_cmbc039_4.sprWfmapView.BlockMode = False
                        Else
                            If CInt(vNowMap) = CInt(vFromMap) Then
                                .GetText 8, lngRoopCnt, vWFStatus1         'サンプル_WF状態
                                sWFStatusWfmap = CStr(vWFStatus1)
                                .GetText 10, lngRoopCnt, vGetSampId
                            Else
                                .GetText 8, lngRoopCnt + 1, vWFStatus1       'サンプル_WF状態
                                sWFStatusWfmap = CStr(vWFStatus1)
                                .GetText 10, lngRoopCnt + 1, vGetSampId
                            End If
                            '変更範囲(サンプルあり）
                            If sWFStatusWfmap = "共有" Or sWFStatusWfmap = "通常" Then
                                If vGetSampId = "共有" Then
                                    vStatusWfmap = sWFStatusWfmap & "(1)"
                                Else
                                    vStatusWfmap = sWFStatusWfmap & "(0)"
                                End If
                                f_cmbc039_4.sprWfmapView.SetText 1, lngMapRoopCnt, vStatusWfmap
                                f_cmbc039_4.sprWfmapView.col = 1
                                f_cmbc039_4.sprWfmapView.col2 = 40
                                f_cmbc039_4.sprWfmapView.row = lngMapRoopCnt
                                f_cmbc039_4.sprWfmapView.row2 = lngMapRoopCnt
                                f_cmbc039_4.sprWfmapView.BlockMode = True
                                f_cmbc039_4.sprWfmapView.backColor = &H80FF80
                                f_cmbc039_4.sprWfmapView.BlockMode = False
                            Else            'サンプルありで通常で（抜試指示WF）
                                vStatusWfmap = sWFStatusWfmap & "(0)"
                                f_cmbc039_4.sprWfmapView.SetText 1, lngMapRoopCnt, vStatusWfmap
                                f_cmbc039_4.sprWfmapView.col = 1
                                f_cmbc039_4.sprWfmapView.col2 = 40
                                f_cmbc039_4.sprWfmapView.row = lngMapRoopCnt
                                f_cmbc039_4.sprWfmapView.row2 = lngMapRoopCnt
                                f_cmbc039_4.sprWfmapView.BlockMode = True
                                f_cmbc039_4.sprWfmapView.backColor = vbYellow
                                f_cmbc039_4.sprWfmapView.BlockMode = False
                            End If
                        End If
                    '抜試行以外
                    ElseIf CInt(vNowMap) > CInt(vFromMap) And CInt(vNowMap) < CInt(vToMap) Then
                        If vHinban = "Z" Then
                            vStatusWfmap = "欠落" & "(4)"
                            f_cmbc039_4.sprWfmapView.SetText 1, lngMapRoopCnt, vStatusWfmap
                            f_cmbc039_4.sprWfmapView.col = 1
                            f_cmbc039_4.sprWfmapView.col2 = 40
                            f_cmbc039_4.sprWfmapView.row = lngMapRoopCnt
                            f_cmbc039_4.sprWfmapView.row2 = lngMapRoopCnt
                            f_cmbc039_4.sprWfmapView.BlockMode = True
                            f_cmbc039_4.sprWfmapView.backColor = &H8080FF
                            f_cmbc039_4.sprWfmapView.BlockMode = False
'Del Start 2011/03/11 SMPK Miyata
'                        Else
'                            '変更範囲(サンプルなし）
'                            vStatusWfmap = "通常" & "(0)"
'                            f_cmbc039_4.sprWfmapView.SetText 1, lngMapRoopCnt, vStatusWfmap
'                            f_cmbc039_4.sprWfmapView.col = 1
'                            f_cmbc039_4.sprWfmapView.col2 = 40
'                            f_cmbc039_4.sprWfmapView.row = lngMapRoopCnt
'                            f_cmbc039_4.sprWfmapView.row2 = lngMapRoopCnt
'                            f_cmbc039_4.sprWfmapView.BlockMode = True
'                            f_cmbc039_4.sprWfmapView.backColor = &H80FF80
'                            f_cmbc039_4.sprWfmapView.BlockMode = False
'Del End   2011/03/11 SMPK Miyata
                        End If
                    End If
                End If
            Next lngMapRoopCnt
        Next lngRoopCnt
    End With
End Function

'*******************************************************************************************
'*    関数名        : sub_LOADDISP_ADD_CHECKBOX
'*
'*    処理概要      : 1.同一SXL同一品番のブロック境界に抜試有無のチェックボックス表示
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Sub sub_LOADDISP_ADD_CHECKBOX()
    Dim i As Integer

    On Error Resume Next
    With sprExamine
        For i = 1 To .MaxRows
            CheckGetSampleID (i)
        Next i
    End With
End Sub

'*******************************************************************************************
'*    関数名        : fnc_ErrDispCheck
'*
'*    処理概要      : 1.抜試スプレッドに表示されたデータが、正常であるかチェック
'*                      ①既存抜試行にサンプルIDが入っているか
'*                      ②抜試のTOP位置とBOTTOM位置が、SXL位置と一致しているか
'*
'*    パラメータ    : 変数名    ,IO ,型      　　 ,説明
'*　　      　　      sMsg      ,O  ,String       ,表示メッセージ文字列
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function fnc_ErrDispCheck(sMsg As String) As FUNCTION_RETURN
    Dim intLoopCnt      As Integer
    Dim vNukisiFlg      As Variant
    Dim vGetSample      As Variant
    Dim vDispSIngotP    As Variant
    Dim vDispEIngotP    As Variant
    Dim vWfSta          As Variant

    With sprExamine
        'サンプルID入力チェック
        For intLoopCnt = 1 To .MaxRows

            'エピ先行評価追加
            .GetText 37, intLoopCnt, vNukisiFlg
            .GetText 8, intLoopCnt, vWfSta
            If vNukisiFlg = 1 Then
                .GetText 10, intLoopCnt, vGetSample
                If Trim(vGetSample) = vbNullString Then
                    cmdF(6).Enabled = False
                    cmdF(7).Enabled = False
                    cmdF(8).Enabled = False
                    cmdF(9).Enabled = False
                    cmdF(10).Enabled = False
                    cmdF(12).Enabled = False
                    fnc_ErrDispCheck = FUNCTION_RETURN_FAILURE
                    sMsg = "ENSP4"
                    Exit Function
                End If
            ElseIf vNukisiFlg = 3 And vWfSta <> gsWF_STA_NORMAL And vWfSta <> gsWF_STA_STA_K Then
                .GetText 10, intLoopCnt, vGetSample
                If Trim(vGetSample) = vbNullString Then
                    cmdF(6).Enabled = False
                    cmdF(7).Enabled = False
                    cmdF(8).Enabled = False
                    cmdF(9).Enabled = False
                    cmdF(10).Enabled = False
                    cmdF(12).Enabled = False
                    fnc_ErrDispCheck = FUNCTION_RETURN_FAILURE
                    sMsg = "ENSP4"
                    Exit Function
                End If
            End If
        Next

        '抜試のTOP位置とBOTTOM位置のチェック
        .GetText 5, 1, vDispSIngotP
        .GetText 5, .MaxRows, vDispEIngotP

        'tuku +-1のズレはOKとする。⇒WF幅間に入力可能値が２つある場合の対応
        fnc_ErrDispCheck = FUNCTION_RETURN_SUCCESS
    End With
End Function

'*******************************************************************************************
'*    関数名        : sub_FurikaeMotoDataSet
'*
'*    処理概要      : 1.振替元品番をセーブし、振替内容設定をクリアする
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_FurikaeMotoDataSet()
    Dim lngCnt  As Long
    Dim i       As Long
    Dim m       As Long

    ReDim MotoHinban((UBound(tExamine) + 1) / 2)

    For i = 1 To ((UBound(tExamine) + 1) / 2)
        With MotoHinban(i)
            For lngCnt = (i * 2 - 1) To (i * 2)
                If lngCnt Mod 2 = 0 Then
                .MOTOICHIE = tExamine(lngCnt - 1).RITOP_POS
                Else
                .MOTOICHIS = tExamine(lngCnt - 1).RITOP_POS
                .MOTOHIN.hinban = tExamine(lngCnt - 1).hinban
                .MOTOHIN.factory = tExamine(lngCnt - 1).factory
                .MOTOHIN.opecond = tExamine(lngCnt - 1).opecond
                .MOTOHIN.mnorevno = tExamine(lngCnt - 1).REVNUM
                End If
            Next lngCnt
        End With
    Next i

    ReDim FurikaeNaiyou(0)
    ReDim FurikaeNaiyouWK(0)

    TokuCntWK = 0

    ReDim TokusaiBangou(TokuCnt)
End Sub

'*******************************************************************************************
'*    関数名        : sub_FurikaeKouho
'*
'*    処理概要      : 1.振替可能候補品番画面より振替先品番を取得する
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_FurikaeKouho()
    With sprExamine
        .row = .ActiveRow
        .col = .ActiveCol
        If .text <> "" Then
            .col = 5
            If .text = "" Then
                lblMsg.Caption = "結晶位置がありません"
                Exit Sub
            End If
            If sub_FurikaeKouho_Check() = False Then
                Exit Sub
            End If
        Else
            lblMsg.Caption = "振替元の品番が指定されていません"
            Exit Sub
        End If
    End With

    '' 振替可能候補品番画面呼び出し
    f_cmzcFKKH.Show 1

    '' 振替先品番を表示する
    If FKKH_SakiHinban <> "" Then
        With sprExamine
            .row = .ActiveRow
            .col = 2
            .text = left$(FKKH_SakiHinban, 8)
            .col = 3
            .text = Right$(FKKH_SakiHinban, 4)

            '' 処理が変更された時
            Call sprExamine_Change(2, .row)
        End With
    End If
End Sub

'*******************************************************************************************
'*    関数名        : sub_FurikaeKouho_Check
'*
'*    処理概要      : 1.品番未選択のチェックを行う
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : Boolean（チェックOK:True、チェックNG:False）
'*
'*******************************************************************************************
Public Function sub_FurikaeKouho_Check() As Boolean
    Dim intIchi         As Integer
    Dim lngCnt          As Long
    Dim sCkHinban       As String
    Dim sCkHinbanRev    As String

    sub_FurikaeKouho_Check = False

    With sprExamine
        '' 選択品番の設定
        .row = .ActiveRow
        .col = 5            ' 位置
        If .text = "" Then
            .row = .ActiveRow - 1
        End If
        intIchi = .text

        '' 振替元品番チェック
        For lngCnt = 1 To UBound(MotoHinban)
            If MotoHinban(lngCnt).MOTOICHIS <= intIchi And intIchi < MotoHinban(lngCnt).MOTOICHIE Then
                sCkHinban = MotoHinban(lngCnt).MOTOHIN.hinban
                sCkHinbanRev = Format$(MotoHinban(lngCnt).MOTOHIN.mnorevno, "00") & MotoHinban(lngCnt).MOTOHIN.factory & MotoHinban(lngCnt).MOTOHIN.opecond
                Exit For
            End If
        Next

        If (Trim$(sCkHinban) = "G" Or Trim$(sCkHinban) = "Z" Or _
            Trim$(sCkHinban) = "") Then
            .col = 2           ' 品番
            .backColor = COLOR_NG
            lblMsg.Caption = "選択の品番は振替可能品番ではありません"
            Exit Function
        End If

        '' 必要データ設定
        FKKH_Proccd = "CW760"
        FKKH_MotoHinban = sCkHinban & sCkHinbanRev
        FKKH_Crynum = txtCryNum.text
    End With

    sub_FurikaeKouho_Check = True
End Function

'*******************************************************************************************
'*    関数名        : fnc_Furikae_Check
'*
'*    処理概要      : 1.振替可否チェック
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : Boolean（チェックOK:True、チェックNG:False）
'*
'*******************************************************************************************
Private Function fnc_Furikae_Check() As Boolean
    Dim intLp           As Integer
    Dim lngCnt          As Long
    Dim lngCnt2         As Long
    Dim lngCnt3         As Long
    Dim lngCnt4         As Long
    Dim udtSakiHin      As tFullHinban
    Dim intIchi         As Integer
    Dim intIchiE        As Integer      '品番終了位置
    Dim intUmuFlg       As Integer
    Dim intTokuFlg      As Integer
    Dim intOikomiFlg    As Integer
    Dim intRet          As Integer
    Dim intErrCode      As Integer
    Dim intErrMsg       As String
    Dim sRepidT         As String
    Dim sRepidB         As String
    Dim sVgetlotid      As String
    Dim sSmpllID1       As String
    Dim intSmpGetFlg    As Integer
    Dim m, intLp2       As Integer
    Dim blWKchkFlg      As Boolean      'Warp/合成角度のﾁｪｯｸ有無
    Dim udtCAhin()      As typ_XSDCA    'ﾌﾞﾛｯｸ内品番取得用
    Dim sSqlWhere       As String       'WHERE句
    Dim udtHin          As tFullHinban  '品番
    Dim udtChkHin()     As tFullHinban  '組合せﾁｪｯｸ用品番
    Dim intHinRow()     As Integer      '品番行位置
    Dim intHinCnt       As Integer      '品番数
    Dim intHinGrp       As Integer      '品番ｸﾞﾙｰﾌﾟ(ﾌﾞﾛｯｸ単位)
    Dim sBlkChk         As String       'ﾌﾞﾛｯｸID
    Dim intOiChk        As Integer      '追込みﾁｪｯｸ
    Dim intNGrow        As Integer      '組合せﾁｪｯｸNG品番行位置

    intHinCnt = 0
    intHinGrp = 0
    intOiChk = 0
    ReDim udtChkHin(0)
    ReDim intHinRow(0)

    Dim nInpos      As Integer          ' 結晶内位置（代表サンプルＩＤ取得の為）

    fnc_Furikae_Check = False

    With sprExamine
        ReDim FurikaeNaiyouWK(.MaxRows)
        m = 0       '06/01/12 ooba
        '' 振替チェック
        For intLp = 1 To .MaxRows - 1
            blWKchkFlg = False

            .row = intLp
            .col = 5                    ' 位置
            If .text = "" Then          '下に行を挿入した場合には初期表示時に位置が設定されていないので上の位置を見る
                .row = intLp - 1
            End If
            intIchi = .text

            '品番終了位置取得
            .row = intLp + 1
            intIchiE = .text
            .row = intLp
            intOiChk = 0      '追込みﾁｪｯｸﾌﾗｸﾞ初期化

            .col = 2           ' 品番
            udtSakiHin.hinban = .text

            Call GetLastHinban(udtSakiHin.hinban, udtSakiHin) 'フル桁品番を取得

            .col = 3        ' 品番リビジョン
            .text = Format$(udtSakiHin.mnorevno, "00") & udtSakiHin.factory & udtSakiHin.opecond

            If Trim(udtSakiHin.hinban) <> "" And _
               Trim(udtSakiHin.hinban) <> "G" And _
               Trim(udtSakiHin.hinban) <> "Z" Then

                m = m + 1
                tMapHinG = tMapHin(m)
                blWKchkFlg = True
            End If

            ''品番組合せﾁｪｯｸ対応
            '上追込みﾁｪｯｸ
            If intIchi <> tExamine(0).RITOP_POS Then intOiChk = intOiChk + 1

            '下追込みﾁｪｯｸ
            If intIchiE <> tExamine(UBound(tExamine)).RITOP_POS Then intOiChk = intOiChk + 1

            'G/Z品以外
            If Trim(udtSakiHin.hinban) <> "" And Trim(udtSakiHin.hinban) <> "G" And _
               Trim(udtSakiHin.hinban) <> "Z" Then
                intHinCnt = intHinCnt + 1
                ReDim Preserve udtChkHin(intHinCnt)
                ReDim Preserve intHinRow(intHinCnt)

                '品番ｾｯﾄ
                udtChkHin(intHinCnt) = udtSakiHin
                intHinRow(intHinCnt) = intLp

                'ﾌﾞﾛｯｸID表示有なら別ﾌﾞﾛｯｸ
                    .col = 1:   sBlkChk = .text
                    If Trim(sBlkChk) <> "" Then intHinGrp = intHinGrp + 1

                'ﾌﾞﾛｯｸ単位Noｾｯﾄ
                If intHinGrp < 10 Then
                    udtChkHin(intHinCnt).Hinkubun = Chr$(intHinGrp + vbKey0)
                Else
                    udtChkHin(intHinCnt).Hinkubun = Chr$(intHinGrp - 10 + vbKeyA)
                End If
            End If

            'エピ先行評価追加対応
            .col = 39
            .row = intLp
            sVgetlotid = .text

            ' サンプルIDの取得
            .col = 5
            nInpos = CInt(.text)
            Call fnc_Furikae_Set_RepID(sRepidT, sRepidB, nInpos)

            '' nブロック1SXL用にクリアする
            .row = intLp
            .col = 10
            If Trim(.text) = "" Then sRepidT = ""
            .row = intLp + 1
            .col = 10
            If Trim(.text) = "" Then sRepidB = ""

            .row = intLp


            '' 振替有無チェック
            intUmuFlg = 0
            intOikomiFlg = 0
            For lngCnt = 1 To UBound(MotoHinban)
                If MotoHinban(lngCnt).MOTOICHIS <= intIchi And intIchi < MotoHinban(lngCnt).MOTOICHIE Then
                    If MotoHinban(lngCnt).MOTOHIN.hinban <> udtSakiHin.hinban Or _
                       MotoHinban(lngCnt).MOTOHIN.mnorevno <> udtSakiHin.mnorevno Or _
                       MotoHinban(lngCnt).MOTOHIN.factory <> udtSakiHin.factory Or _
                       MotoHinban(lngCnt).MOTOHIN.opecond <> udtSakiHin.opecond Then
                        intUmuFlg = 1
                        If MotoHinban(lngCnt).MOTOICHIS <> intIchi Then
                            intOikomiFlg = 1
                        End If
                    End If
                    .col = 2
                    .backColor = vbWhite
                    Exit For
                End If
            Next

            '' 振替可否チェック
            If intUmuFlg = 1 Then
                '' 振替内容設定データ
                FurikaeNaiyouWK(intLp).FURIUMU = intUmuFlg
                FurikaeNaiyouWK(intLp).ICHI = tblNukishi(intLp).INPOSCW         '2003/10/27 SystemBrain
                FurikaeNaiyouWK(intLp).MOTOHIN = MotoHinban(lngCnt).MOTOHIN
                FurikaeNaiyouWK(intLp).SAKIHIN = udtSakiHin
                FurikaeNaiyouWK(intLp).TREPID = sRepidT
                FurikaeNaiyouWK(intLp).BREPID = sRepidB

                '' 位置の再設定
                If tblNukishi(intLp).INPOSCW = -1 Then
                    FurikaeNaiyouWK(intLp).ICHI = intIchi
                End If

                '' 既存の特採配列の位置とｻﾝﾌﾟﾙIDを見直す
                For lngCnt3 = 1 To TokuCntWK - 1
                    For lngCnt4 = 1 To UBound(FurikaeNaiyouWK)
                        If TokusaiBangou(lngCnt3).SAKIHIN.hinban = FurikaeNaiyouWK(lngCnt4).SAKIHIN.hinban And _
                           TokusaiBangou(lngCnt3).SAKIHIN.mnorevno = FurikaeNaiyouWK(lngCnt4).SAKIHIN.mnorevno And _
                           TokusaiBangou(lngCnt3).SAKIHIN.factory = FurikaeNaiyouWK(lngCnt4).SAKIHIN.factory And _
                           TokusaiBangou(lngCnt3).SAKIHIN.opecond = FurikaeNaiyouWK(lngCnt4).SAKIHIN.opecond Then
                            TokusaiBangou(lngCnt3).ICHI = FurikaeNaiyouWK(lngCnt4).ICHI
                            TokusaiBangou(lngCnt3).TREPID = FurikaeNaiyouWK(lngCnt4).TREPID   '代表サンプルID
                            TokusaiBangou(lngCnt3).BREPID = FurikaeNaiyouWK(lngCnt4).BREPID   '代表サンプルID
                        End If
                    Next lngCnt4
                Next lngCnt3

                '' 特採番号入力チェック
                intTokuFlg = 0
                For lngCnt2 = 1 To TokuCntWK
                    If FurikaeNaiyouWK(intLp).ICHI = TokusaiBangou(lngCnt2).ICHI And _
                       FurikaeNaiyouWK(intLp).MOTOHIN.hinban = TokusaiBangou(lngCnt2).MOTOHIN.hinban And _
                       FurikaeNaiyouWK(intLp).MOTOHIN.mnorevno = TokusaiBangou(lngCnt2).MOTOHIN.mnorevno And _
                       FurikaeNaiyouWK(intLp).MOTOHIN.factory = TokusaiBangou(lngCnt2).MOTOHIN.factory And _
                       FurikaeNaiyouWK(intLp).MOTOHIN.opecond = TokusaiBangou(lngCnt2).MOTOHIN.opecond And _
                       FurikaeNaiyouWK(intLp).SAKIHIN.hinban = TokusaiBangou(lngCnt2).SAKIHIN.hinban And _
                       FurikaeNaiyouWK(intLp).SAKIHIN.mnorevno = TokusaiBangou(lngCnt2).SAKIHIN.mnorevno And _
                       FurikaeNaiyouWK(intLp).SAKIHIN.factory = TokusaiBangou(lngCnt2).SAKIHIN.factory And _
                       FurikaeNaiyouWK(intLp).SAKIHIN.opecond = TokusaiBangou(lngCnt2).SAKIHIN.opecond Then
                        intTokuFlg = 1
                        ' 特採番号データありで特採番号なしの場合
                        If TokusaiBangou(lngCnt2).BANGOU = "" Then
                            intTokuFlg = 2
                        End If
                        Exit For
                    End If
                Next

                '' 特採番号入力の場合はチェックなし
                If intTokuFlg = 0 Then
                    For lngCnt2 = 1 To TokuCntWK
                        If FurikaeNaiyouWK(intLp).MOTOHIN.hinban = TokusaiBangou(lngCnt2).MOTOHIN.hinban And _
                           FurikaeNaiyouWK(intLp).MOTOHIN.mnorevno = TokusaiBangou(lngCnt2).MOTOHIN.mnorevno And _
                           FurikaeNaiyouWK(intLp).MOTOHIN.factory = TokusaiBangou(lngCnt2).MOTOHIN.factory And _
                           FurikaeNaiyouWK(intLp).MOTOHIN.opecond = TokusaiBangou(lngCnt2).MOTOHIN.opecond And _
                           FurikaeNaiyouWK(intLp).SAKIHIN.hinban = TokusaiBangou(lngCnt2).SAKIHIN.hinban And _
                           FurikaeNaiyouWK(intLp).SAKIHIN.mnorevno = TokusaiBangou(lngCnt2).SAKIHIN.mnorevno And _
                           FurikaeNaiyouWK(intLp).SAKIHIN.factory = TokusaiBangou(lngCnt2).SAKIHIN.factory And _
                           FurikaeNaiyouWK(intLp).SAKIHIN.opecond = TokusaiBangou(lngCnt2).SAKIHIN.opecond Then
                            '' 先の特採番号データより設定
                            TokuCntWK = TokuCntWK + 1
                            ReDim Preserve TokusaiBangou(TokuCntWK)
                            TokusaiBangou(TokuCntWK).ICHI = FurikaeNaiyouWK(intLp).ICHI
                            TokusaiBangou(TokuCntWK).MOTOHIN = FurikaeNaiyouWK(intLp).MOTOHIN
                            TokusaiBangou(TokuCntWK).SAKIHIN = FurikaeNaiyouWK(intLp).SAKIHIN
                            TokusaiBangou(TokuCntWK).BANGOU = TokusaiBangou(lngCnt2).BANGOU
                            TokusaiBangou(TokuCntWK).RIYUU = TokusaiBangou(lngCnt2).RIYUU
                            TokusaiBangou(TokuCntWK).ERRRIYUU = TokusaiBangou(lngCnt2).ERRRIYUU
                            intTokuFlg = 3
                            Exit For
                        End If
                    Next

                    If intTokuFlg = 0 Then
                        If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                            intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
'                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2)
                            intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                        Else
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                            intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
'                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2)
                            intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                        End If

                        If Trim(udtSakiHin.hinban) <> "" And _
                           Trim(udtSakiHin.hinban) <> "G" And _
                           Trim(udtSakiHin.hinban) <> "Z" Then

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                        End If

                        ''複数品番対応
                        If intRet = 0 Then
                            sSqlWhere = "where crynumca = '" & Mid(tblSXL.CRYNUM, 1, 9) & Trim(sVgetlotid) & "' "
                            sSqlWhere = sSqlWhere & "and livkca = '0' "

                            'ﾌﾞﾛｯｸ内品番取得
                            If DBDRV_GetXSDCA(udtCAhin(), sSqlWhere) = FUNCTION_RETURN_FAILURE Then
                                lblMsg.Caption = GetMsgStr("EGET2", "XSDCA")
                                Exit Function
                            End If

                            '' 1-4(Cs)チェック
                            For lngCnt4 = 1 To UBound(udtCAhin)
                                udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                udtHin.opecond = udtCAhin(lngCnt4).OPECA
                                'Csの仕様ﾁｪｯｸ(1-4)についてはﾌﾞﾛｯｸ内全品番で行う
                                If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                    intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
'                                                              udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                              typ_b, typ_CType, intSmpGetFlg, _
'                                                              sSmpllID1, "", 0, 0, 3)
                                    intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
                                                              udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                              typ_b, typ_CType, intSmpGetFlg, _
                                                              sSmpllID1, "", 0, 0, 3, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                Else
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                    intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
'                                                              udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                              typ_b, typ_CType, intSmpGetFlg, _
'                                                              sSmpllID1, "", 0, 0, 3)
                                    intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
                                                              udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                              typ_b, typ_CType, intSmpGetFlg, _
                                                              sSmpllID1, "", 0, 0, 3, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                End If
                                If intRet <= 0 Then Exit For
                            Next lngCnt4

                            '' 1-4(EPD)チェック
                            If intRet = 0 Then
                                For lngCnt4 = 1 To UBound(udtCAhin)
                                    udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                    udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                    udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                    udtHin.opecond = udtCAhin(lngCnt4).OPECA
                                    'EPDの仕様ﾁｪｯｸ(1-4)についてはﾌﾞﾛｯｸ内全品番で行う
                                    If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                        intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
'                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                                  typ_b, typ_CType, intSmpGetFlg, _
'                                                                  sSmpllID1, "", 0, 0, 4)
                                        intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                                  typ_b, typ_CType, intSmpGetFlg, _
                                                                  sSmpllID1, "", 0, 0, 4, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                    Else
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                        intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
'                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                                  typ_b, typ_CType, intSmpGetFlg, _
'                                                                  sSmpllID1, "", 0, 0, 4)
                                        intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                                  typ_b, typ_CType, intSmpGetFlg, _
                                                                  sSmpllID1, "", 0, 0, 4, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                    End If
                                    If intRet <= 0 Then Exit For
                                Next lngCnt4
                            End If

                            '' 1-4(LT)チェック
                            If intRet = 0 Then
                                For lngCnt4 = 1 To UBound(udtCAhin)
                                    udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                    udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                    udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                    udtHin.opecond = udtCAhin(lngCnt4).OPECA

                                    'LTの仕様ﾁｪｯｸ(1-4)についてはﾌﾞﾛｯｸ内全品番で行う
                                    If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                        intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
'                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                                  typ_b, typ_CType, intSmpGetFlg, _
'                                                                  sSmpllID1, "", 0, 0, 5)
                                        intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, _
                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                                  typ_b, typ_CType, intSmpGetFlg, _
                                                                  sSmpllID1, "", 0, 0, 5, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                    Else
'2012/01/12 Update START DCS)Shoryuji 結晶ボトム側払出規制追加対応
'                                        intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
'                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
'                                                                  typ_b, typ_CType, intSmpGetFlg, _
'                                                                  sSmpllID1, "", 0, 0, 5)
                                        intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, _
                                                                  udtHin, udtSakiHin, intErrCode, intErrMsg, _
                                                                  typ_b, typ_CType, intSmpGetFlg, _
                                                                  sSmpllID1, "", 0, 0, 5, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                                    End If
                                    If intRet <= 0 Then Exit For
                                Next lngCnt4
                            End If
                        End If

                        If intRet = 1 Then  '振替NGと振替OK
                            '' 特採番号データ
                            TokuCntWK = TokuCntWK + 1
                            ReDim Preserve TokusaiBangou(TokuCntWK)
                            TokusaiBangou(TokuCntWK).ICHI = FurikaeNaiyouWK(intLp).ICHI
                            TokusaiBangou(TokuCntWK).MOTOHIN = FurikaeNaiyouWK(intLp).MOTOHIN
                            TokusaiBangou(TokuCntWK).SAKIHIN = FurikaeNaiyouWK(intLp).SAKIHIN
                            TokusaiBangou(TokuCntWK).BANGOU = ""
                            TokusaiBangou(TokuCntWK).RIYUU = ""
                            TokusaiBangou(TokuCntWK).ERRRIYUU = left$(intErrMsg, 25)
                            TokusaiBangou(TokuCntWK).TREPID = FurikaeNaiyouWK(intLp).TREPID   '代表サンプルID
                            TokusaiBangou(TokuCntWK).BREPID = FurikaeNaiyouWK(intLp).BREPID   '代表サンプルID

                            .col = 2           ' 品番
                            .backColor = COLOR_NG
                            lblMsg.Caption = intErrMsg
                            If bTokuKengenFlag = True Then      ' 特採権限ありの場合
                                cmdF(5).Enabled = True
                                cmdF(5).backColor = vbRed
                            End If
                            Exit Function
                        ElseIf intRet < 0 Then '振替エラー
                            lblMsg.Caption = intErrMsg
                            Exit Function
                        ElseIf intRet = 0 Then
                            .col = 2
                            .backColor = vbWhite
                        End If
                    Else
                        '振替ﾁｪｯｸ対象外のWarp/合成角度判定
                        For intLp2 = 1 To 2
                            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                                intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                          tMapHinG.HIN, tMapHinG.HIN, _
                                                          intErrCode, intErrMsg, _
                                                          typ_b, typ_CType, 0)

                                tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                                tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            End If
                        Next intLp2
                    End If

                '' 特採番号入力の場合は設定
                ElseIf intTokuFlg = 2 Then
                    For lngCnt3 = 1 To TokuCntWK
                        If TokusaiBangou(lngCnt2).MOTOHIN.hinban = TokusaiBangou(lngCnt3).MOTOHIN.hinban And _
                           TokusaiBangou(lngCnt2).MOTOHIN.mnorevno = TokusaiBangou(lngCnt3).MOTOHIN.mnorevno And _
                           TokusaiBangou(lngCnt2).MOTOHIN.factory = TokusaiBangou(lngCnt3).MOTOHIN.factory And _
                           TokusaiBangou(lngCnt2).MOTOHIN.opecond = TokusaiBangou(lngCnt3).MOTOHIN.opecond And _
                           TokusaiBangou(lngCnt2).SAKIHIN.hinban = TokusaiBangou(lngCnt3).SAKIHIN.hinban And _
                           TokusaiBangou(lngCnt2).SAKIHIN.mnorevno = TokusaiBangou(lngCnt3).SAKIHIN.mnorevno And _
                           TokusaiBangou(lngCnt2).SAKIHIN.factory = TokusaiBangou(lngCnt3).SAKIHIN.factory And _
                           TokusaiBangou(lngCnt2).SAKIHIN.opecond = TokusaiBangou(lngCnt3).SAKIHIN.opecond Then
                            '' 先の特採番号データありで特採番号ありの場合
                            If TokusaiBangou(lngCnt3).BANGOU <> "" Then
                                TokusaiBangou(lngCnt2).BANGOU = TokusaiBangou(lngCnt3).BANGOU
                                intTokuFlg = 4
                                Exit For
                            End If
                        End If
                    Next

                    '' 特採番号データを作成せずにエラー
                    If intTokuFlg = 2 Then
                        .col = 2           ' 品番
                        .backColor = COLOR_NG
                        lblMsg.Caption = "振替番号の入力がありません"
                        If bTokuKengenFlag = True Then      ' 特採権限ありの場合
                            cmdF(5).Enabled = True
                            cmdF(5).backColor = vbRed
                        End If
                        Exit Function
                    End If

                    '振替ﾁｪｯｸ対象外のWarp/合成角度判定
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                        End If
                    Next intLp2
                Else
                    '振替ﾁｪｯｸ対象外のWarp/合成角度判定
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                        End If
                    Next intLp2
                End If
            Else
                '振替ﾁｪｯｸ対象外のWarp/合成角度判定
                If blWKchkFlg Then
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '合成角度振替ﾁｪｯｸﾌﾗｸﾞｾｯﾄ
                            '判定NG
                            If intRet = 1 Then
                                If Not tMapHinG.KAKUFLG Then
                                    lblMsg.Caption = "Warp判定エラー　品番振替を行ってください。"
                                Else
                                    lblMsg.Caption = "合成角度判定エラー　品番振替を行ってください。"
                                End If
                                Exit Function
                            '振替ﾁｪｯｸｴﾗｰ
                            ElseIf intRet < 0 Then
                                lblMsg.Caption = intErrMsg
                                Exit Function
                            End If
                        End If
                    Next intLp2
                End If
            End If
            intLp = intLp + 1 '1行あとの行
        Next intLp

        '' Z品番のみの場合は品番組合せチェックは不要
        If UBound(udtChkHin) > 0 Then
            ''品番組合せﾁｪｯｸ対応
            '追込み
            If intOiChk = 2 Then
                intRet = funChkKumiHinban("CW762", tblSXL.CRYNUM, _
                                        udtChkHin(), intHinRow(), intNGrow, intErrCode, intErrMsg)
            '振替
            Else
                intRet = funChkKumiHinban("CW761", tblSXL.CRYNUM, _
                                        udtChkHin(), intHinRow(), intNGrow, intErrCode, intErrMsg)
            End If
            If intRet = 1 Then
                .row = intNGrow
                .col = 2
                .backColor = COLOR_NG
                '品番の組合せが不正です(%s)
                lblMsg.Caption = GetMsgStr("EHIN7", intErrMsg)
                Exit Function
            ElseIf intRet < 0 Then
                lblMsg.Caption = intErrMsg
                Exit Function
            End If
        End If
    End With

    fnc_Furikae_Check = True
End Function

'*******************************************************************************************
'*    関数名        : fnc_Furikae_Set_RepID
'*
'*    処理概要      : 1.振替可否チェック
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      sRepidT   ,O  ,String       ,代表サンプルID(Top)
'*　　      　　      sRepidB   ,O  ,String       ,代表サンプルID(Bot)
'*　　      　　      intInpos      ,O  ,Integer      ,結晶内位置（代表サンプルＩＤ取得の為）
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub fnc_Furikae_Set_RepID(sRepidT As String, sRepidB As String, intInpos As Integer)
    Dim i As Integer

    ' 初期化しておく
    sRepidT = ""
    sRepidB = ""
    For i = 0 To UBound(tKensa()) - 1 Step 2
        ' 結晶内位置が上下サンプルの範囲内にあるか？
        If tKensa(i).INPOSCW <= intInpos And tKensa(i + 1).INPOSCW >= intInpos Then
            sRepidT = tKensa(i).REPSMPLIDCW
            sRepidB = tKensa(i + 1).REPSMPLIDCW
            Exit For
        End If
    Next
End Sub

'*******************************************************************************************
'*    関数名        : sub_TokusaiInput
'*
'*    処理概要      : 1.特採番号入力画面にて特採番号を入力する
'*
'*    パラメータ    : 変数名      ,IO ,型      　　 ,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_TokusaiInput()
    '' 振替元品番と振替先品番の設定
    With TokusaiBangou(TokuCntWK)
        TBN_MotoHinban = .MOTOHIN.hinban & Format$(.MOTOHIN.mnorevno, "00") & .MOTOHIN.factory & .MOTOHIN.opecond
        TBN_SakiHinban = .SAKIHIN.hinban & Format$(.SAKIHIN.mnorevno, "00") & .SAKIHIN.factory & .SAKIHIN.opecond
        TBN_Bangou = .BANGOU
        TBN_Riyuu = .RIYUU
    End With

    '' 特採番号入力画面呼び出し
    f_cmzcTBN.Show 1

    '' 特採番号を設定する
    If TBN_Bangou <> "" Then
        TokusaiBangou(TokuCntWK).BANGOU = TBN_Bangou
        TokusaiBangou(TokuCntWK).RIYUU = TBN_Riyuu
    End If
End Sub

'*******************************************************************************************
'*    関数名        : fnc_RirekiKanriDB_Touroku
'*
'*    処理概要      : 1.履歴管理ＤＢの登録を行う
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      p_ErrMsg    ,O  ,string    ,ERRメッセージ
'*
'*    戻り値        : Boolean (True:OK False:NG)
'*
'*******************************************************************************************
Public Function fnc_RirekiKanriDB_Touroku(sP_ErrMsg As String) As Boolean
    Dim cnt             As Long
    Dim cnt2            As Long
    Dim intLp           As Integer
    Dim intKCNTC3       As Integer
    Dim udtDataXSDC3()  As typ_XSDC3    '工程実績配列
    Dim sWhere          As String       'SQL文字列
    Dim sDBName         As String
    Dim intTokuFlg      As Integer
    Dim sErrMsg         As String
    Dim sCGetlotid      As String
    Dim sVgetlotid      As String
    Dim intINPOSC3      As Integer

    fnc_RirekiKanriDB_Touroku = False

    cnt = 0
    ReDim FurikaeRireki(cnt)

    For intLp = 1 To UBound(FurikaeNaiyou)
        '' 振替履歴の登録データ設定
        If FurikaeNaiyou(intLp).FURIUMU = 1 Then
            cnt = cnt + 1
            ReDim Preserve FurikaeRireki(cnt)

            'ブロックID
            sprExamine.col = 39

            sprExamine.row = intLp
            sVgetlotid = sprExamine.text

            '' 工程連番取得
            intKCNTC3 = GetKCNTC3(Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid), "CW760")
            If intKCNTC3 = 0 Then
                sP_ErrMsg = "工程連番を取得できませんでした"
                Exit Function
            End If

            'WHERE条件
            sWhere = "WHERE CRYNUMC3 = '" & Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid) & "' "
            sWhere = sWhere & "AND INPOSC3 = " & FurikaeNaiyou(intLp).ICHI & " "    '結晶内開始位置
            sWhere = sWhere & "AND KCNTC3  = " & intKCNTC3 & " "

            '' 工程実績(XSDC3)データ取得
            sDBName = "XSDC3"
            If DBDRV_GetXSDC3(udtDataXSDC3, sWhere) = FUNCTION_RETURN_FAILURE Then
                sP_ErrMsg = GetMsgStr("EGET ", vbNullString, sDBName)
                Exit Function
            End If

            intINPOSC3 = -1
            If UBound(udtDataXSDC3) = 0 Then
                '' 結晶内開始位置
                intINPOSC3 = GetINPOSC3(Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid), FurikaeNaiyou(intLp).ICHI, intKCNTC3)
                If intINPOSC3 = -1 Then
                    sP_ErrMsg = "結晶内開始位置を取得できませんでした"
                    Exit Function
                End If

                'WHERE条件
                sWhere = "WHERE CRYNUMC3 = '" & Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid) & "' "
                sWhere = sWhere & "AND INPOSC3 = " & intINPOSC3 & " "
                sWhere = sWhere & "AND KCNTC3  = " & intKCNTC3 & " "

                '' 工程実績(XSDC3)データ取得
                sDBName = "XSDC3"
                If DBDRV_GetXSDC3(udtDataXSDC3, sWhere) = FUNCTION_RETURN_FAILURE Then
                    sP_ErrMsg = GetMsgStr("EGET ", vbNullString, sDBName)
                    Exit Function
                End If

                If UBound(udtDataXSDC3) = 0 Then
                    sP_ErrMsg = "工程実績(XSDC3)を取得できませんでした"
                    Exit Function
                End If
            End If

            '' 特採番号取得
            intTokuFlg = 0
            For cnt2 = 1 To TokuCntWK
                If FurikaeNaiyou(intLp).ICHI = TokusaiBangou(cnt2).ICHI And _
                   FurikaeNaiyou(intLp).MOTOHIN.hinban = TokusaiBangou(cnt2).MOTOHIN.hinban And _
                   FurikaeNaiyou(intLp).MOTOHIN.mnorevno = TokusaiBangou(cnt2).MOTOHIN.mnorevno And _
                   FurikaeNaiyou(intLp).MOTOHIN.factory = TokusaiBangou(cnt2).MOTOHIN.factory And _
                   FurikaeNaiyou(intLp).MOTOHIN.opecond = TokusaiBangou(cnt2).MOTOHIN.opecond And _
                   FurikaeNaiyou(intLp).SAKIHIN.hinban = TokusaiBangou(cnt2).SAKIHIN.hinban And _
                   FurikaeNaiyou(intLp).SAKIHIN.mnorevno = TokusaiBangou(cnt2).SAKIHIN.mnorevno And _
                   FurikaeNaiyou(intLp).SAKIHIN.factory = TokusaiBangou(cnt2).SAKIHIN.factory And _
                   FurikaeNaiyou(intLp).SAKIHIN.opecond = TokusaiBangou(cnt2).SAKIHIN.opecond Then
                    intTokuFlg = 1
                    Exit For
                End If
            Next

            With FurikaeRireki(cnt)
                .CRYNUMCE = udtDataXSDC3(1).CRYNUMC3                    ' ブロックID・結晶番号←XSDC3より
                .INPOSCE = FurikaeNaiyou(intLp).ICHI                    ' 結晶内開始位置
                If intINPOSC3 <> -1 Then
                    .INPOSCE = intINPOSC3                               ' 結晶内開始位置
                End If
                .KCNTCE = udtDataXSDC3(1).KCNTC3                        ' 工程連番
                .HINBCE = FurikaeNaiyou(intLp).SAKIHIN.hinban           ' 振替先品番
                .REVNUMCE = FurikaeNaiyou(intLp).SAKIHIN.mnorevno       ' 製品番号改訂番号(振替先)
                .FACTORYCE = FurikaeNaiyou(intLp).SAKIHIN.factory       ' 工場(振替先)
                .OPECE = FurikaeNaiyou(intLp).SAKIHIN.opecond           ' 操業条件(振替先)
                .MOTHINCE = FurikaeNaiyou(intLp).MOTOHIN.hinban         ' 振替元品番
                .MREVNUMCE = FurikaeNaiyou(intLp).MOTOHIN.mnorevno      ' 製品番号改訂番号(振替元)
                .MFACTORYCE = FurikaeNaiyou(intLp).MOTOHIN.factory      ' 工場(振替元)
                .MOPECE = FurikaeNaiyou(intLp).MOTOHIN.opecond          ' 操業条件(振替元)
                .SXLIDCE = udtDataXSDC3(1).SXLIDC3                      ' SXLID
                .WKKTCE = udtDataXSDC3(1).WKKTC3                        ' 工程
                .KNKTCE = udtDataXSDC3(1).KNKTC3                        ' 管理工程
                .REPSMPLIDTCE = FurikaeNaiyou(intLp).TREPID             ' 代表サンプルID(TOP)
                .REPSMPLIDBCE = FurikaeNaiyou(intLp).BREPID             ' 代表サンプルID(BOT)
                If intTokuFlg = 0 Then
                    .TOKNUMCE = ""                                      ' 特採番号
                    .TOKCAUSECE = ""                                    ' 特採理由
                    .ERRCAUSECE = ""                                    ' エラー理由
                Else
                    .TOKNUMCE = TokusaiBangou(cnt2).BANGOU              ' 特採番号
                    .TOKCAUSECE = TokusaiBangou(cnt2).RIYUU             ' 特採理由
                    .ERRCAUSECE = TokusaiBangou(cnt2).ERRRIYUU          ' エラー理由
                End If
                .TOKCODECE = ""                                         ' 特採理由コード
                .HULCE = udtDataXSDC3(1).TOLC3                          ' 振替長さ
                .HUWCE = udtDataXSDC3(1).TOWC3                          ' 振替重量
                .HUMCE = udtDataXSDC3(1).TOMC3                          ' 振替枚数 ←XSDC3より
                .TSTAFFCE = txtStaffID.text                             ' 登録社員ID
                .KSTAFFCE = txtStaffID.text                             ' 更新社員ID
                .SNDKCE = "0"                                           ' 送信フラグ (0:未送信)
                .SNDDAYCE = ""                                          ' 送信日付 (ブランク)
            End With
        End If
    Next

    ' nブロック1SXLのデータの振替履歴は１件
    intLp = 1
    Do While True
        If UBound(FurikaeRireki) >= intLp + 1 Then
            If FurikaeRireki(intLp).HINBCE = FurikaeRireki(intLp + 1).HINBCE And _
               FurikaeRireki(intLp).REVNUMCE = FurikaeRireki(intLp + 1).REVNUMCE And _
               FurikaeRireki(intLp).FACTORYCE = FurikaeRireki(intLp + 1).FACTORYCE And _
               FurikaeRireki(intLp).OPECE = FurikaeRireki(intLp + 1).OPECE And _
               Trim(FurikaeRireki(intLp).REPSMPLIDBCE) = "" And _
               Trim(FurikaeRireki(intLp + 1).REPSMPLIDTCE) = "" Then
                Call fnc_FurikaeRireki_Join(intLp)
                Call fnc_FurikaeRireki_Move(intLp + 1)
                ReDim Preserve FurikaeRireki(UBound(FurikaeRireki) - 1)
            Else
                intLp = intLp + 1
            End If
        Else
            Exit Do
        End If
    Loop

    For intLp = 1 To UBound(FurikaeRireki)
        '' 振替履歴の登録
        If CreateXSDCE(FurikaeRireki(intLp), sErrMsg) = False Then
            sP_ErrMsg = sErrMsg
            Exit Function
        End If
        '特採入力通知メール送信  2011/07/04 Kameda
        If intTokuFlg = 1 Then
            Call SendMailMain(FurikaeRireki(intLp))
        End If
    Next

    fnc_RirekiKanriDB_Touroku = True
End Function

'*******************************************************************************************
'*    関数名        : fnc_FurikaeRireki_Join
'*
'*    処理概要      : 1.下品番を次の履歴からコピー
'*　　　　　　　　　　2.次の履歴との合算
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      intIdx　　     ,I  ,Integer   ,振替履歴データ添え字
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Public Sub fnc_FurikaeRireki_Join(intIdx As Integer)
    ' 下品番を次の履歴からコピー
    FurikaeRireki(intIdx).REPSMPLIDBCE = FurikaeRireki(intIdx + 1).REPSMPLIDBCE
    ' 合算
    FurikaeRireki(intIdx).HULCE = CStr(CInt(FurikaeRireki(intIdx).HULCE) + CInt(FurikaeRireki(intIdx + 1).HULCE))    ' 振替長さ
    FurikaeRireki(intIdx).HUWCE = CStr(CLng(FurikaeRireki(intIdx).HUWCE) + CLng(FurikaeRireki(intIdx + 1).HUWCE))    ' 振替重量
    FurikaeRireki(intIdx).HUMCE = CStr(CInt(FurikaeRireki(intIdx).HUMCE) + CInt(FurikaeRireki(intIdx + 1).HUMCE))    ' 振替枚数
End Sub

'*******************************************************************************************
'*    関数名        : fnc_FurikaeRireki_Move
'*
'*    処理概要      : 1.履歴データを１件詰める
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      intIdx　　     ,I  ,Integer   ,振替履歴データ添え字
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Public Sub fnc_FurikaeRireki_Move(intIdx As Integer)
    ' 履歴データを１件詰める
    Dim intLp As Integer

    For intLp = intIdx To UBound(FurikaeRireki) - 1
        FurikaeRireki(intLp) = FurikaeRireki(intLp + 1)
    Next
End Sub

'*******************************************************************************************
'*    関数名        : sub_ClearTokusai
'*
'*    処理概要      : 1.特採情報クリア
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_ClearTokusai()

    '' 特採バッファクリア
    TokuCntWK = TokuCnt
    ReDim Preserve TokusaiBangou(TokuCntWK)
    '' 特採ボタンクリア
    cmdF(5).Enabled = False
    cmdF(5).backColor = &H8000000F
    'メッセージエリアクリア
    lblMsg.Caption = ""
End Sub

'*******************************************************************************************
'*    関数名        : sub_DispSumple_Hanei_Ep_1
'*
'*    処理概要      : 1.サンプル反映(エピ)
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      i  　　     ,I  ,Integer   ,Spreadの行に使用している添字
'*　　      　　      intNukisiRow　,I  ,Integer   ,抜試指示テーブル位置
'*　　      　　      intSmpkbn   ,I  ,Integer   ,サンプル区分が代表サンプルか判断する
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_DispSumple_Hanei_Ep_1(i As Integer, intNukisiRow As Integer, intSmpkbn As Integer)
    With sprExamine
        'BMD1E
        .col = 29
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDB1 = "3" Then
                If tblNukishi(i).EPINDB1CW = "1" And tblNukishi(i + 1).EPINDB1CW = "1" Then
                    tblNukishi(i + 1).EPINDB1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                    tblNukishi(i + 1).EPRESB1CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDB1CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDB1CW = "1") Then
            intSmpkbn = 1
        End If

        'BMD2E
        .col = 30
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDB2 = "3" Then
                If tblNukishi(i).EPINDB2CW = "1" And tblNukishi(i + 1).EPINDB2CW = "1" Then
                    tblNukishi(i + 1).EPINDB2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                    tblNukishi(i + 1).EPRESB2CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDB2CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDB2CW = "1") Then
            intSmpkbn = 1
        End If

        'BMD3E
        .col = 31
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDB3 = "3" Then
                If tblNukishi(i).EPINDB3CW = "1" And tblNukishi(i + 1).EPINDB3CW = "1" Then
                    tblNukishi(i + 1).EPINDB3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                    tblNukishi(i + 1).EPRESB3CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDB3CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDB3CW = "1") Then
            intSmpkbn = 1
        End If

        'OSF1E
        .col = 32
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDL1 = "3" Then
                If tblNukishi(i).EPINDL1CW = "1" And tblNukishi(i + 1).EPINDL1CW = "1" Then
                    tblNukishi(i + 1).EPINDL1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
                    tblNukishi(i + 1).EPRESL1CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDL1CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDL1CW = "1") Then
            intSmpkbn = 1
        End If

        'OSF2E
        .col = 33
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDL2 = "3" Then
                If tblNukishi(i).EPINDL2CW = "1" And tblNukishi(i + 1).EPINDL2CW = "1" Then
                    tblNukishi(i + 1).EPINDL2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
                    tblNukishi(i + 1).EPRESL2CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDL2CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDL2CW = "1") Then
            intSmpkbn = 1
        End If

        'OSF3E
        .col = 34
        .row = i
        .backColor = vbWhite
        If .text <> "2" Then
            If tblWafInd(intNukisiRow).SMP.EPIINDL3 = "3" Then
                If tblNukishi(i).EPINDL3CW = "1" And tblNukishi(i + 1).EPINDL3CW = "1" Then
                    tblNukishi(i + 1).EPINDL3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
                    tblNukishi(i + 1).EPRESL3CW = "0"
                    .row = i + 1
                    .text = "2"
                Else
                    intSmpkbn = 1
                End If
            End If
        ElseIf (.text = "2" And tblNukishi(i + 1).EPINDL3CW = "2") Or (.text = "2" And tblNukishi(i + 1).EPINDL3CW = "1") Then
            intSmpkbn = 1
        End If
    End With
End Sub

'*******************************************************************************************
'*    関数名        : sub_DispSumple_Hanei_Ep_2
'*
'*    処理概要      : 1.サンプル反映(エピ)
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      i  　　     ,I  ,Integer   ,Spreadの行に使用している添字
'*　　      　　      intNukisiRow　,I  ,Integer   ,抜試指示テーブル位置
'*　　      　　      skensa1   ,I  ,Integer   ,検査用
'*　　      　　      skensa2   ,I  ,Integer   ,検査用
'*　　      　　      intZkbn　   ,I  ,Integer   ,Z区分
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_DispSumple_Hanei_Ep_2(i As Integer, intNukisiRow As Integer, skensa1() As String, skensa2() As String, intZkbn As Integer)
    With sprExamine
         '' BMD1E
         .col = 29
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDB1CW = "1" Then
                     tblNukishi(i).EPINDB1CW = "2"
                     tblNukishi(i).EPSMPLIDB1CW = ""
                     tblNukishi(i).EPRESB1CW = "1"
                 Else
                     'コピー
                     tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                     tblNukishi(i + 1).EPRESB1CW = tblNukishi(i).EPRESB1CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDB1 = "3" Then
                     If tblNukishi(i).EPINDB1CW = "1" And tblNukishi(i + 1).EPINDB1CW = "1" Then
                         If skensa1(19) = "2" Then
                             tblNukishi(i).EPINDB1CW = "2"
                             tblNukishi(i).EPSMPLIDB1CW = tblNukishi(i + 1).EPSMPLIDB1CW
                             tblNukishi(i).EPRESB1CW = "0"
                             .text = skensa1(19)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(19) = "0" Then
             tblNukishi(i + 1).EPINDB1CW = "0"
             tblNukishi(i + 1).EPSMPLIDB1CW = ""
             tblNukishi(i + 1).EPRESB1CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If

         '' BMD2E
         .col = 30
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDB2CW = "1" Then
                     tblNukishi(i).EPINDB2CW = "2"
                     tblNukishi(i).EPSMPLIDB2CW = ""
                     tblNukishi(i).EPRESB2CW = "1"
                 Else
                     'コピー
                     tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                     tblNukishi(i + 1).EPRESB2CW = tblNukishi(i).EPRESB2CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDB2 = "3" Then
                     If tblNukishi(i).EPINDB2CW = "1" And tblNukishi(i + 1).EPINDB2CW = "1" Then
                         If skensa1(20) = "2" Then
                             tblNukishi(i).EPINDB2CW = "2"
                             tblNukishi(i).EPSMPLIDB2CW = tblNukishi(i + 1).EPSMPLIDB2CW
                             tblNukishi(i).EPRESB2CW = "0"
                             .text = skensa1(20)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(20) = "0" Then
             tblNukishi(i + 1).EPINDB2CW = "0"
             tblNukishi(i + 1).EPSMPLIDB2CW = ""
             tblNukishi(i + 1).EPRESB2CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If

         '' BMD3E
         .col = 31
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDB3CW = "1" Then
                     tblNukishi(i).EPINDB3CW = "2"
                     tblNukishi(i).EPSMPLIDB3CW = ""
                     tblNukishi(i).EPRESB3CW = "1"
                 Else
                     'コピー
                     tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                     tblNukishi(i + 1).EPRESB3CW = tblNukishi(i).EPRESB3CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDB3 = "3" Then
                     If tblNukishi(i).EPINDB3CW = "1" And tblNukishi(i + 1).EPINDB3CW = "1" Then
                         If skensa1(21) = "2" Then
                             tblNukishi(i).EPINDB3CW = "2"
                             tblNukishi(i).EPSMPLIDB3CW = tblNukishi(i + 1).EPSMPLIDB3CW
                             tblNukishi(i).EPRESB3CW = "0"
                             .text = skensa1(21)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(21) = "0" Then
             tblNukishi(i + 1).EPINDB3CW = "0"
             tblNukishi(i + 1).EPSMPLIDB3CW = ""
             tblNukishi(i + 1).EPRESB3CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If

         'OSF1E
         .col = 32
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDL1CW = "1" Then
                     tblNukishi(i).EPINDL1CW = "2"
                     tblNukishi(i).EPSMPLIDL1CW = ""
                     tblNukishi(i).EPRESL1CW = "1"
                 Else
                     tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
                     tblNukishi(i + 1).EPRESL1CW = tblNukishi(i).EPRESL1CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDL1 = "3" Then
                     If tblNukishi(i).EPINDL1CW = "1" And tblNukishi(i + 1).EPINDL1CW = "1" Then
                         If skensa1(22) = "2" Then
                             tblNukishi(i).EPINDL1CW = "2"
                             tblNukishi(i).EPSMPLIDL1CW = tblNukishi(i + 1).EPSMPLIDL1CW
                             tblNukishi(i).EPRESL1CW = "0"
                             .text = skensa1(22)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(22) = "0" Then
             tblNukishi(i + 1).EPINDL1CW = "0"
             tblNukishi(i + 1).EPSMPLIDL1CW = ""
             tblNukishi(i + 1).EPRESL1CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If

        'OSF2E
         .col = 33
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDL2CW = "1" Then
                     tblNukishi(i).EPINDL2CW = "2"
                     tblNukishi(i).EPSMPLIDL2CW = ""
                     tblNukishi(i).EPRESL2CW = "1"
                 Else
                     tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
                     tblNukishi(i + 1).EPRESL2CW = tblNukishi(i).EPRESL2CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDL2 = "3" Then
                     If tblNukishi(i).EPINDL2CW = "1" And tblNukishi(i + 1).EPINDL2CW = "1" Then
                         If skensa1(23) = "2" Then
                             tblNukishi(i).EPINDL2CW = "2"
                             tblNukishi(i).EPSMPLIDL2CW = tblNukishi(i + 1).EPSMPLIDL2CW
                             tblNukishi(i).EPRESL2CW = "0"
                             .text = skensa1(23)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(23) = "0" Then
             tblNukishi(i + 1).EPINDL2CW = "0"
             tblNukishi(i + 1).EPSMPLIDL2CW = ""
             tblNukishi(i + 1).EPRESL2CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If

        'OSF3E
         .col = 34
         .backColor = vbWhite
         If .text <> "2" And .text <> "" Then
             If intZkbn = 4 Or intZkbn = 2 Then
                 If tblNukishi(i + 1).EPINDL3CW = "1" Then
                     tblNukishi(i).EPINDL3CW = "2"
                     tblNukishi(i).EPSMPLIDL3CW = ""
                     tblNukishi(i).EPRESL3CW = "1"
                 Else
                     tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
                     tblNukishi(i + 1).EPRESL3CW = tblNukishi(i).EPRESL3CW
                 End If
             ElseIf intZkbn = 0 Then 'Zじゃないとき
                 If tblWafInd(intNukisiRow).SMP.EPIINDL3 = "3" Then
                     If tblNukishi(i).EPINDL3CW = "1" And tblNukishi(i + 1).EPINDL3CW = "1" Then
                         If skensa1(24) = "2" Then
                             tblNukishi(i).EPINDL3CW = "2"
                             tblNukishi(i).EPSMPLIDL3CW = tblNukishi(i + 1).EPSMPLIDL3CW
                             tblNukishi(i).EPRESL3CW = "0"
                             .text = skensa1(24)
                         End If
                     End If
                 End If
             End If
         ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then
             tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
         End If
         If (intZkbn = 4 Or intZkbn = 2) And skensa2(24) = "0" Then
             tblNukishi(i + 1).EPINDL3CW = "0"
             tblNukishi(i + 1).EPSMPLIDL3CW = ""
             tblNukishi(i + 1).EPRESL3CW = "0"
             .row = i + 1
             .text = ""
             .row = i
         End If
     End With
End Sub

'*******************************************************************************************
'*    関数名        : sub_DispSumple_Hanei_Ep_3
'*
'*    処理概要      : 1.サンプル反映(エピ)
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      i  　　     ,I  ,Integer   ,Spreadの行に使用している添字
'*　　      　　      intNukisiRow　,I  ,Integer   ,抜試指示テーブル位置
'*　　      　　      skensa1   ,I  ,Integer   ,検査用
'*　　      　　      skensa2   ,I  ,Integer   ,検査用
'*　　      　　      intZkbn　   ,I  ,Integer   ,Z区分
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_DispSumple_Hanei_Ep_3(i As Integer, intNukisiRow As Integer, skensa1() As String, skensa2() As String, intZkbn As Integer)
    With sprExamine
        '' BMD1E
        .col = 29
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDB1CW = "1" Then
                    tblNukishi(i + 1).EPINDB1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                    tblNukishi(i + 1).EPRESB1CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB1CW = tblNukishi(i + 1).EPSMPLIDB1CW
                    tblNukishi(i).EPRESB1CW = tblNukishi(i + 1).EPRESB1CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDB1 = "3" Then
                    If tblNukishi(i).EPINDB1CW = "1" And tblNukishi(i + 1).EPINDB1CW = "1" Then
                        If skensa2(19) = "2" Then
                            tblNukishi(i + 1).EPINDB1CW = "2"
                            tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                            tblNukishi(i + 1).EPRESB1CW = "0"
                            .text = skensa2(19)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
            tblNukishi(i).EPSMPLIDB1CW = tblNukishi(i + 1).EPSMPLIDB1CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(19) = "0" Then
            tblNukishi(i).EPINDB1CW = "0"
            tblNukishi(i).EPSMPLIDB1CW = ""
            tblNukishi(i).EPRESB1CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If

        '' BMD2E
        .col = 30
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDB2CW = "1" Then
                    tblNukishi(i + 1).EPINDB2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                    tblNukishi(i + 1).EPRESB2CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB2CW = tblNukishi(i + 1).EPSMPLIDB2CW
                    tblNukishi(i).EPRESB2CW = tblNukishi(i + 1).EPRESB2CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDB2 = "3" Then
                    If tblNukishi(i).EPINDB2CW = "1" And tblNukishi(i + 1).EPINDB2CW = "1" Then
                        If skensa2(20) = "2" Then
                            tblNukishi(i + 1).EPINDB2CW = "2"
                            tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                            tblNukishi(i + 1).EPRESB2CW = "0"
                            .text = skensa2(20)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
            tblNukishi(i).EPSMPLIDB2CW = tblNukishi(i + 1).EPSMPLIDB2CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(20) = "0" Then
            tblNukishi(i).EPINDB2CW = "0"
            tblNukishi(i).EPSMPLIDB2CW = ""
            tblNukishi(i).EPRESB2CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If

        '' BMD3E
        .col = 31
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDB3CW = "1" Then
                    tblNukishi(i + 1).EPINDB3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                    tblNukishi(i + 1).EPRESB3CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB3CW = tblNukishi(i + 1).EPSMPLIDB3CW
                    tblNukishi(i).EPRESB3CW = tblNukishi(i + 1).EPRESB3CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDB3 = "3" Then
                    If tblNukishi(i).EPINDB3CW = "1" And tblNukishi(i + 1).EPINDB3CW = "1" Then
                        If skensa2(21) = "2" Then
                            tblNukishi(i + 1).EPINDB3CW = "2"
                            tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                            tblNukishi(i + 1).EPRESB3CW = "0"
                            .text = skensa2(21)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
            tblNukishi(i).EPSMPLIDB3CW = tblNukishi(i + 1).EPSMPLIDB3CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(21) = "0" Then
            tblNukishi(i).EPINDB3CW = "0"
            tblNukishi(i).EPSMPLIDB3CW = ""
            tblNukishi(i).EPRESB3CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If

        ''OSF1E
        .col = 32
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDL1CW = "1" Then
                    tblNukishi(i + 1).EPINDL1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
                    tblNukishi(i + 1).EPRESL1CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL1CW = tblNukishi(i + 1).EPSMPLIDL1CW
                    tblNukishi(i).EPRESL1CW = tblNukishi(i + 1).EPRESL1CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDL1 = "3" Then
                    If tblNukishi(i).EPINDL1CW = "1" And tblNukishi(i + 1).EPINDL1CW = "1" Then
                        If skensa2(22) = "2" Then
                            tblNukishi(i + 1).EPINDL1CW = "2"
                            tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
                            tblNukishi(i + 1).EPRESL1CW = "0"
                            .text = skensa2(22)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                tblNukishi(i).EPSMPLIDL1CW = tblNukishi(i + 1).EPSMPLIDL1CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(22) = "0" Then
            tblNukishi(i).EPINDL1CW = "0"
            tblNukishi(i).EPSMPLIDL1CW = ""
            tblNukishi(i).EPRESL1CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If

        ''OSF2E
        .col = 33
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDL2CW = "1" Then
                    tblNukishi(i + 1).EPINDL2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
                    tblNukishi(i + 1).EPRESL2CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL2CW = tblNukishi(i + 1).EPSMPLIDL2CW
                    tblNukishi(i).EPRESL2CW = tblNukishi(i + 1).EPRESL2CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDL2 = "3" Then
                    If tblNukishi(i).EPINDL2CW = "1" And tblNukishi(i + 1).EPINDL2CW = "1" Then
                        If skensa2(23) = "2" Then
                            tblNukishi(i + 1).EPINDL2CW = "2"
                            tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
                            tblNukishi(i + 1).EPRESL2CW = "0"
                            .text = skensa2(23)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                tblNukishi(i).EPSMPLIDL2CW = tblNukishi(i + 1).EPSMPLIDL2CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(23) = "0" Then
            tblNukishi(i).EPINDL2CW = "0"
            tblNukishi(i).EPSMPLIDL3CW = ""
            tblNukishi(i).EPRESL2CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If

        ''OSF3E
        .col = 34
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '上側Z
                If tblNukishi(i).EPINDL3CW = "1" Then
                    tblNukishi(i + 1).EPINDL3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
                    tblNukishi(i + 1).EPRESL3CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL3CW = tblNukishi(i + 1).EPSMPLIDL3CW
                    tblNukishi(i).EPRESL3CW = tblNukishi(i + 1).EPRESL3CW
                End If
            ElseIf intZkbn = 0 Then 'Zではない
                If tblWafInd(intNukisiRow).SMP.EPIINDL3 = "3" Then
                    If tblNukishi(i).EPINDL3CW = "1" And tblNukishi(i + 1).EPINDL3CW = "1" Then
                        If skensa2(24) = "2" Then
                            tblNukishi(i + 1).EPINDL3CW = "2"
                            tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
                            tblNukishi(i + 1).EPRESL3CW = "0"
                            .text = skensa2(24)
                        End If
                    End If
                End If
            End If
        ElseIf (intZkbn = 3 And .text = "2") Or (intZkbn = 1 And .text = "2") Then
                tblNukishi(i).EPSMPLIDL3CW = tblNukishi(i + 1).EPSMPLIDL3CW
        End If
        If (intZkbn = 3 Or intZkbn = 1) And skensa1(24) = "0" Then
            tblNukishi(i).EPINDL3CW = "0"
            tblNukishi(i).EPSMPLIDL3CW = ""
            tblNukishi(i).EPRESL3CW = "0"
            .row = i
            .text = ""
            .row = i + 1
        End If
    End With
End Sub

'*******************************************************************************************
'*    関数名        : fnc_ChkMukesaki
'*
'*    処理概要      : 1.選択された品番に対する向先をチェックする
'*
'*    パラメータ    : 変数名      ,IO ,型      　,説明
'*　　      　　      なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function fnc_ChkMukesaki() As FUNCTION_RETURN
    Dim i As Integer
    Dim sMukesaki As String
    Dim sHinban As String

    fnc_ChkMukesaki = FUNCTION_RETURN_FAILURE

    With sprExamine
        For i = 1 To .MaxRows Step 2
            .row = i
            .col = 2
            sHinban = .text

            If Trim(sHinban) <> "Z" And Trim(sHinban) <> "G" And Trim(sHinban) <> "" Then
                If ChkMukesaki_E001(Trim(sHinban)) = FUNCTION_RETURN_FAILURE Then
                    .row = i + 1
                    .backColor = vbRed
                    Exit Function
                Else
                    .row = i + 1
                    .backColor = vbWhite
                End If
            End If
        Next i
    End With

    fnc_ChkMukesaki = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************************
'*    関数名        : sub_cmbc039_3_ChangeHinSpec
'*
'*    処理概要      : 1.WF仕様⇔エピ仕様の表示切替
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　intCategory　 ,I  ,Integer         ,表示カテゴリ(0:WF仕様,1:エピ仕様)
'*
'*    戻り値        : なし
'*
'*******************************************************************************************
Private Sub sub_cmbc039_3_ChangeHinSpec(Optional intCategory As Integer = 0)
    Dim i       As Long

    On Error Resume Next

    With f_cmbc039_3.sprSpec

        .ReDraw = False

        Select Case intCategory
        Case 0          ' WF仕様データの表示
            ' Rs(6列目)～AO(21列目)
            For i = 6 To 21 Step 1
                .ColWidth(i) = 2.75
            Next i

            ' OT1(22列目)
            i = 22:   .ColWidth(i) = 3

            ' GD(23列目)
            i = 23:   .ColWidth(i) = 2.75

            ' B1E(24列目)～OT2(30列目)
            For i = 24 To 30 Step 1
                .ColWidth(i) = 0      ' 非表示
            Next i
        Case 1          ' エピ仕様データの表示
            ' Rs(6列目)～GD(23列目)
            For i = 6 To 23 Step 1
                .ColWidth(i) = 0      ' 非表示
            Next i

            ' B1E(24列目)～OT2(30列目)
            For i = 24 To 30 Step 1
                .ColWidth(i) = 3
            Next i
        Case Else

        End Select

        .ReDraw = True
    End With
End Sub

'***********************************************************************************
'*    関数名        : fnc_GetMukesaki_XSDCB
'*
'*    処理概要      : 1.分割結晶（SXL）のSXLIDに対する向先を表示
'*
'*    パラメータ    : 変数名        ,IO ,型              ,説明
'*　　　　　　　　　　sSXLid 　　　 ,I  ,String          ,SXL ID
'*
'*    戻り値        : String(向先)
'*
'***********************************************************************************
Private Function fnc_GetMukesaki_XSDCB(sSXLID As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long      'レコード数
    Dim i           As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039.bas -- Function Getstaffauthority"

    fnc_GetMukesaki_XSDCB = FUNCTION_RETURN_FAILURE

    sSql = "Select PLANTCATCB "
    sSql = sSql & "from XSDCB "
    sSql = sSql & "where SXLIDCB = '" & Trim(sSXLID) & "' "

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    lngRecCnt = rs.RecordCount

    If lngRecCnt = 0 Then
        Exit Function
    End If

    If IsNull(rs("PLANTCATCB")) = False Then
        For i = 0 To UBound(s_MukesakiBase)
            If s_MukesakiBase(i).sMukeCode = rs("PLANTCATCB") Then
               f_cmbc039_3.lblMukesaki.Caption = s_MukesakiBase(i).sMukeName
            End If
        Next i
    End If

    rs.Close

    fnc_GetMukesaki_XSDCB = FUNCTION_RETURN_SUCCESS
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :検査項目のスプレッドの状態からサンプル有をカウントする
'　2008/03/21 aoyagi
Private Function sub_DispSample_SCnt01(IRow As Integer) As Integer
Dim strKval     As String
Dim i           As Integer '列
Dim iCnt        As Integer '件数

    iCnt = 0
    
    'スプレッドの検査項目(""：白=vbWhite、1：黒=vbBlack、2：黄色=vbYellow)
    '                   1or2：ｸﾞﾚｰ=COLOR_CryJitsu
    With sprExamine
    
        .row = IRow
        
        ''残存酸素検査項目追加による変更　03/12/09 ooba
        For i = 11 To 26
            .col = i
            If .backColor = vbBlack Then
                iCnt = iCnt + 1
            End If
            If .backColor = vbYellow Then
                iCnt = iCnt + 1
            End If
            '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
            If i = 19 Then
                If .backColor = COLOR_CryJitsu Then
                    iCnt = iCnt + 1
                End If
            End If
            '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END
        Next i

        ''残存酸素検査項目追加による変更　03/12/09 ooba
        '--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        For i = 27 To 27
            .col = i
            If .backColor = vbBlack Then
                iCnt = iCnt + 1
            End If
            If .backColor = vbYellow Then
                iCnt = iCnt + 1
            End If
        Next i

        'GD抜試指示表示追加　05/02/10 ooba
        .col = 28
        If .backColor = vbBlack Then
            iCnt = iCnt + 1
        End If
        If .backColor = vbYellow Then
            iCnt = iCnt + 1
        End If
        
        ''BM1E～OF3E
        For i = 29 To 34
            .col = i
            If .backColor = vbBlack Then
                iCnt = iCnt + 1
            End If
            If .backColor = vbYellow Then
                iCnt = iCnt + 1
            End If
        Next i
        
        ''OT2
        .col = 35
        If .backColor = vbBlack Then
            iCnt = iCnt + 1
        End If
        If .backColor = vbYellow Then
            iCnt = iCnt + 1
        End If
        
        .col = 37
        If .text = 1 Then       '初期表示行はチェック外OKとする
            iCnt = 9999
        End If
        
        .col = 8
        If .text = "欠落" Then  '欠落はチェック外OKとする
            iCnt = 9999
        End If
        
        .col = 10
        If Trim(.text) = "" Then    'サンプルID無はチェック外OKとする　08/08/01 ooba
            iCnt = 9999
        End If
        
    End With

    '件数を戻す
    sub_DispSample_SCnt01 = iCnt
    
End Function

'概要      :検査項目のスプレッドの状態から有効SXL数をカウントする
'　2008/03/21 aoyagi
Private Function sub_DispSample_SCnt02(IRow As Integer) As Integer
Dim i           As Integer '列
Dim iCnt        As Integer '件数
Dim old_bid     As String
Dim now_bid     As String
Dim flg    As String
    
    iCnt = 0
    
    With sprExamine
    
    .row = IRow
    .col = 39
    old_bid = .text
    now_bid = .text
    
    For i = IRow To .MaxRows Step 2
        .row = i
        .col = 39
        now_bid = .text
        
        If now_bid <> old_bid Then
            Exit For
        End If
        
''        .col = 37
''        flg = .text
''        If flg <> 9 Then
            iCnt = iCnt + 1
''        End If
    Next i
    
    '件数を戻す
    sub_DispSample_SCnt02 = iCnt
    
    End With

End Function

'概要      :検査項目のスプレッドの状態から同一ブロック行数をカウントする
'　2008/03/21 aoyagi
Private Function sub_DispSample_SCnt03(IRow As Integer) As Integer
Dim i           As Integer '列
Dim iCnt        As Integer '件数
Dim old_bid     As String
Dim now_bid     As String
Dim flg    As String
    
    iCnt = 0
    
    With sprExamine
    
    .row = IRow
    .col = 39
    
    old_bid = .text
    now_bid = .text
    
    For i = IRow To .MaxRows
        .row = i
        .col = 39
        now_bid = .text
        
        If now_bid <> old_bid Then
            Exit For
        End If
        
        iCnt = iCnt + 1
    Next i
    
    '件数を戻す
    sub_DispSample_SCnt03 = iCnt
    
    End With

End Function




