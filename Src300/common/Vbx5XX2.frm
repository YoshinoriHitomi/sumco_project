VERSION 5.00
Object = "{72C40FEC-8630-11D1-A417-00606704CC2B}#6.0#0"; "KeyAction.ocx"
Begin VB.Form frmVBX5XX2 
   BorderStyle     =   3  'ŒÅ’èÀÞ²±Û¸Þ
   Caption         =   "VBX5XX2"
   ClientHeight    =   8565
   ClientLeft      =   1425
   ClientTop       =   1500
   ClientWidth     =   11925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "‚l‚r ƒSƒVƒbƒN"
      Size            =   12
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   11925
   Begin VB.TextBox txtChokkei 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   18
      Top             =   4180
      Width           =   615
   End
   Begin ƒL[ˆÚ“®.KeyAction KeyAction1 
      Left            =   240
      Top             =   840
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.TextBox txtTeikouritsu 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   23
      Top             =   5055
      Width           =   975
   End
   Begin VB.TextBox txtTeikou 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   26
      Top             =   5500
      Width           =   975
   End
   Begin VB.TextBox txtSanso 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   28
      Top             =   5945
      Width           =   975
   End
   Begin VB.TextBox txtOi 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   31
      Top             =   6390
      Width           =   975
   End
   Begin VB.TextBox txtOrg 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   5220
      MaxLength       =   6
      TabIndex        =   33
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtDendo 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   19
      Top             =   4625
      Width           =   375
   End
   Begin VB.TextBox txtKakuage 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   3720
      MaxLength       =   1
      TabIndex        =   16
      Top             =   3750
      Width           =   375
   End
   Begin VB.TextBox txtGoki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   4800
      MaxLength       =   3
      TabIndex        =   15
      Top             =   3320
      Width           =   615
   End
   Begin VB.TextBox txtGoki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   14
      Top             =   3320
      Width           =   615
   End
   Begin VB.TextBox txtSeizoKbn 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2416
      Width           =   375
   End
   Begin VB.TextBox txtKisy 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   2
      Top             =   1114
      Width           =   495
   End
   Begin VB.TextBox txtHinban 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   7800
      MaxLength       =   11
      TabIndex        =   1
      Top             =   680
      Width           =   1735
   End
   Begin VB.TextBox txtHinban 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   5640
      MaxLength       =   11
      TabIndex        =   0
      Top             =   680
      Width           =   1735
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   5
      Left            =   8040
      MaxLength       =   2
      TabIndex        =   13
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   4
      Left            =   7560
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   3
      Left            =   7080
      MaxLength       =   2
      TabIndex        =   11
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   2
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   10
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   6120
      MaxLength       =   2
      TabIndex        =   9
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtMokuteki 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2850
      Width           =   495
   End
   Begin VB.TextBox txtPgid 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   5640
      MaxLength       =   8
      TabIndex        =   6
      Top             =   1982
      Width           =   1245
   End
   Begin VB.TextBox txtOrg 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   32
      Top             =   6840
      Width           =   975
   End
   Begin VB.TextBox txtOi 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   30
      Top             =   6390
      Width           =   975
   End
   Begin VB.TextBox txtTeikoKbn 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   24
      Top             =   5055
      Width           =   315
   End
   Begin VB.TextBox txtSanso 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   27
      Top             =   5945
      Width           =   975
   End
   Begin VB.TextBox txtTeikou 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   25
      Top             =   5500
      Width           =   975
   End
   Begin VB.TextBox txtHikiageX 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   2
      Left            =   6360
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1548
      Width           =   375
   End
   Begin VB.TextBox txtHikiageX 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   1
      Left            =   6000
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1548
      Width           =   375
   End
   Begin VB.TextBox txtHikiageX 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   5640
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1548
      Width           =   375
   End
   Begin VB.TextBox txtTeikouritsu 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   6
      TabIndex        =   22
      Top             =   5055
      Width           =   975
   End
   Begin VB.TextBox txtChokkei 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Index           =   0
      Left            =   3720
      MaxLength       =   3
      TabIndex        =   17
      Top             =   4180
      Width           =   615
   End
   Begin VB.TextBox txtDoba 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   6600
      MaxLength       =   2
      TabIndex        =   20
      Top             =   4625
      Width           =   495
   End
   Begin VB.TextBox txtHoui 
      Height          =   375
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   9240
      MaxLength       =   3
      TabIndex        =   21
      Top             =   4625
      Width           =   615
   End
   Begin VB.TextBox txtSansoKbn 
      Height          =   360
      IMEMode         =   3  'µÌŒÅ’è
      Left            =   6240
      MaxLength       =   1
      TabIndex        =   29
      Top             =   5945
      Width           =   315
   End
   Begin VB.Frame FraTitle 
      Height          =   735
      Left            =   30
      TabIndex        =   39
      Top             =   -90
      Width           =   11895
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4800
         TabIndex        =   40
         Top             =   270
         Width           =   6975
      End
      Begin VB.Label lblTitle 
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   270
         Width           =   4335
      End
   End
   Begin VB.Frame fraF 
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   41
      Top             =   7200
      Width           =   11895
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12] ŽÀs"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   12
         Left            =   10800
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11] ‘O‰æ–Ê"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   11
         Left            =   9840
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   10
         Left            =   8880
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚X]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   9
         Left            =   7920
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚W]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   8
         Left            =   6960
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚V]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   7
         Left            =   6000
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚U]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   6
         Left            =   5040
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚T]  @"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   4080
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚S] C³"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3120
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚R]·¬Ý¾Ù"
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2160
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚Q]»ÌÞÒÆ­°"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   1200
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F‚P]Ò²ÝÒÆ­°"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "‚l‚r ƒSƒVƒbƒN"
            Size            =   9.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label43 
      Caption         =   "i‚P.‚b‚yA‚Q.‚l‚b‚yA‚R.‚r‚l‚b‚yj"
      Height          =   255
      Index           =   3
      Left            =   6960
      TabIndex        =   34
      Top             =   1620
      Width           =   4665
   End
   Begin VB.Label Label6 
      Caption         =   "`"
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   35
      Top             =   4230
      Width           =   255
   End
   Begin VB.Label Label43 
      Caption         =   "i‚P.ƒvƒ‰ƒCƒ€A‚Q.ƒ‚ƒjƒ^A‚R.ƒ_ƒ~[A‚X.‚»‚Ì‘¼)"
      BeginProperty Font 
         Name            =   "‚l‚r ƒSƒVƒbƒN"
         Size            =   9.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   73
      Top             =   2490
      Width           =   5265
   End
   Begin VB.Label Label43 
      Caption         =   "i‚PD¸Ø½ÀÙ¶ÀÛ¸Þ‚ðœ‚­j"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   74
      Top             =   3795
      Width           =   3225
   End
   Begin VB.Label Label6 
      Caption         =   "`"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   36
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label40 
      Caption         =   "`"
      Height          =   255
      Left            =   4800
      TabIndex        =   37
      Top             =   6930
      Width           =   375
   End
   Begin VB.Label Label42 
      Caption         =   "`"
      Height          =   255
      Left            =   4800
      TabIndex        =   38
      Top             =   6585
      Width           =   375
   End
   Begin VB.Label Label39 
      Caption         =   "`"
      Height          =   255
      Left            =   4800
      TabIndex        =   42
      Top             =   5970
      Width           =   375
   End
   Begin VB.Label Label38 
      Caption         =   "`"
      Height          =   255
      Left            =   4800
      TabIndex        =   43
      Top             =   5535
      Width           =   375
   End
   Begin VB.Label Label18 
      Caption         =   "’ïR(ƒŒƒ“ƒW)"
      Height          =   255
      Left            =   2200
      TabIndex        =   69
      Top             =   5535
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Šiã‹æ•ª"
      Height          =   255
      Index           =   2
      Left            =   2200
      TabIndex        =   65
      Top             =   3795
      Width           =   1335
   End
   Begin VB.Label lblGen 
      Caption         =   "†@@‹@"
      Height          =   255
      Index           =   0
      Left            =   2200
      TabIndex        =   64
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "»•i‹æ•ª"
      Height          =   255
      Index           =   1
      Left            =   2200
      TabIndex        =   62
      Top             =   2490
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "‹@@@Ží"
      Height          =   255
      Index           =   1
      Left            =   2200
      TabIndex        =   59
      Top             =   1185
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "‚o‚f|‚h‚c"
      Height          =   255
      Left            =   2200
      TabIndex        =   61
      Top             =   2055
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ˆøã•û–@"
      Height          =   255
      Left            =   2200
      TabIndex        =   60
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Žg—p–Ú“I "
      Height          =   255
      Left            =   2200
      TabIndex        =   63
      Top             =   2925
      Width           =   1335
   End
   Begin VB.Label Label46 
      Caption         =   "iãŒÀ’l^‰ºŒÀ’lj"
      Height          =   255
      Left            =   6600
      TabIndex        =   80
      Top             =   6495
      Width           =   2385
   End
   Begin VB.Label Label45 
      Caption         =   "i‚PD‰ºŒÀ’lÆ‡ –¢“ü—Í‚ÍãŒÀ’lj"
      Height          =   255
      Left            =   6600
      TabIndex        =   79
      Top             =   5970
      Width           =   4335
   End
   Begin VB.Label Label44 
      Caption         =   "iãŒÀ’l^‰ºŒÀ’lj"
      Height          =   285
      Left            =   6600
      TabIndex        =   78
      Top             =   5535
      Width           =   2385
   End
   Begin VB.Label Label43 
      Caption         =   "i‚PD‰ºŒÀ’lÆ‡ –¢“ü—Í‚ÍãŒÀ’lj"
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   77
      Top             =   5100
      Width           =   4305
   End
   Begin VB.Label Label37 
      Caption         =   "`"
      Height          =   255
      Left            =   4800
      TabIndex        =   44
      Top             =   5100
      Width           =   375
   End
   Begin VB.Label Label21 
      Caption         =   "‚n ‚q ‚f"
      Height          =   255
      Left            =   2200
      TabIndex        =   72
      Top             =   6900
      Width           =   1335
   End
   Begin VB.Label Label20 
      Caption         =   "‚n‚‰(ƒŒƒ“ƒW)"
      Height          =   315
      Left            =   2200
      TabIndex        =   71
      Top             =   6405
      Width           =   1575
   End
   Begin VB.Label Label19 
      Caption         =   "Ž_‘f”Z“x    "
      Height          =   255
      Left            =   2200
      TabIndex        =   70
      Top             =   5970
      Width           =   1335
   End
   Begin VB.Label Label17 
      Caption         =   "’ï R —¦     "
      Height          =   255
      Left            =   2200
      TabIndex        =   68
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Label Label16 
      Caption         =   "ƒh[ƒpƒ“ƒg "
      Height          =   255
      Left            =   5280
      TabIndex        =   75
      Top             =   4665
      Width           =   1335
   End
   Begin VB.Label Label15 
      Caption         =   "Ž²•ûˆÊ "
      Height          =   285
      Left            =   8280
      TabIndex        =   76
      Top             =   4665
      Width           =   795
   End
   Begin VB.Label Label14 
      Caption         =   "“` “± Œ^"
      Height          =   255
      Left            =   2200
      TabIndex        =   67
      Top             =   4665
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "’¼Œa‹æ•ª"
      Height          =   255
      Left            =   2200
      TabIndex        =   66
      Top             =   4230
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "`"
      Height          =   255
      Index           =   0
      Left            =   7440
      TabIndex        =   45
      Top             =   720
      Width           =   255
   End
   Begin VB.Label lblGen 
      Caption         =   "•i@@”Ô "
      Height          =   255
      Index           =   1
      Left            =   2200
      TabIndex        =   58
      Top             =   750
      Width           =   1335
   End
End
Attribute VB_Name = "frmVBX5XX2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' @(h) frmVBX5XX2.frm         Ver 1.0 ( '00.01.12 ˜aò‘ò )
' @(s)
Option Explicit

Private Sub Form_Activate()
    Call MsgOut(0, "’ŠoðŒ‚ð“ü—Í‚µAC³ƒ{ƒ^ƒ“‰Ÿ‰º")
End Sub

' @(f)
'
' ‹@”\      : ƒtƒ@ƒ“ƒNƒVƒ‡ƒ“ƒL[§Œäˆ—
'
' •Ô‚è’l    : ‚È‚µ
'
' ˆø‚«”    : keyCode   -   ƒL[ƒR[ƒh
'
' ‹@”\à–¾  : “n‚³‚ê‚½ƒL[ƒR[ƒh‚ðˆ—ŠÖ”‚É“n‚·
'
' ”õl      :
'
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim exclCtl(0)  As Object
    Dim invCtl(0)   As Object
    Set invCtl(0) = KeyAction1
    ''OCXŒÄo‚µiƒRƒ“ƒgƒ[ƒ‹§ŒäˆË—Šj
    KeyAction1.Action KeyCode, 0, exclCtl(), 1, invCtl()
    ''ƒtƒ@ƒ“ƒNƒVƒ‡ƒ“ƒL[‚È‚ç‚ÎŠY“–ˆ—ŽÀs
    'If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        'Call KeyActionVbx5XX2(KeyCode)
    'End If
End Sub

' @(f)
'
' ‹@”\      : ƒL[“ü—Íˆ—
'
' •Ô‚è’l    : ‚È‚µ
'
' ˆø‚«”    : KeyAscii   -   ƒL[ƒR[ƒh
'
' ‹@”\à–¾  : ƒŠƒ^[ƒ“ƒL[‚ð‰Ÿ‚³‚ê‚½ê‡‚Íƒr[ƒv‰¹‚ðÁ‚·‚½‚ßKeyAscii‚É‚O‚ð‘ã“ü‚·‚é
'
' ”õl      :
'
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

' @(f)
'
' ‹@”\      : ƒtƒH[ƒ€ƒ[ƒh
'
' •Ô‚è’l    : ‚È‚µ
'
' ˆø‚«”    :
'
' ‹@”\à–¾  : ƒCƒxƒ“ƒgŠÖ”
'
' ”õl      : ‰æ–Ê‚Ì‹N“®ˆ—
'
'Private Sub Form_Load()
    'gbFlgVbx5xx2 = False
    'Call FrmCenter(Me)      ''ƒEƒBƒ“ƒhƒEˆÊ’uÝ’èi’†‰›j
    'Call InitVbx5XX2(True)
    
'End Sub

' @(f)
'
' ‹@”\      : ƒ{ƒ^ƒ“ƒNƒŠƒbƒN
'
' •Ô‚è’l    : ‚È‚µ
'
' ˆø‚«”    :
'
' ‹@”\à–¾  : ƒCƒxƒ“ƒgŠÖ”
'
' ”õl      : Šeƒ{ƒ^ƒ“‚Ìˆ—
'
'Private Sub cmdF_Click(Index As Integer)
    ''ƒ{ƒ^ƒ“‚É‚æ‚éˆ—
    'Select Case Index
    'Case 3
        ''ƒLƒƒƒ“ƒZƒ‹ˆ—
        'Call KeyActionVbx5XX2(vbKeyF3)
    'Case 4
        ''C³ˆ—
        'Call KeyActionVbx5XX2(vbKeyF4)
    'End Select
'End Sub

