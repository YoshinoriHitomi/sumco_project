VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_2 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "f_cmbc039_2(CW750) - 300mm�������ƃV�X�e��"
   ClientHeight    =   10875
   ClientLeft      =   1875
   ClientTop       =   2820
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
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
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   1018
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.TextBox txtDKTmpMid 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   82
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   81
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   80
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   79
      Top             =   7425
      Width           =   990
   End
   Begin VB.TextBox txtRRGMid 
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   75
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   74
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   73
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   72
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   71
      Top             =   7425
      Width           =   990
   End
   Begin VB.OptionButton optPosSelMid 
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Style           =   1  '���̨���
      TabIndex        =   70
      Top             =   7425
      Width           =   990
   End
   Begin VB.TextBox txtKisei 
      Alignment       =   2  '��������
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "PNG�ۑ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2265
      Style           =   1  '���̨���
      TabIndex        =   53
      Tag             =   "WF"
      Top             =   1605
      Width           =   855
   End
   Begin VB.TextBox txtANTempTop 
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      IMEMode         =   3  '�̌Œ�
      Left            =   1680
      MaxLength       =   7
      TabIndex        =   17
      Top             =   780
      Width           =   972
   End
   Begin VB.Frame fraF 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F10]�@�@�@WFϯ��"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F11]�@�@�O���"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�W]�@�@������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�V]�@�@������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�U]�@ �p��"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�T]�@�@�Ĕ���"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�S]�@�@������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�X]�@�@������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�P]�@�@Ҳ��ƭ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�Q]�@�@����ƭ�"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F�R]�@�@������"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "[F12]�@�@���s"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
            Name            =   "�l�r �o�S�V�b�N"
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
            Name            =   "�l�r �o�S�V�b�N"
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
         Alignment       =   2  '��������
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
         Caption         =   "�v�e�Z���^�[��������"
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "DK���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "AN���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���с^�K�v���F"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���e�����F"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�����P�ʁF"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���Ԕ����i���i�ۏ؁j"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���Ԕ������b�Z�[�W"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���R"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�Ĕ����ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "Mid �ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�����ʒu(����)"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      BackStyle       =   0  '����
      Caption         =   "���o�K��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�F��F����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "DK���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "DK���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '����
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "AN���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "AN���x"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���R"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�Ĕ����ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�s������ �ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���R"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�Ĕ����ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "��SXL�|ID"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�S���҃R�[�h"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      BackStyle       =   0  '����
      Caption         =   "�]���d�l"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "�s���� �ʒu"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "���s�ΐ�"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
Private MsComment As String                '�R�����g   07/10/05 miyatake ���F�@�\�ǉ�
'>>>>> add 2011/07/14 Marushita
''  �E�B���h�E�̕\���ʒu�E��ԕύX
Private Declare Function SetWindowPos Lib "user32" _
        (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal x As Long, ByVal Y As Long, ByVal cx As Long, _
         ByVal cy As Long, ByVal wFlags As Long) As Long

'�E�C���h�E�摜�̃f�o�C�X�R���e�L�X�g�擾
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'�f�o�C�X�R���e�L�X�g�̉��
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

Private Const SWP_NOSIZE = &H1              ''�T�C�Y���w�肵�Ȃ�
Private Const SWP_NOMOVE = &H2              ''�ʒu���w�肵�Ȃ�
Private Const HWND_TOPMOST = -1             ''��Ɏ�O
Private Const HWND_NOTOPMOST = -2           ''�őO�ʕ\������
Private Const SRCCOPY = &HCC0020
Private Const SCALEPER = 85                 ''�k����
'<<<<< add 2011/07/14 Marushita

'*******************************************************************************
'*    �֐���        : CmdChangeWF_EP_Click
'*
'*    �����T�v      : 1.�v�e�̃G�s�؊����{�^���N���b�N����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub CmdChangeWF_EP_Click()
    Dim i       As Integer              'Add 2011/03/23 SMPK Miyata

'    On Error Resume Next
    
    CmdChangeWF_EP.Enabled = False
    
    '�e����N���A
    CtrlEnabled txtSXLTop, CTRL_DISABLE, True       'SXL�g�b�v�ʒu
    CtrlEnabled txtCutPosTop, CTRL_DISABLE, True    '�ŃJ�b�g�ʒu�i�g�b�v�j
    CtrlEnabled txtRRGTop, CTRL_DISABLE, True       'RRG�i�g�b�v�j
    CtrlEnabled txtSXLTail, CTRL_DISABLE, True      'SXL�e�C���ʒu
    CtrlEnabled txtCutPosTail, CTRL_DISABLE, True   '�ŃJ�b�g�ʒu�i�e�C���j
    CtrlEnabled txtRRGTail, CTRL_DISABLE, True      'RRG�i�e�C���j
    CtrlEnabled txtJHAll, CTRL_DISABLE, True        '���s�ΐ͑S��
    CtrlEnabled txtANTempTop, CTRL_DISABLE, True    'AN���x�i�g�b�v�j
    CtrlEnabled txtANTempTail, CTRL_DISABLE, True   'AN���x�i�e�C���j
    CtrlEnabled txtRoJdg, CTRL_DISABLE, True        '�F��F����     *2008/08/28 kameda
    CtrlEnabled txtKisei, CTRL_DISABLE, True        '���o�K��       *2010/02/15 kameda
'Add Start 2011/03/23 SMPK Miyata
    CtrlEnabled txtSXLMid, CTRL_DISABLE_SKY, True      'SXL���Ԉʒu
    CtrlEnabled txtCutPosMid, CTRL_DISABLE_SKY, True   '�ŃJ�b�g�ʒu�i���ԁj
    CtrlEnabled txtRRGMid, CTRL_DISABLE_SKY, True      'RRG�i���ԁj
'Add End   2011/03/23 SMPK Miyata
'Add Start 2011/08/25 Y.Hitomi
    CtrlEnabled txtANTempMid, CTRL_DISABLE_SKY, True   'AN���x�i����)
    CtrlEnabled txtDKTmpMid, CTRL_DISABLE_SKY, True    'DK���x�i����)
'Add End   2011/08/25 Y.Hitomi

    '���R���N���A
    SpCtrlBlockEnabled Me.spdMeasTop, 1, 1, 1, 5, CTRL_DISABLE, True
    SpCtrlBlockEnabled Me.spdMeasTail, 1, 1, 1, 5, CTRL_DISABLE, True
'Add Start 2011/03/23 SMPK Miyata
    SpCtrlBlockEnabled spdMeasMid, 1, 1, spdMeasMid.MaxCols, spdMeasMid.MaxRows, CTRL_DISABLE_SKY, True
'Add End   2011/03/23 SMPK Miyata

    '���я��N���A
    SpCtrlBlockEnabled Me.spdKensaTop, 1, -1, 12, -1, CTRL_DISABLE
    SpCtrlBlockEnabled Me.spdKensaTail, 1, -1, 12, -1, CTRL_DISABLE
'Add Start 2011/07/21 Y.Hitomi
    SpCtrlBlockEnabled Me.spdKensaMid, 1, -1, 12, -1, CTRL_DISABLE_SKY
'Add End   2011/07/21 Y.Hitomi

    '' WF �� EP
    If CmdChangeWF_EP.Tag = "WF" Then
        '�d�l�\��
        Call sub_cmbc061_2_ChangeHinSpec(1)
        '���я��\��
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlTop039
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlTail039

        CmdChangeWF_EP.Tag = "EP"
        CmdChangeWF_EP.Caption = "�G�s >>"
        EPSiyouSansyouFlg = True
        cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039)
        cmdF(5).Enabled = (txtJfName.text <> "")
    '' EP �� WF
    Else
        '�d�l�\��
        Call sub_cmbc061_2_ChangeHinSpec(0)
        '���R���\��
        sub_PutRs
        'AN���x�FDKAN��3�`6����AN���x
        Me.txtANTempTop.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTop039, WFRES).DKAN, 3, 4), "0") 'AN���x
        '�`�F�b�NNG�̎��͔w�i�F��ς���
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "TOP") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    If Not (typ_CType.JudgAntnp(SxlTop039)) Then
                        CtrlEnabled Me.txtANTempTop, CTRL_DISABLE_WARNING, False  'AN���x
                    End If
                End If
            End If
        End If
        'AN���x�FDKAN��3�`6����AN���x
        Me.txtANTempTail.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTail039, WFRES).DKAN, 3, 4), "0") 'AN���x
        '�`�F�b�NNG�̎��͔w�i�F��ς���
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "BOT") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTail039).WFRESRS1CW = "1" Then
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled Me.txtANTempTail, CTRL_DISABLE_WARNING, False  'AN���x
                    End If
                End If
            End If
        End If
        '���я��\��
        sub_PutRslt typ_CType.typ_rslt(), SxlTop039
        sub_PutRslt typ_CType.typ_rslt(), SxlTail039

        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "�v�e >>"
    End If
'Add Start 2011/03/23 SMPK Miyata
    '���Ԉʒu�I���{�^�������[�v
    For i = optPosSelMid.LBound To optPosSelMid.UBound
        If optPosSelMid(i).Value = True Then
            '���Ԕ����T���v���ʒu�{�^���N���b�N�������s��
            Call optPosSelMid_Click(i)
            Exit For
        End If
    Next i
'Add End   2011/03/23 SMPK Miyata

    CmdChangeWF_EP.Enabled = True
End Sub

'*******************************************************************************
'*    �֐���        : Form_Unload
'*
'*    �����T�v      : 1.Form_Unload����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    Cancel        ,I  ,Integer
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Unload WFCJudgDialog
End Sub



'Add Start 2011/03/09 SMPK Miyata
'*******************************************************************************
'*    �֐���        : optPosSelMid_Click
'*
'*    �����T�v      : 1.���Ԕ����T���v���ʒu�{�^���N���b�N
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub optPosSelMid_Click(Index As Integer)
    
    If CmdChangeWF_EP.Tag = "WF" Then
        '���R�l�\��(���Ԕ���)
        Call sub_PutRsMid(Index + 1)
        
        sub_PutRslt typ_CType.typ_rslt(), SxlMidl039 + Index
    Else
        sub_PutRslt_EP typ_CType_EP.typ_rslt(), SxlMidl039 + Index
    End If
    
End Sub
'Add End   2011/03/09 SMPK Miyata

'*******************************************************************************
'*    �֐���        : txtJfName_Change
'*
'*    �����T�v      : 1.�S���҂��ύX�ɂȂ����ꍇ�A�Ĕ����Ǝ��s�{�^���̌�����ύX
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub txtJfName_Change()
    Dim StopFlgF5 As Boolean
    Dim StopFlgF12 As Boolean
    
    'add ������~�`�F�b�N�ǉ� SETkimizuka Start
    StopFlgF5 = CheckXODY4(WATCH_PROCCD_NUKISI, "", txtSxlId.text)
    StopFlgF12 = CheckXODY4(WATCH_PROCCD, "", txtSxlId.text)
    'add ������~�`�F�b�N�ǉ� SETkimizuka End
    
    
'upd ������~�`�F�b�N�ǉ� SETkimizuka Start
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
''    cmdF(5).Enabled = (txtJfName.Text <> "")
'    cmdF(5).Enabled = ((txtJfName.text <> "") _
'                        And ((typ_CType.typ_si.HEPHS = False) _
'                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)))
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    cmdF(5).Enabled = ((txtJfName.text <> "") _
                        And ((typ_CType.typ_si.HEPHS = False) _
                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)) And (StopFlgF5 = True))
'upd ������~�`�F�b�N�ǉ� SETkimizuka End
'    cmdF(6).Enabled = (txtJfName.Text <> "")
''2001/12/18 S.Sano    cmdF(12).Enabled = ((txtJfName.Text <> "") And TotalJudg)
'�v�e�T���v�������ύX 2003.05.20 yakimura
'    cmdF(12).Enabled = ((txtJfName.Text <> "") And (TotalJudg Or bPPlus Or bNPlus)) ''2001/12/18 S.Sano
'upd ������~�`�F�b�N�ǉ� SETkimizuka Start
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
''    cmdF(12).Enabled = ((txtJfName.Text <> "") And TotalJudg039)
'    '' �G�s�d�l�A���сA���茋��NG���Q�ƍς݂̏ꍇ�͎��s�{�^���������\�ɂ���
'    cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039 _
'                        And ((typ_CType.typ_si.HEPHS = False) _
'                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)))
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
'�v�e�T���v�������ύX 2003.05.20 yakimura
    '' �G�s�d�l�A���сA���茋��NG���Q�ƍς݁A������~�̏ꍇ�͎��s�{�^���������\�ɂ���
    cmdF(12).Enabled = ((txtJfName.text <> "") And TotalJudg039 _
                        And ((typ_CType.typ_si.HEPHS = False) _
                            Or (typ_CType.typ_si.HEPHS = True And EPSiyouSansyouFlg = True)) And (StopFlgF12 = True))
'upd ������~�`�F�b�N�ǉ� SETkimizuka End

    '������~�̏ꍇ�̓��b�Z�[�W�\������ 2010/06/16 SETsw kubota
    If StopFlgF12 = False Then
        Call MsgOut(0, PROCD_WFC_SOUGOUHANTEI & "�H���ł̗�����~�i�ł��B(F12�s��)", DEBUG_DISP)
    End If
    If StopFlgF5 = False Then
        Call MsgOut(0, left$(WATCH_PROCCD_NUKISI, Len(WATCH_PROCCD_NUKISI) - 1) & "�H���ł̗�����~�i�ł��B(F5,F12�s��)", DEBUG_DISP)
    End If

End Sub

'*******************************************************************************
'*    �֐���        : txtStaffID_Change
'*
'*    �����T�v      : 1.�S���R�[�h�ύX����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub txtStaffID_Change()
    If STAFFIDBUFF <> Trim(txtStaffID.text) Then
        txtJfName.text = ""
    End If
End Sub

'*******************************************************************************
'*    �֐���        : txtStaffID_KeyDown
'*
'*    �����T�v      : 1.�S���҃R�[�h���̓`�F�b�N����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    KeyCode       ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift         ,I  ,Integer  ,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub txtStaffID_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim FuncAns As FUNCTION_RETURN

'    '' ��ʕ\�����b�Z�[�W�N���A
'    lblMsg.Caption = ""
    
    If KeyCode = vbKeyReturn And txtStaffID.Locked <> True Then
        '' ��ʕ\�����b�Z�[�W�N���A
        lblMsg.Caption = ""
        FuncAns = StaffIDCheck(txtStaffID, txtJfName, lblMsg)
    End If
End Sub

'*******************************************************************************
'*    �֐���        : cmdF_Click
'*
'*    �����T�v      : 1.�t�@���N�V�����{�^�����N���b�N���ꂽ��A�e�����ɕ��򂷂�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    Index       ,I  ,Integer�@,�R���g���[���z��̓Y��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub cmdF_Click(intIndex As Integer)
    
    Dim sErrMsg   As String
    
    '' ��������
    Select Case intIndex
        Case 1          '' �e�P�L�[�i���C�����j���[�j
            '' �v���O�����I������
             GotoMainMenu
        Case 2          '' �e�Q�L�[�i�T�u���j���[�j
            '' �T�u���j���[�ɖ߂�
            GotoSubMenu
        Case 11 ''�e11�i�O��ʁj
            '' �O��ʂɖ߂�
            intModoru = 1
            Unload Me
            f_cmbc039_1.Visible = True
            CloseFormProc f_cmbc039_1, f_cmbc039_2
        Case 5 '' �e5�L�[�i�Ĕ����j
            '>>>>> add 2011/07/14 Marushita
            '�L���v�`���̕ۑ�
            Call saveCapture_BMP(Me, pic_Png)
            '<<<<< add 2011/07/14 Marushita
            
            '' �����Ď��`�F�b�N add 09/03/17 SETkimizuka
            If CheckXODY4(WATCH_PROCCD_NUKISI, "", txtSxlId.text) = False Then
                lblMsg.Caption = Y4_STOP_ERR
                Exit Sub
            End If
            
            '' ���s�������s��
            typ_CType.StrStaffId = txtStaffID.text
            typ_CType.strStaffName = txtJfName.text
            If fnc_ExecutionProcess(intIndex) = FUNCTION_RETURN_FAILURE Then
                Exit Sub
            End If
                    
            '' �ăJ�b�g��ʂɑJ��
            CloseFormProc f_cmbc039_3, f_cmbc039_2
        Case 10
            'WFϯ�ߊǗ�ð��ق����ް����擾
            If SelWFmap(vbNullString, SelectSxlID039, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_2.lblMsg.Caption = sErrMsg
                Exit Sub
            End If
            
            '���گ�ނ��ް���\��
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
        Case 12       '' �e12�L�[�i���s�j
            '' �S����ID�̃`�F�b�N
            If f_cmzcChkUser.CanExec(Me.Name, txtStaffID.text) = False Then
                lblMsg.Caption = GetMsgStr("EUSR0")
                Exit Sub
            End If
            
            '' �����Ď��`�F�b�N add 09/03/17 SETkimizuka
            If CheckXODY4(WATCH_PROCCD_ENT, "", txtSxlId.text) = False Then
                lblMsg.Caption = Y4_STOP_ERR
                Exit Sub
            End If
            
            If MsgBox(GetMsgStr("PIN01"), vbOKCancel, "WF��������") = vbOK Then
                
                ' ���F�@�\�ǉ��ɂ��C��  2007/10/05 miyatake ===================> START
                '' �R�����g����
                If Me.chk_Png = 1 Then
                    If f_comment.GetComment(MsComment) <> vbOK Then
                        Exit Sub
                    End If
                    Call SetForceForegroundWindow(Me.hwnd)
                End If
                ' ���F�@�\�ǉ��ɂ��C��  2007/10/05 miyatake ===================> START
                
                BeginProcess '' �v���Z�X�J�n
                '' ���s�������s��
                If fnc_ExecutionProcess(intIndex) = FUNCTION_RETURN_FAILURE Then
                    EndProcess '' �v���Z�X�I��
                    Exit Sub
                End If
                EndProcess '' �v���Z�X�I��
                        
                '' �O��ʂɖ߂�
                intModoru = 2
                Unload Me
                f_cmbc039_1.Visible = True
            End If
    End Select
End Sub

'*******************************************************************************
'*    �֐���        : Form_KeyDown
'*
'*    �����T�v      : 1.�L�[�{�[�h��������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer�@,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    '' ��ʕ\�����b�Z�[�W�N���A
    lblMsg.Caption = ""
    '' �t�@���N�V�����L�[���L���Ȃ�
    If KeyCode >= 112 And KeyCode <= 123 Then
        If cmdF(KeyCode - 111).Enabled = True Then
            '' �t�@���N�V�����L�[�������������s����
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
'*    �֐���        : Form_Load
'*
'*    �����T�v      : 1.Form_Load����
'*                    2.Warp����p�ް��擾
'*                    3.�U�։ۃ`�F�b�N�i�d�l�j
'*                    4.Warp/�����p�x���\��
'*                    5.�K�i���\��
'*                    6.���R���\��
'*                    7.���я��\��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Load()
    Me.Hide
    Me.Show
    DoEvents

    Load WFCJudgDialog
    CtrlEnabled txtStaffID, CTRL_ENABLE, True       '�S���҃R�[�h
    CtrlEnabled txtJfName, CTRL_DISABLE, True       '�S���Җ�
    CtrlEnabled txtSxlId, CTRL_DISABLE, True      '�u���b�NID
    txtStaffID.text = typ_AType.StrStaffId ' �X�^�b�tID
    txtJfName.text = typ_AType.strStaffName ' �X�^�b�t��
    txtSxlId.text = SelectSxlID039 ' �u���b�NID�̕\��
    SpCtrlInit spdKensaTop, 0
    SpCtrlInit spdKensaTail, 0
    sprWarp.MaxRows = 0             '05/12/15 ooba
    
    'Add Start 2011/04/28 SMPK Miyata (�����\�����ɂ�����h�~)
    '�\����ʃN���A
    sub_InitDisp
    'Add End   2011/04/28 SMPK Miyata

    ' ���ݓ����̕\��
    '' �������ԃZ�b�g
    SetPresentTime lblTime

    ' �o�[�W�������̕\��
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    '' �t�H�[���ʒu�Z�b�g
    CenterForm Me
    
    f_cmbc039_2.Enabled = False
    BeginProcess '' �v���Z�X�J�n
    lblMsg.Caption = GetMsgStr(PWAIT)
    DoEvents
        
    'Del Start 2011/04/28 SMPK Miyata (�����\�����ɂ�����h�~) ��������ֈړ�
    ''�\����ʃN���A
    'sub_InitDisp
    'Del End   2011/04/28 SMPK Miyata
    
    Dim intErrCode As Integer
    Dim strErrMsg As String
    Dim intRet As Integer
    
'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    Dim wkXsdcw     As typ_XSDCW
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------
'>>>>> add start 2011/06/30 Marushita
    Dim iMinMidCnt      As Integer       '���Ԕ����̕K�v��
    Dim iRstMidCnt      As Integer       '���Ԕ����̌���
    Dim iMSMPTANI       As Integer       '���Ԕ����P��(mm)
'<<<<< add end 2011/06/30 Marushita
        
    'Add Start 2011/09/29 Y.Hitomi
    Dim sSXLIDFLG       As Integer       '�r�w�k�h�c�m��ۃt���O
    'Add End   2011/09/29 Y.Hitomi
        
'���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
' ���̃v���V�[�W�����ŃN���A�p�̕ϐ���錾����ƁAVB�̗e�ʐ����Ɉ���������̂Ť
' �ʃv���V�[�W���ŃN���A����
    'typ_Ctype��������
'    Dim clear_typeC As typ_AllTypesC
'    typ_CType = clear_typeC
    Call Crear_type_Siyou_Spv
'���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------

'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    typ_CType.JudgDkTmp(SxlTop) = JUDG_OK
    typ_CType.JudgDkTmp(SxlTail) = JUDG_OK
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------
    
    Call InitHensu2(typ_CType)   '2003-11-01 SystemBrain �ǉ�
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    Call InitHensu2_EP(typ_CType_EP)
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

    '2009/10/15 Kameda
    ReDim gNinteiro_Data(0)
'Del Start 2012/07/09 Y.Hitomi
'        2011/05/31 Kameda
'    ReDim tbl_chk2_5.MLTJDG(1)
'Del End 2012/07/09 Y.Hitomi
    
    '��ʏ��ݒ�

    ''Warp����Ή��@06/01/11 ooba START ==================================>
    'Warp����p�ް��擾
    If fnc_LoadData_Warp() = FUNCTION_RETURN_FAILURE Then
        f_cmbc039_2.Enabled = True
        f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        EndProcess
        Exit Sub
    End If
    
'--------------- 2008/08/25 INSERT START  By Systeh ---------------
    ' DK���x(����)�擾
    typ_CType.DkTmpJsk(SxlTop) = GetWfDKTmpCode(False, typ_AType.typ_Param.WFSMP(SxlTop))
    typ_CType.DkTmpJsk(SxlTail) = GetWfDKTmpCode(False, typ_AType.typ_Param.WFSMP(SxlTail))
    ' DK���x(�d�l)�擾
    wkXsdcw.HINBCW = typ_AType.typ_Param.HINBCA
    wkXsdcw.REVNUMCW = typ_AType.typ_Param.REVNUMCA
    wkXsdcw.FACTORYCW = typ_AType.typ_Param.FACTORYCA
    wkXsdcw.OPECW = typ_AType.typ_Param.OPECA
    typ_CType.DkTmpSiyo = GetWfDKTmpCode(True, wkXsdcw)
'--------------- 2008/08/25 INSERT  END   By Systeh ---------------

    ReDim tWarpMeasG(0)
    ReDim tKakuMeasG(0)
    'Add Start 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    tMapHinG = tMapHin(1)
    tNew_Hinban = tMapHinG.HIN
    ''Warp����Ή��@06/01/11 ooba END ====================================>
    
    '�Ăяo���֐�������Ă�������A�ύX������B2003/10/08 SystemBrain MM
    If funChkFurikaeShiyou(PROCD_WFC_SOUGOUHANTEI, txtSxlId.text, tOld_Hinban, tNew_Hinban, _
                           intErrCode, strErrMsg, typ_b, typ_CType, 0) < FUNCTION_RETURN_SUCCESS Then
        f_cmbc039_2.Enabled = True
        f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
        lblMsg.Caption = strErrMsg
        EndProcess '' �v���Z�X�I��
'        Exit Function
        'Add Start 2011/05/10 SMPK Miyata
        lblMidMsg.Caption = typ_CType.sMidErrMsg
        'Add End   2011/05/10 SMPK Miyata
        Exit Sub
    End If
    'Add Start 2011/05/10 SMPK Miyata
    lblMidMsg.Caption = typ_CType.sMidErrMsg
    'Add End   2011/05/10 SMPK Miyata

' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
    EPSiyouSansyouFlg = False
    If typ_CType.typ_si.HEPHS = True Then
        '' �G�s��s�]�����ڂɔ���NG������ꍇ�́A�؊����{�^����ԐF�ŕ\������
        '' ���̏ꍇ�A�G�s���茋�ʂ��Q�Ƃ���܂Ŏ��s�{�^���������s�Ƃ���
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
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    
    tMapHin(1).WARPFLG = tMapHinG.WARPFLG       'Warp�U�������׸޾�ā@06/01/12 ooba
    tMapHin(1).KAKUFLG = tMapHinG.KAKUFLG       '�����p�x�U�������׸޾�ā@06/01/12 ooba
    
    If intErrCode = 0 Then
        TotalJudg039 = True
    Else
        TotalJudg039 = False
    End If
    
    '�U�����������{�i�Ԃ�Warp/�����p�x����@06/01/11 ooba START ========================>
    lblMsg.Caption = ""
    Dim i, j    As Integer
    For i = 1 To UBound(tMapHin)
        '�U���������{�̊m�F
        tMapHinG = tMapHin(i)
        For j = 1 To 2
            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                intRet = funChkFurikaeShiyou("CW763", txtSxlId.text, tMapHinG.HIN, _
                                             tMapHinG.HIN, intErrCode, strErrMsg, _
                                             typ_b, typ_CType, 0)

                tMapHin(i).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                '����NG
                If intRet = 1 Then
                    TotalJudg039 = False
                '�U�������װ
                ElseIf intRet < 0 Then
                    f_cmbc039_2.Enabled = True
                    f_cmbc039_2.txtStaffID.Locked = True
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                    f_cmbc039_2.CmdChangeWF_EP.Enabled = False
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    lblMsg.Caption = strErrMsg
                    EndProcess
                    Exit Sub
                End If
            End If
        Next j
    Next i
    'Warp����NG�ł��d�l���Ȃ��ꍇ�͑�������OK�Ƃ���
    'Nr�Z�x�ǉ��@06/06/08 ooba
    '2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
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
    'Warp����NG�̏ꍇ�װү���ޕ\��
    For i = 1 To UBound(tWarpMeasG)
        If tWarpMeasG(i).EXISTFLG >= 0 Then
            If Not tWarpMeasG(i).Judg Then
                lblMsg.Caption = "Warp����G���[�@�i�ԐU�ւ��s���Ă��������B"
                Exit For
            End If
        End If
    Next i
    'Warp/�����p�x���\��
    Call WarpKakuDisp(Me)
    '�U�����������{�i�Ԃ�Warp/�����p�x����@06/01/11 ooba END ==========================>
                
'��������typ_A����typ_C���g�p���邱��
    
    '�K�i���\��
    sub_PutSeihinTop        '��i
    sub_PutSeihinCenter     '���i
    sub_PutSeihinTail       '���i
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)hama -s-
    If typ_CType.typ_si.HEPHS = True Then
        sub_PutSeihinEpi        '�G�s
    End If
' 06/08/15 Add �G�s��s�]���ǉ��Ή� SMP)hama -e-
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'SP�\�� �� SPV(Fe)�ASPV(�g�U��)�ASPV(Nr)�\���ɂ��ύX
    sub_PutSeihinTail2      '���i2
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    'typ_rtInit
    
    '���R���\��
    sub_PutRs

    '���я��\��
    'sub_PutRslt typ_AType.typ_rslt(), SxlTop039
    sub_PutRslt typ_CType.typ_rslt(), SxlTop039
    
    'sub_PutRslt typ_AType.typ_rslt(), SxlTail039
    sub_PutRslt typ_CType.typ_rslt(), SxlTail039

'Add Start 2011/03/09 SMPK Miyata
    sub_PutRslt typ_CType.typ_rslt(), SxlMidl039

    '���Ԕ����T���v���ʒu�{�^���ݒ�
    sub_SampleMidlePosBtnSet
    
'Add End   2011/03/09 SMPK Miyata
    
'Add Start 2011/08/10 Y.Hitomi
    If typ_CType.typ_si.MSMPFLG = "1" Or typ_CType.typ_si.MSMPFLG = "2" Or typ_CType.typ_si.MSMPFLG = "3" Then
        
        If typ_CType.typ_si.MSMPFLG = "1" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "���Ԕ���(���i�ۏ�)"
        ElseIf typ_CType.typ_si.MSMPFLG = "2" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "���Ԕ���(����Q�l)"
        ElseIf typ_CType.typ_si.MSMPFLG = "3" Then
            lblMSMP_FLG.Visible = True
            lblMSMP_FLG.Caption = "���Ԕ���(����ۏ�)"
        End If
        
        '���Ԕ����P��(���Ԕ������e�l(����)/(mm))
        lblMSMP_TANI.Visible = True
        '���Ԕ����P��(mm)���擾
        If getMSMPTANI(tNew_Hinban, iMSMPTANI) = FUNCTION_RETURN_FAILURE Then
            iMSMPTANI = 0
        End If
        lblMSMP_TANI.Caption = "�����P�ʁF" & vbCrLf & _
        CInt(typ_CType.typ_si.MSMPTANIMAI) & "��"
        '���Ԕ����K�v��
        lblMSMP_SUU.Visible = True
        '���Ԕ����̕K�v�� = (SXL��WF���� - ���Ԕ������e�l(����)) / ���Ԕ����P��(����)
        iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
        '�}�C�i�X�̏ꍇ�A�O�Ƃ���
        If iMinMidCnt < 0 Then iMinMidCnt = 0
        '���Ԕ����̌���
        iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
        lblMSMP_SUU.Caption = "����/�K�v���F" & vbCrLf & _
        CInt(iRstMidCnt) & "/" & CInt(iMinMidCnt) & "��"
        '���Ԕ����P��(����)
        lblMSMP_JOSU.Visible = True
        lblMSMP_JOSU.Caption = "���e�����F" & vbCrLf & _
        CInt(typ_CType.typ_si.MSMPCONSTMAI) & "��"
    Else
        lblMSMP_FLG.Visible = False
        lblMSMP_TANI.Visible = False
        lblMSMP_SUU.Visible = False
        lblMSMP_JOSU.Visible = False
    End If
'Add End 2011/08/10 Y.Hitomi
    

'end ���ʃ��W���[���֕ύX�ɂȂ�

    '����وُ폈���ǉ��@06/10/19 ooba
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
                lblMsg.Caption = "�T���v���ُ� (" & .REPSMPLIDCW & ")"
                EndProcess
                Exit Sub
            End If
        End With
    Next i
        
    '�F��F����    *2008/08/28 kameda    mod 2009/10/15 Kameda
    'If gNinteiro_Data.JUDGRO = "0" Then
    If gNinteiro_Data(1).JUDGRO = "0" Then
        txtRoJdg.text = "OK"
    'ElseIf gNinteiro_Data.JUDGRO = "-1" Then
    ElseIf gNinteiro_Data(1).JUDGRO = "-1" Then
        TotalJudg039 = False
        txtRoJdg.text = "NG"
        CtrlEnabled txtRoJdg, CTRL_DISABLE_WARNING, False
    End If
        
    '���o���K���ǉ�   *2010/02/15 Kameda
    PutAllData_Haraidashi
    
    '�}���`���グ�K�p����  2011/05/31 Kameda
    If tbl_chk2_5.MLTJDG(1) = "-1" Then
        TotalJudg039 = False
        'Add Start 2012/07/09 Y.Hitomi
        lblMsg.Caption = "�}���`�K�p�s�i�Ԃׁ̈A�����ł��܂���B"
        'lblMsg.Caption = "�}���`���グ�K�p�G���["
        'Add End 2012/07/09 Y.Hitomi
    End If
    
    
     '>>>>> Mod Start 2012/09/07 SETsw Marushita WF10���ȉ��𗬓��Ƃ���
'    'Add Start 2010/08/26 Y.Hitomi WF10���ȉ��́A�����s�Ƃ���
'    With typ_CType.typ_Param
'        If .COUNT <= 10 Then
'            TotalJudg039 = False
'            lblMsg.Caption = "WF������10���ȉ��ł��B"
'            CtrlEnabled lblIchi, CTRL_DISABLE_WARNING, False
'        End If
'    End With
'    'Add End  2010/08/26 Y.Hitomi
     '<<<<< Mod End 2012/09/07 SETsw Marushita WF10���ȉ��𗬓��Ƃ���
    
    'Add Start 2011/09/28 Y.Hitomi SXLID�m��ۃt���O�`�F�b�N�Ή�
    If getSXLIDFLG(tNew_Hinban, sSXLIDFLG) = FUNCTION_RETURN_SUCCESS Then
        If sSXLIDFLG = "1" Then
            TotalJudg039 = False
            lblMsg.Caption = "SXLID�m��s�i�Ԃׁ̈A�����ł��܂���B"
        End If
    Else
        TotalJudg039 = False
        lblMsg.Caption = "SXLID�m��ۃ`�F�b�N�G���["
    End If
    'Add End  2011/09/28 Y.Hitomi
    
    EndProcess '' �v���Z�X�I��
    f_cmbc039_2.Enabled = True
    
    lblMukesaki.Caption = sCmbMukeName
    
    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> START
    ''PNG�ۑ��`�F�b�N�{�b�N�XON/OFF
    If Trim(GetCodeFieldA9(SWS_CHK_KEY1, SWS_CHK_KEY2, SWS_CHK_KEY3, SWS_CHK_COLUMN)) = SWS_CHK_VALUE_ON Then
        Me.chk_Png = 1
    ElseIf Trim(GetCodeFieldA9(SWS_CHK_KEY1, SWS_CHK_KEY2, SWS_CHK_KEY3, SWS_CHK_COLUMN)) = SWS_CHK_VALUE_OFF Then
        Me.chk_Png = 0
    Else
        Me.chk_Png = 0
    End If
    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> END
    
    '�����ʒu�\��  2010/02/15 Kameda
    lblIchi.Caption = GetXtalPos(txtSxlId.text)
    
    '�t�H�[�J�X�Z�b�g�i�S���ҁj
    txtStaffID.SetFocus
End Sub

'*******************************************************************************
'*    �֐���        : Sub_SetParamData
'*
'*    �����T�v      : 1.�O��ʂ���̈�����ݒ肷��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub Sub_SetParamData()
    Call Sub_S_SetParamData
End Sub

'*******************************************************************************
'*    �֐���        : sub_InitDisp
'*
'*    �����T�v      : 1.��ʂ̏���������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Public Sub sub_InitDisp()
    intEnCmd = 0                                    '�{�^���\���W��
    lblMsg.Caption = ""
    lblMidMsg.Caption = ""                          'Add 2011/05/10 SMPK Miyata

    CtrlEnabled txtSXLTop, CTRL_DISABLE, True       'SXL�g�b�v�ʒu
    CtrlEnabled txtCutPosTop, CTRL_DISABLE, True    '�ŃJ�b�g�ʒu�i�g�b�v�j
    CtrlEnabled txtRRGTop, CTRL_DISABLE, True       'RRG�i�g�b�v�j
    CtrlEnabled txtSXLTail, CTRL_DISABLE, True      'SXL�e�C���ʒu
    CtrlEnabled txtCutPosTail, CTRL_DISABLE, True   '�ŃJ�b�g�ʒu�i�e�C���j
    CtrlEnabled txtRRGTail, CTRL_DISABLE, True      'RRG�i�e�C���j
    CtrlEnabled txtJHAll, CTRL_DISABLE, True        '���s�ΐ͑S��
    CtrlEnabled txtRoJdg, CTRL_DISABLE, True        '�F��F���� *2008/08/28 kameda
    CtrlEnabled txtKisei, CTRL_DISABLE, True        '���o�K���@ *2010/02/15 kameda
'Add Start 2011/03/09 SMPK Miyata
    CtrlEnabled txtSXLMid, CTRL_DISABLE_GRAY, True      'SXL���Ԕ����ʒu
    CtrlEnabled txtCutPosMid, CTRL_DISABLE_GRAY, True   '�ŃJ�b�g�ʒu�i���Ԕ����j
    CtrlEnabled txtRRGMid, CTRL_DISABLE_GRAY, True      'RRG�i���Ԕ����j
'Add End   2011/03/09 SMPK Miyata
'Add Start 2011/08/25 Y.Hitomi
    CtrlEnabled txtANTempMid, CTRL_DISABLE_GRAY, True   'AN���x�i����)
    CtrlEnabled txtDKTmpMid, CTRL_DISABLE_GRAY, True    'DK���x�i����)
'Add End   2011/08/25 Y.Hitomi

    Call InitHensu(typ_AType)
    
    With f_cmbc039_2
        '�K�i�V�[�g��i
        ''2001/07/27 �C��
        SpCtrlInit .spdHinbanTop, 1
    '���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        'AN���x�ǉ�
        SpCtrlBlockEnabled .spdHinbanTop, 1, 1, 5, 2, CTRL_DISABLE
    '���ύX �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        '�K�i�V�[�g���i
        ''2001/07/27 �C��
        SpCtrlInit .spdHinbanCen, 1
        SpCtrlBlockEnabled .spdHinbanCen, 1, 1, 9, 2, CTRL_DISABLE
        '�K�i�V�[�g���i
        ''2001/07/27 �C��
        SpCtrlInit .spdHinbanTail, 1
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 7, 2, CTRL_DISABLE
    '*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ��ǉ�
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 10, 2, CTRL_DISABLE    'GD�d�l�\���ǉ��@05/02/04 ooba
'���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'SP�\�� �� SPV(Fe)�ASPV(�g�U��)�ASPV(Nr)�\���ɂ��ύX
'        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 11, 2, CTRL_DISABLE
        SpCtrlBlockEnabled .spdHinbanTail, 1, 1, 10, 2, CTRL_DISABLE
'���ύX SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
    '*** UPDATE �� Y.SIMIZU 2005/10/1 GDײݐ��ǉ�
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'SP�\�� �� SPV(Fe)�ASPV(�g�U��)�ASPV(Nr)�\���ɂ��ύX
        '�K�i�V�[�g���i2
        SpCtrlInit .spdHinbanTail2, 1
'        SpCtrlBlockEnabled .spdHinbanTail2, 1, 1, 3, 2, CTRL_DISABLE
        SpCtrlBlockEnabled .spdHinbanTail2, 1, 1, 13, 2, CTRL_DISABLE   '08/03/12 ooba
'���ǉ� SPV���菈���ǉ� 2006/06/12 SMP)kondoh ---------------
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        '�K�i�V�[�g�G�s
        SpCtrlInit .spdHinbanCenEpi, 1
        SpCtrlBlockEnabled .spdHinbanCenEpi, 1, 1, 6, 2, CTRL_DISABLE
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        
        '���R(TOP)
        ''2001/07/27 �C��
        SpCtrlInit .spdMeasTop, 5
        SpCtrlBlockEnabled .spdMeasTop, 1, 1, 1, 5, CTRL_DISABLE
        
        '���R(TAIL)
        ''2001/07/27 �C��
        SpCtrlInit .spdMeasTail, 5
        SpCtrlBlockEnabled .spdMeasTail, 1, 1, 1, 5, CTRL_DISABLE

'Add Start 2011/03/10 SMPK Miyata
        '���R(MIDLE)
        ''2001/07/27 �C��
        SpCtrlInit .spdMeasMid, 5
        SpCtrlBlockEnabled .spdMeasMid, 1, 1, .spdMeasMid.MaxCols, .spdMeasMid.MaxRows, CTRL_DISABLE_GRAY, True
'Add End   2011/03/10 SMPK Miyata

        '�������(TOP)
        ''2001/07/27 �C��
        SpCtrlInit .spdKensaTop, 0
        
        '�������(TAIL)
        ''2001/07/27 �C��
        SpCtrlInit .spdKensaTail, 0
        
        '�������(MIDLE)
        SpCtrlInit spdKensaMid, 0           'Add 2011/03/10 SMPK Miyata

    End With
End Sub

'*******************************************************************************
'*    �֐���        : fnc_ExecutionProcess
'*
'*    �����T�v      : 1.���͉�ʂɂ����Ă̓��͂��ꂽ�l��o�^����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    Index       ,I  ,Integer�@,Cmd�{�^���z��̓Y��
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_ExecutionProcess(Index As Integer) As FUNCTION_RETURN
    Dim udtImgData              As typImgData           '07/10/05 miyatake START ================>
    Dim udtImgData_Detail(0)    As typImgData_Detail
    Dim sErrMsg                 As String
    
    udtImgData.detail = udtImgData_Detail     '07/10/05 miyatake END ==================>
    
    '' �p�����[�^������
    fnc_ExecutionProcess = FUNCTION_RETURN_FAILURE

    '' �p�����[�^���菈�����s��
    If StaffIDCheck(txtStaffID, txtJfName, lblMsg) = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

    typ_CType.StrStaffId = Trim(txtStaffID.text) ' �X�^�b�tID
    typ_CType.strStaffName = Trim(txtJfName.text) ' �X�^�b�t��
    typ_CType.typ_Param.SXLID = SelectSxlID039 ' �u���b�NID
     
    ''�f�[�^�o�^���s��
    Select Case Index
        Case 12
            BeginProcess '' �v���Z�X�J�n
    
            If TotalJudg039 Then
                OraDB.BeginTrans
                If RegWfSogoRsltOK() <> FUNCTION_RETURN_SUCCESS Then
                    OraDB.Rollback
                    EndProcess '' �v���Z�X�I��
                    Exit Function
                End If
                Debug.Print "�VDB�����ݏ����J�n"
                If MakeParameter(WF_HANTEI_FORM) <> FUNCTION_RETURN_SUCCESS Then
                    OraDB.Rollback
                    Debug.Print "�VDB�����ݏ����ُ�I��"
                    Call clearType  '�\���̏�����
                    EndProcess '' �v���Z�X�I��
                    Exit Function
                End If
    '            OraDB.Rollback
    
                ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> START
                ''PNG�t�@�C���쐬
                If Me.chk_Png = 1 Then
                    udtImgData.xtal = BlkNow.XTALC2
                    udtImgData.STAFFID = txtStaffID
                    udtImgData.SXLID = txtSxlId
                    udtImgData.memo = MsComment
'                    If FileCreate_PNG(PROCD_WFC_SOUGOUHANTEI, udtImgData, Me, sErrMsg, Nothing, pic_Png) = FUNCTION_RETURN_FAILURE Then
                    If FileCreate_PNG(PROCD_WFC_SOUGOUHANTEI, udtImgData, Me, sErrMsg, Nothing, pic_Png) = False Then 'upd 09/02/04 SETmiyatake
                        OraDB.Rollback
                        lblMsg.Caption = sErrMsg
                        Debug.Print "PNG�t�@�C���쐬�����ُ�I��"
                        Call clearType  '�\���̏�����
                        EndProcess '' �v���Z�X�I��
                        Exit Function
                    End If
                End If
                ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> END
    
                Call clearType  '�\���̏�����
                
                OraDB.CommitTrans
                Debug.Print "�VDB�����ݏ�������I��"
                
                ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> START
                If Me.chk_Png = 1 Then
                    ''PNG�t�@�C�����M
                    Call FileReSend_PNG(PROCD_WFC_SOUGOUHANTEI)
                End If
                ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> END
            Else
                EndProcess '' �v���Z�X�I��
                lblMsg.Caption = GetMsgStr(TJE01)
                Exit Function
            End If
    End Select
    
    '' ��������I��
    fnc_ExecutionProcess = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    �֐���        : fnc_LoadData_Warp
'*
'*    �����T�v      : 1.Warp/�����p�x����p�ް��擾
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_LoadData_Warp() As FUNCTION_RETURN

    Dim i, j, k, m, n       As Integer
    Dim RET                 As FUNCTION_RETURN
    Dim udtWarpMapData()    As type_DBDRV_Nukisi
    Dim udtTmp_Y018()       As typ_WarpKakuData     '�W�������ް�(TBCMY018)�擾�p
    
    fnc_LoadData_Warp = FUNCTION_RETURN_FAILURE
    
    ReDim tSXLID(0)
    tSXLID(0).SXLID = txtSxlId.text
    '�֘A��ۯ�ID�擾
    If DBDRV_BLOCKIDGET() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("ESXL2")
        Exit Function
    End If
    
    'WFϯ���ް��擾
    If DBDRV_WARPMAPGET(udtWarpMapData()) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EGET2", "Y011")
        Exit Function
    End If
    
    'Warp/�����p�x�ް��擾
    ReDim tWarpInitG(0)
    ReDim tKakuInitG(0)
    'Add Start 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    ReDim tKakuXInitG(0)
    ReDim tKakuYInitG(0)
    'Add End 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    bMapWarpFlg = False
    
    For i = 0 To UBound(tSXLID)
        '�����p�x�ް��擾
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "ORIENT", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '�����p�x�ް����
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuInitG(m + n)
            
            For j = 1 To n
                tKakuInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        
        'Add Start 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
        '��(X)�p�x�ް��擾
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "XKAKU", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '��(X)�p�x�ް����
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuXInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuXInitG(m + n)

            For j = 1 To n
                tKakuXInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        
        '�c(Y)�p�x�ް��擾
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "YKAKU", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        '�c(Y)�p�x�ް����
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tKakuYInitG)
            n = UBound(udtTmp_Y018)
            ReDim Preserve tKakuYInitG(m + n)

            For j = 1 To n
                tKakuYInitG(m + j) = udtTmp_Y018(j)
            Next j
        End If
        'Add End 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�

        'Warp�ް��擾
        ReDim udtTmp_Y018(0)
        RET = funGet_TBCMY018(tSXLID(i).LOTID, "WARP", udtTmp_Y018())
        If RET = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr("EGET2", "Y018")
            Exit Function
        End If
        'Warp�ް����
        If UBound(udtTmp_Y018) > 0 Then
            m = UBound(tWarpInitG)
            n = UBound(udtTmp_Y018)
            k = 0
            Call fnc_MapWarpChk(udtTmp_Y018())
            For j = 1 To n
                'WFϯ�߂ɕR�t���Ȃ��ް��;�Ă��Ȃ�
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
    
    'WFϯ�ߏ�̕i�ԏ��擾
    ReDim tMapHin(0)
    m = 0
    For i = 1 To UBound(udtWarpMapData) Step 2
        If udtWarpMapData(i).hinban <> vbNullString And _
           Trim(udtWarpMapData(i).hinban) <> "Z" And _
           Trim(udtWarpMapData(i).hinban) <> "G" Then
           
            m = m + 1
            ReDim Preserve tMapHin(m)
            '�i��
            tMapHin(m).HIN.hinban = udtWarpMapData(i).hinban
            tMapHin(m).HIN.mnorevno = udtWarpMapData(i).REVNUM
            tMapHin(m).HIN.factory = udtWarpMapData(i).factory
            tMapHin(m).HIN.opecond = udtWarpMapData(i).opecond
            '��ۯ�ID
            tMapHin(m).BLOCKID = udtWarpMapData(i).LOTID
            '��ۯ����A��(Start)
            tMapHin(m).BLKSEQ_S = CInt(udtWarpMapData(i).BLOCKSEQ)
            '��ۯ����A��(End)
            tMapHin(m).BLKSEQ_E = CInt(udtWarpMapData(i + 1).BLOCKSEQ)
            '�U�������׸�
            tMapHin(m).WARPFLG = False
            tMapHin(m).KAKUFLG = False
        End If
    Next i
    
    fnc_LoadData_Warp = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    �֐���        : fnc_MapWarpChk
'*
'*    �����T�v      : 1.WFϯ�߂�Warp���т̕R�t������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^ �@�@�@�@�@      ,����
'*                    udtChkWarp()  ,I  ,typ_WarpKakuData   ,�W�������ް�(Warp����)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Sub fnc_MapWarpChk(udtChkWarp() As typ_WarpKakuData)

    Dim i, j, k, m, n   As Integer
    
    k = 1
    m = UBound(sWrpLOTID)
    n = UBound(udtChkWarp)
    
    For i = 1 To n
        udtChkWarp(i).EXISTFLG = 0        'WFϯ�ߔ͈͊O�̎��т�����ΏۂƂ���B
        For j = k To m
            'WFϯ�߂�Warp���т���ۯ�ID�^��ۯ����A�Ԃ���v����ΕR�t���L��
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
'*    �֐���        : fnc_CheckHWS
'*
'*    �����T�v      : 1.�������@���`�F�b�N���Č����̗L����Ԃ�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^ �@�@�@,����
'*                    sHWS  �@�@�@,I  ,String �@,�������@
'*
'*    �߂�l        : Boolean �����̗L��
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
'*    �֐���        : spdKensaTail_TopLeftChange
'*
'*    �����T�v      : 1.���ѕ\���ꗗTOP/BOT�ԂŁA���X�N���[����A��������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^ �@�@�@,����
'*                    oldLeft  �@ ,I  ,String �@,�X�N���[���O�̍ō���̗�ԍ�
'*                    oldTop   �@ ,I  ,String �@,�X�N���[���O�̍ŏ�s�̍s�ԍ�
'*                    NewLeft  �@ ,I  ,String �@,�X�N���[����̍ō���̗�ԍ�
'*                    NewTop   �@ ,I  ,String �@,�X�N���[����̍ŏ�s�̍s�ԍ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub spdKensaTail_TopLeftChange(ByVal oldLeft As Long, ByVal oldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdKensaTop.LeftCol = spdKensaTail.LeftCol
End Sub

'*******************************************************************************
'*    �֐���        : spdKensaTop_TopLeftChange
'*
'*    �����T�v      : 1.���ѕ\���ꗗTOP/BOT�ԂŁA���X�N���[����A��������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^ �@�@�@,����
'*                    oldLeft  �@ ,I  ,String �@,�X�N���[���O�̍ō���̗�ԍ�
'*                    oldTop   �@ ,I  ,String �@,�X�N���[���O�̍ŏ�s�̍s�ԍ�
'*                    NewLeft  �@ ,I  ,String �@,�X�N���[����̍ō���̗�ԍ�
'*                    NewTop   �@ ,I  ,String �@,�X�N���[����̍ŏ�s�̍s�ԍ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub spdKensaTop_TopLeftChange(ByVal oldLeft As Long, ByVal oldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    spdKensaTail.LeftCol = spdKensaTop.LeftCol
End Sub

'*******************************************************************************
'*    �֐���        : sub_PutSeihinTop
'*
'*    �����T�v      : 1.���i�V�[�g�\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutSeihinTop()
    Dim i As Integer, j As Integer      ' ٰ�� ����

    With f_cmbc039_2
    '���ύX �M�������f�����ǉ�
    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
        'AN���x�ǉ�
        For i = 1 To 5
            .spdHinbanTop.col = i
            .spdHinbanTop.row = 1
            Select Case i
            Case 1
                '�i��
                '�i��12���\��-------Start SystemBrain 2003/10/05
                .spdHinbanTop.Value = typ_CType.typ_Param.hinban & Format(typ_CType.typ_Param.REVNUM, "00") & typ_CType.typ_Param.factory & typ_CType.typ_Param.opecond
                '.spdHinbanTop.Value = typ_CType.typ_Param.HINBCA
            Case 2
                '�^�C�v
                .spdHinbanTop.Value = typ_CType.typ_si.HWFTYPE
            Case 3
                '����
                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDIR
            Case 4
                '�����h�[�v
                .spdHinbanTop.Value = typ_CType.typ_si.HWFCDOP
                
                '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
                '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                'AN���x�ǉ�
            Case 5
                'AN���x
                .spdHinbanTop.Value = typ_CType.typ_si.HWFANTNP
            End Select
        Next i
    End With
End Sub

'*******************************************************************************
'*    �֐���        : sub_PutSeihinCenter
'*
'*    �����T�v      : 1.���i�V�[�g�\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutSeihinCenter()
    Dim i As Integer, j As Integer      ' ٰ�� ����

    'CENTER��
    With f_cmbc039_2
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
''        For i = 1 To 9
        For i = 1 To 10
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
            .spdHinbanCen.col = i
            .spdHinbanCen.row = 1

            Select Case i
            Case 1
                '���R
                .spdHinbanCen.Value = toRsStr_nl(typ_CType.typ_si.HWFRMIN, typ_CType.typ_si.HWFRMAX)
                .spdHinbanCen.backColor = IIf(typ_CType.typ_si.MSMPFLGWFR = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata

            Case 2
'--------------- 2008/08/25 UPDATE START  By Systech ---------------
                'DK���x
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
                '�ׂ��搔�ύX
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
                'Change 2010/01/17 SIRD�Ή��@Y.Hitomi
                'SIRD
                .spdHinbanCen.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSIRDMX, "##0")
            End Select
        Next i
    End With
End Sub

'*******************************************************************************
'*    �֐���        : sub_PutSeihinTail
'*
'*    �����T�v      : 1.���i�V�[�g�\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutSeihinTail()
    Dim i As Integer, j As Integer      ' ٰ�� ����

    'TAIL��
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
                ''�c���_�f�d�l�\���ǉ�
                Case 7
                    'AO
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFZOMIN, "0.0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFZOMAX, "0.0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFAOI = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                ''GD�d�l�\���ǉ�
                'GDײݐ��ǉ�
                Case 8
                    'GDײݐ�
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFGDLINE, "")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 9
                    'Den
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDENMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFDENMX, "0")
                    .spdHinbanTail.backColor = IIf(typ_CType.typ_si.MSMPFLGWFGD = "1", COLOR_SKY, COLOR_DISABLE)  'Add 2011/04/25 SMPK Miyata
                Case 10
                    'L/DL
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech Start
''                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFLDLMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLMX, "0")
                    .spdHinbanTail.Value = DBData2DispData_nl(typ_CType.typ_si.HWFLDLMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLMX, "0") & " , " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLRMN, "0") & " - " & _
                                          DBData2DispData_nl(typ_CType.typ_si.HWFLDLRMX, "0")
'' 2008/10/01 L/DL,OSF����ۼޯ��ǉ� UPD By Systech End
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
'*    �֐���        : sub_PutSeihinTail2
'*
'*    �����T�v      : 1.���i�V�[�g�\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutSeihinTail2()
    Dim i As Integer, j As Integer      ' ٰ�� ����

    'TAIL2��
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
''                'SP(�g�U��)
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
        
        'SPV�\���ύX�@08/03/13 ooba START ===============================================>
        For i = 1 To 13
            .spdHinbanTail2.col = i
            .spdHinbanTail2.row = 1
            Select Case i
            'SP(Fe)
            Case 1      '���
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVMX, "0.00")
            Case 2      'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUG, "0.00")
            Case 3      'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVPUR, "0.000")
            Case 4      '�W���΍�
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFSPVSTD, "0.000")
            'SP(�g�U��)
            Case 5      '����
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMIN, "0.0")
            Case 6      '���
            
                '�ۏؕ��@���ނ��uL�v(AVE+MIN)�ȊO�̏ꍇ
                If typ_CType.typ_si.HWFDLHWT <> "L" Then
                    .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0.0")
                End If
            Case 7      'AVE����(���)
            
                '�ۏؕ��@���ނ��uL�v(AVE+MIN)�̏ꍇ�͏����AVE�����Ƃ���
                If typ_CType.typ_si.HWFDLHWT = "L" Then
                    .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLMAX, "0.0")
                End If
            Case 8      'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLPUG, "0.00")
            Case 9      'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFDLPUR, "0.000")
            'SP(Nr)
            Case 10     '���
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRMX, "0.00")
            Case 11     'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRPUG, "0.00")
            Case 12     'PUA��
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRPUR, "0.000")
            Case 13     '�W���΍�
                .spdHinbanTail2.Value = DBData2DispData_nl(typ_CType.typ_si.HWFNRSTD, "0.000")
            End Select
        Next i
        'SPV�\���ύX�@08/03/13 ooba END =================================================>
    End With
End Sub

'*******************************************************************************
'*    �֐���        : sub_PutSeihinEpi
'*
'*    �����T�v      : 1.���i�V�[�g�\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutSeihinEpi()
    Dim i As Integer, j As Integer      ' ٰ�� ����

    '�G�s��
    With f_cmbc039_2
        'Chg Start 2011/04/28 SMPK Miyata�@(OSF3E�܂ŕ\�����Ă��Ȃ��̂ŏC��)
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
                'BMD3E(�O��)�@09/05/07 ooba
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
'*    �֐���        : sub_PutRs
'*
'*    �����T�v      : 1.���R�l�\��(TOP,TAIL)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutRs()
    '���R�l�\��(TOP��)
    sub_PutRsTop

    '���R�l�\��(TAIL��)
    sub_PutRsTail

'Add Start 2011/03/09 SMPK Miyata
    '���R�l�\��(MIDLE��)
    Call sub_PutRsMid(1)
'Add End   2011/03/09 SMPK Miyata

End Sub

'*******************************************************************************
'*    �֐���        : sub_PutRsTop
'*
'*    �����T�v      : 1.���R�l�\��(TOP)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutRsTop()
    Dim blJudg  As Boolean  '���茋��
    Dim dblScut As Double   '�ăJ�b�g�ʒu
    Dim dblCoef As Double   '���s�ΐ�

    dblScut = typ_CType.dblScut(SxlTop039)
    dblCoef = typ_CType.COEF(SxlTop039)

    With f_cmbc039_2
        '' WF�����w���iRs)*****************************************************************
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        'DK���x
        .txtDKTmp(SxlTop).text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(SxlTop)))
        If Not typ_CType.JudgDkTmp(SxlTop) Then
            .txtDKTmp(SxlTop).backColor = COLOR_NG
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        '�ۏؕ��@�����ǉ�
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "TOP") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then

                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '�ʒu

                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG

                '���ǉ� �M�������f�����ǉ�
                '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                    'AN���x�ǉ�
                    '���ځFDKAN��3�`6����AN���x
                    .txtANTempTop.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTop039, WFRES).DKAN, 3, 4), "0") 'AN���x
                    
                    '�`�F�b�NNG�̎��͔w�i�F��ς���
                    If Not (typ_CType.JudgAntnp(SxlTop039)) Then
                        CtrlEnabled .txtANTempTop, CTRL_DISABLE_WARNING, False  'AN���x
                    End If
                    
                    '�v�e�T���v�������ύX
                    If Not (typ_CType.JudgRrg(SxlTop039)) Then
                        CtrlEnabled .txtRRGTop, CTRL_DISABLE_WARNING, False  'RRG
                    End If
                    
                    If dblCoef = -1 Or dblCoef = -9999 Then

                        .txtJHAll.text = ""         '���s�ΐ̓u���b�N
                    Else
                        .txtJHAll.text = DBData2DispData(dblCoef, "0.000")         '���s�ΐ̓u���b�N
                    End If

                    '�ăJ�b�g�ʒu
                    '�v�e�T���v�������ύX
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
                        CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP�ăJ�b�g
                        intEnCmd = 1
                    End If
                        
                    '���R
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
                            .Value = "�d�l�L"
                            
                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "�����L"
                            
                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "���і�"
                        End With
                    End If
                End If
            Else
                .txtSXLTop.text = ""            '�ʒu
                .txtRRGTop.text = ""            'RRG
                .txtJHAll.text = ""
                
                '�ăJ�b�g�ʒu
                '�v�e�T���v�������ύX
                .txtCutPosTop.text = "NG"
                CtrlEnabled .txtCutPosTop, CTRL_DISABLE_WARNING, False  'TOP�ăJ�b�g

                '���R
                With .spdMeasTop
                    .col = 1
                    .row = 1:
                    .CellType = CellTypeStaticText
                    .Value = "�d�l�L"
                    
                    .row = 2:
                    .CellType = CellTypeStaticText
                    .Value = "������"
                    .row = 3: .Value = ""
                    .row = 4: .Value = ""
                    .row = 5: .Value = ""
                End With
            End If
        Else
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTop039).WFINDRSCW) <> 0 Then
                If typ_CType.typ_Param.WFSMP(SxlTop039).WFRESRS1CW = "1" Then
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '�ʒu
                    
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG
                    If dblCoef = -1 Or dblCoef = -9999 Then
                        .txtJHAll.text = ""         '���s�ΐ̓u���b�N
                    Else
                        .txtJHAll.text = DBData2DispData(dblCoef, "0.000")         '���s�ΐ̓u���b�N
                    End If

                    '�ăJ�b�g�ʒu
                    .txtCutPosTop.text = "OK"

                    '���R
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
                                    .Value = "�d�l��"
                            .row = 2:
                                    .CellType = CellTypeStaticText
                                    .Value = "�����L"
                            .row = 3:
                                    .CellType = CellTypeStaticText
                                    .Value = "���і�"
                        End With
                    End If
                Else
                    .txtSXLTop.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS, "0")            '�ʒu
                    
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTop.text = DBData2DispData(typ_CType.typ_y013(SxlTop039, WFRES).MESDATA6, "0.000000")  'RRG
                    .txtJHAll.text = ""
                    
                    '�ăJ�b�g�ʒu
                    .txtCutPosTop.text = "OK"
                    
                    '���R
                    With .spdMeasTop
                        .col = 1
                        .row = 1:
                        .CellType = CellTypeStaticText
                        .Value = "�d�l��"
                        
                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "�����L"
                        
                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "���і�"
                        
                        .row = 4: .Value = ""
                        .row = 5: .Value = ""
                    End With
                End If
            End If
        End If
    End With
End Sub

'*******************************************************************************
'*    �֐���        : sub_PutRsTail
'*
'*    �����T�v      : 1.���R�l�\��(TAIL)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutRsTail()
    Dim blJudg  As Boolean  '���茋��
    Dim dblScut As Double   '�ăJ�b�g�ʒu

    dblScut = typ_CType.dblScut(SxlTail039)

    With f_cmbc039_2
        '' WF�����w���iRs)*****************************************************************
'--------------- 2008/08/25 INSERT START  By Systech ---------------
        'DK���x
        .txtDKTmp(SxlTail).text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(SxlTail)))
        If Not typ_CType.JudgDkTmp(SxlTail) Then
            .txtDKTmp(SxlTail).backColor = COLOR_NG
        End If
'--------------- 2008/08/25 INSERT  END   By Systech ---------------
        
        '�ۏؕ��@�����ǉ�
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "BOT") Then
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then

                If typ_CType.typ_Param.WFSMP(SxlTail039).WFRESRS1CW = "1" Then
                    .txtSXLTail.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH, "0")           '�ʒu
                    
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.00")  'RRG
                    .txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.000000")  'RRG
                    
                    '���ǉ� �M�������f�����ǉ�
                    '2.1.3 AN���x ���є��f�`�F�b�N�ǉ�
                    'AN���x�ǉ�
                    '���ځFDKAN��3�`6����AN���x
                    .txtANTempTail.text = DBData2DispData(Mid(typ_CType.typ_y013(SxlTail039, WFRES).DKAN, 3, 4), "0") 'AN���x
                    '�`�F�b�NNG�̎��͔w�i�F��ς���
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled .txtANTempTail, CTRL_DISABLE_WARNING, False  'AN���x
                    End If

                    '�v�e�T���v�������ύX
                    If Not (typ_CType.JudgRrg(SxlTail039)) Then
                        CtrlEnabled .txtRRGTail, CTRL_DISABLE_WARNING, False  'RRG
                    End If


                    '�v�e�T���v�������ύX
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
                        
                        CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'tail�ăJ�b�g
                        intEnCmd = 1
                    End If

                    '���R
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
                            .Value = "�d�l�L"
                            
                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "�����L"
                            
                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "���і�"
                        End With
                    End If
                End If
            Else
                .txtSXLTail.text = ""            '�ʒu
                .txtRRGTail.text = ""            'RRG
                .txtJHAll.text = ""
                '�ăJ�b�g�ʒu

                '�v�e�T���v�������ύX
                .txtCutPosTail.text = "NG"
                CtrlEnabled .txtCutPosTail, CTRL_DISABLE_WARNING, False  'Tail�ăJ�b�g

                '���R
                With .spdMeasTail
                    .col = 1
                    .row = 1:
                            .CellType = CellTypeStaticText
                            .Value = "�d�l�L"
                    .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "������"
                    .row = 3: .Value = ""
                    .row = 4: .Value = ""
                    .row = 5: .Value = ""
                End With
            End If
        Else
            If InStr("123", typ_CType.typ_Param.WFSMP(SxlTail039).WFINDRSCW) <> 0 Then
                .txtSXLTail.text = DBData2DispData(typ_CType.typ_Param.INGOTPOS + typ_CType.typ_Param.LENGTH, "0")            '�ʒu
                
                'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                '.txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.00")  'RRG
                .txtRRGTail.text = DBData2DispData(typ_CType.typ_y013(SxlTail039, WFRES).MESDATA6, "0.000000")  'RRG

                '�ăJ�b�g�ʒu
                .txtCutPosTail.text = "OK"

                '���R
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
                        .Value = "�d�l��"
                        
                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "�����L"
                        
                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "���і�"
                    End With
                End If
            End If
        End If
    End With
End Sub

'Add Start 2011/03/09 SMPK Miyata
'*******************************************************************************
'*    �֐���        : sub_PutRsMid
'*
'*    �����T�v      : 1.���R�l�\��(���Ԕ���)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                  : iMidNo        ,I  ,Integer  ,���Ԕ���No(1-10)
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutRsMid(iMidNo As Integer)
    Dim blJudg  As Boolean  '���茋��
    Dim dblScut As Double   '�ăJ�b�g�ʒu
    Dim tt      As Integer  'Top Tail Midle����p

    tt = SxlMidl + iMidNo - 1

    If tt < 1 Or tt > UBound(typ_CType.typ_Param.WFSMP) Then
        Exit Sub
    End If
    
    dblScut = typ_CType.dblScut(tt)
    
    With f_cmbc039_2
        
        '�ۏؕ��@����(�i�v�e���R�����p�x�Q��) �L��̏ꍇ
        If JudgSW.rs And CheckKHN(typ_CType.typ_si.HWFRKHNN, 1, "MID") Then
            '���FLG(Rs)�� 1�F�ʏ�A2�F���f�A3�F����̏ꍇ
            If InStr("123", typ_CType.typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
                '����FLG1(Rs)��1�F���т���̏ꍇ
                If typ_CType.typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then
                    
                    'Add Start 2011/08/25 Y.Hitomi
                    'DK���x
                    txtDKTmpMid.text = GetDKTmpDispName("" & GetGPCodeCont(DKTMP_TBCME033CODE, typ_CType.DkTmpJsk(tt)))
                    If Not typ_CType.JudgDkTmp(tt) Then
                        txtDKTmpMid.backColor = COLOR_NG
                    End If
                    txtSXLMid.backColor = COLOR_SKY         'SXL���Ԉʒu
                    txtCutPosMid.backColor = COLOR_SKY      '�ŃJ�b�g�ʒu�i���ԁj
                    txtRRGMid.backColor = COLOR_SKY         'RRG�i���ԁj
                    txtANTempMid.backColor = COLOR_SKY      'AN���x�i���ԁj
                    txtDKTmpMid.backColor = COLOR_SKY       'DK���x�i���ԁj
                    SpCtrlBlockEnabled spdMeasMid, 1, 1, spdMeasMid.MaxCols, spdMeasMid.MaxRows, CTRL_DISABLE_SKY, True
                    'Add End   2011/08/25 Y.Hitomi

                    'Mid �ʒu
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")

                    'Add Start 2011/08/11 Y.Hitomi
                    'AN���x�ǉ� �FDKAN��3�`6����AN���x
                    .txtANTempMid.text = DBData2DispData(Mid(typ_CType.typ_y013(tt, WFRES).DKAN, 3, 4), "0") 'AN���x
                    '�`�F�b�NNG�̎��͔w�i�F��ς���
                    If Not (typ_CType.JudgAntnp(SxlTail039)) Then
                        CtrlEnabled .txtANTempMid, CTRL_DISABLE_WARNING, False  'AN���x
                    End If
                    'Add End   2011/08/11 Y.Hitomi
                    
                    '�v�e�T���v�������ύX
                    If Not (typ_CType.JudgRrg(tt)) Then
                        'RRG�F�ύX
                        CtrlEnabled .txtRRGMid, CTRL_DISABLE_WARNING, False
                    End If

                    '���Ԕ��� �ăJ�b�g�ʒu
                    '�v�e�T���v�������ύX
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

                    '���R
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
                            .Value = "�d�l�L"

                            .row = 2:
                            .CellType = CellTypeStaticText
                            .Value = "�����L"

                            .row = 3:
                            .CellType = CellTypeStaticText
                            .Value = "���і�"
                        End With
                    End If
                End If
            Else
                '���FLG(Rs)�������Ȃ�(1�F�ʏ�A2�F���f�A3�F����ȊO)�̏ꍇ

                .txtSXLMid.text = ""            '�ʒu
                .txtRRGMid.text = ""            'RRG

                '���Ԕ��� �ăJ�b�g�ʒu
                '�v�e�T���v�������ύX
'                .txtCutPosMid.text = "NG"
'                CtrlEnabled .txtCutPosMid, CTRL_DISABLE_WARNING, False
'
            End If
        Else
            '�ۏؕ��@����(�i�v�e���R�����p�x�Q��) �Ȃ��̏ꍇ

            '���FLG(Rs)�� 1�F�ʏ�A2�F���f�A3�F����̏ꍇ
            If InStr("123", typ_CType.typ_Param.WFSMP(tt).WFINDRSCW) <> 0 Then
                '����FLG1(Rs)��1�F���т���̏ꍇ
                If typ_CType.typ_Param.WFSMP(tt).WFRESRS1CW = "1" Then

                    'Mid �ʒu
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")
                    
                    '�ăJ�b�g�ʒu
                    .txtCutPosMid.text = "OK"

                    '���R
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
                                    .Value = "�d�l��"
                            .row = 2:
                                    .CellType = CellTypeStaticText
                                    .Value = "�����L"
                            .row = 3:
                                    .CellType = CellTypeStaticText
                                    .Value = "���і�"
                        End With
                    End If
                Else
                    'Mid �ʒu
                    .txtSXLMid.text = DBData2DispData(typ_CType.typ_Param.WFSMP(tt).INPOSCW, "0")
                    'RRG
                    'RRG�̏���������6��(7���ڎl�̌ܓ�)�ɕύX 2011/11/25 SETsw kubota
                    '.txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.00")
                    .txtRRGMid.text = DBData2DispData(typ_CType.typ_y013(tt, WFRES).MESDATA6, "0.000000")

                    '�ăJ�b�g�ʒu
                    .txtCutPosMid.text = "OK"

                    '���R
                    With .spdMeasMid
                        .col = 1
                        .row = 1:
                        .CellType = CellTypeStaticText
                        .Value = "�d�l��"

                        .row = 2:
                        .CellType = CellTypeStaticText
                        .Value = "�����L"

                        .row = 3:
                        .CellType = CellTypeStaticText
                        .Value = "���і�"

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
'*    �֐���        : sub_PutRslt
'*
'*    �����T�v      : 1.���ђl�\��(TOP)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^           ,����
'*                    udt_rslt()    ,I  ,typ_ALLRSLT  ,���я��\����
'*                    tt            ,I  ,Integer      ,TopTail����p
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_PutRslt(udt_rslt() As typ_ALLRSLT, tt As Integer)
    Dim i, j        As Integer
    Dim spdVa       As vaSpread
    Dim lngSpMaxLine   As Long

'    '�ő�s���擾
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
                        '�ʒu
                        spdVa.Value = DBData2DispData(CVar(.pos), "0")
                    Case 2
                        '���e
                        If left(.NAIYO, 3) = "BMD" Then
                            spdVa.Value = .NAIYO & "(�~E4)"
                        Else
                            spdVa.Value = .NAIYO
                        End If
                    Case 3
                        '���P
                        spdVa.Value = .INFO1
                    Case 4
                        '���Q
                        spdVa.Value = .INFO2
                    Case 5
                        '���R
                        spdVa.Value = .INFO3
                    Case 6
                        '���S
                        spdVa.Value = .INFO4
                    Case 7
                        '���T
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO5
                    Case 8
                        '���V
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO6
                    Case 9
                        '���W
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO7
                    Case 10
                        '���W
                        spdVa.Value = typ_rslt_ex(tt, i - 1).INFO8
                    Case 11
                        '����
                        '�v�e�T���v�������ύX
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
                        '�ʒu
                        spdVa.Value = CStr(DBData2DispData(.SMPLID, "0"))
                End Select
            Next j
        End With
        i = i + 1
    Loop

    '�\�[�g����
    If i <> 1 Then
        With spdVa
            .MaxRows = i - 1                      '�@�i�ԁi�s���j
            .row = 1                            ' �Z���u���b�N��ݒ�
            .col = 1
            .row2 = i - 1
            .col2 = 12
            .SortBy = SS_SORT_BY_ROW

            .SortKey(1) = 11                    ' ��P�\�[�g�L�[��ݒ�

            ' �����ɕ��בւ�
            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            .Action = SS_ACTION_SORT
        End With
    End If
End Sub

'*******************************************************************************************
'*    �֐���        : sub_PutRslt_EP
'*
'*    �����T�v      : 1.WF�d�l�̃G�s�d�l�̕\���ؑ�
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@�@udt_rslt �@�@ ,I  ,typ_ALLRSLT     ,���я��\����
'*�@�@�@�@�@�@�@�@�@�@tt       �@�@ ,I  ,Integer         ,TopTail����p
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_PutRslt_EP(udt_rslt() As typ_ALLRSLT_EX, tt As Integer)
    Dim i, j            As Integer
    Dim spdVa           As vaSpread
    Dim lngSpMaxLine    As Long

    ''�ő�s���擾
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
                    '�ʒu
                    spdVa.Value = DBData2DispData(CVar(.pos), "0")
                Case 2
                    '���e
                    spdVa.Value = .NAIYO
                Case 3
                    '���P
                    spdVa.Value = .INFO1
                Case 4
                    '���Q
                    spdVa.Value = .INFO2
                Case 5
                    '���R
                    spdVa.Value = .INFO3
                Case 6
                    '���S
                    spdVa.Value = .INFO4
                Case 7
                    '���T
                    spdVa.Value = .INFO5
                Case 8
                    '���V
                    spdVa.Value = .INFO6
                Case 9
                    '���W
                    spdVa.Value = .INFO7
                Case 10
                    '���W
                    spdVa.Value = .INFO8
                Case 11
                        If .OKNG = "NG" Then
                            SpCtrlEnabled spdVa, spdVa.col, spdVa.row, CTRL_DISABLE_WARNING
                            intEnCmd = 1
                        End If
                        spdVa.Value = .OKNG
                Case 12
                    '�ʒu
                    spdVa.Value = CStr(DBData2DispData(.SMPLID, "0"))
                End Select
            Next j
        End With
        i = i + 1
    Loop

    '�\�[�g����
    If i <> 1 Then
        With spdVa
            .MaxRows = i - 1                    '�@�i�ԁi�s���j
            .row = 1                            ' �Z���u���b�N��ݒ�
            .col = 1
            .row2 = i - 1
            .col2 = 12
            .SortBy = SS_SORT_BY_ROW
            .SortKey(1) = 11                    ' ��P�\�[�g�L�[��ݒ�
            
            ' �����ɕ��בւ�
            .SortKeyOrder(1) = SS_SORT_ORDER_ASCENDING
            .Action = SS_ACTION_SORT
        End With
    End If
End Sub

''*******************************************************************************
''*    �֐���        : RegWfSogoRsltOK
''*
''*    �����T�v      : 1.����������ё}��
''*                    2.WF_GD����(TBCMJ015)�X�V����
''*                    3.SXL�Ǘ��X�V
''*                    4.WF�T���v���Ǘ��X�V
''*
''*    �p�����[�^    : �ϐ���        ,IO ,�^           ,����
''*
''*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
''*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
''*
''*******************************************************************************
'Private Function RegWfSogoRsltOK() As FUNCTION_RETURN
'    Dim udt_soz         As typ_TBCMW005                             ' WF�����������
'    Dim udt_sxl         As type_DBDRV_scmzc_fcmlc001c_UpdSXL1       ' SXL�Ǘ�
'    Dim udt_WFSmp(2)    As type_DBDRV_scmzc_fcmlc001c_UpdWfCrySmp
'    Dim i               As Long
'    Dim intCnt          As Integer
'
'    'WF�����������
'    With udt_soz
'        .CRYNUM = typ_CType.typ_Param.CRYNUM                                ' �����ԍ�
'        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' �C���S�b�g�ʒu
'        .CRYLEN = typ_CType.typ_Param.LENGTH                                ' ����
'        .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI                                 ' �Ǘ��H���R�[�h
'        .PROCCODE = PROCD_WFC_SOUGOUHANTEI                                  ' �H���R�[�h
'        .SXLID = NtoS(typ_CType.typ_Param.SXLID)                                  ' SXLID
'        .CODE = "0"                                                         ' �敪�R�[�h
'        .TSTAFFID = typ_CType.strStaffID                                    ' �o�^�Ј�ID
'    End With
'
'    'WF����������ё}��
'    If DBDRV_scmzc_fcmlc001c_InsWfSougou(udt_soz) <> FUNCTION_RETURN_SUCCESS Then
'        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "W005")
'        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    '' WF_GD����(TBCMJ015)�X�V����
'    If UBound(typ_J015_WFGDUpd) > 0 Then
'        '�ް�����UPDATE
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
'    'SXL�Ǘ�
'    With udt_sxl
'        .CRYNUM = NtoS(typ_AType.typ_Param.CRYNUMCA)                        ' �����ԍ�
'        .INGOTPOS = typ_CType.typ_Param.INGOTPOS                            ' �������J�n�ʒu
'        .NOWPROC = PROCD_SXL_KAKUTEI                                        ' ���ݍH��
'        .LASTPASS = PROCD_WFC_SOUGOUHANTEI                                  ' �ŏI�ʉߍH��
'    End With
'
'    'SXL�Ǘ��X�V
'    If DBDRV_scmzc_fcmlc001c_UpdSXL1(udt_sxl) <> FUNCTION_RETURN_SUCCESS Then
'        f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E042")
'        RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'        Exit Function
'    End If
'
'    'WF�T���v���Ǘ������݂���ꍇ�͊m��敪�R�[�h��1�𗧂Ă�
'    '�G�s��s�]���ǉ��Ή�
'    If (UBound(typ_CType.typ_y013top) <> 0 Or UBound(typ_CType_EP.typ_y022top) <> 0) _
'        And (UBound(typ_CType.typ_y013tail) <> 0 Or UBound(typ_CType_EP.typ_y022tail) <> 0) Then
'
'        'WF�T���v���Ǘ�
'        udt_WFSmp(1).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' �����ԍ�
'        udt_WFSmp(1).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTop039).INPOSCW    ' �������J�n�ʒu
'        udt_WFSmp(1).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTop039).SMPKBNCW     ' �T���v���敪
'        udt_WFSmp(2).CRYNUM = NtoS(typ_CType.typ_Param.CRYNUM)                  ' �����ԍ�
'        udt_WFSmp(2).INGOTPOS = typ_CType.typ_Param.WFSMP(SxlTail039).INPOSCW   ' �������J�n�ʒu
'        udt_WFSmp(2).SMPKBN = typ_CType.typ_Param.WFSMP(SxlTail039).SMPKBNCW    ' �T���v���敪
'
'        'WF�T���v���Ǘ��X�V
'        If DBDRV_scmzc_fcmlc001c_UpdWfCrySmp(udt_WFSmp) <> FUNCTION_RETURN_SUCCESS Then
'            f_cmbc039_2.lblMsg.Caption = GetMsgStr("EGET2", "E044")
'            RegWfSogoRsltOK = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'    End If
'End Function

'*******************************************************************************************
'*    �֐���        : sub_cmbc061_2_ChangeHinSpec
'*
'*    �����T�v      : 1.WF�d�l�̃G�s�d�l�̕\���ؑ�
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@�@intCategory�@ ,I  ,Integer         ,�\���J�e�S��(0:WF�d�l,1:�G�s�d�l)
'*
'*    �߂�l        : �Ȃ�
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
            Case 0          ' WF�d�l�f�[�^�̕\��
                'WF�f�[�^�̃X�v���b�h��\��
                .spdHinbanCen.Visible = True
                .spdHinbanTail.Visible = True
                .spdHinbanTail2.Visible = True
                .spdHinbanHed.Visible = True        '08/03/12 ooba
                .spdHinbanCenEpi.Visible = False
    
                'AN���x
                f_cmbc039_2.spdHinbanTop.Value = typ_CType.typ_si.HWFANTNP
            Case 1          ' �G�s�d�l�f�[�^�̕\��
                '�G�s�f�[�^�̃X�v���b�h��\��
                .spdHinbanCen.Visible = False
                .spdHinbanTail.Visible = False
                .spdHinbanTail2.Visible = False
                .spdHinbanHed.Visible = False       '08/03/12 ooba
                .spdHinbanCenEpi.Visible = True
    
                'AN���x(���i�d�l�G�s�f�[�^1)
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
'*    �֐���        : sub_SampleMidlePosBtnSet
'*
'*    �����T�v      : 1.���Ԕ����T���v���ʒu�{�^���ݒ�
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@:�@�@           ,   ,                ,
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_SampleMidlePosBtnSet()
    Dim i       As Long
    Dim k       As Long
    Dim blnOK   As Boolean

    '�{�^�������[�v
    For i = optPosSelMid.LBound To optPosSelMid.UBound

        '���Ԕ��������邩�T���v���Ǘ��L���Ŕ��f
        If SxlMidl + i <= UBound(typ_CType.typ_Param.WFSMP) Then
            '���Ԕ������L��̏ꍇ
            
            '�{�^���ݒ�
            With optPosSelMid(i)
                .Enabled = True
                .Caption = typ_CType.typ_Param.WFSMP(SxlMidl + i).INPOSCW
                
                
                '����NG���ڂ����邩����
                blnOK = True
                ' ���R����
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
                
                '�\���ʒu�̐F�ݒ�@����OK�F���@����NG�F��
                If blnOK = True Then
                    .ForeColor = vbBlack
                Else
                    .ForeColor = vbRed
                End If
                If i = optPosSelMid.LBound Then .Value = True
                
            End With
        Else
            '���Ԕ����������̏ꍇ
            
            '�{�^���ݒ�
            With optPosSelMid(i)
                .Enabled = False
                .Caption = ""
                .ForeColor = vbBlack
            End With
        
        End If
    Next

End Sub

'>>>>> add 2011/07/13 Marushita
'��ʃL���v�`������
Public Function saveCapture_BMP(ByRef frm As Form, ByRef picData As PictureBox) As Boolean
    
    Dim lRetVal As Long
    Dim lDC As Long
    
On Error GoTo Err:
    
    '��O�ɕ\��
    Call SetWindowPos(frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE)
    DoEvents
    
    ' �E�B���h�E���A�N�e�B�u�ɂ���
    Call SetForceForegroundWindow(frm.hwnd)
    DoEvents
    
    '�n���h������f�o�C�X�R���e�L�X�g���擾
    lDC = GetDC(frm.hwnd)
    
    picData.AutoRedraw = True
    picData.Width = frm.ScaleWidth + 10
    picData.Height = frm.ScaleHeight + 29
    
    'lRetVal = BitBlt(picData.hdc, 0, 0, picData.Width, picData.Height, lDC, -3, -22, SRCCOPY)
    lRetVal = StretchBlt(picData.hdc, 0, 0, picData.Width * CLng(SCALEPER) / 100, picData.Height * CLng(SCALEPER) / 100, lDC, 0, 0, picData.Width, picData.Height, SRCCOPY)
    
    DoEvents
    '�N���b�v�{�[�h���Ƀr�b�g�}�b�v�`���̃f�[�^�����邩���ׂ�
    If lRetVal <> 0 Then
        '�t�@�C�����𐶐�
        SavePicture pic_Png.Image, App.Path & CAP_FNAME
    Else
        '���s
        Call MsgBox("���s")
    End If
    
    'DC�J��
    Call ReleaseDC(frm.hwnd, lDC)
    
    '��O�ɕ\��������
    SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    
    saveCapture_BMP = True
    Exit Function
    
Err:
    Call MsgOut(0, "��ʃL���v�`���ۑ��Ɏ��s���܂���" & vbCrLf _
                 & Err.Number & ":" & Err.Description, ERR_DISP)
    
    '��O�ɕ\��������
    SetWindowPos frm.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOSIZE + SWP_NOMOVE
    
    saveCapture_BMP = False

End Function
'<<<<< add 2011/07/13 Marushita
