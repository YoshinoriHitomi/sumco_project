VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "f_cmbc039_3(CW750) - 300mm�������ƃV�X�e��"
   ClientHeight    =   10875
   ClientLeft      =   0
   ClientTop       =   750
   ClientWidth     =   15270
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "�l�r �S�V�b�N"
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
   ScaleMode       =   3  '�߸��
   ScaleWidth      =   1018
   StartUpPosition =   2  '��ʂ̒���
   Begin VB.CommandButton cmdDisp 
      Caption         =   "WFC��������Q�ƕ\��"
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
      Caption         =   "PNG�ۑ�"
      Height          =   375
      Left            =   690
      TabIndex        =   58
      Top             =   8535
      Width           =   1335
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
         Name            =   "�l�r �S�V�b�N"
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
      BorderStyle     =   0  '�Ȃ�
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
      BorderStyle     =   0  '�Ȃ�
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
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   8175
      Width           =   1095
   End
   Begin VB.TextBox txtTopRsltR 
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   1650
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   7935
      Width           =   1095
   End
   Begin VB.TextBox txtCryP 
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   8745
      Width           =   855
   End
   Begin VB.TextBox txtBlkP 
      Alignment       =   1  '�E����
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   8475
      Width           =   855
   End
   Begin VB.TextBox txtBlkID 
      Alignment       =   2  '��������
      BackColor       =   &H0080FF80&
      Height          =   270
      Left            =   8310
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   8205
      Width           =   855
   End
   Begin VB.TextBox txtTarget 
      Alignment       =   1  '�E����
      Height          =   270
      IMEMode         =   3  '�̌Œ�
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
         Caption         =   "[F11]�@�@�O���"
         Height          =   735
         Index           =   11
         Left            =   12680
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]�@�@�����"
         Height          =   735
         Index           =   10
         Left            =   11448
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�W]�@�@�폜"
         Height          =   735
         Index           =   8
         Left            =   8984
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�V]�@�@�}��"
         Height          =   735
         Index           =   7
         Left            =   7752
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�U]�@�@WFϯ��"
         Height          =   735
         Index           =   6
         Left            =   6520
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�T]�@�@�U��"
         Enabled         =   0   'False
         Height          =   735
         Index           =   5
         Left            =   5288
         Style           =   1  '���̨���
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�S]�@�@���"
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�X]�@�@�����Ұ�ޕ\��"
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�P]�@�@������"
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
         Caption         =   "[F�Q]�@�@����ƭ�"
         Height          =   735
         Index           =   2
         Left            =   1592
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F�R]�@�@��ݾ�"
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]�@�@���s"
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
         TabIndex        =   4
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
         Left            =   1800
         TabIndex        =   3
         Top             =   210
         Width           =   8535
      End
      Begin VB.Label lblTitle 
         Caption         =   "�Ĕ����w��"
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
      Height          =   255
      Left            =   10755
      TabIndex        =   62
      Top             =   1650
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblNukishi 
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�֘A�u���b�N"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Alignment       =   2  '��������
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '����
      Caption         =   "4��"
      Height          =   255
      Left            =   6600
      TabIndex        =   56
      Top             =   1515
      Width           =   390
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFF00&
      Caption         =   "�]�����ʎ�M��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "����ٖ�(��������)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "����ٖ�(���f)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "����ٗL"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "����ٖ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�F1SXL"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�FSXL����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�����ޯ��"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "����P"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "��ۯ�P"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "��ۯ�ID"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "�˂炢��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
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
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "B�����у�"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   690
      TabIndex        =   14
      Top             =   8175
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  '��������
      Appearance      =   0  '�ׯ�
      BackColor       =   &H80000005&
      BackStyle       =   0  '����
      BorderStyle     =   1  '����
      Caption         =   "T�����у�"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   690
      TabIndex        =   12
      Top             =   7935
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "���r�w�k�|�h�c"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�����ԍ�"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
      Caption         =   "�S���҃R�[�h"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
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
' �Ĕ����w�����
' �T�v    :
'===============================================================================

Private orgXl As c_cmzcXl                         ' �ǂݍ��ݎ��_�̌������
Private tblHinNum() As tFullHinban                ' �i�ԃe�[�u��
Private InMaxRow As Integer                       ' �ő�s�ێ�
Private bSampFlag As Boolean                      ' �T���v���擾�t���O
Private CutCntFlg As Integer                      ' �T���v���̂Ȃ�SXL
Private orgSXL As c_cmzcSxls                      ' ������Ԃ�SXL�\��
Private giFKeyFlg As Integer                      '�T���v���E���s�{�^���t���O
Private bJituChkFlg As Boolean                    ' �����ް������׸�
Private iBetuRow()  As Integer                    ' ���L(��)����ٍs
Private CpyCrySmpl  As typ_CpyJisseki             ' �������ш��p���ް�

' �U�փ`�F�b�N�ǉ��ɂ��C��
Private Type tbl_MotoHin
    MOTOICHIS As Integer                          ' �U�֌��J�n�ʒu
    MOTOICHIE As Integer                          ' �U�֌��I���ʒu
    MOTOHIN As tFullHinban                        ' �U�֌��i��
End Type
Private MotoHinban() As tbl_MotoHin               ' �U�֌��i�ԃf�[�^

Private Type tbl_FuriNaiyou
    FURIUMU As Integer                            ' �U�֗L��(0:�����A1:�L��)
    ICHI    As Integer                            ' �ʒu
    MOTOHIN As tFullHinban                        ' �U�֌��i��
    SAKIHIN As tFullHinban                        ' �U�֐�i��
    TREPID As String                              ' ��\�T���v��ID(TOP)
    BREPID As String                              ' ��\�T���v��ID(BOT)
End Type
Private FurikaeNaiyou() As tbl_FuriNaiyou         ' �U�֓��e�ݒ�f�[�^
Private FurikaeNaiyouWK() As tbl_FuriNaiyou       ' �U�֓��e�ݒ�f�[�^(Work�p)

Private Type tbl_TokusaiBan
    ICHI     As Integer                           ' �ʒu
    MOTOHIN  As tFullHinban                       ' �U�֌��i��
    SAKIHIN  As tFullHinban                       ' �U�֐�i��
    BANGOU   As String                            ' ���̔ԍ�
    RIYUU    As String                            ' ���̗��R
    ERRRIYUU As String                            ' �G���[���R
    TREPID As String                              '��\�T���v��ID(TOP)
    BREPID As String                              '��\�T���v��ID(BOT)
End Type

Private TokusaiBangou() As tbl_TokusaiBan         ' ���̔ԍ��f�[�^
Private TokuCnt As Integer                        ' ���̔ԍ��f�[�^�J�E���^
Private TokuCntWK As Integer                      ' ���̔ԍ��f�[�^�J�E���^(Work�p)

Private FurikaeRireki() As typ_XSDCE_Update       ' �U�֗����f�[�^

Private bTokuKengenFlag As Boolean                ' ���̌����t���O
Private tblKns() As typ_XSDCW                     ' �������ڂ��Ƃ��Ă������߂̍\����
Private tKensa() As typ_XSDCW                     ' �������ڎ擾�p
Private tWafk() As typeSprWFmap                   ' �E�G�n�[�Z���^�[���ɏ��̃f�[�^���Ƃ��Ă���

Private sComment As String                        ' �R�����g   07/10/05 miyatake ���F�@�\�ǉ�
Private bSampleBtn        As Boolean              ' �T���v���{�^����� 2010/12/08 Marushita

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
    CmdChangeWF_EP.Enabled = False

    If CmdChangeWF_EP.Tag = "WF" Then
        Call sub_cmbc039_3_ChangeHinSpec(1)
        CmdChangeWF_EP.Tag = "EP"
        CmdChangeWF_EP.Caption = "�G�s >>"

    Else
        Call sub_cmbc039_3_ChangeHinSpec(0)
        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "�v�e >>"
    End If

    CmdChangeWF_EP.Enabled = True
End Sub

'>>>>> 2011/07/14 Marushita
'WFC���������ʕ\���{�^���������ɃL���v�`����ʂ�\������
Private Sub cmdDisp_Click()
    '��ʕ\��
    f_hanteiS.picDisp.Picture = LoadPicture(App.Path & CAP_FNAME)
    Call f_hanteiS.Show
End Sub
'<<<<< 2011/07/14 Marushita

'*******************************************************************************
'*    �֐���        : cmdF_Click
'*
'*    �����T�v      : 1.�t�@���N�V�����{�^�����N���b�N���ꂽ��A�e�����ɕ��򂷂�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    intIndex    ,I  ,Integer�@,�R���g���[���z��̓Y��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub cmdF_Click(intIndex As Integer)
    Dim sErrMsg     As String
    Dim sBlkId      As String
    Dim sSXLID      As String
    Dim blnsflg     As Boolean

    lblMsg.Caption = ""
    Tokusai = ""

    '' ��������
    Select Case intIndex
        Case 2          '' �e�Q�L�[�i�T�u���j���[�j
            '>>>>> 2011/07/14 Marushita
            '�������蔻��\�����͕���
            If f_hanteiS.Visible = True Then
                Unload f_hanteiS
            End If
            '<<<<< 2011/07/14 Marushita
            GotoSubMenu
        Case 3          '' �e�R�L�[�i�L�����Z���j
            ' �{�^���A�ł�h����
            cmdF(3).Enabled = False

            Call sub_DispClear
            lblMsg.Caption = GetMsgStr("PWAIT")
            Call sub_LoadAndDisp

            '2003/04/21 F�L�[�t���O�ǉ��@0:������� 1:�T���v���{�^�� 2:���s�{�^��
            giFKeyFlg = 0 '�������
            ReDim FurikaeNaiyouWK(0)
            TokuCntWK = TokuCnt
            ReDim Preserve TokusaiBangou(TokuCntWK)

            ' �{�^���A�ł�h����
            cmdF(3).Enabled = True
        Case 4
            '���{�^��
            Call sub_FurikaeKouho
        Case 5
            '���̃{�^��
            Call sub_TokusaiInput
            Tokusai = "1"
            lblMsg.Caption = TBN_Msg
        Case 6
'Add Start 2011/03/11 SMPK Miyata
            'WFϯ�ߊǗ�ð��ق����ް����擾
            If SelWFmap(vbNullString, SelectSxlID039, sErrMsg) = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_2.lblMsg.Caption = sErrMsg
                Exit Sub
            End If
            
            '���گ�ނ��ް���\��
            If SetWFmapData = FUNCTION_RETURN_FAILURE Then
                f_cmbc039_4.lblMsg.Caption = sErrMsg
                f_cmbc039_4.sprWfmapView.MaxRows = 0
                Exit Sub
            End If

            '�T���v���{�^���������͉�ʂ̏�Ԃ���ԂƐF�Ŕ��f����
            If giFKeyFlg = 1 Then
                fnc_DispChangeDataWfmap '��Ԃ�Map�ɕ\��
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
'                '�������擾
'                sSXLID = Trim(txtKSXLID.text)
'                lblSxlId.Caption = sSXLID
'                sBlkId = vbNullString
'
'                'WFϯ�ߊǗ�ð��ق����ް����擾
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
'            '�T���v���{�^���������͉�ʂ̏�Ԃ���ԂƐF�Ŕ��f����
'            If giFKeyFlg = 1 Then
'                fnc_DispChangeDataWfmap '��Ԃ�Map�ɕ\��
'            End If
'
'            '���گ���ް��\�[�g
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

        Case 7          '' �e�V�L�[�i�s�}���j
            sprExamine.SetFocus
            Call sprExamine_KeyDown(vbKeyF7, 0)
        Case 8          '' �e�W�L�[�i�s�폜�j
            sprExamine.SetFocus
            Call sprExamine_KeyDown(vbKeyF8, 0)
        Case 9          '' �e�X�L�[�i�����C���[�W�\���j
            If sub_DispSample(blnsflg) = True Then
                If blnsflg = True Then '�f�[�^�`�F�b�N�ŃG���[�ɂȂ�����Ԃł̓t���O�𗧂ĂȂ��悤�ɏC��
                    giFKeyFlg = 1
                End If
                Call sub_DrawImage
            End If

            ' �U�փ`�F�b�N�ǉ��ɂ��C��
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
'----- �T���v���{�^���Q�񉟉��s�Ή��@2010/12/08 Marushita
            If bSampleBtn = True And giFKeyFlg = 1 Then
                bSampleBtn = False
                cmdF(7).Enabled = False
                cmdF(8).Enabled = False
                cmdF(10).Enabled = False
            End If
        Case 10         '' �e10�L�[�i�T���v���j
            If sub_DispSample(blnsflg) = False Then
                Exit Sub
            End If

            If blnsflg = True Then '�f�[�^�`�F�b�N�ŃG���[�ɂȂ�����Ԃł̓t���O�𗧂ĂȂ��悤�ɏC��
                giFKeyFlg = 1
            End If

            '2003/04/21 F�L�[�t���O�ǉ��@0:������� 1:�T���v���{�^�� 2:���s�{�^��
            If fnc_DispHinSpec(1) = False Then
                Exit Sub
            End If

            ' �U�փ`�F�b�N�ǉ��ɂ��C��
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
'----- �T���v���{�^���Q�񉟉��s�Ή��@2010/12/08 Marushita
            If bSampleBtn = True And giFKeyFlg = 1 Then
                bSampleBtn = False
                cmdF(7).Enabled = False
                cmdF(8).Enabled = False
                cmdF(10).Enabled = False
            End If
        Case 11         '' �e11�L�[�i�O��ʁj
            '>>>>> 2011/07/14 Marushita
            '��������Q�ƕ\�����͕���
            If f_hanteiS.Visible = True Then
                Unload f_hanteiS
            End If
            '<<<<< 2011/07/14 Marushita
            CloseFormProc f_cmbc039_2, f_cmbc039_3
        Case 12         '' �e12�L�[�i���s�j
            '' �S����ID�̃`�F�b�N
            If f_cmzcChkUser.CanExec(Me.Name, txtStaffID.text) = False Then
                lblMsg.Caption = GetMsgStr("EUSR0")
                Exit Sub
            End If

            '��SaveData��̑O�Ɉړ�
            giFKeyFlg = 2 '���s�{�^��

            Call sub_SaveData

            ' �U�փ`�F�b�N�ǉ��ɂ��C��
            ReDim FurikaeNaiyou(UBound(FurikaeNaiyouWK))
            FurikaeNaiyou = FurikaeNaiyouWK
    End Select
End Sub

'*******************************************************************************
'*    �֐���        : cmdF_GotFocus
'*
'*    �����T�v      : 1.�R���g���[�����t�H�[�J�X���擾�������Ɏ��s���鏈��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    Index       ,I  ,Integer�@,�R���g���[���z��̓Y��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub cmdF_GotFocus(Index As Integer)
    '' �s�폜�^�s�}���t�@���N�V�����L�[�łȂ����
    If Index <> 7 And Index <> 8 Then
        '' �s�폜�^�s�}���t�@���N�V�����L�[�𖳌��ɂ���
        cmdF(7).Enabled = False
        cmdF(8).Enabled = False
    End If

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
    Dim intIndex As Integer

    '' �t�@���N�V�����L�[���L���Ȃ�
    If KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12 Then
        '' ��ʕ\�����b�Z�[�W�N���A
        lblMsg.Caption = ""

        intIndex = KeyCode - (vbKeyF1 - 1)
        If KeyCode = vbKeyF5 Then
            Call cmdF_Click(intIndex)
        End If
        If cmdF(intIndex).Visible = True And cmdF(intIndex).Enabled = True Then
            '' �t�@���N�V�����L�[�������������s����
            If KeyCode <> vbKeyF7 And KeyCode <> vbKeyF8 Then
                Call cmdF_Click(intIndex)
            End If
        End If
    End If
End Sub

'*******************************************************************************
'*    �֐���        : Form_Load
'*
'*    �����T�v      : 1.SXL�\���̏�����Ԃ�ۑ�����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Load()
    Dim sxl As c_cmzcSxl

    Call Pic_Disp(0)

    '' ��ʃN���A
    Call sub_DispClear

'Del Start 2011/03/11 SMPK Miyata
'    '' �t�H�[���ʒu�Z�b�g
'    CenterForm Me
'    Me.Height = 8580
'Del End   2011/03/11 SMPK Miyata
    Me.Show

    ' �o�[�W�������̕\��
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision


    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    '�G�s��s�]���ǉ��Ή�
    sprExamine.ColsFrozen = 35

    CutCntFlg = 0
    sprExamine.col = 5
    sprExamine.TypeComboBoxWidth = 1

    '' �f�[�^��������
    lblMsg.Caption = GetMsgStr("PWAIT")
    DoEvents
    Call sub_InitData

    '' �f�[�^�̃��[�h�ƕ\��
    lblMsg.Caption = GetMsgStr("PWAIT")
    DoEvents
    Call sub_LoadAndDisp

    '' SXL�\���̏�����Ԃ�ۑ�����
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

'2003/04/21 F�L�[�t���O�ǉ��@0:������� 1:�T���v���{�^�� 2:���s�{�^��
    giFKeyFlg = 0 '�������

    ReDim tblWafInd(0)
    ReDim tblNukishi(0) '�����ύX�p�\����
    ' ���̌����ǉ��ɂ��C��
            '' �S����ID�̃`�F�b�N
            If f_cmzcChkUser.CanExec("TOKUSAI", txtStaffID.text) Then
                bTokuKengenFlag = True
            Else
                bTokuKengenFlag = False
            End If

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
End Sub

'*******************************************************************************
'*    �֐���        : Form_Unload
'*
'*    �����T�v      : 1.�t�H�[�����A�����[�h����鎞�Ɏ��s���鏈��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    Cancel      ,I  ,Integer
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    Set orgSXL = Nothing

    '' �����}�E�B���h�E�̃A�����[�h
    Unload f_cmzc003a
End Sub

'*******************************************************************************
'*    �֐���        : sprExamine_Change
'*
'*    �����T�v      : 1.�i�ԕύX���Ɍ�������Ď擾
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    col         ,I  ,Long     ,�I���
'*                    Row         ,I  ,Long     ,�I���s
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sprExamine_Change(ByVal col As Long, ByVal row As Long)
    Dim blAns           As Boolean
    Dim vTemp           As Variant
    Dim vTemp1          As Variant
    Dim vTemp2          As String
    Dim udtFullHinban   As tFullHinban

    '' �ʒu�ƕi�Ԃ��ύX�̏ꍇ
    If (col = 2) Or (col = 4) Then
        '' ���̏��N���A
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
    '' �i�Ԃ��ύX�̏ꍇ
    If (col = 2) Then
        sprExamine.col = 2
        sprExamine.row = row
        sHinban = sprExamine.text

        sprExamine.col = 2
        sprExamine.row = row + 1

        sprExamine.text = sCmbMukeName

        '' ������擾
        sNewMuke = GetMukesaki(sHinban)
        sprExamine.text = sNewMuke

        If InStr(1, sNewMuke, left(sBaseMukesaki, 1), vbTextCompare) = 0 Then
            lblMsg.Caption = "���悪�ύX����Ă��܂��B"
            sprExamine.backColor = vbRed
        Else
            lblMsg.Caption = ""
            sprExamine.backColor = vbWhite
        End If
    End If
    '�T���v���{�^��������f�[�^������������ꂽ�ꍇ�A���s�ł��Ȃ����邽��2003/04/25 okazaki
    bSampFlag = False
End Sub

'*******************************************************************************
'*    �֐���        : sprExamine_ComboSelChange
'*
'*    �����T�v      : 1.�R���{�{�b�N�X�̑I�����ڂ��ύX���ꂽ���Ɏ��s���鏈��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    col         ,I  ,Long     ,�R���{�{�b�N�X�̗�ԍ�
'*                    Row         ,I  ,Long     ,�R���{�{�b�N�X�̍s�ԍ�
'*
'*    �߂�l        : �Ȃ�
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
'*    �֐���        : sprExamine_GotFocus
'*
'*    �����T�v      : 1.�R���g���[�����t�H�[�J�X���擾�������Ɏ��s���鏈��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sprExamine_GotFocus()
    '' �s�폜�^�s�}���t�@���N�V�����L�[��L���ɂ���
'----- �T���v���{�^�������s�Ή��@2010/12/08 Marushita
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
'*    �֐���        : sprExamine_KeyDown
'*
'*    �����T�v      : 1.�����w���ꗗ��ҏW����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer�@,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
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
    Dim blZhinFlg(2)    As Boolean      'Z�i�ԍs�̗L��
    Dim lngsRow         As Long         '�����J�n�s
    Dim lngeRow         As Long         '�����I���s

    '' �L�[�R�[�h����
    Select Case KeyCode
    Case vbKeyF7        '' F7�L�[�i�s�}���j
        'F7�L�[���g�p�s�̂Ƃ������𔲂���@2010/12/14�@Marushita
        If cmdF(7).Enabled = False Then
            Exit Sub
        End If
        Call sub_F_InsertRow

        '' �T���v���t���O�I�t
        bSampFlag = False
        cmdF(12).Enabled = False

        '' ���̏��N���A
        Call sub_ClearTokusai
    Case vbKeyF8        '' F8�L�[�i�s�폜�j
        'F8�L�[���g�p�s�̂Ƃ������𔲂���@2010/12/14�@Marushita
        If cmdF(8).Enabled = False Then
            Exit Sub
        End If
        vDeleteFlag = ""
        '�A�N�e�B�u�Z���̈ʒu�ɂ��Q�s�폜(�t���O�̗����Ă�����͕̂s��)
        '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/06
'        If (sprExamine.GetText(29, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/15 ooba
'        If (sprExamine.GetText(30, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        'GD�ǉ��ɂ��ύX�@05/02/17 ooba
        'If (sprExamine.GetText(31, sprExamine.ActiveRow, vDeleteFlag) = True) Then
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
        If (sprExamine.GetText(37, sprExamine.ActiveRow, vDeleteFlag) = True) Then
            If (vDeleteFlag = "1") Or (vDeleteFlag = "3") Then
                Exit Sub
            End If
        End If
        With sprExamine
            intResult = .ActiveRow Mod 2
            If intResult = 0 Then     '�����s
                lngNewRow = .ActiveRow
                lngNewRow2 = .ActiveRow + 1
            Else                      '�
                lngNewRow = .ActiveRow - 1
                lngNewRow2 = .ActiveRow
            End If

            .GetText 2, lngNewRow2, vGetHinban
            If intResult = 0 Then '�J�[�\���s�������̎��͂P�s��̕i�Ԃ������������ꍇ�Ȃ̂ŕi�Ԃ�����������
                .SetText 2, lngNewRow - 1, vGetHinban
            End If
            .row = lngNewRow
            .row2 = lngNewRow2
            .col = (-1)
            .BlockMode = True
            .FormulaSync = True
            .Action = ActionDeleteRow

            '�ő�s���폜
            .MaxRows = sprExamine.MaxRows - 2

            '�i�Ԃ̐F��K�p����
            .GetText 2, lngNewRow - 1, vGetHinban

            .row = lngNewRow - 1
            .col = 4
           vBColor = .backColor

            If (vGetHinban = "Z" Or vGetHinban = "�y") And vBColor = &H8080FF Then
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

                '' Z�s�݂̂̏����Ƃ���=================================>
                '�폜�s�̏オZ�i�Ԃ��ǂ���
                .col = 11: .row = lngNewRow - 1
                If .backColor = &H8080FF Then blZhinFlg(1) = True Else blZhinFlg(1) = False
                '�폜�s�̉���Z�i�Ԃ��ǂ���
                .col = 11: .row = lngNewRow
                If .backColor = &H8080FF Then blZhinFlg(2) = True Else blZhinFlg(2) = False
                'Z�i�ԍs�̏ꍇ
                If blZhinFlg(1) Or blZhinFlg(2) Then
                    '�����J�n�s�ݒ�
                    If blZhinFlg(1) Then lngsRow = lngNewRow - 1 Else lngsRow = lngNewRow
                    '�����I���s�ݒ�
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

                    ''�������ڗ��̔w�i�ĕ\��
                    For intRowCnt = lngsRow To lngeRow
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

        '' �T���v���t���O�I�t
        bSampFlag = False
        cmdF(12).Enabled = False

        '' ���̏��N���A
        Call sub_ClearTokusai
    End Select
End Sub

'*******************************************************************************
'*    �֐���        : sprExamine_KeyPress
'*
'*    �����T�v      : 1.�����w���T���v������͂���
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyAscii    ,I  ,Integer�@,�����R�[�h
'*
'*    �߂�l        : �Ȃ�
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

        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
'*    �֐���        : sprSpec_GotFocus
'*
'*    �����T�v      : 1.�R���g���[�����t�H�[�J�X���擾������
'*                      �s�폜�^�s�}���t�@���N�V�����L�[�𖳌��ɂ���
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sprSpec_GotFocus()
    '' �s�폜�^�s�}���t�@���N�V�����L�[��L���ɂ���
    cmdF(7).Enabled = False
    cmdF(8).Enabled = False
End Sub

'*******************************************************************************
'*    �֐���        : txtStaffID_KeyDown
'*
'*    �����T�v      : 1.�S���҃R�[�hKeyDown����
'*                      �i���̋@�\�ǉ��ɂ��C���j
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer  ,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub txtStaffID_KeyDown(KeyCode As Integer, Shift As Integer)
    '���̋@�\�ǉ��ɂ��C��
    If f_cmzcChkUser.CanExec("TOKUSAI", txtStaffID.text) Then
        bTokuKengenFlag = True
    Else
        bTokuKengenFlag = False
    End If
End Sub

'*******************************************************************************
'*    �֐���        : txtTarget_KeyDown
'*
'*    �����T�v      : 1.�˂炢�ςŃ��^�[���L�[����
'*                      �i�˂炢�ς������ۯ�ID�A��ۯ�P�A����P�̎擾�j
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    KeyCode     ,I  ,Integer�@,�L�[�R�[�h
'*                    Shift       ,I  ,Integer  ,Shift�L�[�̏��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub txtTarget_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim udtTmpResPosCal As type_ResPosCal
    Dim udtCof          As type_Coefficient
    Dim dblMenseki      As Double
    Dim dblTopWght      As Double
    Dim intPos          As Integer
    Dim lngWgtCharge    As Long     '�ΐ͌v�Z�p�p�����[�^
    Dim dblWgtTop       As Double   '�ΐ͌v�Z�p�p�����[�^
    Dim dblWgtTopCut    As Double   '�ΐ͌v�Z�p�p�����[�^
    Dim dblDM           As Double   '�ΐ͌v�Z�p�p�����[�^
    Dim sBlkId          As String
    Dim intBlkPos       As Integer
    Dim dblCalcPos      As Double

    lblMsg.Caption = ""
    txtBlkID = ""
    txtBlkP = ""
    txtCryP.text = ""

    '' ���^�[���L�[�������̂ݏ������s
    If KeyCode = vbKeyReturn Then
        If ChkTextBox(txtTarget, CHK_NUMBER, 5, 5) = FUNCTION_RETURN_FAILURE Then
            '' �G���[���b�Z�[�W��\������
            lblMsg.Caption = GetMsgStr("EINPM")
            Exit Sub
        End If
        txtTarget.text = toRsStr(val(txtTarget.text))
        udtCof.TOPSMPLPOS = tblSXL.INGOTPOS
        udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT
        udtCof.TOPRES = tblTotal.typ_y013(1, WFRES).MESDATA5
        udtCof.BOTRES = tblTotal.typ_y013(2, WFRES).MESDATA5
        
        ''�ΐ͌v�Z�p�p�����[�^�擾 �}���`����Ή� �Q�Ɗ֐��ύX 2008/05/22 SETsw Nakada
        If GetCoeffParams_new(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
'        If GetCoeffParams(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
            Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
        End If
        dblMenseki = AreaOfCircle(dblDM)
        dblTopWght = dblWgtTop + dblWgtTopCut
        udtCof.DUNMENSEKI = dblMenseki                      ' �f�ʐ�
        udtCof.CHARGEWEIGHT = lngWgtCharge                  ' �`���[�W��
        udtCof.TOPWEIGHT = dblTopWght                       ' �g�b�v�d��
        udtCof.TOPSMPLPOS = tblSXL.INGOTPOS                 ' �g�b�v�ʒu
        udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT  ' �{�g���ʒu

        udtTmpResPosCal.COEFFICIENT = CoefficientCalculation(udtCof)
        udtTmpResPosCal.DUNMENSEKI = udtCof.DUNMENSEKI
        udtTmpResPosCal.CHARGEWEIGHT = udtCof.CHARGEWEIGHT
        udtTmpResPosCal.TOPWEIGHT = udtCof.TOPWEIGHT
        udtTmpResPosCal.TOPSMPLPOS = udtCof.TOPSMPLPOS
        udtTmpResPosCal.TOPRES = udtCof.TOPRES
        udtTmpResPosCal.target = val(txtTarget.text)

        '' �ΐ͌v�Z�֐��̌Ăяo��
        dblCalcPos = PosCalculation(udtTmpResPosCal)
        If dblCalcPos <= -9999 Or dblCalcPos > 9999 Then
            lblMsg.Caption = GetMsgStr("ECLC2")
            Exit Sub
        End If
        intPos = dblCalcPos

        '' �v�Z�ʒu��SXL�̒��Ɋ܂܂�Ă���ꍇ
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
'*    �֐���        : sub_Top_Btm_TEIKOU
'*
'*    �����T�v      : 1.��ʔ͈͂̂˂炢�ς��Z�o
'*                      �i����P�̂���ς��擾�j
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    ingot     �@,I  ,Integer�@,����P
'*                    dblTeikou   ,O  ,Double   ,�˂炢��
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_Top_Btm_TEIKOU(ingot As Integer, dblTeikou As Double)
    Dim udtTmpResPosCal As type_ResPosCal
    Dim udtCof          As type_Coefficient
    Dim dblMenseki      As Double
    Dim dblTopWght      As Double
    Dim intPos          As Integer
    Dim lngWgtCharge    As Long     '�ΐ͌v�Z�p�p�����[�^
    Dim dblWgtTop       As Double   '�ΐ͌v�Z�p�p�����[�^
    Dim dblWgtTopCut    As Double   '�ΐ͌v�Z�p�p�����[�^
    Dim dblDM           As Double   '�ΐ͌v�Z�p�p�����[�^
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
    
    ''�ΐ͌v�Z�p�p�����[�^�擾 �}���`����Ή� �Q�Ɗ֐��ύX 2008/05/22 SETsw Nakada
    If GetCoeffParams_new(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
'    If GetCoeffParams(txtCryNum.text, lngWgtCharge, dblWgtTop, dblWgtTopCut, dblDM) = FUNCTION_RETURN_FAILURE Then
        Debug.Print "�ΐ͌v�Z�p�p�����[�^�̎擾�Ɏ��s����"
    End If

    dblMenseki = AreaOfCircle(dblDM)
    dblTopWght = dblWgtTop + dblWgtTopCut
    udtCof.DUNMENSEKI = dblMenseki                      ' �f�ʐ�
    udtCof.CHARGEWEIGHT = lngWgtCharge                  ' �`���[�W��
    udtCof.TOPWEIGHT = dblTopWght                       ' �g�b�v�d��
    udtCof.TOPSMPLPOS = tblSXL.INGOTPOS                 ' �g�b�v�ʒu
    udtCof.BOTSMPLPOS = tblSXL.INGOTPOS + tblSXL.COUNT  ' �{�g���ʒu

    udtTmpResPosCal.COEFFICIENT = CoefficientCalculation(udtCof)
    udtTmpResPosCal.DUNMENSEKI = udtCof.DUNMENSEKI
    udtTmpResPosCal.CHARGEWEIGHT = udtCof.CHARGEWEIGHT
    udtTmpResPosCal.TOPWEIGHT = udtCof.TOPWEIGHT
    udtTmpResPosCal.TOPSMPLPOS = udtCof.TOPSMPLPOS
    udtTmpResPosCal.TOPRES = udtCof.TOPRES
    udtTmpResPosCal.target = val(ingot)

    '' �ΐ͌v�Z�֐��̌Ăяo��
    dblCalcPos = ResCalculation(udtTmpResPosCal)
    If dblCalcPos <= -9999 Or dblCalcPos > 9999 Then
        lblMsg.Caption = GetMsgStr("ECLC2")
        Exit Sub
    End If

    dblTeikou = dblCalcPos
    Debug.Print "�v�Z��R�l = " & dblTeikou
End Sub

'*******************************************************************************
'*    �֐���        : sub_DispClear
'*
'*    �����T�v      : 1.��ʂ̑S�f�[�^���N���A����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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
    cmdF(5).Enabled = False '���̃{�^��
    cmdF(5).backColor = &H8000000F
    cmdF(12).Enabled = False

    Call Pic_Disp(0)
'----- �T���v���{�^����ԏ������@2010/12/08 Marushita
    bSampleBtn = True
    cmdF(10).Enabled = True
End Sub

'*******************************************************************************
'*    �֐���        : fnc_DispHinSpec
'*
'*    �����T�v      : 1.���i�d�l�f�[�^��\������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    intsyoki    ,I  ,Integer  ,0:�����\���A1:�T���v���{�^��
'*
'*    �߂�l        : �Ȃ�
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
        '' �i�ԃ`�F�b�N
        For i = 1 To sprExamine.MaxRows - 1
            '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX-------start iida 2003/09/06
'            sprExamine.GetText 29, i, vNukisiFlg
            ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/15 ooba
'            sprExamine.GetText 30, i, vNukisiFlg
            'GD�ǉ��ɂ��ύX�@05/02/17 ooba
            'sprExamine.GetText 31, i, vNukisiFlg
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
            sprExamine.GetText 37, i, vNukisiFlg
            If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Then    '�����s��������
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                sprExamine.GetText 37, i + 1, vOldNukisiFlg
                If (vNukisiFlg = 1) And (vOldNukisiFlg = 1) Then  '�Q�s�����ď����\�������s��������
                    '�������Ȃ�
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

        '' �i�Ԃ̐ݒ�
        m = .MaxRows
        ReDim tblHinbanRs(m)
        If m = 0 Then
            Exit Function
        End If
        j = 0
        For i = 1 To m - 1
            .row = i

            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή�
            sprExamine.GetText 37, i, vNukisiFlg

            If intsyoki = 0 Then  '�����\���̂Ƃ��̏���
                '�����s��������
                If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Then
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                        If vNukisiFlg = 1 Then      '2003/11/14 SystemBrain �ǂ�����Ȃ����ǁA�������ɁE�E�E ��
                            .row = i + 1                    '�ł��A������ۯ�1SXL�̎��A���߂ł���
                        Else
                            .row = i + 2
                        End If                      '2003/11/14 SystemBrain �ǂ�����Ȃ����ǁA�������ɁE�E�E ��
                        If i = 1 Or i <> m Then
                            .col = 4
                            intNLen = val(.text) - intNLen

                            '�����12���i�Ԃ����ɑ��݂��邩�������A���ɂ���΍쐬���Ȃ� 2003/11/14 SystemBrain ��
                            blSameHin = False
                            For s = 1 To j
                                If tblHinbanRs(s).HIN.hinban = sHin And tblHinbanRs(s).HIN.mnorevno = val(Mid(sRev, 1, 2)) And _
                                   tblHinbanRs(s).HIN.factory = Mid(sRev, 3, 1) And tblHinbanRs(s).HIN.opecond = Mid(sRev, 4, 1) Then
                                    tblHinbanRs(s).LENGHT = tblHinbanRs(s).LENGHT + intNLen    '���Ԃ���Z �� vNukisiFlg=0���ް��͂�����ʂ�Ȃ��̂Œ���������Ȃ�
                                    blSameHin = True
                                    Exit For
                                End If
                            Next s

                            '�����12���i�Ԃ����ɑ��݂��邩�������A���ɂ���΍쐬���Ȃ� 2003/11/14 SystemBrain ��
                            If Not blSameHin Then
                                j = j + 1

                                With tblHinbanRs(j)                                                 '�\�����Ă�i�Ԃ̎d�l�Ȃ񂾂��� 2003/11/14 ��
                                    .CRYNUM = tblSXL.CRYNUM
                                    .HIN.hinban = sHin                        ' �i��
                                    .HIN.mnorevno = val(Mid(sRev, 1, 2))      ' ���i�ԍ������ԍ�
                                    .HIN.factory = Mid(sRev, 3, 1)            ' �H��
                                    .HIN.opecond = Mid(sRev, 4, 1)            ' ���Ə���
                                    .LENGHT = intNLen                            ' ����
                                End With                                                            '�\�����Ă�i�Ԃ̎d�l�Ȃ񂾂��� 2003/11/14 ��
                            End If
                        End If
                    End If
                End If
            Else
'                '������ǉ�(2�u���b�N1SXL�d�l�̂Ȃ��i�ԂɐU�֑Ή�) �����\���ȊO
                If (vNukisiFlg = 1) Or (vNukisiFlg = 2) Or (i Mod 2 = 0 And vNukisiFlg = 3 And UBound(tblWafInd()) > 2) Then
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                        If vNukisiFlg = 1 Then      '2003/11/14 SystemBrain �ǂ�����Ȃ����ǁA�������ɁE�E�E ��
                            .row = i + 1                    '�ł��A������ۯ�1SXL�̎��A���߂ł���
                        Else
                            .row = i + 2
                        End If                      '2003/11/14 SystemBrain �ǂ�����Ȃ����ǁA�������ɁE�E�E ��
                        If i = 1 Or i <> m Then
                            .col = 4
                            intNLen = val(.text) - intNLen

                            '�����12���i�Ԃ����ɑ��݂��邩�������A���ɂ���΍쐬���Ȃ� 2003/11/14 SystemBrain ��
                            blSameHin = False
                            For s = 1 To j
                                If tblHinbanRs(s).HIN.hinban = sHin And tblHinbanRs(s).HIN.mnorevno = val(Mid(sRev, 1, 2)) And _
                                   tblHinbanRs(s).HIN.factory = Mid(sRev, 3, 1) And tblHinbanRs(s).HIN.opecond = Mid(sRev, 4, 1) Then
                                    tblHinbanRs(s).LENGHT = tblHinbanRs(s).LENGHT + intNLen    '���Ԃ���Z �� vNukisiFlg=0,3���ް��͂�����ʂ�Ȃ��̂Œ���������Ȃ�
                                    blSameHin = True
                                    Exit For
                                End If
                            Next s

                            '�����12���i�Ԃ����ɑ��݂��邩�������A���ɂ���΍쐬���Ȃ� 2003/11/14 SystemBrain ��
                            If Not blSameHin Then
                                j = j + 1
                                With tblHinbanRs(j)                                                 '�\�����Ă�i�Ԃ̎d�l�Ȃ񂾂��� 2003/11/14 ��
                                    .CRYNUM = tblSXL.CRYNUM
                                    .HIN.hinban = sHin                        ' �i��
                                    .HIN.mnorevno = val(Mid(sRev, 1, 2))      ' ���i�ԍ������ԍ�
                                    .HIN.factory = Mid(sRev, 3, 1)            ' �H��
                                    .HIN.opecond = Mid(sRev, 4, 1)            ' ���Ə���
                                    .LENGHT = intNLen                            ' ����
                                End With                                                            '�\�����Ă�i�Ԃ̎d�l�Ȃ񂾂��� 2003/11/14 ��
                            End If
                        End If
                    End If
                End If
            End If
        Next i
        ReDim Preserve tblHinbanRs(j)
    End With

    '' DB����f�[�^��ǂݍ���
    '' �d�l�����̃`�F�b�N
    If DBDRV_scmzc_fcmlc001d_DispSiyou(tblHinbanRs, tblsiyou, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = sErrMsg
        Exit Function
    End If

    '' ���i�d�l�f�[�^�̕\��
    m = UBound(tblHinbanRs())
    With sprSpec
        j = 0
        .MaxRows = m
        For i = 1 To m
            If Trim$(tblHinbanRs(i).HIN.hinban) <> "" Then
                j = j + 1
                .row = j
                .col = 1
                .text = tblHinbanRs(i).HIN.hinban   ' �i��
                .col = 2
                .text = Format(tblHinbanRs(j).HIN.mnorevno, "00") & tblHinbanRs(j).HIN.factory & tblHinbanRs(j).HIN.opecond
                .col = 3                            ' ���R
                .text = toRsStr_nl(tblsiyou(i).HWFRMIN, tblsiyou(i).HWFRMAX) '2001/12/26 S.Sano
                .col = 4
                .text = tblsiyou(i).KEIKAKUL        ' �v�撷
                .col = 5
                .text = tblHinbanRs(i).LENGHT       ' ���蒷
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

                '�g�U��,Nr�Z�x�ǉ��@06/06/08 ooba START ===================>
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
                .text = tblsiyou(i).HWFZOHWS        ' AO        ''�c���_�f�ǉ�
                ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/15 ooba

                '�d�l����ީOT1�OT2�Ɏw���̗L����\������
                If tblsiyou(i).HWFOT1 = "1" Then
                    .col = 22
                    .text = "�L"
                ElseIf tblsiyou(i).HWFOT1 = "0" Then
                    .col = 22
                    .text = "��"
                End If
                If tblsiyou(i).HWFOT2 = "1" Then
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                    .col = 30
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    .text = "�L"
                ElseIf tblsiyou(i).HWFOT2 = "0" Then
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                    .col = 30
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                    .text = "��"
                End If

'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                .col = 23
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                If tblsiyou(i).HWFDENHS = "H" Or tblsiyou(i).HWFLDLHS = "H" Or _
                        tblsiyou(i).HWFDVDHS = "H" Then
                    .text = "H"
                ElseIf tblsiyou(i).HWFDENHS = "S" Or tblsiyou(i).HWFLDLHS = "S" Or _
                        tblsiyou(i).HWFDVDHS = "S" Then
                    .text = "S"
                Else
                    .text = " "
                End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
                .col = 24:  .text = tblsiyou(i).HEPBM1HS
                .col = 25:  .text = tblsiyou(i).HEPBM2HS
                .col = 26:  .text = tblsiyou(i).HEPBM3HS
                .col = 27:  .text = tblsiyou(i).HEPOF1HS
                .col = 28:  .text = tblsiyou(i).HEPOF2HS
                .col = 29:  .text = tblsiyou(i).HEPOF3HS
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
                '>>>>> ���Ԕ����K�i�Z�b�g�ǉ� 2011/07/15 Marushita
                .col = 31:  .text = tblsiyou(i).CHUTAN
                .col = 32:  .text = tblsiyou(i).CHUKYO
                .col = 33:  .text = tblsiyou(i).CHUFLG
                '<<<<< ���Ԕ����K�i�Z�b�g�ǉ� 2011/07/15 Marushita
            End If
        Next i
        .MaxRows = j

'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
        '' WF�d�l�������\������
        Call sub_cmbc039_3_ChangeHinSpec(0)
        CmdChangeWF_EP.Tag = "WF"
        CmdChangeWF_EP.Caption = "�v�e >>"
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-
    End With

    '�d�l���Ȃ��f�[�^�ւ̐U�ւ��s�����Ƃ��̓G���[�ɂ��Ȃ�
    '�����\���s�̕i�Ԃ�����ύX�����Ƃ�
    '�����\�������̂Ƃ��̓`�F�b�N�����Ȃ�
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
        '�g�U��,Nr�Z�x�ǉ�
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
        ''�c���_�f�ǉ�
        If tblsiyou(i).HWFZOHWS <> "H" And tblsiyou(i).HWFZOHWS <> "S" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOT1 = "0" Then
            intCnt = intCnt + 1
        End If
        If tblsiyou(i).HWFOT2 = "0" Then
            intCnt = intCnt + 1
        End If
        'GD�ǉ�
        If tblsiyou(i).HWFDENHS <> "H" And tblsiyou(i).HWFDENHS <> "S" And _
                tblsiyou(i).HWFLDLHS <> "H" And tblsiyou(i).HWFLDLHS <> "S" And _
                tblsiyou(i).HWFDVDHS <> "H" And tblsiyou(i).HWFDVDHS <> "S" Then
            intCnt = intCnt + 1
        End If
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -s-
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
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh -e-

'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        If intCnt = 25 Then
            If UBound(tblWafInd()) > 2 Then
                lblMsg.Caption = "�d�l������܂���"
                cmdF(12).Enabled = False
                Exit Function
            End If
        End If
    Next
    End If
    fnc_DispHinSpec = True
End Function

'*******************************************************************************
'*    �֐���        : sub_DispExamineData
'*
'*    �����T�v      : 1.�����w���f�[�^��\������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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

    '' �����w���f�[�^�̕\��
    With sprExamine
        .MaxRows = 2
        For i = 1 To 2
            '' �u���b�NID
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

            '' �u���b�NP
            .col = 2
            intREBlockPos = DBData2DispData(tblsmp(i).INGOTPOS - intBlockStPos, "0")
            .text = intREBlockPos

            '' ����P
            .col = 3
            .text = tblsmp(i).INGOTPOS

            '' �敪�R���{�̐ݒ�
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

            '' �i��
            .col = 5
            If i = 1 Then
                .text = Trim(tblTotal.typ_Param.hinban)
            Else
                '' SXL�̉��[�͓����̕i�Ԃ�\��������
                .text = tblsmp(1).hinban
            End If
        Next i

        '' WF��������ɂčĔ����̎w�����ł��ꍇ�A�s�ǉ�
        For i = 1 To 2
            If tblTotal.bOKNG(i) = False And _
               tblTotal.dblScut(i) > tblSXL.INGOTPOS And _
               tblTotal.dblScut(i) < tblSXL.INGOTPOS + tblSXL.COUNT Then
                .MaxRows = .MaxRows + 1
                .row = .ActiveRow
                .Action = ActionInsertRow

                '' �u���b�NID
                intBlockStPos = orgXl.Blks.UpperPos(tblTotal.dblScut(i))
                Set Blk = orgXl.Blks(CStr(intBlockStPos))

                '' �u���b�NID�R���{�̐ݒ�
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

                '' �u���b�NP
                .col = 2
                intREBlockPos = DBData2DispData(tblTotal.dblScut(i) - orgXl.Blks.UpperPos(tblTotal.dblScut(i)), "0")
                .text = intREBlockPos
                .col = 3
                .text = DBData2DispData(tblTotal.dblScut(i), "0")

                '' �敪�R���{�̐ݒ�
                m = UBound(tblPrcList)
                sList = GetGPCodeDspStr(tblPrcList(1).CODE, tblPrcList(1).INFO1)
                For j = 2 To m
                    sList = sList & vbTab & GetGPCodeDspStr(tblPrcList(j).CODE, tblPrcList(j).INFO1)
                Next j
                .col = 4
                .TypeComboBoxList = sList
                .TypeComboBoxCurSel = 0

                '' �T���v��ID�̐ݒ�
                '' �ǉ��s�̕i�Ԑݒ�(������Z�i�Ԃ͒ǉ��s�ɑ΂���Z�����Ă邽�߁A���̃��[�v���Őݒ肷��)
                .col = 5
                If i = 1 Then
                    .text = tblSXL.hinban
                ElseIf i = 2 Then
                    .text = "Z"
                End If
                InMaxRow = InMaxRow + 1
            End If
        Next i

        '' �����ʒu�ɂă\�[�g��������
        .col = 1
        .col2 = .MaxCols
        .row = 1
        .row2 = .MaxRows
        .SortBy = SortByRow
        .SortKey(1) = 3
        .SortKeyOrder(1) = SortKeyOrderAscending
        .Action = ActionSort

        '' �i�Ԃɔp�����ݒ�(�㑤NG�̏ꍇ�́A1�s�ڂ�Z�ɂ���)
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
'*    �֐���        : sub_DispSample
'*
'*    �����T�v      : 1.�K��l�̌����w���T���v���f�[�^��\������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    blnsflg�@�@ ,I  ,Boolean  ,�f�[�^�`�F�b�N�����t���O
'*
'*    �߂�l        : �Ȃ�
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
    Dim intModori       As Integer  '�߂�l
    Dim sGHin           As String   '�i��
    Dim blnhflg         As Boolean  '���f����t���O
    Dim intZkbn         As Integer  'Z�敪
    Dim udtFHin         As tFullHinban
    Dim intSmpkbn       As Integer '�T���v���敪����\�T���v�������f����
    Dim sHinRev         As String
    Dim intBcnt         As Integer  '���L(��)�T���v����
    
    Dim flg          As Integer
    Dim now_bid      As String
    Dim old_bid      As String
    Dim iJCnt1      As Integer
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
    Dim sird1stBlockSet As Boolean          '������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END
     

    sub_DispSample = False
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
    sird1stBlockSet = False             '������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END
    
    intBcnt = 1
    ReDim iBetuRow(sprExamine.MaxRows)

    blNukisiOk = False
    '' �Ĕ����̐V�K�s�����݂��Ȃ���Ώ����𔲂���&�S���p����Z�i�Ԃ̂ݗL���Ƃ���

    With sprExamine
        For intLoopCnt = 1 To .MaxRows
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

    '' �����w���ꗗ�̓��̓`�F�b�N�iWFϯ�߁j
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

'�i�Ԃ̂m�t�k�k�`�F�b�N�͂��łɂ͂����Ă���B
'�����łm�t�k�k���͂����Ă���̂́A��ʃ��C�A�E�g�ύX�ɂ��A�m�t�k�k�s(�i�ԓ��͕s��)�����݂��邽�߁B
    '' �i�ԃ`�F�b�N
    For c0 = 1 To sprExamine.MaxRows - 1
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg

        If (vNukisiFlg = "1") Or (vNukisiFlg = "2") Then    '�����s��������
            If (vNukisiFlg = "1") Then
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                sprExamine.GetText 37, c0 + 1, vNukisiFlg

                If vNukisiFlg = "1" Then
                    '�ǉ������̖����s�͉������Ȃ�
                ElseIf vNukisiFlg = "2" Then
                    sprExamine.GetText 2, c0, vTemp             '�e�s�ɕi�Ԃ������Ă���킯�ł͂Ȃ��A�_�~�[�ł͑S�s�ɕi�Ԃ������Ă���̂ł��������
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
                        ''�c���_�f�d�l�`�F�b�N
                        Else
                            iChkAoi = ChkAoiSiyou(udtFullHinban)
                            If iChkAoi < 0 Then
                                lblMsg.Caption = "�c���_�f(AOi)�d�l�G���["
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
                    ''�c���_�f�d�l�`�F�b�N
                    Else
                        iChkAoi = ChkAoiSiyou(udtFullHinban)
                        If iChkAoi < 0 Then
                            lblMsg.Caption = "�c���_�f(AOi)�d�l�G���["
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next c0

   '�����w���p�\����
    ReDim Preserve tblNukishi(m)

    '' �敪�G���[�`�F�b�N
    If fnc_CheckHinbanZ() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EHIN6")
        Exit Function
    End If

    '' �G���[�`�F�b�N
    If fnc_CheckBlockP() = FUNCTION_RETURN_FAILURE Then
        Exit Function
    End If

    '�����ް��s���\���`�F�b�N
    If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr(sMsg)
        Exit Function
    End If

    '�f�[�^�`�F�b�N������������t���O�𗧂Ă�
    blnsflg = True

    '' �����w���m�F���b�Z�[�W
    If MsgBox(GetMsgStr("PSMP1"), vbOKCancel, "�����w���T���v������") = vbCancel Then
        sub_DispSample = False
        Exit Function
    End If

    lblMsg.Caption = GetMsgStr("PWAIT")

     '�\���̃N���A
    ReDim tblKns(0) '�������ڐؑ֗p�̍\���̂��N���A����

    DoEvents

    bSampFlag = False
    bMotoGDcpyFlg(1) = False
    bMotoGDcpyFlg(2) = False

    '' �Ĕ����w���e�[�u���̍X�V
    If fnc_UpdateData() = FUNCTION_RETURN_FAILURE Then
        cmdF(12).Enabled = False
        bSampFlag = False
    Else
        cmdF(12).Enabled = True
        bSampFlag = True
    End If

    '' �����w���f�[�^�̕\��
    With sprExamine
        .ReDraw = False
        If .MaxRows = 0 Then
            Exit Function
        End If

       '' �������ڂ̓��e��ݒ�
        m = .MaxRows
        n = UBound(tblWafInd)
        intNukisiRow = 0
        intCngBlpnt = 1
        For i = 1 To m
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
            .GetText 37, i, vNukisiFlg

    ''''''' '�����\���̃u���b�N�̋�
            If i Mod 2 = 0 And vNukisiFlg = "3" Then
                '�T���v���N���A
                .SetText 10, i, vbNullString
                .SetText 10, i + 1, vbNullString
                .SetText 8, i, gsWF_STA_NORMAL
                .SetText 8, i + 1, gsWF_STA_NORMAL
                '�������ڃN���A
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

                '' �����w���̐ݒ�
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
                        .SMP.CRYINDAO = GetWFSamp(.HINUP, .HINDN, 18)    '�c���_�f�ǉ�
                        .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)    'GD�ǉ�

                        'GDײ������@�\�ǉ�
                        If Trim(.SMP.CRYINDGD) = "3" Then
                            .SMP.CRYINDGD2 = GetWFSamp(.HINUP, .HINDN, 26)
                        Else
                            .SMP.CRYINDGD2 = ""
                        End If

                        '--- 2006/08/15 Add �G�s��s�]���ǉ��Ή�
                        .SMP.EPIINDB1 = GetWFSamp(.HINUP, .HINDN, 20)
                        .SMP.EPIINDB2 = GetWFSamp(.HINUP, .HINDN, 21)
                        .SMP.EPIINDB3 = GetWFSamp(.HINUP, .HINDN, 22)
                        .SMP.EPIINDL1 = GetWFSamp(.HINUP, .HINDN, 23)
                        .SMP.EPIINDL2 = GetWFSamp(.HINUP, .HINDN, 24)
                        .SMP.EPIINDL3 = GetWFSamp(.HINUP, .HINDN, 25)
                    End With
                Next j

                '�������ڂ̕��ޕ\��
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                .GetText 37, i, vNukisiFlg

                If CheckGetSampleID(i) = True Then
                    intNukisiRow = intNukisiRow + 1
                    '�\���F��x����
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

                    '���f���菈���ǉ��̂��ߏC��
                    If GetSampleID(intNukisiRow, sSampID1, sSampID2, CInt(vNukisiFlg)) = True Then
                    '���L�T���v�������݂��A�؂�ւ����ł���ꍇ
                        blKirikaeflg = True  '2003/05/26 okazaki

                        If Right(sSampID1, 1) = "U" Then
                            vViewSmpId = sSampID1
                            vViewSmpId2 = vbNullString
                        Else
                            vViewSmpId = sSampID2
                            vViewSmpId2 = vbNullString
                        End If

                        intZkbn = 0 'Z�敪��������

                        '���f����`�F�b�N�֐��Ăяo��
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
'''                        Call sub_Hanei(i, intNukisiRow, intZkbn)
                        Call sub_Hanei(i, intNukisiRow, intZkbn, sird1stBlockSet)           '���Ұ��ǉ��F������SIRD����َw���ݒ�L��
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP END

                         intSmpkbn = 0 '������
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

                        '�ۏؕ��@�ύX�Ή�
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

                        If .text <> "2" Then                                                                '���ʊ֐��Ŕ��f���ł��Ȃ��Ƃ�
                            If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                '�d�l������Ƃ�
                                If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then '�����Ƃ�����
                                    tblNukishi(i + 1).WFINDOICW = "2"                                       '���̃f�[�^�𔽉f�ɂ���
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

                        ''�c���_�f�ǉ�
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

                        '�G�s��s�]���ǉ��Ή�
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

                        ''GD�ǉ�
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

                        '�G�s��s�]���ǉ��Ή�
                        ' VB��̐���(�v���V�[�W���e��64k����)�̂��ߤ�G�s���͕ʊ֐��ŏ�������
                        Call sub_DispSumple_Hanei_Ep_1(i, intNukisiRow, intSmpkbn)

                        ''���R�ŕK�����т𗧂Ă鏈����ǉ�
                        .col = 11
                        .row = i
                        Dim intCntrs As Integer
                        If .text = "2" Then
                            '�G�s��s�]���ǉ��Ή�
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
                            '�G�s��s�]���ǉ��Ή�
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

                        '���f�[�^�`�F�b�N���� ���f�[�^�����ƐF��h�鏈����2�s�܂Ƃ߂čs��
                          Call sub_Jitu(i + 1, blKirikaeflg, blnhflg, intZkbn)

                        '���f�[�^�`�F�b�N����
                        '�F��h��
                        Call sub_Paint(i + 1)
                   Else
                        '�؂�ւ��Ȃ��̏ꍇ
                        If Right(sSampID1, 1) = "U" Then
                            vViewSmpId = sSampID1
                            vViewSmpId2 = sSampID2
                        Else
                            vViewSmpId = sSampID2
                            vViewSmpId2 = sSampID1
                        End If

                        intZkbn = 0
                        '��i�Ԃ܂��͉��i�Ԃ�Z�������ꍇ�̏���������
                        '��i�Ԃ�Z
                          If Trim(tblWafInd(intNukisiRow).HINUP.hinban) = "Z" Then
                            intZkbn = 1
                                If tblWafInd(intNukisiRow).HINDN.hinban = tblSXL.hinban Then  '�U�ւ����Ă��Ȃ��Ƃ��ƐU�ւ����i�Ԃ����ɖ߂��Ƃ�
                                    intZkbn = 3
                                End If

                                udtFHin.factory = tblSXL.factory
                                udtFHin.hinban = tblSXL.hinban
                                udtFHin.mnorevno = tblSXL.REVNUM
                                udtFHin.opecond = tblSXL.opecond
                            With tblWafInd(intNukisiRow)
                            ''''''''''''''''''''''''''�U�֑O�̕i��-------------------,�U�֌�̕i��
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
                            .SMP.CRYOTHER1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 16)  '03/05/26 �㓡
                            .SMP.CRYOTHER2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 17)   '03/05/28 �㓡
                            .SMP.CRYINDAO = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 18)    '�c���_�f�ǉ��@03/12/15 ooba
                            .SMP.CRYINDGD = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 19)    'GD�ǉ��@05/02/18 ooba

                            '�G�s��s�]���ǉ��Ή�
                            .SMP.EPIINDB1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 20)
                            .SMP.EPIINDB2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 21)
                            .SMP.EPIINDB3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 22)
                            .SMP.EPIINDL1 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 23)
                            .SMP.EPIINDL2 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 24)
                            .SMP.EPIINDL3 = GetWFSamp(udtFHin, tblWafInd(intNukisiRow).HINDN, 25)
                            End With
                        '���i�Ԃ�Z
                        ElseIf Trim(tblWafInd(intNukisiRow).HINDN.hinban) = "Z" Then
                            intZkbn = 2
                                If tblWafInd(intNukisiRow).HINUP.hinban = tblSXL.hinban Then '�U�ւ����Ă��Ȃ��Ƃ��ƐU�ւ����i�Ԃ����ɖ߂��Ƃ�
                                    intZkbn = 4
                                End If
                                udtFHin.factory = tblSXL.factory
                                udtFHin.hinban = tblSXL.hinban
                                udtFHin.mnorevno = tblSXL.REVNUM
                                udtFHin.opecond = tblSXL.opecond
                            With tblWafInd(intNukisiRow)

                            ''''''''''''''''''''''''''�U�֌�̕i��---------------,�U�֑O�̕i��
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
                            .SMP.CRYOTHER1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 16)  '03/05/26 �㓡
                            .SMP.CRYOTHER2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 17)   '03/05/28 �㓡
                            .SMP.CRYINDAO = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 18)    '�c���_�f�ǉ��@03/12/15 ooba
                            .SMP.CRYINDGD = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 19)    'GD�ǉ��@05/02/18 ooba

                            '�G�s��s�]���ǉ��Ή�
                            .SMP.EPIINDB1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 20)
                            .SMP.EPIINDB2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 21)
                            .SMP.EPIINDB3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 22)
                            .SMP.EPIINDL1 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 23)
                            .SMP.EPIINDL2 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 24)
                            .SMP.EPIINDL3 = GetWFSamp(tblWafInd(intNukisiRow).HINUP, udtFHin, 25)
                            End With
                        End If

                        '���������̂��߂ɋ��L(��)�̔����s��Ă���
                        If intZkbn = 0 Then
                            iBetuRow(intBcnt) = i + 1
                            intBcnt = intBcnt + 1
                        End If

                        '���f����`�F�b�N�֐��Ăяo��
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
'''                        Call sub_Hanei(i, intNukisiRow, intZkbn)
                        Call sub_Hanei(i, intNukisiRow, intZkbn, sird1stBlockSet)           '���Ұ��ǉ��F������SIRD����َw���ݒ�L��
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP END

                        '�G�s��s�]���ǉ��Ή�
                        Dim skensa1(24) As String
                        Dim skensa2(24) As String
                         Call sub_Betu(i, intNukisiRow, skensa1(), skensa2(), intZkbn, blnhflg)

                        'Z�̂Ƃ�
                          '2
                          '1
                          '1
                          '2
                        '���L�̔���
                        .row = i
                        .col = 11
                        .backColor = vbWhite

                        '�ύX---
                        'Rs�ׂ͗̃f�[�^�̔��f
                        .text = skensa1(0)
                        If intZkbn = 4 Or intZkbn = 2 Then
                            If skensa2(0) = "2" Or skensa2(0) = "1" Then  '���L��
                                .text = "1" '���������
                                tblNukishi(i).WFINDRSCW = "1"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                                tblNukishi(i).WFRESRS1CW = "0" '���і�
                                '�R�s�[
                                tblNukishi(i + 1).WFINDRSCW = "2"  '�����𗧂ĂȂ�
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i).WFSMPLIDRSCW
                                tblNukishi(i + 1).WFRESRS1CW = tblNukishi(i).WFRESRS1CW  '���уt���O�������ĂȂ�����0
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
                        ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
                            .text = skensa1(0)
                            If .text = "1" Then
                                tblNukishi(i).WFINDRSCW = "1"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                            ElseIf .text = "2" Then
                                tblNukishi(i).WFINDRSCW = "2"
                                tblNukishi(i).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                                tblNukishi(i).WFRESRS1CW = "0" '���і�
                            End If
                        End If

                        .col = 12
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then '���ʊ֐��Ŕ��f�s��
                            If intZkbn = 4 Or intZkbn = 2 Then '
                                If tblNukishi(i + 1).WFINDOICW = "1" Then '2�������Ă���
                                    tblNukishi(i).WFINDOICW = "2"
                                    tblNukishi(i).WFSMPLIDOICW = ""
                                    tblNukishi(i).WFRESOICW = "1" '���ї��Ă�
                                Else
                                    '���ʊ֐��Ŕ��肵�Ă��錋�ʂ�
                                    '�R�s�[
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                    tblNukishi(i + 1).WFRESOICW = tblNukishi(i).WFRESOICW
                                End If
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
                                If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                 '�d�l��������
                                    If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then  '���������̏ꍇ
                                          If skensa1(1) = "2" Then                                             '�d�l�̌������Ȃ��ق������f
                                            tblNukishi(i).WFINDOICW = "2"
                                            tblNukishi(i).WFSMPLIDOICW = tblNukishi(i + 1).WFSMPLIDOICW
                                            tblNukishi(i).WFRESOICW = "0"
                                            .text = skensa1(1)
                                          End If
                                    End If
                                End If
                            End If
                        ElseIf (intZkbn = 4 And .text = "2") Or (intZkbn = 2 And .text = "2") Then  '���ʊ֐��Ŕ��f�̂Ƃ��̃T���v��ID���R�s�[
                            tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                        End If

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                                    '�R�s�[
                                    tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB2CW = tblNukishi(i).WFRESB2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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

                        '�ۏؕ��@�ύX�Ή�
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

                        '�ۏؕ��@�ύX�Ή�
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

                        '�ۏؕ��@�ύX�Ή�
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

                        '�ۏؕ��@�ύX�Ή�
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(14) = "0" Then
                            tblNukishi(i + 1).WFINDDO3CW = "0"
                            tblNukishi(i + 1).WFSMPLIDDO3CW = ""
                            tblNukishi(i + 1).WFRESDO3CW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        ''�c���_�f�ǉ�
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

                        '�ۏؕ��@�ύX�Ή�
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(15) = "0" Then
                            tblNukishi(i + 1).WFINDAOICW = "0"
                            tblNukishi(i + 1).WFSMPLIDAOICW = ""
                            tblNukishi(i + 1).WFRESAOICW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        .col = 27 'OT1(������)
                        .backColor = vbWhite

                        '�ύX��
                        If intZkbn = 4 Or intZkbn = 2 Then
                            .text = skensa1(16)
                            If skensa1(16) = "1" Then '�d�l�L��
                                tblNukishi(i).WFINDOT1CW = "1"
                                tblNukishi(i).WFSMPLIDOT1CW = tblNukishi(i).REPSMPLIDCW
                                '�R�s�[
                                tblNukishi(i + 1).WFINDOT1CW = "0"  'OT�͔��f�͂Ȃ�
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

                        '�ύX��
                        .col = 35
                        .backColor = vbWhite
                        If intZkbn = 4 Or intZkbn = 2 Then
                            .text = skensa1(17)
                            If skensa1(17) = "1" Then '�d�l�L��
                                tblNukishi(i).WFINDOT2CW = "1"
                                tblNukishi(i).WFSMPLIDOT2CW = tblNukishi(i).REPSMPLIDCW

                                '�R�s�[
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

                        '�G�s��s�]���ǉ��Ή�
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

                        '�ۏؕ��@�ύX�Ή�
                        If (intZkbn = 4 Or intZkbn = 2) And skensa2(18) = "0" Then
                            tblNukishi(i + 1).WFINDGDCW = "0"
                            tblNukishi(i + 1).WFSMPLIDGDCW = ""
                            tblNukishi(i + 1).WFRESGDCW = "0"
                            tblNukishi(i + 1).WFHSGDCW = "0"
                            .row = i + 1
                            .text = ""
                            .row = i
                        End If

                        '�G�s��s�]���ǉ��Ή�
                        ' VB��̐���(�v���V�[�W���e��64k����)�̂��ߤ�G�s���͕ʊ֐��ŏ�������
                        Call sub_DispSumple_Hanei_Ep_2(i, intNukisiRow, skensa1(), skensa2(), intZkbn)
                        .row = i + 1
                        .col = 11 'Rs
                        .backColor = vbWhite

                        '�ύX-------
                        .text = skensa2(0)
                        If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                            If skensa1(0) = "2" Or skensa1(0) = "1" Then  '���L�ł���
                                .text = "1"
                                tblNukishi(i + 1).WFINDRSCW = "1" '���������
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW '�C��
                                tblNukishi(i + 1).WFRESRS1CW = "0"
                                '�R�s�[
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
                        ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
                            If .text = "2" Then
                                tblNukishi(i + 1).WFINDRSCW = "2"
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i).REPSMPLIDCW
                               '���уt���O�͗��ĂȂ�
                            ElseIf .text = "1" Then
                                tblNukishi(i + 1).WFINDRSCW = "1"
                                tblNukishi(i + 1).WFSMPLIDRSCW = tblNukishi(i + 1).REPSMPLIDCW
                            End If
                        End If

                        '-----�ύX
                        .col = 12
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then '���f����֐��Ŕ��fOK�ȊO
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDOICW = "1" Then  '1�������Ă͂����Ȃ�(���f�̂͂�)
                                    tblNukishi(i + 1).WFINDOICW = "2"
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDOICW
                                    tblNukishi(i + 1).WFRESOICW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDOICW = tblNukishi(i + 1).WFSMPLIDOICW
                                    tblNukishi(i).WFRESOICW = tblNukishi(i + 1).WFRESOICW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
                                If tblWafInd(intNukisiRow).SMP.CRYINDOI = "3" Then                                 '�d�l��������
                                    If tblNukishi(i).WFINDOICW = "1" And tblNukishi(i + 1).WFINDOICW = "1" Then '���������̏ꍇ
                                        If skensa2(1) = "2" Then                                           '�d�l�̌������Ȃ��ق������f
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDB1CW = "1" Then
                                    tblNukishi(i + 1).WFINDB1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB1CW = tblNukishi(i).WFSMPLIDB1CW
                                    tblNukishi(i + 1).WFRESB1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB1CW = tblNukishi(i + 1).WFSMPLIDB1CW
                                    tblNukishi(i).WFRESB1CW = tblNukishi(i + 1).WFRESB1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDB2CW = "1" Then
                                    tblNukishi(i + 1).WFINDB2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB2CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB2CW = tblNukishi(i + 1).WFSMPLIDB2CW
                                    tblNukishi(i).WFRESB2CW = tblNukishi(i + 1).WFRESB2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDB3CW = "1" Then
                                    tblNukishi(i + 1).WFINDB3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDB3CW = tblNukishi(i).WFSMPLIDB2CW
                                    tblNukishi(i + 1).WFRESB3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDB3CW = tblNukishi(i + 1).WFSMPLIDB3CW
                                    tblNukishi(i).WFRESB3CW = tblNukishi(i + 1).WFRESB3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDL1CW = "1" Then
                                    tblNukishi(i + 1).WFINDL1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL1CW = tblNukishi(i).WFSMPLIDL1CW
                                    tblNukishi(i + 1).WFRESL1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL1CW = tblNukishi(i + 1).WFSMPLIDL1CW
                                    tblNukishi(i).WFRESL1CW = tblNukishi(i + 1).WFRESL1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDL2CW = "1" Then
                                    tblNukishi(i + 1).WFINDL2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL2CW = tblNukishi(i).WFSMPLIDL2CW
                                    tblNukishi(i + 1).WFRESL2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL2CW = tblNukishi(i + 1).WFSMPLIDL2CW
                                    tblNukishi(i).WFRESL2CW = tblNukishi(i + 1).WFRESL2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDL3CW = "1" Then
                                    tblNukishi(i + 1).WFINDL3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL3CW = tblNukishi(i).WFSMPLIDL3CW
                                    tblNukishi(i + 1).WFRESL3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL3CW = tblNukishi(i + 1).WFSMPLIDL3CW
                                    tblNukishi(i).WFRESL3CW = tblNukishi(i + 1).WFRESL3CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDL4CW = "1" Then
                                    tblNukishi(i + 1).WFINDL4CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDL4CW = tblNukishi(i).WFSMPLIDL4CW
                                    tblNukishi(i + 1).WFRESL4CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDL4CW = tblNukishi(i + 1).WFSMPLIDL4CW
                                    tblNukishi(i).WFRESL4CW = tblNukishi(i + 1).WFRESL4CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDDSCW = "1" Then
                                    tblNukishi(i + 1).WFINDDSCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDSCW = tblNukishi(i).WFSMPLIDDSCW
                                    tblNukishi(i + 1).WFRESDSCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDSCW = tblNukishi(i + 1).WFSMPLIDDSCW
                                    tblNukishi(i).WFRESDSCW = tblNukishi(i + 1).WFRESDSCW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDDZCW = "1" Then
                                    tblNukishi(i + 1).WFINDDZCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDZCW = tblNukishi(i).WFSMPLIDDZCW
                                    tblNukishi(i + 1).WFRESDZCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDZCW = tblNukishi(i + 1).WFSMPLIDDZCW
                                    tblNukishi(i).WFRESDZCW = tblNukishi(i + 1).WFRESDZCW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDSPCW = "1" Then
                                    tblNukishi(i + 1).WFINDSPCW = "2"
                                    tblNukishi(i + 1).WFSMPLIDSPCW = tblNukishi(i).WFSMPLIDSPCW
                                    tblNukishi(i + 1).WFRESSPCW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDSPCW = tblNukishi(i + 1).WFSMPLIDSPCW
                                    tblNukishi(i).WFRESSPCW = tblNukishi(i + 1).WFRESSPCW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDDO1CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO1CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO1CW = tblNukishi(i).WFSMPLIDDO1CW
                                    tblNukishi(i + 1).WFRESDO1CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO1CW = tblNukishi(i + 1).WFSMPLIDDO1CW
                                    tblNukishi(i).WFRESDO1CW = tblNukishi(i + 1).WFRESDO1CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDDO2CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO2CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDOICW = tblNukishi(i).WFSMPLIDDO2CW
                                    tblNukishi(i + 1).WFRESDO2CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO2CW = tblNukishi(i + 1).WFSMPLIDDO2CW
                                    tblNukishi(i).WFRESDO2CW = tblNukishi(i + 1).WFRESDO2CW
                                End If
                            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDDO3CW = "1" Then
                                    tblNukishi(i + 1).WFINDDO3CW = "2"
                                    tblNukishi(i + 1).WFSMPLIDDO3CW = tblNukishi(i).WFSMPLIDDO3CW
                                    tblNukishi(i + 1).WFRESDO3CW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDDO3CW = tblNukishi(i + 1).WFSMPLIDDO3CW
                                    tblNukishi(i).WFRESDO3CW = tblNukishi(i + 1).WFRESDO3CW
                                End If
                            ElseIf intZkbn = 0 Then  'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(14) = "0" Then
                            tblNukishi(i).WFINDDO3CW = "0"
                            tblNukishi(i).WFSMPLIDDO3CW = ""
                            tblNukishi(i).WFRESDO3CW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        ''�c���_�f�ǉ�
                        .col = 26
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                                If tblNukishi(i).WFINDAOICW = "1" Then
                                    tblNukishi(i + 1).WFINDAOICW = "2"
                                    tblNukishi(i + 1).WFSMPLIDAOICW = tblNukishi(i).WFSMPLIDAOICW
                                    tblNukishi(i + 1).WFRESAOICW = "1"
                                Else
                                    tblNukishi(i).WFSMPLIDAOICW = tblNukishi(i + 1).WFSMPLIDAOICW
                                    tblNukishi(i).WFRESAOICW = tblNukishi(i + 1).WFRESAOICW
                                End If
                            ElseIf intZkbn = 0 Then  'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
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
                                    '�R�s�[
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

                        '�G�s��s�]���ǉ��Ή�
                        .col = 35
                        .backColor = vbWhite
                        If .text <> "2" Then
                            If intZkbn = 3 Or intZkbn = 1 Then
                                .text = skensa2(17)
                                If .text = "1" Then
                                    tblNukishi(i + 1).WFINDOT2CW = "1"
                                    tblNukishi(i + 1).WFSMPLIDOT2CW = tblNukishi(i + 1).REPSMPLIDCW
                                    '�R�s�[
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

                        '�G�s��s�]���ǉ��Ή�
                        .col = 28
                        .backColor = vbWhite
                        If .text <> "2" And .text <> "" Then
                            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
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
                            ElseIf intZkbn = 0 Then  'Z�ł͂Ȃ�
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

                        '�ۏؕ��@�ύX�Ή�
                        If (intZkbn = 3 Or intZkbn = 1) And skensa1(18) = "0" Then
                            tblNukishi(i).WFINDGDCW = "0"
                            tblNukishi(i).WFSMPLIDGDCW = ""
                            tblNukishi(i).WFRESGDCW = "0"
                            tblNukishi(i).WFHSGDCW = "0"
                            .row = i
                            .text = ""
                            .row = i + 1
                        End If

                        '�G�s��s�]���ǉ��Ή�
                        ' VB��̐���(�v���V�[�W���e��64k����)�̂��ߤ�G�s���͕ʊ֐��ŏ�������
                        Call sub_DispSumple_Hanei_Ep_3(i, intNukisiRow, skensa1(), skensa2(), intZkbn)

                        ''���R�ŕK�����т𗧂Ă鏈����ǉ�
                        With tblNukishi(i)
                            If .WFINDRSCW = "2" Then
                                '�G�s��s�]���ǉ��Ή�
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
                                '�G�s��s�]���ǉ��Ή�
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

                        '���f�[�^�̃`�F�b�N
                         Call sub_Jitu(i + 1, blKirikaeflg, blnhflg, intZkbn)

                        '�F��h��
                        Call sub_Paint(i + 1)
                        '�F��h��
                End If
                    If i Mod 2 = 0 Then
                        ReDim Preserve gtSprWfMap(i + 1)
                        gtSprWfMap(i).ADD_FLG = 2
                        gtSprWfMap(i + 1).ADD_FLG = 0
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                        .GetText 39, i, vgetlotid
                        gtSprWfMap(i).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                        .GetText 39, i + 1, vGetLotId2
                        gtSprWfMap(i + 1).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vGetLotId2)

                        .GetText 4, i, vGetBlkP
                        gtSprWfMap(i).blockp = Trim(vGetBlkP)
                        gtSprWfMap(i + 1).blockp = Trim(vGetBlkP)
                        .GetText 4, i + 1, vGetBlkP
                        If vGetBlkP <> "" And vNukisiFlg <> 2 Then    '�ǉ��s�͏������Ȃ��@04/06/07 ooba
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
                        'TBCMY011����Y���f�[�^�擾
                        vSample1 = Trim(Right(sSampID1, 1))
                        vSample2 = Trim(Right(sSampID2, 1))
                        If i <= m - 2 Then
                            .GetText 4, i + 2, vNextBlkP
                            intNextBlkP = CInt(Trim(vNextBlkP))
                        End If

                        'WF�����A�T���v��ID�擾
                        If DBDRV_GET_WFMAP(gtSprWfMap(i).LOTID, tblSXL.SXLID, gtSprWfMap(i).blockp, _
                                            vGetBlkP, vGetIngotP, sNextIngotP, vGetBlkSeq, vGetBlkSeq2, _
                                            vSample1, vSample2, intNextBlkP, vGetWfNum) = FUNCTION_RETURN_FAILURE Then
                            lblMsg.Caption = GetMsgStr("EWFM1") '03/06/06 �㓡
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
                        '���L�A�T���v���̑I���\�̏ꍇ�ASampleID��ʉB���J�����ɕۑ�
                        '���L�̏ꍇ�ł��T���v��ID����\�T���v��ID�̂Ƃ��ɏ���������悤�ɕύX�@2003/11/11
                        If blKirikaeflg = True And intSmpkbn <> 1 Then
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .SetText 44, i, vSample1
                            gtSprWfMap(i).SAMPLEID = vSample1
                            .SetText 44, i, vSample1
                            gtSprWfMap(i + 1).SAMPLEID = vSample1
                        End If
                    End If
                Else
                    '�敪�R�Ŕ����𔺂�Ȃ���������
                    If i Mod 2 = 0 Then
                        ReDim Preserve gtSprWfMap(i + 1)
                        gtSprWfMap(i).ADD_FLG = 3
                        gtSprWfMap(i + 1).ADD_FLG = 3
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                        .GetText 39, i, vgetlotid
                        gtSprWfMap(i).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                         .GetText 39, i + 1, vgetlotid
                        gtSprWfMap(i + 1).LOTID = Mid(tblSXL.SXLID, 1, 9) & Trim(vgetlotid)
                        .GetText 4, i, vGetBlkP
                        gtSprWfMap(i).blockp = Trim(vGetBlkP)

                        .GetText 4, i + 1, vGetBlkP
                        gtSprWfMap(i + 1).blockp = Trim(vGetBlkP)
                        '�����\���̌����ʒu,�A�Ԃ͉�ʂ���@2003/04/23
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
                intNukisiRow = intNukisiRow + 1 '1�s�ځ��ŏI�s�ŁAINDEX+1���Ă���
                ReDim Preserve gtSprWfMap(i)
                If vNukisiFlg = 1 Then
                    gtSprWfMap(i).ADD_FLG = 1
                Else
                    gtSprWfMap(i).ADD_FLG = 3
                End If
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

                ''�S�U�֎��̌���GD���p���Ή�
                .row = i
                .col = 28
                'GD�̎w�����Ȃ��ꍇ
                If .text <> "1" And .text <> "2" Then
                    With tblWafInd(intNukisiRow)
                        '�����s��TOP���ŕi�Ԃ�U�ւ����ꍇ
                        If i = 1 And Trim(.HINDN.hinban) <> "Z" And _
                                    (.HINDN.hinban <> tblSXL.hinban Or _
                                    .HINDN.mnorevno <> tblSXL.REVNUM Or _
                                    .HINDN.factory <> tblSXL.factory Or _
                                    .HINDN.opecond <> tblSXL.opecond) Then

                            .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)       'GD
                            'GD�̎w��������ꍇ
                            If .SMP.CRYINDGD <> "0" And .SMP.CRYINDGD <> "2" And _
                                        IsNumeric(CpyCrySmpl.TsmplidGD) Then

                                .SMP.WFHSGD = "1"
                                sprExamine.text = CpyCrySmpl.TindGD
                                sprExamine.backColor = COLOR_CryJitsu
                                sprExamine.ForeColor = COLOR_CryJitsu
                                bMotoGDcpyFlg(1) = True
                            End If

                        '�����s��BOT���ŕi�Ԃ�U�ւ����ꍇ
                        ElseIf i = m And Trim(.HINUP.hinban) <> "Z" And _
                                    (.HINUP.hinban <> tblSXL.hinban Or _
                                    .HINUP.mnorevno <> tblSXL.REVNUM Or _
                                    .HINUP.factory <> tblSXL.factory Or _
                                    .HINUP.opecond <> tblSXL.opecond) Then

                            .SMP.CRYINDGD = GetWFSamp(.HINUP, .HINDN, 19)       'GD
                            'GD�̎w��������ꍇ
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

        '��ʕ\��
        m = .MaxRows
        intNukisiRow = 0
        For i = 1 To m
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
            .GetText 37, i, vNukisiFlg
            If CheckGetSampleID(i) = True Then  '3�̏�Ƌ����s
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
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .SetText 44, i, gtSprWfMap(i).SAMPLEID
                    If blKirikaeflg = True And intSmpkbn <> 1 Then  '���ʂ̂Ƃ�����
                        If i Mod 2 = 1 Then
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .SetText 38, i, gtSprWfMap(i - 1).SAMPLEID
                        Else
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .SetText 38, i, gtSprWfMap(i + 1).SAMPLEID
                        End If
                    End If
                Else
                    .SetText 10, i, Right(gtSprWfMap(i).LOTID, 3) & "-" & Right(gtSprWfMap(i).SAMPLEID, 4)
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .SetText 38, i, gtSprWfMap(i).SAMPLEID
                    .SetText 8, i, gsWF_STA_SIJI
                    .SetText 44, i, gtSprWfMap(i).SAMPLEID
                End If

                blKensaLock = False '�������`�F�b�N
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                For intLoopCnt = 11 To 35
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    If intLoopCnt <> 27 And intLoopCnt <> 35 Then
                        .col = intLoopCnt
                        .row = i + 1
                        If .text = "1" Then '����������
                            blKensaLock = True
                        End If
                    End If
                Next
                If blKirikaeflg = True And intSmpkbn <> 1 Then       '����
                    .SetText 10, i + 1, gsWF_SMPL_JOINT
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                Else
                    If blKensaLock = False Then  '��������
                        .SetText 10, i + 1, Right(gtSprWfMap(i + 1).LOTID, 3) & "-" & Right(gtSprWfMap(i + 1).SAMPLEID, 4)
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                        .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 44, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 8, i + 1, gsWF_STA_SIJI
                    Else                            '�����L��
                        .SetText 10, i + 1, Right(gtSprWfMap(i + 1).LOTID, 3) & "-" & Right(gtSprWfMap(i + 1).SAMPLEID, 4)
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                        .SetText 38, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 44, i + 1, gtSprWfMap(i + 1).SAMPLEID
                        .SetText 8, i + 1, gsWF_STA_SIJI
                    End If
                End If
            End If
        Next

        'WF�����v�Z
        For i = 1 To .MaxRows - 2
            If i Mod 2 = 0 Then

                '����0���̓u���b�N���E �܂���Z�i��
                If CInt(Trim$(gtSprWfMap(i).wfnum)) = 0 Then
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .GetText 43, i - 1, vAllNum
                    If i < .MaxRows Then
                        .GetText 1, i + 1, vBlockId
                        If vBlockId = "" And vAllNum <> "" Then
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                    If CInt(gtSprWfMap(i).wfnum) <> 0 Then  '�u���b�N���E�ւ̑Ώ�
                        .SetText 7, i + 1, Trim$(gtSprWfMap(i).wfnum)
                    End If
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .GetText 43, i - 1, vAllNum
                    .GetText 1, intCngBlpnt, vBlockId
                    If vBlockId = Mid(gtSprWfMap(i).LOTID, 10, 3) Then
                        If (gtSprWfMap(i).LOTID) = (gtSprWfMap(i + 1).LOTID) Then '�u���b�N���E�ւ̑Ώ�    'upd 2003/04/28 hitec)matsumoto WF���0�̎��́A�u���b�N�������łȂ�
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .GetText 43, i - 1, vAllNum
                            .SetText 7, intCngBlpnt, vAllNum
                            intCngBlpnt = i + 1
                        End If

                        '���̃u���b�N�̌����ʒu�͌�̃f�[�^�W�J�̂��ߓ����ʒu�ɂ��Ă����i4/23���݁j
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
                    .text = "����"
                End If
            ElseIf i = .MaxRows Then
                .GetText 2, i - 1, vGetHinban
                If vGetHinban = "Z" Then
                    .col = 8
                    .row = i - 1
                    .text = "����"
                End If
            Else
                If i Mod 2 = 1 And i <> .MaxRows - 1 And i <> .MaxRows Then
                    .GetText 2, i, vGetHinban
                    If vGetHinban = "Z" Then
                        .col = 8
                        .row = i
                        .text = "����"
                        .row = i + 1
                        .text = "����"
                    End If
                End If
            End If
        Next i

        vKeturaku = "����"
        vNull = ""
        vZERO = "0"
        For i = 1 To .MaxRows Step 2
            .GetText 2, i, vGetHinban
            If vGetHinban = "Z" Then
                .SetText 7, i, vZERO
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                .GetText 37, i, vNukisiFlg
                If vNukisiFlg <> "1" Then
                    .SetText 10, i, vNull
                    For intLoopCnt = i To 1 Step -1  'add 2003/06/11 hitec)matsumoto z�i�Ԏ��̃T���v��ID�o�b�N�A�b�v
                        If intLoopCnt = 1 Then
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
                            .SetText 38, i, vGetToSamp
                            Exit For
                        End If
                    Next
                End If

                ' 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)kondoh
                .GetText 37, i + 1, vNukisiFlg

                If vNukisiFlg <> "1" Then
                    .SetText 10, i + 1, vNull
                    For intLoopCnt = i + 1 To .MaxRows Step 1  'add 2003/06/11 hitec)matsumoto z�i�Ԏ��̃T���v��ID�o�b�N�A�b�v
                        If intLoopCnt = .MaxRows Then
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                            .GetText 38, intLoopCnt, vGetToSamp
                            .SetText 38, i + 1, vGetToSamp
                            Exit For
                        End If
                    Next
                End If

                '�����\���s�i1,�ŏI�s�j�̃T���v��ID�͏������Ȃ�
                .col = 10
                .row = i
                .ForeColor = &H8080FF
                .col = 10
                .row = i + 1
                .ForeColor = &H8080FF
                ''Z�i�Ԃ�ԕ\���ɂ���
                .row = i + 1

                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
               ''�������ڂ̕�����Ԃɂ���
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

                '�u���b�N�o�̗�
                .row = i + 1
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
        .row = intRowNum - 1  '' 04/24 �㓡
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

'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
        'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
        For i = 1 To .MaxRows Step 2
            '�i�Ԏ擾
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .col = 10
                'WF��Ԏ擾(TOP)
                .GetText 8, i, vTemp
                'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
                If Trim(vTemp) = gsWF_STA_SIJI_KEKKA Then
                    .row = i
                    .backColor = f_cmbc039_3.Label12.backColor
                End If
                'WF��Ԏ擾(BOTTOM)
                .GetText 8, i + 1, vTemp
                'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
                If Trim(vTemp) = gsWF_STA_SIJI_KEKKA Then
                    .row = i + 1
                    .backColor = f_cmbc039_3.Label12.backColor
                End If

            End If
        Next i
'���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------

        'add start 2003/04/28 hitec)matsumoto�@�������Ă���WF���d�����Ă��Ȃ����`�F�b�N���A�d�����Ă���ꍇ�̓G���[���b�Z�[�W��\��������----------------------
        For i = 2 To .MaxRows Step 2    '�����s�̂݃��[�v����
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
            .GetText 37, i, vNukisiFlg
            If vNukisiFlg = "2" Then                '�ǉ������s�̏ꍇ
                .GetText 6, i, vGetUpWf    '�}�b�v�ʒu�擾
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                .GetText 39, i, vGetUpBlk  '�u���b�NID
                .GetText 8, i - 1, vGetWfNum   '����
                If vGetWfNum = 0 Then
                    If vGetHinban <> "Z" Then
                        cmdF(12).Enabled = False
                        bSampFlag = False
                        lblMsg.Caption = GetMsgStr("EWFM2") '03/06/06 �㓡
                        Exit Function
                    End If
                End If

                For intWfChkLoop = i + 2 To .MaxRows Step 2   '�d���`�F�b�N�̂��߁A���̋����s������
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    .GetText 37, intWfChkLoop, vNukisiFlg

                    If vNukisiFlg = "2" Then                '�ǉ������s�̏ꍇ
                        .GetText 6, intWfChkLoop, vGetDnWf    '�}�b�v�ʒu�擾
                        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                        .GetText 39, intWfChkLoop, vGetDnBlk  '�u���b�NID
                        If Trim(vGetUpWf) = Trim(vGetDnWf) Then     '�}�b�v�ʒu���A�ŏ��Ɏ擾�����l�Ɠ�����������A
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

        ''Warp����Ή�
        'WFϯ�ߏ�̕i�ԏ��擾
        ReDim tMapHin(0)
        m = 0
        For i = 1 To .MaxRows Step 2
            '�i�Ԏ擾
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                If GetLastHinban(CStr(vGetHinban), udtFullHinban) = FUNCTION_RETURN_FAILURE Then
                    lblMsg.Caption = "�i�ԓ��̓G���["
                    Exit Function
                End If

                m = m + 1
                ReDim Preserve tMapHin(m)

                tMapHin(m).HIN = udtFullHinban

                '��ۯ�ID�擾
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                .GetText 39, i, vBlockId
                tMapHin(m).BLOCKID = left(txtCryNum.text, 9) & CStr(vBlockId)
                '��ۯ�SEQ�擾
                .GetText 6, i, vTemp
                .GetText 6, i + 1, vTemp1
                tMapHin(m).BLKSEQ_S = CInt(vTemp)
                tMapHin(m).BLKSEQ_E = CInt(vTemp1)
                '�U�������׸�
                tMapHin(m).WARPFLG = False
                tMapHin(m).KAKUFLG = False

                'Add Start 2011/04/25 SMPK Miyata
                tMapHin(m).XTALCS = txtCryNum.text      '�����ԍ�
                .GetText 5, i, vTemp
                .GetText 5, i + 1, vTemp1
                tMapHin(m).INPOSCS_S = CInt(vTemp)      '�������ʒu(Start)
                tMapHin(m).INPOSCS_E = CInt(vTemp1)     '�������ʒu(End)
                'Add End   2011/04/25 SMPK Miyata

            End If
        Next i
    End With

    '' ���i�d�l�̕\��
    cmdF(12).Enabled = False
    bSampFlag = False

    '�\���̂��Ƃ��Ă���
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
    'Add Start 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�

    '�U�փ`�F�b�N-------start iida 2003/09/29 ���ړ� cmdF_Click
    If fnc_Furikae_Check = False Then
         bSampFlag = False

        '�U�����������{�i�Ԃ�Warp/�����p�x����@06/01/12 ooba START ========================>
        For i = 1 To UBound(tMapHin)
            '�U���������{�̊m�F
            tMapHinG = tMapHin(i)
            For j = 1 To 2
                If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                    m = funChkFurikaeShiyou("CW763", txtKSXLID.text, tMapHinG.HIN, _
                                            tMapHinG.HIN, intModori, sMsg, _
                                            typ_b, typ_CType, 0)

                    tMapHin(i).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                    tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                End If
            Next j
        Next i

        'Warp/�����p�x���\��
        Call WarpKakuDisp(Me)
        Exit Function
    End If

    '>>>>> MOD 2012/09/07 SETsw Marushita 10�����������Ή�
    ' �u���b�N�ۏ�`�F�b�N����(�v���V�[�W���T�C�Y�G���[�̂��ߕ���)
    If fnc_BlockHCheck() < 0 Then
        Exit Function
    End If

    'Warp/�����p�x���\���@06/01/12 ooba
    Call WarpKakuDisp(Me)
    
    '�U�փ`�F�b�N-------end iida 2003/09/29
    cmdF(12).Enabled = True
    bSampFlag = True

    sub_DispSample = True
End Function

'*******************************************************************************
'*    �֐���        : sub_BlockHCheck
'*
'*    �����T�v      : �u���b�N�ۏ�`�F�b�N
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*
'*    �߂�l        : Integer
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
    Dim sC_Flg       As String   '�`�F�b�N�t���O�擾�p�@�@�@2012/09/07 SETsw Marushita
    Dim sC_Mai       As String   '�`�F�b�NWafer�����擾�p �@2012/09/07 SETsw Marushita
    Dim iRtn         As Integer  '�����m�F�p�@2012/09/07 SETsw Marushita

    fnc_BlockHCheck = -1

'** ��ۯ��ۏ����� *********************** 2008.03.20 aoyagi *************
    With sprExamine

    flg = 0
    old_bid = ""

    '�S�s�����[�v-------
    For i = 1 To .MaxRows
        .row = i
        .col = 39
        now_bid = .text
        If now_bid <> old_bid Then  '�u���b�NID�ς��---
            old_bid = now_bid
            '�L����SXL��--------
            iJCnt1 = sub_DispSample_SCnt02(i)
            If iJCnt1 <= 1 Then   '1��ۯ�=1SXL�Ȃ̂���ۯ��ۏ��������Ȃ�
                '1��ۯ��̍s��--------
                iJCnt1 = sub_DispSample_SCnt03(i)
                i = i + iJCnt1 - 1  '���ɐi�߂�
            Else
                '��ۯ��ۏ�����-------
                iJCnt1 = sub_DispSample_SCnt01(i)
                If iJCnt1 <= 0 Then  '����قȂ�
                    flg = 1          'NG
                    Exit For
                End If
            End If
        Else
            '��ۯ��ۏ�����-------
            iJCnt1 = sub_DispSample_SCnt01(i)
            If iJCnt1 <= 0 Then  '����قȂ�
                flg = 1          'NG
                Exit For
            End If
        End If
    Next i
    
    If flg = 1 Then
        lblMsg.Caption = "WF�T���v�������͕����ł��܂���B"
        cmdF(12).Enabled = False
    
        Exit Function
    End If
    
    '>>>>> ADD 2012/09/07 SETsw Marushita 10�����������Ή�
    '�`�F�b�N�t���O�̎擾(1:�G���[�`�F�b�N�A2:�A���[���`�F�b�N)
    sC_Flg = GetCodeA9Field("X", "19", "NUKIMAI", "KCODE01A9")
    '�`�F�b�NWafer�����̎擾
    sC_Mai = GetCodeA9Field("X", "19", "NUKIMAI", "CTR01A9")
    '<<<<< ADD 2012/09/07 SETsw Marushita 10�����������Ή�

'Cng Start 2012/02/23 Y.Hitomi
    '�d�|�iWafer10������
    m = .MaxRows
    For i = 1 To m
        .row = i
        .col = 7

        If i Mod 2 <> 0 Then
            .GetText 2, i, vGetHinban
            .GetText 7, i, vGetWfNum
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .backColor = &H80FF80
                '>>>>> Mod 2012/09/07 SETsw Marushita 10���ȉ����b�g�����Ή�
                'Wafer�����̃`�F�b�N
                If CInt(vGetWfNum) < CInt(sC_Mai) Then
                    If Trim(sC_Flg) = "1" Then
                        .backColor = &H8080FF
                        lblMsg.Caption = "Wafer������" & sC_Mai & "�������ׁ̈A���s�ł��܂���B"
                        Exit Function
                    Else
                        If Trim(sC_Flg) = "2" Then
                            iRtn = MsgBox("Wafer������" & sC_Mai & "�������ł��B���s���Ă���낵���ł����H", vbQuestion + vbOKCancel, "�m�F")
                            '�L�����Z���̏ꍇ�͏����𔲂���
                            If iRtn = vbCancel Then
                                Exit Function
                            End If
                        End If
                    End If
                End If
                'If CInt(vGetWfNum) < 11 Then
                '    .backColor = &H8080FF
                '    lblMsg.Caption = "Wafer������10���ȉ��ׁ̈A���s�ł��܂���B"
                '    Exit Function
                'End If
                '<<<<< Mod 2012/09/07 SETsw Marushita 10���ȉ����b�g�����Ή�
            End If
        End If
    Next
'''Add Start SPK 2009/09/14
'    '�d�|�iWafer0������
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
'                    lblMsg.Caption = "Wafer������0���ł��I"
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
'*    �֐���        : sub_Hanei
'*
'*    �����T�v      : 1.���f����`�F�b�N
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    i�@�@       ,I  ,Integer  ,Row
'*                    j�@�@       ,I  ,Integer  ,tblWafind
'*                    i�@�@       ,I  ,Integer  ,Z�敪
'*                    pSird1stBlockSet�@,IO  ,Boolean  ,������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP START  (���Ұ��ǉ��F������SIRD����َw���ݒ�L��)
'''Private Sub sub_Hanei(i As Integer, j As Integer, intZkbn As Integer)
Private Sub sub_Hanei(i As Integer, j As Integer, intZkbn As Integer, ByRef pSird1stBlockSet As Boolean)
'��--- 2010/01/20 SIRD�Ή� SPK habuki REP END
    'i�FRow�Aj:tblWafind
    Dim k               As Integer '����
    Dim sKensa()        As String
    Dim sSampID()       As String
    Dim sJflg()         As String
    Dim s               As Integer
    Dim sTB             As String
    Dim udtFullHinban   As tFullHinban
    Dim sGetSmpllid1    As String
    Dim sGetSmpllid2    As String '���f�Ȃ̂Ŏg�p���Ȃ�
    Dim intHanSuiKBN    As Integer
    Dim intSmpPos       As Integer
    Dim intModori       As Integer
    Dim sKRs            As String
    Dim sKensac()       As String
    Dim sGetHS          As String        '�ۏ��׸ށ@05/02/18 ooba
    Dim sGDChk          As String
    Dim sHanFlg         As String
    
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
    Dim SirdYN          As Boolean          'SIRD����َw���L��[True:��񂠂�AFalse:��񖳂�]
    Dim sirdSmpID       As String           '�����ςݻ����ID�iSIRD�p�j<TBCMJ022>
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END


    sGetHS = "0"        '05/02/24 ooba

    'GDײ������@�\�ǉ�
    sHanFlg = "0"

    CrySampleID = CpyCrySmpl        '�������ш��p���ް��̺�߰�@05/06/13 ooba

    For s = i To i + 1 'Row
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
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
            sKensac(14) = .SMP.CRYINDAO       '�c���_�f�ǉ��@03/12/15 ooba
            sKensac(15) = .SMP.CRYINDGD       'GD�ǉ��@05/02/18 ooba

            '--- 2006/08/15 Add �G�s��s�]���ǉ��Ή�
            sKensac(16) = .SMP.EPIINDB1
            sKensac(17) = .SMP.EPIINDB2
            sKensac(18) = .SMP.EPIINDB3
            sKensac(19) = .SMP.EPIINDL1
            sKensac(20) = .SMP.EPIINDL2
            sKensac(21) = .SMP.EPIINDL3

            'GDײ������@�\�ǉ� 2007/06/25 M.Kaga START   ---
            sGDChk = .SMP.CRYINDGD2
        End With

        If s = i Then
        '�i�Ԃ�Z�̂Ƃ�
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
                udtFullHinban.factory = tblWafInd(j).HINUP.factory '�i��
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
                sTB = "T"  'TB�敪
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
        '1�FD�A2�FU�A3UD 2�s���̔��������

        '--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
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
                If (intZkbn = 3 And s = i) Or (intZkbn = 1 And s = i) Or (intZkbn = 4 And s = i + 1) Or (intZkbn = 2 And s = i + 1) Then   'Z����̂Ƃ��Ɖ��̂Ƃ��͔��f:2�����鋤�ʊ֐��Ă΂Ȃ�
                   sKensa(k) = "2"
                   sJflg(k) = "1"
                   sGetHS = "0"     '05/02/18 ooba
                ElseIf sKensa(k) <> "0" Then
                    'GD�ǉ��ɂ��ύX�@05/02/17 ooba
                    If k >= 14 Then k = k + 2
                    '���f���苤�ʊ֐��̌Ăяo��
                    intModori = funChkWfHanSui(tblSXL.SXLID, sTB, tblSXL.CRYNUM, udtFullHinban, intSmpPos, k + 2, SIngotP, EIngotP, intHanSuiKBN, sGetSmpllid1, sGetSmpllid2, sGetHS)
                    '�߂�l(0:����I��(���f/����OK)�A1�F����I��(���f/����NG)�A-1�F���͈����l�G���[�A-2�F����ȊO�̃G���[)
                    If k >= 16 Then k = k - 2

                    If intModori = 1 Then  '���fNG
                        sKensa(k) = "1"
                        sSampID(k) = tblNukishi(s).REPSMPLIDCW '�������ڂɑ�\�T���v��ID������
                        sGetHS = "0"    '05/02/18 ooba
                    ElseIf intModori = 0 Then
                        If intHanSuiKBN = 0 Then '���fOK
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
'--- 2006/08/15 Add �G�s��s�]���ǉ��Ή� SMP)kondoh
                If k < 15 Then .col = 12 + k Else .col = 13 + k
                'GDײ������̏ꍇ
                If .col = 28 Then
                    If sHanFlg = "0" Then
                        '��i�ԂɎ���
                        If Trim(sGDChk) = "2" Then
                            sKensa(k) = "1"
                        ElseIf Trim(sGDChk) = "1" Then
                            sKensa(k) = "2"
                        End If
                        .text = sKensa(k) '���ʂ��X�v���b�h�ɃZ�b�g
                        sHanFlg = "1"
                    ElseIf sHanFlg = "1" Then
                        '���i�ԂɎ���
                        If Trim(sGDChk) = "2" Then
                            sKensa(k) = "2"
                        ElseIf Trim(sGDChk) = "1" Then
                            sKensa(k) = "1"
                        End If
                        .text = sKensa(k) '���ʂ��X�v���b�h�ɃZ�b�g
                        sHanFlg = "0"
                    End If
                Else
                    '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
'''                    .text = sKensa(k) '���ʂ��X�v���b�h�ɃZ�b�g
                    
                    If .col = 19 Then
                        '<< SIRD >>
                        If (sKensa(k) = "1") Or (sKensa(k) = "2") Or (sKensa(k) = "3") Or (sKensa(k) = "4") Then
                            '<TBCMJ022����>
                            Call fncGetSirdSample(Trim(txtCryNum.text), SirdYN, sirdSmpID)
                            If SirdYN Then
                                '<������ 1st block �ȍ~>
                                .text = "2"          '���L
                                pSird1stBlockSet = True             '������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
                            Else
                                If Not pSird1stBlockSet Then
                                    '<������ 1st block>
                                    .text = "1"      '�擾
                                    pSird1stBlockSet = True         '������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
                                Else
                                    '<������ 1st block �ȍ~>
                                    .text = "2"      '���L
                                    pSird1stBlockSet = True         '������SIRD����َw���ݒ�L��[True:�ݒ�ς݁AFalse:���ݒ�]
                                End If
                            End If
                        Else
'''                            .text = sKensa(k) '���ʂ��X�v���b�h�ɃZ�b�g
                            .text = ""
                        End If
                    Else
                        .text = sKensa(k) '���ʂ��X�v���b�h�ɃZ�b�g
                    End If
                    '��--- 2010/01/20 SIRD�Ή� SPK habuki REP END
                End If
            End With

            With tblNukishi(s) '�\���̂ɃZ�b�g
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
                    Case 14     ''�c���_�f�ǉ��@03/12/15 ooba
                    .WFINDAOICW = IIf(sKensa(k) = "", "0", sKensa(k))
                    .WFSMPLIDAOICW = sSampID(k)
                    .WFRESAOICW = sJflg(k)
                    Case 15     'GD�ǉ��@05/02/18 ooba
                    .WFINDGDCW = IIf(sKensa(k) = "", "0", sKensa(k))    '���FLG(GD)
'                    .WFSMPLIDGDCW = sSampID(k)                          '�����ID(GD)
                    '09/07/22 Y.Hitomi GD����َw���s��Ή���GD�̂�,WF����َw���L�莞,��������ٔ��fNG
                        If .WFINDGDCW <> "1" Then
                            .WFSMPLIDGDCW = sSampID(k)                 '�����ID(GD)
                        ElseIf .WFINDGDCW = "1" Then
                            .WFSMPLIDGDCW = tblNukishi(s).REPSMPLIDCW '�������ڂɑ�\�T���v��ID������
                        End If
                    
                    .WFRESGDCW = sJflg(k)                               '����FLG(GD)
                    .WFHSGDCW = sGetHS                                  '�ۏ�FLG(GD)

                    '�G�s��s�]���ǉ��Ή�
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
'*    �֐���        : sub_Betu
'*
'*    �����T�v      : 1.���L�ʂ̏ꍇ�̐ݒ肷��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    i�@�@       ,I  ,Integer  ,Row
'*                    j�@�@       ,I  ,Integer  ,tblWafind
'*                    sKensa1�@ ,I  ,String   ,�����p
'*                    sKensa2�@ ,I  ,String�@ ,�����p
'*                    intZkbn�@�@ ,I  ,Integer  ,Z�敪
'*                    blnhflg�@�@ ,I  ,Boolean  ,���f����t���O
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_Betu(i As Integer, j As Integer, skensa1() As String, skensa2() As String, intZkbn As Integer, blnhflg As Boolean)
    Dim k           As Integer '����
'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
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
        sKensa(15) = .SMP.CRYINDAO        '�c���_�f�ǉ��@03/12/15 ooba
        sKensa(16) = .SMP.CRYOTHER1
        sKensa(17) = .SMP.CRYOTHER2
        sKensa(18) = .SMP.CRYINDGD        'GD�ǉ��@05/02/18 ooba

        '�G�s��s�]���ǉ��Ή�
        sKensa(19) = .SMP.EPIINDB1        'BMD1E
        sKensa(20) = .SMP.EPIINDB2        'BMD2E
        sKensa(21) = .SMP.EPIINDB3        'BMD3E
        sKensa(22) = .SMP.EPIINDL1        'OSF1E
        sKensa(23) = .SMP.EPIINDL2        'OSF2E
        sKensa(24) = .SMP.EPIINDL3        'OSF3E
    End With

'--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
    For k = 0 To 24
        Select Case sKensa(k)
            Case 1 'Top
               If intZkbn = 2 Or intZkbn = 4 Then
                   '���i��Z�Ȃ猟���Ȃ��@04/04/14 ooba
                   skensa1(k) = "0"
                   skensa2(k) = "0"
               Else
                   skensa1(k) = "0"
                   skensa2(k) = "1"
                   k2 = k2 + 1
               End If
            Case 2 'Tail
               If intZkbn = 1 Or intZkbn = 3 Then
                   '��i��Z�Ȃ猟���Ȃ��@04/04/14 ooba
                   skensa1(k) = "0"
                   skensa2(k) = "0"
               Else
                   skensa1(k) = "1"
                   skensa2(k) = "0"
                   k1 = k1 + 1
               End If
            Case 3 '����
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

            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            '2.1.2 ���L�`�F�b�N�ǉ�
            '��R��case 4 �̎��͖������Ŏ��т𗧂Ă�
            Case 4 '��������(��)
'Cng Start 2011/11/02 Y.Hitomi
                ''��R�̎������������s��
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
            '���ǉ� �M�������f�����ǉ� 2006/02/15 SMP�ΐ� ---------------
            Case 0  '����
               skensa1(k) = ""
               skensa2(k) = ""
        End Select
    Next k

    If intZkbn = 0 Then 'Z�ł͂Ȃ��Ƃ�
        blnhflg = False

        '�G�s��s�]���ǉ��Ή� SMP)kondoh
        For k = 0 To 24
            If skensa1(k) = "1" And skensa2(k) = "1" Then '���L�̂Ƃ�����2�ɂ���
                If k1 < k2 Then '�������ڂ����Ȃ��ق������f�ƂȂ�
                    skensa1(k) = IIf(skensa1(k) = "1", "2", skensa1(k))
                    blnhflg = False '�㋤�L

                Else
                    skensa2(k) = IIf(skensa2(k) = "1", "2", skensa2(k))
                    blnhflg = True  '�����L
                End If
            End If
        Next k
    End If
End Sub

'*******************************************************************************
'*    �֐���        : sub_Paint
'*
'*    �����T�v      : 1.�������ڂ̃X�v���b�h�̏�Ԃɂ��F����������
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    intRow        ,I  ,Integer  ,Row
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_Paint(intRow As Integer)
    Dim sTval       As String
    Dim sKval       As String
    Dim intKensa    As Integer
    Dim i           As Integer '��
    Dim k           As Integer '�s---2�s��F�h��

    '�X�v���b�h�̌�������(""�F���A1�F���A2�F���F)
    intKensa = 0
        With sprExamine
        For k = intRow - 1 To intRow
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
'''                        .backColor = vbYellow
'''                        .ForeColor = vbYellow
'''                        .Lock = False
                        
                        If i = 19 Then
                            '<SIRD>�E�E�E��ڲ
                            .backColor = COLOR_CryJitsu
                            .ForeColor = COLOR_CryJitsu
                            .Lock = False
                        Else
                            '<SIRD�ȊO>�E�E�E���F
                            .backColor = vbYellow
                            .ForeColor = vbYellow
                            .Lock = False
                        End If
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
                    End If
                End If
            Next i

            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
            .col = 28
            .row = k
            sKval = .text
            '�w�����̏ꍇ�͔��\��
            If sKval = "" Then
                .backColor = vbWhite
                .ForeColor = vbWhite
                .Lock = True
            '�����̏ꍇ�͍��\��
            ElseIf sKval = "1" Then
                .backColor = vbBlack
                .ForeColor = vbBlack
                .Lock = True
            '���f�̏ꍇ�͉��F�\��
            ElseIf sKval = "2" And tblNukishi(k).WFHSGDCW <> "1" Then
                .backColor = vbYellow
                .ForeColor = vbYellow
                .Lock = False
            '�������т̏ꍇ�͸�ڰ�\��
            ElseIf sKval = "2" And tblNukishi(k).WFHSGDCW = "1" Then
                .backColor = COLOR_CryJitsu
                .ForeColor = COLOR_CryJitsu
                .Lock = False
            End If
        Next k
    End With
End Sub

'*******************************************************************************************
'*    �֐���        : sub_Jitu
'*
'*    �����T�v      : 1.�������s���Ĕ��f�ƂȂ����s�Ŏ��f�[�^��1���Ȃ��s��T��(���L�͏���)
'*                      (���f�[�^�����݂��邩�̃`�F�b�N������)
'*                    2.���f�[�^�����̏ꍇ���̍s�̒�R(Rs)�Ɏ��f�[�^�𗧂Ă�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^       ,����
'*                    k           ,I  ,Integer  ,Row
'*                    blKirikaeflg ,I  ,Boolean  ,�ؑւ��t���O
'*                    blnhflg�@�@ ,I  ,Boolean  ,���f����t���O
'*                    intZkbn�@�@ ,I  ,Integer  ,Z�敪
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_Jitu(k As Integer, blKirikaeflg As Boolean, blnhflg As Boolean, intZkbn As Integer)
    Dim sTval       As String
    Dim sKval       As String
    Dim intKensa    As Integer
    Dim j           As Integer
    Dim i           As Integer
    Dim blTwoFlg    As Boolean     '2004/01/29 ooba

    '�@�������s���Ĕ��f�ƂȂ����s�Ŏ��f�[�^��1���Ȃ��s��T��(���L�͏���)
    '�A���f�[�^�����̏ꍇ���̍s�̒�R(Rs)�Ɏ��f�[�^�𗧂Ă�
    '���f�[�^�̃`�F�b�N-------start iida 2003/09/12

    For j = k - 1 To k
        blTwoFlg = False     '2004/01/29 ooba

        If blKirikaeflg = True Then
            If j = k Then
                Exit For
            End If
        'Z�̂Ƃ�
        ElseIf intZkbn = 3 Or intZkbn = 4 Or intZkbn = 1 Or intZkbn = 2 Then
            If (intZkbn = 3 And j = k - 1) Or (intZkbn = 1 And j = k - 1) Then
                j = j + 1  '��Ɏ������ĂȂ�
            ElseIf (intZkbn = 4 And j = k) Or (intZkbn = 4 And j = k) Then
                Exit For   '���Ɏ������ĂȂ�
            End If
        '���L�ʂ̂Ƃ�
        Else
            intKensa = 0    '2004/01/28 ooba
            blTwoFlg = True  '2004/01/29 ooba
        End If

        With sprExamine
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                For i = 11 To 35
                    .col = i
                    .row = j
                    sKval = .text

                    If sKval = "1" Then   '��������(����)
                        intKensa = intKensa + 1
                    End If
                    '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                    If i = 35 Then
                        If intKensa = 0 Then
                            .col = 11
                            sKval = .text             '2004/01/28 ooba
                            If sKval = "2" Then       '2004/01/28 ooba
                                .text = "1"
                                .Lock = True
                                '�T���v��ID���\�T���v��ID�ɂ���
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
                                    '2�s��������ꍇ�͊Y���s�̂ݕύX�@2004/01/29 ooba
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
'*    �֐���        : sub_DrawImage
'*
'*    �����T�v      : 1.�����C���[�W�E�B���h�E��\������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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

'>>>>> ��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota ---------------------
'    If Mid(tblSXL.CRYNUM, 1, 1) = "8" Then
'        '' �G���[���b�Z�[�W��\������
'        lblMsg.Caption = GetMsgStr("EKDE1")
''        Screen.MousePointer = 1
'        Exit Sub
'    End If
'<<<<< ��8���w���P�����������Ȃ� 2007/10/10 SETsw kubota ---------------------

'2001/08/30 S.Sano Start
    '' �i�ԃ`�F�b�N
    For c0 = 1 To sprExamine.MaxRows
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg

        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

    '' �����w���ꗗ�̓��̓`�F�b�N�iWFϯ�߁j
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

'2003/03/18 hitec)matsumoto �i�Ԃ̂m�t�k�k�`�F�b�N�͂��łɂ͂����Ă���B�����łm�t�k�k���͂����Ă���̂́A��ʃ��C�A�E�g�ύX�ɂ��A�m�t�k�k�s(�i�ԓ��͕s��)�����݂��邽�߁B
'2001/08/30 S.Sano Start
    '' �i�ԃ`�F�b�N
    For c0 = 1 To sprExamine.MaxRows - 1
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
        sprExamine.GetText 37, c0, vNukisiFlg
        If (vNukisiFlg = "1") Or (vNukisiFlg = "2") Then    '�����s��������
            If (vNukisiFlg = "1") Then
                '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
                sprExamine.GetText 37, c0 + 1, vNukisiFlg
                If vNukisiFlg = "1" Then
                    '�ǉ������̖����s�͉������Ȃ�
                ElseIf vNukisiFlg = "2" Then
                    sprExamine.GetText 2, c0, vTemp             '�e�s�ɕi�Ԃ������Ă���킯�ł͂Ȃ��A�_�~�[�ł͑S�s�ɕi�Ԃ������Ă���̂ł��������
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

    '' �����N���X����u���b�N�����擾���邽�߂�Getxl�֐��Ăяo��
    Set Xl = GetXl(tblSXL.CRYNUM, "f_cmbc039_3")

    With Xl
        Call sub_MakeTBCME042
        '' �Ĕ����ʒu
        m = UBound(tblWafInd) - 1
        For i = 1 To m
            '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SAMPO)sekine
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

    '' �����}��`�悷��
    f_cmzc003a.Draw Xl  '' �����̏���`�悷��
    f_cmzc003a.Show     '' �����}�E�B���h�E��\������
End Sub

'*******************************************************************************
'*    �֐���        : sub_InitData
'*
'*    �����T�v      : 1.�f�[�^�����ݒ�
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*                    �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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

    '' �O��ʂ���̃f�[�^�󂯎��
    tblTotal = typ_CType
    tblSXL = typ_CType.typ_Param

    '' �Ώۃu���b�NID�\���p�̍\���̂��擾����
    '' �s�}�������ۂ̃u���b�NID�̑I���R���{�{�b�N�X�Ɏg�p����
    Set orgXl = GetXl(tblSXL.CRYNUM, Me.Name)

    '' �����N���X����u���b�N���e���擾���ASXL�̑Ώۃu���b�N���𔻕�
    m = orgXl.Blks.COUNT
    ReDim SxlIntoBlock(m)
    i = 0
    For Each Blk In orgXl.Blks
        intSP = tblSXL.INGOTPOS
        intEP = intSP + tblSXL.COUNT
        intSBP = Blk.INGOTPOS
        intEBP = intSBP + Blk.LENGTH
        blFlag = False

        '' �u���b�N��SXL�̒��Ɋ��S�Ɋ܂܂�Ă���ꍇ
        If intSP <= intSBP And intEP >= intEBP Then
            blFlag = True
        '' �u���b�N��SXL�̊J�n�ʒu����ɂ���A���I�[�ʒu���������ꍇ
        ElseIf intSP >= intSBP And intEP <= intEBP Then
            blFlag = True
        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N���㑤�B�������u���b�N�̏I�[��SXL�̊J�n�ʒu����v���Ȃ�����)
        ElseIf intSP > intSBP And intSP < intEBP And intSP <> intEBP Then
            blFlag = True
        '' �u���b�N���ꕔSXL�ɂ������Ă���ꍇ
        '' (�u���b�N�������B������SXL�̏I�[�ƃu���b�N�̊J�n�ʒu����v���Ȃ�����)
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

    '' �u���b�N�Ǘ��e�[�u�����\���̂ɐݒ�
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

    '' �敪�R�[�h�̎擾
    Call GetCodeListSC18("SC", "18", "WF", tblPrcList)
End Sub

'*******************************************************************************
'*    �֐���        : sub_LoadAndDisp
'*
'*    �����T�v      : 1.�f�[�^�̃��[�h�ƕ\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_LoadAndDisp()

    Dim m       As Integer
    Dim n       As Integer
    Dim i       As Integer
    Dim sMsg    As String
    Dim vt      As Variant
'>>>>> add start 2011/06/30 Marushita
    Dim iMinMidCnt      As Integer       '���Ԕ����̕K�v��
    Dim iRstMidCnt      As Integer       '���Ԕ����̌���
'<<<<< add start 2011/06/30 Marushita

    '' �O��ʂ�������n���ꂽ�l��\������
    txtStaffID.text = tblTotal.StrStaffId                                   ' �S���҃R�[�h
    txtJfName.text = tblTotal.strStaffName                                  ' �S���Җ�
    txtKSXLID.text = tblSXL.SXLID                                           ' ��SXLID
    txtCryNum.text = tblSXL.CRYNUM                                          ' �C���S�b�gID
    If sKanrenFlg = "1" Then lblKanren.Visible = True       '�֘A��ۯ��\���@08/01/31 ooba
    txtTopRsltR.text = toRsStr(CDbl(tblTotal.typ_y013(1, WFRES).MESDATA5))  ' T�����у�
    txtBotRsltR.text = toRsStr(CDbl(tblTotal.typ_y013(2, WFRES).MESDATA5))  ' B�����у�

    If fnc_GetMukesaki_XSDCB(Trim(txtKSXLID.text)) = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' �������ԃZ�b�g
    SetPresentTime lblTime

    InMaxRow = 2

    '' �����Ɋ܂܂��i�Ԃ��ׂĂ��擾
    If GetXlHinban(tblSXL.CRYNUM, tblHinNum) = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' �i�Ԃ̐ݒ�
    ReDim tblHinbanRs(1)
    With tblHinbanRs(1)
        .CRYNUM = tblSXL.CRYNUM
        .HIN.hinban = tblSXL.hinban
        .HIN.mnorevno = tblSXL.REVNUM
        .HIN.factory = tblSXL.factory
        .HIN.opecond = tblSXL.opecond
    End With

    '' DB����֘A�f�[�^��ǂݍ���
    If fnc_LoadData = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' �T���v���̂Ȃ�SXL�������ꍇ�A��ʂ̏����͂Ȃɂ��s���Ȃ����̂Ƃ���
    If Trim$(tblSXL.WFSMP(1).XTALCW) <> "" And _
       Trim$(tblSXL.WFSMP(2).XTALCW) <> "" Then
        '' �����w���̕\��
        If fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE Then
            Exit Sub
        End If

        '' ���i�d�l�̕\��
        If fnc_DispHinSpec(0) = False Then
            Exit Sub
        End If

        '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
        With sprExamine
            m = .MaxRows
            n = .MaxCols

            '' �Ώ�SXL�̌����ʒu��ێ�
            .row = 1
            SIngotP = tblsmp(1).INGOTPOS    '�g�b�v�ʒu

            '�G�s��s�]���ǉ��Ή�
            .SetText 42, 1, tblsmp(1).INGOTPOS
'            '' �Ώ�SXL�̌����ʒu��ێ�
            .row = m
            EIngotP = tblsmp(2).INGOTPOS    '�{�g���ʒu

            '�G�s��s�]���ǉ��Ή�
            .SetText 42, m, tblsmp(2).INGOTPOS
        End With

        '�����ް��s���\���`�F�b�N
        If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
            lblMsg.Caption = GetMsgStr(sMsg)
            Exit Sub
        End If
'>>>>> add start 2011/06/30 Marushita
        ' ���Ԕ����i���H
        With sprSpec
            lblNukishi.Caption = ""
            For i = 1 To .MaxRows
                .GetText 33, i, vt
                If CStr(vt) = "1" Then
                    lblNukishi.Caption = "���Ԕ���"
                End If
            Next i
        End With
'        If typ_CType.typ_si.MSMPFLG = "1" Then
'            lblNukishi.Visible = True
'            lblNukishi.Caption = "���Ԕ���"
'            With sprSpec
'                .ColWidth(31) = 3.88      ' �\��
'                .ColWidth(32) = 3.88      ' �\��
'            End With
'            '���Ԕ����P��(���Ԕ������e�l(����)/(mm))
'            sprSpec.SetText 32, 1, CInt(typ_CType.typ_si.MSMPCONSTMAI)
'            '���Ԕ����K�v��
'            lblMSMP_SUU.Visible = True
'            '���Ԕ����̕K�v�� = (SXL��WF���� - ���Ԕ������e�l(����)) / ���Ԕ����P��(����)
'            iMinMidCnt = Fix((typ_CType.typ_Param.COUNT - typ_CType.typ_si.MSMPCONSTMAI) / typ_CType.typ_si.MSMPTANIMAI)
'            '�}�C�i�X�̏ꍇ�A�O�Ƃ���
'            If iMinMidCnt < 0 Then iMinMidCnt = 0
'            '���Ԕ����̌���
'            iRstMidCnt = (UBound(typ_CType.typ_Param.WFSMP) - SxlMidl) + 1
'            lblMSMP_SUU.Caption = "������/�K�v���F " & _
'            CInt(iRstMidCnt) & "/" & CInt(iMinMidCnt) & " ��"
'            '���Ԕ����P��(����)
'            sprSpec.SetText 31, 1, CInt(typ_CType.typ_si.MSMPTANIMAI)
'        Else
'            lblNukishi.Visible = False
'            lblNukishi.Caption = ""
'            lblMSMP_SUU.Visible = False
'            With sprSpec
'                .ColWidth(31) = 0      ' ��\��
'                .ColWidth(32) = 0      ' ��\��
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
'*    �֐���        : fnc_LoadData
'*
'*    �����T�v      : 1.�e�[�u������K�v���R�[�h���擾
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_LoadData() As FUNCTION_RETURN

    Dim udtTmpLackWaf() As typ_LackWaf
    Dim udtTmpBlkInf()  As typ_BlkInf3
    Dim sErrMsg         As String
    Dim m               As Integer
    Dim i               As Integer

    '' DB����f�[�^��ǂݍ���
    If DBDRV_scmzc_fcmlc001d_DispSiyou(tblHinbanRs, tblsiyou, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W�\��
        lblMsg.Caption = sErrMsg
        fnc_LoadData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '' �Ĕ����w���f�[�^�̎擾
    If DBDRV_scmzc_fcmlc001d_DispSmp(tblSXL.SXLID, tblsmp, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W�\��
        lblMsg.Caption = sErrMsg
        fnc_LoadData = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '' �������̎擾
    If DBDRV_scmzc_fcmlc001d_LostInfo(tblSXL.CRYNUM, udtTmpLackWaf) = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W�\��
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

    '' �����E�F�n�[�e�[�u���̍쐬
    If LackMapMake(udtTmpBlkInf, udtTmpLackWaf, 1, m) = FUNCTION_RETURN_FAILURE Then
    End If

    fnc_LoadData = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************
'*    �֐���        : sub_SaveData
'*
'*    �����T�v      : 1.�e�[�u���ւ̒ǉ��A�X�V���s��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : �Ȃ�
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

    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
    Dim SirdYN                  As Boolean      'SIRD����َw���L��[True:��񂠂�AFalse:��񖳂�]
    Dim sirdSmpID               As String       '�����ςݻ����ID�iSIRD�p�j<TBCMJ022>
    Dim sirdSmpID_Inp           As String       '�����w�������ID�iSIRD�p�j
    Dim ix                      As Integer      'SIRD����ٌ����pINDEX
    Dim iy                      As Integer      'SIRD����ٌ����pINDEX
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END

    udtNewData.detail = udtNewData_Detail     '07/10/05 miyatake END ==================>

    lblMsg.Caption = ""

    '' �T���v���{�^����������Ă��邩
    If bSampFlag = False Then
        '' �G���[���b�Z�[�W��\�����ď����𔲂���
        lblMsg.Caption = GetMsgStr("ESAMP")
        Exit Sub
    End If

    '' �����w���ꗗ�̓��̓`�F�b�N�iWFϯ�߁j
    If fnc_CheckDataWfmap() = FUNCTION_RETURN_FAILURE Then
        Exit Sub
    End If

    '' �敪�G���[�`�F�b�N
    If fnc_CheckHinbanZ() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("EHIN6")
        Exit Sub
    End If

    '����]�p�`�F�b�N
    If fnc_ChkMukesaki = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = "���悪�ύX����Ă��܂��B"
        Exit Sub
    End If

    '�����ް��s���\���`�F�b�N
    If fnc_ErrDispCheck(sMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr(sMsg)
        Exit Sub
    End If

    '' �����Ď��`�F�b�N add 09/03/17 SETkimizuka
    If CheckXODY4(WATCH_PROCCD_NUKISI, "", txtKSXLID.text) = False Then
        lblMsg.Caption = Y4_STOP_ERR
        Exit Sub
    End If
        
    '���o�K���`�F�b�N    *2010/02/15 Kameda
    If F_HaraiKisei = False Then
        lblMsg.Caption = GetMsgStr("EREG1")
        Exit Sub
    End If
    
    If MsgBox(GetMsgStr("PIN01"), vbOKCancel, "�Ĕ���") = vbCancel Then
        cmdF(12).Enabled = True
        Exit Sub
    End If

    ' ���F�@�\�ǉ��ɂ��C��  2007/10/05 miyatake ===================> START
    '' �R�����g����
    If Me.chk_Png = 1 Then
        If f_comment.GetComment(sComment) <> vbOK Then
            Exit Sub
        End If
        Call SetForceForegroundWindow(Me.hwnd)
    End If
    DoEvents
    ' ���F�@�\�ǉ��ɂ��C��  2007/10/05 miyatake ===================> START

    '' �Ĕ����e�[�u���̍X�V
    If fnc_UpdateData() = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W��\�����ď����𔲂���
        '���������װ��ү���ޕ\���@2004/01/28 ooba
        If bJituChkFlg = True Then
            lblMsg.Caption = "�T���v���ʒu�Ɏ���������܂���B"
        Else
            lblMsg.Caption = GetMsgStr("SET49")
        End If
        Exit Sub
    End If

    '' U/D���㉺�ɕ���
    Call SeparateUD

    '�T���v���Ǘ���SXL�Ǘ��̏�����ς���
    '' SXL�Ǘ��e�[�u���\���̂̐ݒ�
    Call sub_MakeTBCME042

    '' WF�T���v���Ǘ��e�[�u���\���̂̐ݒ�
    '�V�T���v���Ǘ�(XSDCW)�ɕύX
    Call sub_MakeTBCME044

    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
    '�����ς݂̻���ُ�������<TBCMJ022>
    If Not fncGetSirdSample(Trim(txtCryNum.text), SirdYN, sirdSmpID) Then
        MsgBox "�����ς݂̃T���v�����̌����Ɏ��s���܂���" & vbCrLf & "( TBCMJ022 )", vbInformation + vbOKOnly
        Exit Sub
    End If

    '-------------------------------------------------------------------------
    'SIRD�̔����w�����L��ꍇ�A�������牺�̏��͍���w����������ق����L����
    'SIRD�̔����w���������ꍇ�A���Ɍ����ς݂̻����ID�����L����
    '-------------------------------------------------------------------------
    For ix = 1 To UBound(tblWfSample)
        'SIRD�́u����ً��L�v�ɂ��āA���L��������ID���������u��������i�ʒu����O�ōł��߂�����فj
        If tblWfSample(ix).WFSMP.WFINDL4CW = "2" Then
            
            '����o�^������SIRD�̔����w�������邩��������
            sirdSmpID_Inp = ""
            For iy = 1 To ix - 1
            
                '����o�^������SIRD�̔����w��������ꍇ�A�ʒu�����m�F
                If tblWfSample(iy).WFSMP.WFINDL4CW = "1" Then
                
                    '�����w������ق̈ʒu����������O�ł���λ����ID��Keep
                    If tblWfSample(iy).WFSMP.INPOSCW < tblWfSample(ix).WFSMP.INPOSCW Then
                        sirdSmpID_Inp = tblWfSample(iy).WFSMP.REPSMPLIDCW           '��\�T���v��ID
                    End If
                End If
            Next iy
            
            If sirdSmpID_Inp = "" Then
                '����o�^������SIRD�̔����w�������L����O�ɖ����ꍇ�A<TBCMJ022>�̌����ςݻ����ID���
                tblWfSample(ix).WFSMP.WFSMPLIDL4CW = sirdSmpID
            Else
                '����o�^������SIRD�̔����w�������L����O�ɗL��ꍇ�A�������������ID���
                tblWfSample(ix).WFSMP.WFSMPLIDL4CW = sirdSmpID_Inp
            End If
            
        End If
    Next ix
    '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END
    
    '' SXL�m��w���e�[�u���\���̂̐ݒ�
    Call sub_MakeTBCMY007

    '' WF����������уe�[�u���\���̂̐ݒ�
    Call sub_MakeTBCMW005

    '' �U�֔p�����уe�[�u���\���̂̐ݒ�
    Call sub_MakeTBCMW006

'' ����]�����@�w���e�[�u���\���̂̐ݒ� Start
    '' TBCME041�\���̂ɕi�Ԑݒ�(����]�����@�w���e�[�u���쐬�p)
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
                    ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/15 ooba
                    .col = 27
                    If .backColor = vbWhite Then
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT1CW = "0"
                    Else
                        tblWfSample(intLoopCnt).WFSMP.WFINDOT1CW = IIf(.text = "", "0", .text)
                    End If

                    '�G�s��s�]���ǉ��Ή�
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

    '' TBCME044�\���̂ɐݒ�(����]�����@�w���e�[�u���쐬�p)
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
            tblWafSmp(j).WFINDOT1CW = tblWfSample(i).WFSMP.WFINDOT1CW       ' WF�����w���iOT1)
            tblWafSmp(j).WFINDOT2CW = tblWfSample(i).WFSMP.WFINDOT2CW       ' WF�����w���iOT2)
            tblWafSmp(j).WFINDAOICW = tblWfSample(i).WFSMP.WFINDAOICW       ' �c���_�f�ǉ�
            tblWafSmp(j).WFINDGDCW = tblWfSample(i).WFSMP.WFINDGDCW         ' GD�ǉ�
            tblWafSmp(j).WFHSGDCW = tblWfSample(i).WFSMP.WFHSGDCW           ' �ۏ�FLG(GD)

            '�G�s��s�]���ǉ��Ή�
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

    '�G�s��s�]���ǉ��Ή�
    ReDim udtTmpEpMesInd(0)

    If UBound(tblWafSmp) > 0 Then
        '�G�s��s�]���ǉ��Ή�
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


'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
For i = 1 To UBound(tblWafInd)
    Debug.Print "tblWafInd(" & i & ").BLOCKID=" & tblWafInd(i).BLOCKID & " : " & "tblWafInd(" & i & ").SAMPLEID=" & tblWafInd(i).SAMPLEID & " : " & "tblWafInd(" & i & ").SMP.CRYINDL4=" & tblWafInd(i).SMP.CRYINDL4
Next i
For i = 1 To UBound(tblSokuSizi)
    Debug.Print "tblSokuSizi(" & i & ").SAMPLEID=" & tblSokuSizi(i).SAMPLEID & " : " & "tblSokuSizi(" & i & ").OSITEM=" & tblSokuSizi(i).OSITEM
Next i
'��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END

    '' DB�ɓo�^
    OraDB.BeginTrans

    If DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W��\�����ď����𔲂���
        lblMsg.Caption = GetMsgStr("EWFM4") '03/06/06 �㓡
        OraDB.Rollback
        Exit Sub
    End If

    '' WF_GD����(TBCMJ015)�X�V����
    If UBound(typ_J015_WFGDUpd) > 0 Then
        '�ް�����UPDATE
        For intCnt = 1 To UBound(typ_J015_WFGDUpd)
            If DBDRV_scmzc_fcmlc001c_UpdGDdata(typ_J015_WFGDUpd(intCnt), txtStaffID.text) _
                                        <> FUNCTION_RETURN_SUCCESS Then
                lblMsg.Caption = GetMsgStr("EAPLY") & "J015"
                OraDB.Rollback
            Exit Sub
            End If
        Next
    End If

    '�G�s��s�]���ǉ��Ή�
    RET = DBDRV_scmzc_fcmlc001d_Exec(tblWfSample, tblWfSxlMng, tblWfHantei, tblHuriHai, tblSokuSizi, tblSxlKSiji, udtTmpEpMesInd(), sErrMsg)

    If RET = FUNCTION_RETURN_FAILURE Then
        '' �G���[���b�Z�[�W��\������
        lblMsg.Caption = sErrMsg
        OraDB.Rollback
        Exit Sub
    End If

    '### ��{���� ###
    Debug.Print "�VDB�����ݏ����J�n"
    If MakeParameter(SAINUKISI_FORM) <> FUNCTION_RETURN_SUCCESS Then
        '�G���[���b�Z�[�W�͂��łɏo���Ă���
        OraDB.Rollback
        Debug.Print "�VDB�����ݏ����ُ�I��"
        Call clearType  '�\���̏�����
        EndProcess '' �v���Z�X�I��
        Exit Sub
    End If

    '�����Ǘ�DB�o�^
    '�i�Ԃ̐U�ւ��s�����ꍇ�����Ǘ�DB�Ƀf�[�^��o�^����
     If fnc_RirekiKanriDB_Touroku(sErrMsg) = False Then
     lblMsg.Caption = sErrMsg
        OraDB.Rollback
        Exit Sub
     End If

     Call clearType  '�\���̏�����

    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> START
    ''PNG�t�@�C���쐬
    If Me.chk_Png = 1 Then
        udtNewData.xtal = txtCryNum
        udtNewData.STAFFID = txtStaffID
        udtNewData.SXLID = txtKSXLID
        udtNewData.memo = sComment
        '�H����CC710����CW760�ɕύX 2010/04/30 SETsw kubota
        'If FileCreate_PNG(PROCD_NUKISI_SIJI, udtNewData, Me, sErrMsg, Nothing, pic_Png) = False Then ' upd 09/02/04 SETmiyatake
        If FileCreate_PNG(PROCD_WFC_SAINUKISI, udtNewData, Me, sErrMsg, Nothing, pic_Png) = False Then
            OraDB.Rollback
            lblMsg.Caption = sErrMsg
            Exit Sub
        End If
    End If
    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> END

    'OraDB.Rollback   'test�p�R�����g��߂�
    OraDB.CommitTrans
    Debug.Print "�VDB�����ݏ�������I��"

    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> START
    If Me.chk_Png = 1 Then
        ''PNG�t�@�C�����M
        '�H����CC710����CW760�ɕύX 2010/04/30 SETsw kubota
        'Call FileReSend_PNG(PROCD_NUKISI_SIJI)
        Call FileReSend_PNG(PROCD_WFC_SAINUKISI)
    End If
    ' ���F�@�\�ǉ��ɂ��C��  07/10/05 miyatake ===================> END

    '' �I�����b�Z�[�W��\������
    lblMsg.Caption = GetMsgStr("PPROK")

    ' �I�����b�Z�[�W��\������
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
'*    �֐���        : fnc_CheckBlockP
'*
'*    �����T�v      : 1.���͂��ꂽ�u���b�N�ʒu��������������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@�@�@�@�@�@�@�@�@�Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '' �u���b�N�ʒu�A�����ʒu�̊m��
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With sprExamine
        m = .MaxRows
        If m = 0 Then
            Exit Function
        End If

        For i = 1 To m
            If i <> 1 And i <> m Then
                '�G�s��s�]���ǉ��Ή�
                .GetText 37, i, vNukisiFlg
                If (vNukisiFlg = "2") Then
                    .row = i
                    .col = 1
                    intCmdIndex = .TypeComboBoxCurSel

                    '�G�s��s�]���ǉ��Ή�
                    .col = 39
                    sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                    .col = 4
                    '' �G���[�`�F�b�N(���͂��ꂽ��)
                    If Trim$(.text) = "" Then
                        .backColor = COLOR_NG
                        Exit Function
                    End If
                    intTmpBlockP = .text
                    .col = 3
                    intRetIngotPos = orgXl.Blks.GetPosByID(sTmpBlockID)
                    '' �G���[�`�F�b�N
                    If fnc_CheckBlockP2(intRetIngotPos, sTmpBlockID, intTmpBlockP, i, intBlockPErrFlg) = FUNCTION_RETURN_FAILURE Then
                        '' �G���[�łȂ���Ό���P��ݒ肷��
                        .row = i
                        .col = 4
                        .backColor = COLOR_NG
                    Else
                        .row = i

                        '�G�s��s�]���ǉ��Ή�
                        .col = 39
                        sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                        .col = 4
                        intBlockp = CInt(.text)
                        .row = i + 2
                        .col = 4
                        intNextBlkP = CInt(.text)

                        .row = i
                        vSample1 = "U"
                        vSample2 = " " 'D�̃T���v��ID���쐬����

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
                        tblNukishi(i).REPSMPLIDCW = vSample1 '��\�T���v��ID���쐬
                        tblNukishi(i + 1).REPSMPLIDCW = vSample2

                        '�����ō��ꂽ����P�͕\�������Ȃ� -----------
                        .col = 5
                        .text = CStr(vGetIngotP)
                        .col = 4
                            .backColor = COLOR_OK
                    End If
                'vNukisiFlg = "3"�̂Ƃ��̏�����ǉ�(�T���v��ID�ؑւ̂���)
                ElseIf vNukisiFlg = "3" And i Mod 2 = 0 Then
                    .row = i

                    '�G�s��s�]���ǉ��Ή�
                    .col = 39
                    sTmpBlockID = Mid(Trim(txtCryNum.text), 1, 9) & Trim(.text)
                    .col = 4
                    intBlockp = CInt(.text)
                    .row = i + 2
                    .col = 4
                    intNextBlkP = CInt(.text)

                    .row = i
                    vSample1 = "U"
                    vSample2 = " " 'D�̃T���v��ID���쐬����

                    If DBDRV_GET_WFMAP(sTmpBlockID, tblSXL.SXLID, intBlockp, vGetBlkP, vGetIngotP, sNextIngotP, vGetBlkSeq, vGetBlkSeq2, vSample1, vSample2, intNextBlkP, vGetWfNum) = FUNCTION_RETURN_SUCCESS Then
                        ReDim Preserve tblNukishi(i + 1)
                        tblNukishi(i).REPSMPLIDCW = vSample1 '��\�T���v��ID���쐬
                        tblNukishi(i + 1).REPSMPLIDCW = vSample2
                    End If
                End If
            End If
        Next i

        '' �G���[�������ꍇ�A�����𔲂���
        If intBlockPErrFlg <> 0 Then
            '' �G���[���b�Z�[�W��\�����ď����𔲂���
            Select Case intBlockPErrFlg
                Case 1      ' SXL�͈͊O
                    lblMsg.Caption = GetMsgStr("EBLK2")
                Case 2      ' �ʒu�d��
                    lblMsg.Caption = GetMsgStr("EBLK3")
                Case 3      ' �i�ԃG���[
                    lblMsg.Caption = GetMsgStr("EHIN3")
                Case 4      ' ����
                    lblMsg.Caption = GetMsgStr("EBLK4")
            End Select

            fnc_CheckBlockP = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    End With
End Function

'*******************************************************************************
'*    �֐���        : fnc_CheckBlockP2
'*
'*    �����T�v      : 1.�u���b�NP�G���[�`�F�b�N�i�T�u�j
'*�@�@�@�@�@�@�@�@�@�@  (�u���b�NP���͈͊O���͂���Ă��Ȃ����`�F�b�N)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*          �@�@      intIngotpos   ,I  ,Integer�@,�u���b�N�J�n�ʒu
'*�@�@      �@�@      InBlockId     ,I  ,String �@,�u���b�NID
'*�@�@      �@�@      intBlockp     ,I  ,String �@,�u���b�N�ʒu
'*�@�@      �@�@      intRowCount   ,I  ,String �@,SPREAD�̈ʒu
'*�@�@      �@�@      intErrP       ,O  ,String �@,�G���[���e�t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '' �G���[�`�F�b�N�iSXL�͈͓����ǂ����j
    If intIngotpos + intBlockp > EIngotP Or intIngotpos + intBlockp < SIngotP Then
        fnc_CheckBlockP2 = FUNCTION_RETURN_FAILURE
        intErrP = 1
        Exit Function
    End If

    '' �G���[�`�F�b�N�i�u���b�N�͈͓̔����j
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

    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With sprExamine
        m = .MaxRows

        '' �u���b�N�o�̌����`�F�b�N
        n = UBound(tblLackMap)
        For i = 1 To m
            '�G�s��s�]���ǉ��Ή�
            .GetText 37, i, vNukisiFlg
            If (vNukisiFlg = "2") Then  '�ǉ������s�̂ݏ������s��
                .row = i

                '�G�s��s�]���ǉ��Ή�
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
'*    �֐���        : fnc_CheckHinbanZ
'*
'*    �����T�v      : 1.Z�i�Ԃ��ݒ肳��Ă���Ƃ��ɋ敪0�Ȃ�G���[�Ƃ���
'*�@�@�@�@�@�@�@�@�@�@  (�i�ԃG���[�`�F�b�N)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With sprExamine
        m = .MaxRows
        For i = 1 To m
            If i <> m Then
                '�G�s��s�]���ǉ��Ή�
                .GetText 37, i, vNukisiFlg  '���ݍs
                .GetText 37, i + 1, vNextNukisiFlg  '���s

                If i Mod 2 <> 0 Then    '��s�̂ݏ������s��
                     .GetText 2, i, vHinChk
                    sNowHinban = vHinChk
                    .row = i
                    .col = 9
                    j = .TypeComboBoxCurSel + 1
                    sNowKubun = Trim$(tblPrcList(j).CODE)
                    If sNowHinban = "Z" And sNowKubun = "0" Then
                        .backColor = COLOR_NG  '�F�ς����Ȃ�
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
'*    �֐���        : fnc_UpdateData
'*
'*    �����T�v      : 1.�����w���ꗗ�̓��͓��e�ɂ��S�f�[�^���X�V����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '' �Ĕ����w���f�[�^�̎擾
    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX

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

            '' WF�T���v���w���f�[�^���X�V����
            If i = 1 Or i = m Or CheckGetSampleID(i) = True Or CheckGetSampleID(i) = False Then
                '�G�s��s�]���ǉ��Ή�
                .GetText 37, i, vNukisiFlg

                If (vNukisiFlg = 1) Then   '�����\�������s�̏������s��
                    intSmpRow = intSmpRow + 1
                    ReDim Preserve tblWafInd(intSmpRow)     '�����s�̂ݍ\���̂��쐬����
                    .row = i

                    '�G�s��s�]���ǉ��Ή�
                    .col = 39
                    '�T���v��ID������� UPDATE�̏����Ɏg�p
                    tblWafInd(intSmpRow).BLOCKID = left(tblSXL.SXLID, 9) & .text

                    '�T���v��ID���ς��Ȃ��i�����\���̂P�E�ŏI�s�j�ł͊J�n�E�I���ʒu�͕ς��Ȃ�
                    .col = 4
                    tblWafInd(intSmpRow).BlockPos = val(.text)

                    .col = 5
                    If Not .text = "" Then
                        tblWafInd(intSmpRow).INGOTPOS = .text
                    End If

                    If i = 1 Then   '�P�s�ڂ̌����ʒu�͊����ʒu�ɂȂ�
                        tblWafInd(intSmpRow).INGOTPOS = SIngotP
                        tblNukishi(i).INPOSCW = SIngotP
                    ElseIf i = m Then
                        tblWafInd(intSmpRow).INGOTPOS = EIngotP   '�ŏI�s�������ʒu
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
                        '' �i�Ԑݒ肳��Ă���Ȃ�t�����i�Ԃ��擾
                        If sHin <> "" Then
                            If GetLastHinban(sHin, udtTmpHin) = FUNCTION_RETURN_FAILURE Then
                                '' �G���[����
                                tblWafInd(intSmpRow).ERRDNFLG = True
                                blFlag = True
                            End If
                        End If
                    End If

                    '�Ō�̍s�̉��i�Ԃɂ͏����\���̕i�Ԃ�ݒ�
                    If i = m Then
                        sHin = ""
                    End If

                    With tblWafInd(intSmpRow)
                        If sHin = "Z" Then
                            .HINDN.hinban = sHin                    ' �i��
                            .HINDN.mnorevno = 0                     ' ���i�ԍ������ԍ�
                            .HINDN.factory = ""                     ' �H��
                            .HINDN.opecond = ""                     ' ���Ə���
                        ElseIf sHin = "" Then
                            .HINDN.hinban = tblSXL.hinban           ' �i��
                            .HINDN.mnorevno = tblSXL.REVNUM         ' ���i�ԍ������ԍ�
                            .HINDN.factory = tblSXL.factory         ' �H��
                            .HINDN.opecond = tblSXL.opecond         ' ���Ə���
                        Else
                            .HINDN.hinban = sHin                    ' �i��
                            .HINDN.mnorevno = udtTmpHin.mnorevno    ' ���i�ԍ������ԍ�
                            .HINDN.factory = udtTmpHin.factory      ' �H��
                            .HINDN.opecond = udtTmpHin.opecond      ' ���Ə���
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

                    ''�c���_�f�ǉ�
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = 0
                    If .text <> vbNullString Then   '��������ON�̕����̂݌������ڎ擾
                        If .backColor = vbBlack Then
                            tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                            tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)
                        End If
                    End If

                    '�G�s��s�]���ǉ��Ή�
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = 0
                    If .text <> vbNullString Then   '��������ON�̕����̂݌������ڎ擾
                        If .backColor = vbBlack Then
                            tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                            tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)
                        End If
                    End If

                    'GD�ǉ�
                    '�G�s��s�]���ǉ��Ή�
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)     '����׸�(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)               '����׸�(GD)

                    '����ٖ������܂���Z�i�ԂłȂ��ꍇ
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg) '�ۏ��׸�(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD        '�ۏ��׸�(GD)
                    '����ُ����ς���Z�i�Ԃ̏ꍇ
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW        '�ۏ��׸�(GD)
                    End If

                    '�G�s��s�]���ǉ��Ή�
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

                    '�����\���̐擪�A�Ō�̓TSMPLKBN2���v��UPDATE�̂��ߋ敪�͓���Ȃ�
                    tblWafInd(intSmpRow).SMPLKBN2 = ""
                    tblWafInd(intSmpRow).SMPLKBN1 = ""

                    '�G�s��s�]���ǉ��Ή�
                    .col = 38
                    tblWafInd(intSmpRow).SAMPLEID = .text

                    If i = 1 Then
                        tblNukishi(i).SMPKBNCW = tKensa(0).SMPKBNCW
                    Else
                        tblNukishi(i).SMPKBNCW = tKensa(1).SMPKBNCW
                    End If
                    tblNukishi(i).TBKBNCW = IIf(tblNukishi(i).SMPKBNCW = "B" Or tblNukishi(i).SMPKBNCW = "U", "B", "T") 'TB�敪
                ElseIf vNukisiFlg = 2 Or (vNukisiFlg = 3 And CheckGetSampleID(i) = True) Then  '�ǉ������s�̏������s��
                    intSmpRow = intSmpRow + 1
                    ReDim Preserve tblWafInd(intSmpRow)   'add 2003/04/14 hitec)matsumoto �����s�̂ݍ\���̂��쐬����

                    If vNukisiFlg = 2 Then
                        .row = i
                        '�G�s��s�]���ǉ��Ή�
                        .col = 39
                        tblWafInd(intSmpRow).BLOCKID = left(tblSXL.SXLID, 9) & .text

                        .col = 4
                        tblWafInd(intSmpRow).BlockPos = val(.text)

                        .col = 5
                        tblWafInd(intSmpRow).INGOTPOS = .text
                        tblNukishi(i).INPOSCW = .text
                    ElseIf vNukisiFlg = 3 Then
                        .row = i + 1
                        '�G�s��s�]���ǉ��Ή�
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
                        '' �i�Ԑݒ肳��Ă���Ȃ�t�����i�Ԃ��擾
                        If sHin <> "" Then
                            If GetLastHinban(.text, udtTmpHin) = FUNCTION_RETURN_FAILURE Then
                                '' �G���[����
                                tblWafInd(intSmpRow).ERRDNFLG = True
                                blFlag = True
                            End If
                        End If
                    End If

                    With tblWafInd(intSmpRow)
                        If sHin = "Z" Then
                            .HINDN.hinban = sHin                    ' �i��
                            .HINDN.mnorevno = 0                     ' ���i�ԍ������ԍ�
                            .HINDN.factory = ""                     ' �H��
                            .HINDN.opecond = ""                     ' ���Ə���
                        ElseIf sHin = "" Then
                            .HINDN.hinban = tblSXL.hinban           ' �i��
                            .HINDN.mnorevno = tblSXL.REVNUM         ' ���i�ԍ������ԍ�
                            .HINDN.factory = tblSXL.factory         ' �H��
                            .HINDN.opecond = tblSXL.opecond         ' ���Ə���
                        Else
                            .HINDN.hinban = sHin                    ' �i��
                            .HINDN.mnorevno = udtTmpHin.mnorevno    ' ���i�ԍ������ԍ�
                            .HINDN.factory = udtTmpHin.factory      ' �H��
                            .HINDN.opecond = udtTmpHin.opecond      ' ���Ə���
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

                        '�G�s��s�]���ǉ��Ή�
                        .GetText 38, i + 1, vGetSample2
                        .GetText 44, i, vGetSample1

                        tblWafInd(intSmpRow).SAMPLEID = CStr(vGetSample2)
                        tblWafInd(intSmpRow).SAMPLEID2 = ""
                    ElseIf Trim(vGetSample2) = gsWF_SMPL_JOINT Then
                        tblWafInd(intSmpRow).SMPLKBN1 = Right(vGetSample1, 1)
                        tblWafInd(intSmpRow).SMPLKBN2 = ""

                        '�G�s��s�]���ǉ��Ή�
                        .GetText 38, i, vGetSample1
                        .GetText 44, i + 1, vGetSample2

                        tblWafInd(intSmpRow).SAMPLEID = CStr(vGetSample1)
                        tblWafInd(intSmpRow).SAMPLEID2 = ""
                    Else
                        tblWafInd(intSmpRow).SMPLKBN1 = Right(vGetSample1, 1)
                        tblWafInd(intSmpRow).SMPLKBN2 = Right(vGetSample2, 1)

                        '�G�s��s�]���ǉ��Ή�
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

                    ''�c���_�f�ǉ�
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)

                    '�G�s��s�]���ǉ��Ή�
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)

                    'GD�ǉ��@05/02/18 ooba
                    '�G�s��s�]���ǉ��Ή�
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)     '����׸�(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)               '����׸�(GD)

                    '����ٖ������܂���Z�i�ԂłȂ��ꍇ
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg) '�ۏ��׸�(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD        '�ۏ��׸�(GD)
                    '����ُ����ς���Z�i�Ԃ̏ꍇ
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW        '�ۏ��׸�(GD)
                    End If

                    '�G�s��s�]���ǉ��Ή�
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
                     .row = i      '�����s�������ǉ�

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

                    ''�c���_�f�ǉ�
                    .col = 26
                    tblWafInd(intSmpRow).SMP.CRYINDAO = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDAOICW = IIf(.text = "", "0", .text)

                    .col = 27
                    tblWafInd(intSmpRow).SMP.CRYOTHER1 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT1CW = IIf(.text = "", "0", .text)

                    '�G�s��s�]���ǉ��Ή�
                    .col = 35
                    tblWafInd(intSmpRow).SMP.CRYOTHER2 = IIf(.text = "", "0", .text)
                    tblNukishi(i).WFINDOT2CW = IIf(.text = "", "0", .text)

                    '�G�s��s�]���ǉ��Ή�
                    .col = 28
                    tblWafInd(intSmpRow).SMP.CRYINDGD = IIf(.text = "", "0", .text)   '����׸�(GD)
                    tblNukishi(i).WFINDGDCW = IIf(.text = "", "0", .text)           '����׸�(GD)
                    '����ٖ������܂���Z�i�ԂłȂ��ꍇ
                    If bSampFlag = False Or Trim(tblNukishi(i).HINBCW) <> "Z" Then
                        tblWafInd(intSmpRow).SMP.WFHSGD = fnc_Get_CellHsFlg(vNukisiFlg)   '�ۏ��׸�(GD)
                        tblNukishi(i).WFHSGDCW = tblWafInd(intSmpRow).SMP.WFHSGD      '�ۏ��׸�(GD)
                    '����ُ����ς���Z�i�Ԃ̏ꍇ
                    Else
                        tblWafInd(intSmpRow).SMP.WFHSGD = tblNukishi(i).WFHSGDCW      '�ۏ��׸�(GD)
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

                    '�������������ǉ�
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

        '' �i�ԂɌ�肪����Ȃ珈���𔲂���
        If blFlag = True Then
            fnc_UpdateData = FUNCTION_RETURN_FAILURE
            lblMsg.Caption = GetMsgStr("EHIN3")
            Exit Function
        End If
    End With

    '' WF�Ώەi�ԍ\���̂̐ݒ�
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

    '' �����̍X�V
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
'*    �֐���        : fnc_Get_CellHsFlg
'*
'*    �����T�v      : 1.�ِF����ۏ��׸ނ��擾����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      vNukisi�@�@�@ ,I  ,Variant  ,�����t���O
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_Get_CellHsFlg(vNukisi As Variant) As String

    Dim intRow      As Integer            '�����s
    Dim intCol      As Integer            '������
    Dim sHin        As String             '8���i��
    Dim udtFullHin  As tFullHinban        '12���i��
    Dim sBlkflg     As String             '��ۯ��P�ʕۏ��׸�

    '�ۏ��׸ޢ0��FWF����
    fnc_Get_CellHsFlg = "0"

    With sprExamine
        intRow = .row
        intCol = .col

        '�����ʒu�̕i�Ԏ擾
        If intRow Mod 2 = 1 Then .row = intRow Else .row = intRow - 1
        .col = 2
        sHin = Trim$(.text)
        If GetLastHinban(sHin, udtFullHin) = FUNCTION_RETURN_FAILURE Then Exit Function

        '��ۯ��P�ʕۏ��׸ގ擾
        If chkBlkTanFlg(udtFullHin, sBlkflg) = FUNCTION_RETURN_FAILURE Then Exit Function

        .row = intRow
        .col = intCol

        '��ۯ��P�ʕۏ��׸ނ��������O
        If .backColor = COLOR_CryJitsu Then fnc_Get_CellHsFlg = "1"
    End With
End Function

'*******************************************************************************
'*    �֐���        : fnc_CheckJituData
'*
'*    �����T�v      : 1.���s�������A�ǉ������s�ɑ΂������ް�����������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      intRow�@�@   �@ ,I  ,Integer  ,Spread�̍s
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************
Private Function fnc_CheckJituData(intRow As Integer) As FUNCTION_RETURN
    Dim intChkRow       As Integer      '�������������s
    Dim intChkCol       As Integer      '����������������
    Dim intJituChkCnt   As Integer      '���т��茟�����ڐ�
    Dim blBetuFlg       As Boolean      '���L(��)������׸�
    Dim intCnt          As Integer

    fnc_CheckJituData = FUNCTION_RETURN_FAILURE
    intJituChkCnt = 0
    intCnt = 1
    blBetuFlg = False

    With sprExamine

        '���������s�����L(��)����ق�����
        Do While iBetuRow(intCnt) > 0
            If intRow = iBetuRow(intCnt) Then
                blBetuFlg = True
            End If
            intCnt = intCnt + 1
            If intCnt > .MaxRows Then Exit Do
        Loop

        '���L(��)����ق̏ꍇ��1�s��������������
        If blBetuFlg = True Then
            For intChkRow = intRow - 1 To intRow
                .row = intChkRow

                '�G�s��s�]���ǉ��Ή�
                For intChkCol = 11 To 35
                    .col = intChkCol
                    If .text = "1" Then
                        intJituChkCnt = intJituChkCnt + 1
                    End If
                Next

                '�����ް���1�ł�����ꍇ������OK�B�Ȃ��ꍇ������NG�Ŏ��s�����𒆎~����B
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

                '�G�s��s�]���ǉ��Ή�
                For intChkCol = 11 To 35
                    .col = intChkCol
                    If .text = "1" Then
                        intJituChkCnt = intJituChkCnt + 1
                    End If
                Next
            Next

            '�����ް���1�ł�����ꍇ������OK�B�Ȃ��ꍇ������NG�Ŏ��s�����𒆎~����B
            If intJituChkCnt > 0 Then
                fnc_CheckJituData = FUNCTION_RETURN_SUCCESS
            End If
        End If
    End With
End Function

'*******************************************************************************
'*    �֐���        : sub_GetWafHinban
'*
'*    �����T�v      : 1.�e�[�u���̏㉺�i�Ԃ��擾����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_GetWafHinban()

    Dim sHin    As String
    Dim intPos  As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim i       As Integer
    Dim j       As Integer

    '' �����w���f�[�^�̍X�V
    m = UBound(tblWafInd)
    n = UBound(tblHinNum)
    For i = 1 To m
        With tblWafInd(i)
            If i = 1 Then
                '' ��i�Ԃ̎擾
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
                    '' ��i�Ԃ̎擾
                    .HINUP.hinban = tblWafInd(i - 1).HINDN.hinban
                    .HINUP.mnorevno = tblWafInd(i - 1).HINDN.mnorevno
                    .HINUP.factory = tblWafInd(i - 1).HINDN.factory
                    .HINUP.opecond = tblWafInd(i - 1).HINDN.opecond

                    ' ���i�Ԃ̎擾
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
'*    �֐���        : sub_GetWafHinban01
'*
'*    �����T�v      : 1.�e�[�u���̏㉺�i�Ԃ��擾����
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^       ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_GetWafHinban01(i As Integer)

    Dim sHin    As String
    Dim intPos  As Integer
    Dim m       As Integer
    Dim n       As Integer
    Dim j       As Integer

    '' �����w���f�[�^�̍X�V
    m = UBound(tblWafInd)
    n = UBound(tblHinNum)
    With tblWafInd(i)
        If i = 1 Then
            '' ��i�Ԃ̎擾
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
                '' ��i�Ԃ̎擾
                .HINUP.hinban = tblWafInd(i - 1).HINDN.hinban
                .HINUP.mnorevno = tblWafInd(i - 1).HINDN.mnorevno
                .HINUP.factory = tblWafInd(i - 1).HINDN.factory
                .HINUP.opecond = tblWafInd(i - 1).HINDN.opecond

                ' ���i�Ԃ̎擾
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
'*    �֐���        : fnc_BsmpID
'*
'*    �����T�v      : 1.�����T���v����WF�T���v���R�t���֐��Ăяo��������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      udtWfSample�@ ,I  ,typ_WfSampleGr ,WF�T���v���Ǘ�
'*�@�@      �@�@      j          �@ ,I  ,Integer        ,tblWfSample�\���̔z��p����
'*
'*    �߂�l        : �����T���v���̃u���b�NID
'*
'*******************************************************************************
Private Function fnc_BsmpID(udtWfSample() As typ_WfSampleGr, j As Integer) As String
    Dim udtWfsmp    As typ_Wf_Smpl   '�R�t�����ʊ֐��p(XSDCW) 2003/09/29 �ǉ�
    Dim udtCrsmp    As typ_Cry_Smpl  '�R�t�����ʊ֐��p(XSDCS�߂�l)

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
        '�G���[���b�Z�[�W
        fnc_BsmpID = ""
    End If

    fnc_BsmpID = udtCrsmp.CRYNUMCS
End Function

'*******************************************************************************
'*    �֐���        : sub_Set_SMP_TB
'*
'*    �����T�v      : 1.���L�s����ۯ��̋��E��SXL���������s�̏ꍇ�AUD��TB�ɕύX����B(�e�������ڂ̻����ID)
'*                      (UD��TB�ɕύX(�e�������ڂ̻����ID))
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      udtTmpWafSmp  ,I  ,typ_WfSampleGr ,WF�T���v���Ǘ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_Set_SMP_TB(udtTmpWafSmp() As typ_WfSampleGr)

    Dim intCnt As Integer      'ٰ�߶���

    'udtTmpWafSmp�̐������[�v���� (��)
    For intCnt = 1 To UBound(udtTmpWafSmp)
        With udtTmpWafSmp(intCnt)
            '�e�������ڂ�WFIND���A0�ȊO�Ȃ�TB�ϊ��������s�Ȃ�
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

            '�G�s��s�]���ǉ��Ή�
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
'*    �֐���        : sub_MakeTBCME044
'*
'*    �����T�v      : 1.�e�[�u���֓o�^�����邽�߂ɍ\���̍\�z
'*                      (WF�T���v���Ǘ��e�[�u���̍쐬)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_MakeTBCME044()
    Dim m           As Integer
    Dim i           As Integer
    Dim j           As Integer
    Dim ii          As Integer
    Dim intCnt      As Integer      'SXL�Ǘ��p����
    Dim k           As Integer      'tblWfSample�\���̔z��p����

    Dim intDx1      As Integer
    Dim intDx2      As Integer
    Dim blBlkFlg    As Boolean      'SXL�����ޯ���׸�
    Dim vGetIngotP1 As Variant
    Dim vGetIngotP2 As Variant

    '' WF�T���v���Ǘ��e�[�u���\���̂̐ݒ�
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
                    tblWfSample(k).BLOCKID = .BLOCKID                   ' �u���b�NID
                    tblWfSample(k).blockp = .BlockPos                   ' �u���b�N�ʒu
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(i - 1 * ((j + 1) Mod 2)).SXLID  'SXLID

                    tblWfSample(k).WFSMP.SMPKBNCW = Right(tblNukishi(j).REPSMPLIDCW, 1) ' �T���v���敪
                    tblWfSample(k).WFSMP.TBKBNCW = IIf(tblWfSample(k).WFSMP.SMPKBNCW = "U" Or tblWfSample(k).WFSMP.SMPKBNCW = "B", "B", "T") 'TB�敪
                    tblWfSample(k).WFSMP.REPSMPLIDCW = tblNukishi(j).REPSMPLIDCW '��\�T���v��ID
                    tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM            ' �����ԍ�
                    tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                ' �������ʒu

                    '�T���v��ID���}�b�v��ID�ƈقȂ邱�Ƃւ̑Ή� 2003/04/23
                    If tblWfSample(k).WFSMP.TBKBNCW = "B" Then
                        tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban         ' �i��
                        tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno     ' ���i�ԍ������ԍ�
                        tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory     ' �H��
                        tblWfSample(k).WFSMP.OPECW = .HINUP.opecond         ' ���Ə���
                    Else
                        tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban         ' �i��
                        tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno     ' ���i�ԍ������ԍ�
                        tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory     ' �H��
                        tblWfSample(k).WFSMP.OPECW = .HINDN.opecond         ' ���Ə���
                    End If

                    tblWfSample(k).WFSMP.KTKBNCW = "0"                      ' �m��敪
                    tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k)              '�T���v���u���b�NID���R�t���֐�
                    tblWfSample(k).WFSMP.WFSMPLIDRSCW = IIf(tblNukishi(j).WFINDRSCW = "0", "", tblNukishi(j).WFSMPLIDRSCW) '�T���v��ID
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
                    tblWfSample(k).WFSMP.WFSMPLIDAOICW = IIf(tblNukishi(j).WFINDAOICW = "0", "", tblNukishi(j).WFSMPLIDAOICW)   '�c���_�f�ǉ��@03/12/15 ooba
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = IIf(tblNukishi(j).WFINDGDCW = "0", "", tblNukishi(j).WFSMPLIDGDCW)      'GD�ǉ��@05/02/21 ooba

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPSMPLIDB1CW = IIf(tblNukishi(j).EPINDB1CW = "0", "", tblNukishi(j).EPSMPLIDB1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB2CW = IIf(tblNukishi(j).EPINDB2CW = "0", "", tblNukishi(j).EPSMPLIDB2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB3CW = IIf(tblNukishi(j).EPINDB3CW = "0", "", tblNukishi(j).EPSMPLIDB3CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL1CW = IIf(tblNukishi(j).EPINDL1CW = "0", "", tblNukishi(j).EPSMPLIDL1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL2CW = IIf(tblNukishi(j).EPINDL2CW = "0", "", tblNukishi(j).EPSMPLIDL2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL3CW = IIf(tblNukishi(j).EPINDL3CW = "0", "", tblNukishi(j).EPSMPLIDL3CW)
                    tblWfSample(k).WFSMP.WFINDRSCW = tblNukishi(j).WFINDRSCW        ' WF�����w���iRs)
                    tblWfSample(k).WFSMP.WFINDOICW = tblNukishi(j).WFINDOICW        ' WF�����w���iOi)
                    tblWfSample(k).WFSMP.WFINDB1CW = tblNukishi(j).WFINDB1CW        ' WF�����w���iB1)
                    tblWfSample(k).WFSMP.WFINDB2CW = tblNukishi(j).WFINDB2CW        ' WF�����w���iB2)
                    tblWfSample(k).WFSMP.WFINDB3CW = tblNukishi(j).WFINDB3CW        ' WF�����w���iB3)
                    tblWfSample(k).WFSMP.WFINDL1CW = tblNukishi(j).WFINDL1CW        ' WF�����w���iL1)
                    tblWfSample(k).WFSMP.WFINDL2CW = tblNukishi(j).WFINDL2CW        ' WF�����w���iL2)
                    tblWfSample(k).WFSMP.WFINDL3CW = tblNukishi(j).WFINDL3CW        ' WF�����w���iL3)
                    tblWfSample(k).WFSMP.WFINDL4CW = tblNukishi(j).WFINDL4CW        ' WF�����w���iL4)
                    tblWfSample(k).WFSMP.WFINDDSCW = tblNukishi(j).WFINDDSCW        ' WF�����w���iDS)
                    tblWfSample(k).WFSMP.WFINDDZCW = tblNukishi(j).WFINDDZCW        ' WF�����w���iDZ)
                    tblWfSample(k).WFSMP.WFINDSPCW = tblNukishi(j).WFINDSPCW        ' WF�����w���iSP)
                    tblWfSample(k).WFSMP.WFINDDO1CW = tblNukishi(j).WFINDDO1CW      ' WF�����w���iD1)
                    tblWfSample(k).WFSMP.WFINDDO2CW = tblNukishi(j).WFINDDO2CW      ' WF�����w���iD2)
                    tblWfSample(k).WFSMP.WFINDDO3CW = tblNukishi(j).WFINDDO3CW      ' WF�����w���iD3)
                    tblWfSample(k).WFSMP.WFINDAOICW = tblNukishi(j).WFINDAOICW      ' WF�����w�� (AOi)    '�c���_�f�ǉ��@03/12/15 ooba
                    tblWfSample(k).WFSMP.WFINDGDCW = tblNukishi(j).WFINDGDCW        ' WF�����w�� (GD)     'GD�ǉ��@05/02/21 ooba

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPINDB1CW = tblNukishi(j).EPINDB1CW        ' WF�����w���iB1E)
                    tblWfSample(k).WFSMP.EPINDB2CW = tblNukishi(j).EPINDB2CW        ' WF�����w���iB2E)
                    tblWfSample(k).WFSMP.EPINDB3CW = tblNukishi(j).EPINDB3CW        ' WF�����w���iB3E)
                    tblWfSample(k).WFSMP.EPINDL1CW = tblNukishi(j).EPINDL1CW        ' WF�����w���iL1E)
                    tblWfSample(k).WFSMP.EPINDL2CW = tblNukishi(j).EPINDL2CW        ' WF�����w���iL2E)
                    tblWfSample(k).WFSMP.EPINDL3CW = tblNukishi(j).EPINDL3CW        ' WF�����w���iL3E)
                    tblWfSample(k).WFSMP.WFINDOT1CW = tblNukishi(j).WFINDOT1CW
                    tblWfSample(k).WFSMP.WFINDOT2CW = tblNukishi(j).WFINDOT2CW
                    tblWfSample(k).WFSMP.WFRESRS1CW = IIf(tblNukishi(j).WFRESRS1CW = "1", "1", "0")                ' WF�������сiRs)
                    tblWfSample(k).WFSMP.WFRESOICW = IIf(tblNukishi(j).WFRESOICW = "1", "1", "0")                  ' WF�������сiOi)
                    tblWfSample(k).WFSMP.WFRESB1CW = IIf(tblNukishi(j).WFRESB1CW = "1", "1", "0")                  ' WF�������сiB1)
                    tblWfSample(k).WFSMP.WFRESB2CW = IIf(tblNukishi(j).WFRESB2CW = "1", "1", "0")                  ' WF�������сiB2�j
                    tblWfSample(k).WFSMP.WFRESB3CW = IIf(tblNukishi(j).WFRESB3CW = "1", "1", "0")                  ' WF�������сiB3)
                    tblWfSample(k).WFSMP.WFRESL1CW = IIf(tblNukishi(j).WFRESL1CW = "1", "1", "0")                  ' WF�������сiL1)
                    tblWfSample(k).WFSMP.WFRESL2CW = IIf(tblNukishi(j).WFRESL2CW = "1", "1", "0")                  ' WF�������сiL2)
                    tblWfSample(k).WFSMP.WFRESL3CW = IIf(tblNukishi(j).WFRESL3CW = "1", "1", "0")                  ' WF�������сiL3)
                    tblWfSample(k).WFSMP.WFRESL4CW = IIf(tblNukishi(j).WFRESL4CW = "1", "1", "0")                  ' WF�������сiDS)
                    tblWfSample(k).WFSMP.WFRESDSCW = IIf(tblNukishi(j).WFRESDSCW = "1", "1", "0")                  ' WF�������сiDZ)
                    tblWfSample(k).WFSMP.WFRESDZCW = IIf(tblNukishi(j).WFRESDZCW = "1", "1", "0")                  ' WF�������сiDZ)
                    tblWfSample(k).WFSMP.WFRESSPCW = IIf(tblNukishi(j).WFRESSPCW = "1", "1", "0")                  ' WF�������сiSP)
                    tblWfSample(k).WFSMP.WFRESDO1CW = IIf(tblNukishi(j).WFRESDO1CW = "1", "1", "0")                ' WF�������сiDO1)
                    tblWfSample(k).WFSMP.WFRESDO2CW = IIf(tblNukishi(j).WFRESDO2CW = "1", "1", "0")                ' WF�������сiDO2)
                    tblWfSample(k).WFSMP.WFRESDO3CW = IIf(tblNukishi(j).WFRESDO3CW = "1", "1", "0")                ' WF�������сiDO3)
                    tblWfSample(k).WFSMP.WFRESOT1CW = IIf(tblNukishi(j).WFRESOT1CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESOT2CW = IIf(tblNukishi(j).WFRESOT2CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESAOICW = IIf(tblNukishi(j).WFRESAOICW = "1", "1", "0")                ' WF�������сiAOi)  '�c���_�f�ǉ�
                    tblWfSample(k).WFSMP.WFRESGDCW = IIf(tblNukishi(j).WFRESGDCW = "1", "1", "0")                  ' WF�������сiGD)   'GD�ǉ�

                    tblWfSample(k).WFSMP.WFHSGDCW = IIf(tblNukishi(j).WFHSGDCW = "1", "1", "0")                    ' �ۏ�FLG�iGD)

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPRESB1CW = IIf(tblNukishi(j).EPRESB1CW = "1", "1", "0")                  ' WF�������сiB1E)
                    tblWfSample(k).WFSMP.EPRESB2CW = IIf(tblNukishi(j).EPRESB2CW = "1", "1", "0")                  ' WF�������сiB2E�j
                    tblWfSample(k).WFSMP.EPRESB3CW = IIf(tblNukishi(j).EPRESB3CW = "1", "1", "0")                  ' WF�������сiB3E)
                    tblWfSample(k).WFSMP.EPRESL1CW = IIf(tblNukishi(j).EPRESL1CW = "1", "1", "0")                  ' WF�������сiL1E)
                    tblWfSample(k).WFSMP.EPRESL2CW = IIf(tblNukishi(j).EPRESL2CW = "1", "1", "0")                  ' WF�������сiL2E)
                    tblWfSample(k).WFSMP.EPRESL3CW = IIf(tblNukishi(j).EPRESL3CW = "1", "1", "0")                  ' WF�������сiL3E)
                    tblWfSample(k).WFSMP.TSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.LIVKCW = "0"

                    j = j + 1
                    k = k + 1
                    ReDim Preserve tblWfSample(k)

                    tblWfSample(k).BLOCKID = .BLOCKID                   ' �u���b�NID
                    tblWfSample(k).blockp = .BlockPos                   ' �u���b�N�ʒu

                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(i - 1 * ((j + 1) Mod 2)).SXLID   'SXLID
                    tblWfSample(k).WFSMP.SMPKBNCW = Right(tblNukishi(j).REPSMPLIDCW, 1)         ' �T���v���敪
                    tblWfSample(k).WFSMP.TBKBNCW = IIf(tblWfSample(k).WFSMP.SMPKBNCW = "U" Or tblWfSample(k).WFSMP.SMPKBNCW = "B", "B", "T") 'TB�敪
                    tblWfSample(k).WFSMP.REPSMPLIDCW = tblNukishi(j).REPSMPLIDCW                '��\�T���v��ID
                    tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM                                 ' �����ԍ�
                    tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                                    ' �������ʒu

                    If tblNukishi(j).TBKBNCW = "B" Then
                        tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban     ' �i��
                        tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno ' ���i�ԍ������ԍ�
                        tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory ' �H��
                        tblWfSample(k).WFSMP.OPECW = .HINUP.opecond     ' ���Ə���
                    Else
                        tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban     ' �i��
                        tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno ' ���i�ԍ������ԍ�
                        tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory ' �H��
                        tblWfSample(k).WFSMP.OPECW = .HINDN.opecond     ' ���Ə���
                    End If
                    tblWfSample(k).WFSMP.KTKBNCW = "0"                  ' �m��敪
                    tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k)    ' �T���v���u���b�NID���R�t���֐�

                    tblWfSample(k).WFSMP.WFSMPLIDRSCW = IIf(tblNukishi(j).WFINDRSCW = "0", "", tblNukishi(j).WFSMPLIDRSCW) '�T���v��ID
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
                    tblWfSample(k).WFSMP.WFSMPLIDAOICW = IIf(tblNukishi(j).WFINDAOICW = "0", "", tblNukishi(j).WFSMPLIDAOICW)   '�c���_�f�ǉ��@03/12/15 ooba
                    tblWfSample(k).WFSMP.WFSMPLIDGDCW = IIf(tblNukishi(j).WFINDGDCW = "0", "", tblNukishi(j).WFSMPLIDGDCW)      'GD�ǉ��@05/02/21 ooba

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPSMPLIDB1CW = IIf(tblNukishi(j).EPINDB1CW = "0", "", tblNukishi(j).EPSMPLIDB1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB2CW = IIf(tblNukishi(j).EPINDB2CW = "0", "", tblNukishi(j).EPSMPLIDB2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDB3CW = IIf(tblNukishi(j).EPINDB3CW = "0", "", tblNukishi(j).EPSMPLIDB3CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL1CW = IIf(tblNukishi(j).EPINDL1CW = "0", "", tblNukishi(j).EPSMPLIDL1CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL2CW = IIf(tblNukishi(j).EPINDL2CW = "0", "", tblNukishi(j).EPSMPLIDL2CW)
                    tblWfSample(k).WFSMP.EPSMPLIDL3CW = IIf(tblNukishi(j).EPINDL3CW = "0", "", tblNukishi(j).EPSMPLIDL3CW)

                    tblWfSample(k).WFSMP.WFINDRSCW = tblNukishi(j).WFINDRSCW        ' WF�����w���iRs)
                    tblWfSample(k).WFSMP.WFINDOICW = tblNukishi(j).WFINDOICW        ' WF�����w���iOi)
                    tblWfSample(k).WFSMP.WFINDB1CW = tblNukishi(j).WFINDB1CW        ' WF�����w���iB1)
                    tblWfSample(k).WFSMP.WFINDB2CW = tblNukishi(j).WFINDB2CW        ' WF�����w���iB2)
                    tblWfSample(k).WFSMP.WFINDB3CW = tblNukishi(j).WFINDB3CW        ' WF�����w���iB3)
                    tblWfSample(k).WFSMP.WFINDL1CW = tblNukishi(j).WFINDL1CW        ' WF�����w���iL1)
                    tblWfSample(k).WFSMP.WFINDL2CW = tblNukishi(j).WFINDL2CW        ' WF�����w���iL2)
                    tblWfSample(k).WFSMP.WFINDL3CW = tblNukishi(j).WFINDL3CW        ' WF�����w���iL3)
                    tblWfSample(k).WFSMP.WFINDL4CW = tblNukishi(j).WFINDL4CW        ' WF�����w���iL4)
                    tblWfSample(k).WFSMP.WFINDDSCW = tblNukishi(j).WFINDDSCW        ' WF�����w���iDS)
                    tblWfSample(k).WFSMP.WFINDDZCW = tblNukishi(j).WFINDDZCW        ' WF�����w���iDZ)
                    tblWfSample(k).WFSMP.WFINDSPCW = tblNukishi(j).WFINDSPCW        ' WF�����w���iSP)
                    tblWfSample(k).WFSMP.WFINDDO1CW = tblNukishi(j).WFINDDO1CW      ' WF�����w���iD1)
                    tblWfSample(k).WFSMP.WFINDDO2CW = tblNukishi(j).WFINDDO2CW      ' WF�����w���iD2)
                    tblWfSample(k).WFSMP.WFINDDO3CW = tblNukishi(j).WFINDDO3CW      ' WF�����w���iD3)
                    tblWfSample(k).WFSMP.WFINDOT1CW = tblNukishi(j).WFINDOT1CW
                    tblWfSample(k).WFSMP.WFINDOT2CW = tblNukishi(j).WFINDOT2CW
                    tblWfSample(k).WFSMP.WFINDAOICW = tblNukishi(j).WFINDAOICW      ' WF�����w�� (AOi)     '�c���_�f�ǉ��@03/12/15 ooba
                    tblWfSample(k).WFSMP.WFINDGDCW = tblNukishi(j).WFINDGDCW        ' WF�����w�� (GD)      'GD�ǉ��@05/02/21 ooba

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPINDB1CW = tblNukishi(j).EPINDB1CW        ' WF�����w���iB1)
                    tblWfSample(k).WFSMP.EPINDB2CW = tblNukishi(j).EPINDB2CW        ' WF�����w���iB2)
                    tblWfSample(k).WFSMP.EPINDB3CW = tblNukishi(j).EPINDB3CW        ' WF�����w���iB3)
                    tblWfSample(k).WFSMP.EPINDL1CW = tblNukishi(j).EPINDL1CW        ' WF�����w���iL1)
                    tblWfSample(k).WFSMP.EPINDL2CW = tblNukishi(j).EPINDL2CW        ' WF�����w���iL2)
                    tblWfSample(k).WFSMP.EPINDL3CW = tblNukishi(j).EPINDL3CW        ' WF�����w���iL3)
                    tblWfSample(k).WFSMP.WFRESRS1CW = IIf(tblNukishi(j).WFRESRS1CW = "1", "1", "0")               ' WF�������сiRs)
                    tblWfSample(k).WFSMP.WFRESOICW = IIf(tblNukishi(j).WFRESOICW = "1", "1", "0")                  ' WF�������сiOi)
                    tblWfSample(k).WFSMP.WFRESB1CW = IIf(tblNukishi(j).WFRESB1CW = "1", "1", "0")                  ' WF�������сiB1)
                    tblWfSample(k).WFSMP.WFRESB2CW = IIf(tblNukishi(j).WFRESB2CW = "1", "1", "0")                  ' WF�������сiB2�j
                    tblWfSample(k).WFSMP.WFRESB3CW = IIf(tblNukishi(j).WFRESB3CW = "1", "1", "0")                  ' WF�������сiB3)
                    tblWfSample(k).WFSMP.WFRESL1CW = IIf(tblNukishi(j).WFRESL1CW = "1", "1", "0")                  ' WF�������сiL1)
                    tblWfSample(k).WFSMP.WFRESL2CW = IIf(tblNukishi(j).WFRESL2CW = "1", "1", "0")                  ' WF�������сiL2)
                    tblWfSample(k).WFSMP.WFRESL3CW = IIf(tblNukishi(j).WFRESL3CW = "1", "1", "0")                  ' WF�������сiL3)
                    tblWfSample(k).WFSMP.WFRESL4CW = IIf(tblNukishi(j).WFRESL4CW = "1", "1", "0")                  ' WF�������сiDS)
                    tblWfSample(k).WFSMP.WFRESDSCW = IIf(tblNukishi(j).WFRESDSCW = "1", "1", "0")                  ' WF�������сiDZ)
                    tblWfSample(k).WFSMP.WFRESDZCW = IIf(tblNukishi(j).WFRESDZCW = "1", "1", "0")                  ' WF�������сiDZ)
                    tblWfSample(k).WFSMP.WFRESSPCW = IIf(tblNukishi(j).WFRESSPCW = "1", "1", "0")                  ' WF�������сiSP)
                    tblWfSample(k).WFSMP.WFRESDO1CW = IIf(tblNukishi(j).WFRESDO1CW = "1", "1", "0")                 ' WF�������сiDO1)
                    tblWfSample(k).WFSMP.WFRESDO2CW = IIf(tblNukishi(j).WFRESDO2CW = "1", "1", "0")                ' WF�������сiDO2)
                    tblWfSample(k).WFSMP.WFRESDO3CW = IIf(tblNukishi(j).WFRESDO3CW = "1", "1", "0")                 ' WF�������сiDO3)
                    tblWfSample(k).WFSMP.WFRESOT1CW = IIf(tblNukishi(j).WFRESOT1CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESOT2CW = IIf(tblNukishi(j).WFRESOT2CW = "1", "1", "0")
                    tblWfSample(k).WFSMP.WFRESAOICW = IIf(tblNukishi(j).WFRESAOICW = "1", "1", "0")                 ' WF�������� (AOi)  '�c���_�f�ǉ�
                    tblWfSample(k).WFSMP.WFRESGDCW = IIf(tblNukishi(j).WFRESGDCW = "1", "1", "0")                   ' WF�������� (GD)   'GD�ǉ�
                    tblWfSample(k).WFSMP.WFHSGDCW = IIf(tblNukishi(j).WFHSGDCW = "1", "1", "0")                     ' �ۏ�FLG�iGD)

                    '�G�s��s�]���ǉ��Ή�
                    tblWfSample(k).WFSMP.EPRESB1CW = IIf(tblNukishi(j).EPRESB1CW = "1", "1", "0")                  ' WF�������сiB1E)
                    tblWfSample(k).WFSMP.EPRESB2CW = IIf(tblNukishi(j).EPRESB2CW = "1", "1", "0")                  ' WF�������сiB2E�j
                    tblWfSample(k).WFSMP.EPRESB3CW = IIf(tblNukishi(j).EPRESB3CW = "1", "1", "0")                  ' WF�������сiB3E)
                    tblWfSample(k).WFSMP.EPRESL1CW = IIf(tblNukishi(j).EPRESL1CW = "1", "1", "0")                  ' WF�������сiL1E)
                    tblWfSample(k).WFSMP.EPRESL2CW = IIf(tblNukishi(j).EPRESL2CW = "1", "1", "0")                  ' WF�������сiL2E)
                    tblWfSample(k).WFSMP.EPRESL3CW = IIf(tblNukishi(j).EPRESL3CW = "1", "1", "0")                  ' WF�������сiL3E)
                    tblWfSample(k).WFSMP.TSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text
                    tblWfSample(k).WFSMP.LIVKCW = "0"

                    '��ۯ��̋��E��SXL���������ꍇUD��TB�ɕύX����
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
                        ' �ύX�Ώۂɒǉ�
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

                tblWfSample(k).BLOCKID = ""                                 ' �u���b�NID
                tblWfSample(k).WFSMP.XTALCW = tblSXL.CRYNUM                 ' �����ԍ�
                tblWfSample(k).WFSMP.INPOSCW = .INGOTPOS                    ' �������ʒu
                If i = 1 Then
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(1).SXLID     'SXLID
                    tblWfSample(k).WFSMP.HINBCW = .HINDN.hinban             ' �i��
                    tblWfSample(k).WFSMP.REVNUMCW = .HINDN.mnorevno         ' ���i�ԍ������ԍ�
                    tblWfSample(k).WFSMP.FACTORYCW = .HINDN.factory         ' �H��
                    tblWfSample(k).WFSMP.OPECW = .HINDN.opecond             ' ���Ə���
                Else
                    tblWfSample(k).WFSMP.SXLIDCW = tblWfSxlMng(UBound(tblWfSxlMng)).SXLID  'SXLID �Ō��
                    tblWfSample(k).WFSMP.HINBCW = .HINUP.hinban             ' �i��
                    tblWfSample(k).WFSMP.REVNUMCW = .HINUP.mnorevno         ' ���i�ԍ������ԍ�
                    tblWfSample(k).WFSMP.FACTORYCW = .HINUP.factory         ' �H��
                    tblWfSample(k).WFSMP.OPECW = .HINUP.opecond             ' ���Ə���
                End If
                tblWfSample(k).WFSMP.SMPKBNCW = tblNukishi(j).SMPKBNCW
                tblWfSample(k).WFSMP.TBKBNCW = tblNukishi(j).TBKBNCW        'TB�敪
                tblWfSample(k).WFSMP.SMCRYNUMCW = fnc_BsmpID(tblWfSample, k) '�T���v���u���b�NID���R�t���֐�
                tblWfSample(k).WFSMP.KSTAFFCW = txtStaffID.text             '�X�V�Ј�ID

                ''�S�U�֎��̌���GD���p���Ή�
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

    '��ۯ��̋��E��SXL���������ꍇUD��TB�ɕύX����
    If UBound(CngSmpID_UD) > 0 Then
        Call sub_Set_SMP_TB(tblWfSample())
    End If
End Sub

'*******************************************************************************
'*    �֐���        : sub_MakeTBCME042
'*
'*    �����T�v      : 1.SXL�Ǘ��e�[�u���֓o�^�����邽�߂ɍ\���̍\�z
'*                      (SXL�Ǘ��e�[�u���̍쐬)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************
Private Sub sub_MakeTBCME042()

    Dim udtTmpHin       As tFullHinban  ' �t�����i�Ԏ擾�p�\����
    Dim sNowHinban      As String       ' ���ݍs�̕i��
    Dim intNowIngotPos  As Integer      ' ���ݍs�̃C���S�b�g�ʒu
    Dim intOldIngotPos  As Integer
    Dim sTmpSXLID       As String       ' SXLID
    Dim intNowSxlLen    As Integer      ' SXL�̒���
    Dim intSPoint       As Integer      ' �s�ʒu�ۑ�
    Dim intFlg          As Integer      ' ���ʃt���O
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
                .SXLID = tblSXL.SXLID     'SXL���ʒu����쐬���Ȃ����ƁASXLID������Ă��܂��ׁA��SXLID�����̂܂܎g���悤�C��
                .CRYNUM = ""                        ' �����ԍ�
            Else
                .SXLID = Mid(tblWafInd(i).BLOCKID, 1, 10) & GetWafPos(tblWafInd(i).INGOTPOS)
                .CRYNUM = tblSXL.CRYNUM              ' �����ԍ�
            End If

            '�擪SXL�̊J�n�ʒu�͑O�H���̊J�n�ʒu�ƂȂ�
            If i = 1 Then
                .INGOTPOS = SIngotP
            Else
                .INGOTPOS = tblWafInd(i).INGOTPOS               ' �������J�n�ʒu
            End If
            .LENGTH = tblWafInd(i).LENGTH                       ' ����
            .hinban = tblWafInd(i).HINDN.hinban                 ' �i��
            .REVNUM = tblWafInd(i).HINDN.mnorevno               ' ���i�ԍ������ԍ�
            .factory = tblWafInd(i).HINDN.factory               ' �H��
            .opecond = tblWafInd(i).HINDN.opecond               ' ���Ə���
            .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI                 ' �Ǘ��H��
            .NOWPROC = PROCD_WFC_SOUGOUHANTEI                   ' ���ݍH��
            .LPKRPROCCD = MGPRCD_WFC_SOUGOUHANTEI               ' �ŏI�ʉߊǗ��H��
            .LASTPASS = PROCD_WFC_SOUGOUHANTEI                  ' �ŏI�ʉߍH��

            '' �y�i�ԂȂ�폜�敪�ƍŏI��ԋ敪��ς���
            If Trim(.hinban) = "Z" Then
                .DELCLS = "1"                                   ' �폜�敪
                .LSTATCLS = "H"                                 ' �ŏI��ԋ敪
            Else
                .DELCLS = "0"                                   ' �폜�敪
                .LSTATCLS = "T"                                 ' �ŏI��ԋ敪
            End If
            .HOLDCLS = "0"                                      ' �z�[���h�敪

            '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
            sprExamine.row = i
            sprExamine.col = 9
            j = sprExamine.TypeComboBoxCurSel + 1
            .BDCAUS = Trim$(tblPrcList(j).CODE)                 ' �s�Ǘ��R
            .COUNT = "0"                                        ' ����
        End With
    Next i
End Sub

'*******************************************************************************
'*    �֐���        : sub_MakeTBCMY007
'*
'*    �����T�v      : 1.SXL�m��w���e�[�u���֓o�^�����邽�߂ɍ\���̍\�z
'*                      (SXL�m��w���e�[�u���̍쐬)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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
    Dim sSmpId(2)       As String       '���L�����ID
    Dim k               As Integer
    Dim intCntSXL       As Integer      'SXLں��ސ�
    Dim udtRsHIN        As tFullHinban  '���R(Rs)�d�l�擾�i��
    Dim sRsData(10)     As String       '���R(Rs)�ް�
    Dim sRsPtn          As String       '���R�ް��擾�����
    Dim intSxlPtn       As Integer      '�o�^SXL�����
    Dim sHosho          As String       '�ۏؕ��@
'Add Start 2012/03/08 Y.Hitomi
    Dim iBlockHoshoCnt  As String       '�u���b�N�m��i�p���j��
'Add End   2012/03/08 Y.Hitomi

    
'                                       ��1 : �S�p��SXL
'                                       ��2 : ��Ǎ���SXL
'                                       ��3 : ���Ǎ���SXL
'                                       ��4 : SXL�̊Ԃ�Z
'                                       ��0 : �擾�ް��Ȃ�

    sSmpId(1) = ""       '�����ID(From)������
    sSmpId(2) = ""       '�����ID(To)������
    intCntSXL = UBound(tblWfSample)
'Add Start 2012/03/08 Y.Hitomi
    iBlockHoshoCnt = 0
'Add End   2012/03/08 Y.Hitomi

    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With sprExamine
        m = .MaxRows
        ReDim tblSxlKSiji(m)
        j = 0
        For i = 1 To m - 1 Step 2
            '' Z�i�ԂȂ�e�[�u���쐬
            .row = i

            '�G�s��s�]���ǉ��Ή�
            .GetText 37, i, vFlg
            .GetText 2, i, vHinban

            If vHinban = "Z" Then
                j = j + 1

                '�G�s��s�]���ǉ��Ή�
                .col = 39
                sNowBlockID = Mid(tblSXL.CRYNUM, 1, 9) & .text

                If i = 1 Then   '�����ʒu����SXLID�͍��Ȃ�
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
                    If sHosho = 1 Then ''��ۯ��ۏ�
                    
                        'Add Start 2012/03/08 Y.Hitomi�@�u���b�N�m��i�p���j�́A1��݂̂Ƃ���B
                        If iBlockHoshoCnt = 0 Then
                            iBlockHoshoCnt = iBlockHoshoCnt + 1
                        Else
                            j = j - 1
                            Exit For
                        End If
                        'Add End   2012/03/08 Y.Hitomi
                        
                        '' �u���b�NID
                        tblSxlKSiji(j).BLOCKID = sNowBlockID
                        '' �T���v��ID(From)
                        tblSxlKSiji(j).SAMPLE_FROM = ""
                        '' �T���v��ID(To)
                        tblSxlKSiji(j).SAMPLE_TO = ""
                    ElseIf sHosho = 2 Then 'WF�ۏ�
                        '' �u���b�NID
                        tblSxlKSiji(j).BLOCKID = ""
                        '' �T���v��ID(From)
                        If i = 1 Then
                            tblSxlKSiji(j).SAMPLE_FROM = tblSXL.WFSMP(1).REPSMPLIDCW
                        Else
                            '�G�s��s�]���ǉ��Ή�
                            .GetText 38, i, vChgSmpId
                            tblSxlKSiji(j).SAMPLE_FROM = vChgSmpId  '�T���v��ID����蒼���K�v�͂Ȃ�
                        End If
    
                        '�֘A��ۯ��Ή� 08/08/25 ooba
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
                        
                        '' �T���v��ID(To)
                        If i = m - 1 Then
                            tblSxlKSiji(j).SAMPLE_TO = tblSXL.WFSMP(2).REPSMPLIDCW
                        Else
                            '�G�s��s�]���ǉ��Ή�
                            .GetText 38, i + 1, vChgSmpId
                            tblSxlKSiji(j).SAMPLE_TO = vChgSmpId
                        End If
    
                        ''�p�������ID���L�Ή�
                        sSmpId(1) = tblSxlKSiji(j).SAMPLE_FROM
                        sSmpId(2) = tblSxlKSiji(j).SAMPLE_TO
    
                        '�o�^�ݸ�ق̊e�������ڂɑ΂������ٗL������������
                        If chkComSAMPL(tblSXL.SXLID, sSmpId(1), sSmpId(1)) = FUNCTION_RETURN_SUCCESS Then
                            '�������ЂƂ��Ȃ��ꍇ�A���L�����ID��o�^����
                            If sSmpId(1) <> tblSxlKSiji(j).SAMPLE_FROM Then
                                tblSxlKSiji(j).SAMPLE_FROM = sSmpId(1)
                            End If
                        End If
    
                        '�o�^�ݸ�ق̊e�������ڂɑ΂������ٗL������������
                        If chkComSAMPL(tblSXL.SXLID, sSmpId(2), sSmpId(2)) = FUNCTION_RETURN_SUCCESS Then
                            '�������ЂƂ��Ȃ��ꍇ�A���L�����ID��o�^����
                            If sSmpId(2) <> tblSxlKSiji(j).SAMPLE_TO Then
                                tblSxlKSiji(j).SAMPLE_TO = sSmpId(2)
                            End If
                        End If
                        ''�p�������ID���L�Ή�
                    End If
                End If

                '' �m��i��
                tblSxlKSiji(j).hinban = tblSXL.hinban & Format(tblSXL.REVNUM, "00")

                '' �敪�R�[�h
                .row = i
                .col = 9
                intFuryoCmb = .TypeComboBoxCurSel + 1
                tblSxlKSiji(j).KUBUN = Trim$(tblPrcList(intFuryoCmb).CODE)       ' �s�Ǘ��R

                '' �g�����U�N�V����ID
                tblSxlKSiji(j).TXID = "TX853I"

                ''���R�ް��擾
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
'*    �֐���        : sub_MakeTBCMW005
'*
'*    �����T�v      : 1.�e�[�u���o�^�p��WF����������уe�[�u���\���̂̐ݒ�
'*                      (WF����������уe�[�u���\���̐ݒ�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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
        .CRYNUM = tblSXL.CRYNUM             ' �����ԍ�
        .INGOTPOS = tblSXL.INGOTPOS         ' �C���S�b�g�ʒu
        .CRYLEN = tblSXL.COUNT              ' ����
        .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI ' �Ǘ��H���R�[�h
        .PROCCODE = PROCD_WFC_SOUGOUHANTEI  ' �H���R�[�h
        .SXLID = tblSXL.SXLID               ' SXLID
        .CODE = sCode                       ' �敪�R�[�h
        .TSTAFFID = txtStaffID.text         ' �o�^�Ј�ID
        .KSTAFFID = ""                      ' �X�V�Ј�ID
    End With
End Sub

'*******************************************************************************
'*    �֐���        : sub_MakeTBCMW006
'*
'*    �����T�v      : 1.�e�[�u���o�^�p�ɐU�֔p�����уe�[�u���\���̂̐ݒ�
'*                      (�U�֔p�����уe�[�u���\���̂̐ݒ�)
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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
        ReDim tblHuriHai(0) '�\���̂͊Y���s�̎��̂ݑ��₷
        intRowCnt = 0

        For i = 1 To m
            .row = i

            '�G�s��s�]���ǉ��Ή�
            sprExamine.GetText 37, i, vNukisiFlg

            If (vNukisiFlg = 1) Or CheckGetSampleID(i) = True Then    '�����s��������
                If i = m Then
                    Exit For
                End If
                intRowCnt = intRowCnt + 1

                ReDim Preserve tblHuriHai(intRowCnt) '�\���̂͊Y���s�̎��̂ݑ��₷
                .col = 5

                '�擪SXL�̊J�n�ʒu�͑O�H���̊J�n�ʒu�ƂȂ�
                If i = 1 Then
                    tblHuriHai(intRowCnt).INGOTPOS = SIngotP
                    intNowIngotPos = SIngotP
                Else
                    tblHuriHai(intRowCnt).INGOTPOS = .text
                    intNowIngotPos = .text
                End If

                '�ŏISXL�̏I���ʒu�͑O�H���̏I���ʒu�ƂȂ�
                If i = m - 2 Then
                    intDwnIngotPos = EIngotP
                Else
                    .row = i + 1
                    intDwnIngotPos = .text
                    .row = i
                End If

                .col = 2
                If .row Mod 2 = 0 Then
                    Call GetLastHinban("", udtTmpHin)           ' �t���i��
                    sHin = ""
                Else
                    Call GetLastHinban(.text, udtTmpHin)        ' �t���i��
                    sHin = Trim$(.text)
                End If

                With tblHuriHai(intRowCnt)
                    .CRYNUM = tblSXL.CRYNUM                     ' �����ԍ�
                    If UBound(tblWafInd()) = 2 Then             '2�s�̎������A�����擾�����ύX
                        .CRYLEN = EIngotP - SIngotP
                    Else
                        .CRYLEN = intDwnIngotPos - intNowIngotPos     ' ����
                    End If
                    .KRPROCCD = MGPRCD_WFC_SOUGOUHANTEI         ' �Ǘ��H���R�[�h
                    .PROCCODE = PROCD_WFC_SOUGOUHANTEI          ' �H���R�[�h
                    .TRANCLS = IIf(sHin = "Z", "1", "0")        ' �����敪
                    .DUOGNUM = tblSXL.hinban                    ' �]�p���i��
                    .DUOGREV = tblSXL.REVNUM                    ' �]�p���i�� ���i�ԍ������ԍ�
                    .DUOGFACT = tblSXL.factory                  ' �]�p���i�� �H��
                    .DUOGOPCD = tblSXL.opecond                  ' �]�p���i�� ���Ə���
                    .DUNWNUM = sHin                             ' �]�p��i��
                    .DUNWREV = udtTmpHin.mnorevno               ' �]�p��i�� ���i�ԍ������ԍ�
                    .DUNWFACT = udtTmpHin.factory               ' �]�p��i�� �H��
                    .DUNWOPCD = udtTmpHin.opecond               ' �]�p��i�� ���Ə���
                    .TSTAFFID = txtStaffID.text                 ' �o�^�Ј�ID
                    .KSTAFFID = ""                              ' �X�V�Ј�ID
                    .MUKESAKI = sCmbMukesaki
                End With
            End If
        Next i
    End With
End Sub

'*******************************************************************************
'*    �֐���        : fnc_Nukisi_LOAD_DISP
'*
'*    �����T�v      : 1.SXLID����u���b�N�h�c�i�͂ݏo�������j�擾
'*                    2.�u���b�N�ASXLID���̃u���b�N�r�d�p�i�u���b�N���A�ԁj��
'*                      �ő�A�ŏ��l�Ƃ��̑����擾
'*                    3.�������ڎ擾
'*                    4.�U������(�����\��)
'*                    5.Warp/�����p�x���\��
'*                    6.�`�F�b�N�{�b�N�X�������\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '�O��ʂ����SXLID���擾
     ReDim tSXLID(0)
     tSXLID(0).SXLID = tblSXL.SXLID

    ''SXLID����u���b�N�h�c�i�͂ݏo�������j�擾
    If DBDRV_BLOCKIDGET() = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr("ESXL2")
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''�u���b�N�ASXLID���̃u���b�N�r�d�p�i�u���b�N���A�ԁj�̍ő�A�ŏ��l�Ƃ��̑����擾
    If DBDRV_MIN_MAX_SEQGET(intWfNum) = FUNCTION_RETURN_FAILURE Then
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '�������ڎ擾
    If DVDRV_KENSA_KOUMOKU(tKensa()) = FUNCTION_RETURN_FAILURE Then
        fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '�f��ʕ\��
    Call sub_ExamineDisp(tKensa(), intIngotpos(), intWfNum)

    '�U������(�����\��)
    lblMsg.Caption = ""
    ReDim tWarpMeasG(0)
    ReDim tKakuMeasG(0)
    'Add Start 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�
    ReDim tKakuXMeasG(0)
    ReDim tKakuYMeasG(0)
    'Add End 2011/07/21 SMPK Nakamura �����ʌX���`�F�b�N�ǉ��Ή�

    For i = 1 To UBound(tMapHin)
        tMapHinG = tMapHin(i)
        For j = 1 To 2
            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                intRtn = funChkFurikaeShiyou("CW763", txtKSXLID.text, _
                                            tMapHinG.HIN, tMapHinG.HIN, _
                                            intErrCode, intErrMsg, typ_b, typ_CType, 0)

                tMapHin(i).WARPFLG = tMapHinG.WARPFLG       'Warp�U�������׸޾��
                tMapHin(i).KAKUFLG = tMapHinG.KAKUFLG       '�����p�x�U�������׸޾��

                '����NG
                If intRtn = 1 Then
                    If Not tMapHinG.KAKUFLG Then
                        lblMsg.Caption = "Warp����G���[�@�i�ԐU�ւ��s���Ă��������B"
                    Else
                        lblMsg.Caption = "�����p�x����G���[�@�i�ԐU�ւ��s���Ă��������B"
                    End If

                '�U�������װ
                ElseIf intRtn < 0 Then
                    fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_FAILURE
                    lblMsg.Caption = intErrMsg
                    Exit Function
                End If
            End If
        Next j
    Next i

    If bMapWarpFlg Then lblMsg.Caption = "WF�}�b�v��Warp���т̕s��v�G���["

    'Warp/�����p�x���\��
    Call WarpKakuDisp(Me)

    '�`�F�b�N�{�b�N�X�������\��
    Call sub_LOADDISP_ADD_CHECKBOX

    '�U�փ`�F�b�N�ǉ��ɂ��C��
    Call sub_FurikaeMotoDataSet

    Erase intIngotpos

    fnc_Nukisi_LOAD_DISP = FUNCTION_RETURN_SUCCESS

End Function

'*******************************************************************************
'*    �֐���        : sub_ExamineDisp
'*
'*    �����T�v      : 1.sprExamine�ɕ\������
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^             ,����
'*�@�@      �@�@      udtKensa        ,I  ,typ_XSDCW      ,�V�T���v���Ǘ�(SXL)
'*�@�@      �@�@      intIngotpos()   ,I  ,Integer        ,
'*�@�@      �@�@      intWfNum        ,I  ,Integer        ,WF����
'*
'*    �߂�l        : �Ȃ�
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
    Dim udtChkHin       As tFullHinban  '�����p12���i��
    Dim sBlkSmplID      As String       '��������َ���
    Dim sCryindFlg      As String       '�������FLG
    Dim vGetHinban      As Variant      '�i��
    Dim vGetBlockID     As Variant      '��ۯ�ID
    Dim vGetBlockSEQ_S  As Variant      '��ۯ�SEQ(Start)
    Dim vGetBlockSEQ_E  As Variant      '��ۯ�SEQ(End)
    Dim vlvGetWFstatus  As Variant      'WF���
    Dim sBuf            As String
    Dim sHinban         As String

    intUCount = UBound(tExamine)

    '�i�Ԃ�1��ǉ����邱�Ƃɂ���̕ύX
    With sprExamine
        ''���[�v�J�n
        .MaxRows = 0
        .MaxRows = 1

        '�������ш��p���ް�������
        CpyCrySmpl.TsmplidGD = ""           'TOP_�����ID(GD)
        CpyCrySmpl.TindGD = ""              'TOP_���FLG(GD)
        CpyCrySmpl.BsmplidGD = ""           'BOT_�����ID(GD)
        CpyCrySmpl.BindGD = ""              'BOT_���FLG(GD)

        For i = 0 To intUCount
            blSampleChk = False
            .BlockMode = True
            .row = i + 1

            '' �u���b�NID
            If i = 0 Then   'i = 0�����ƁAelseif���A�Ɠ�������������
                vInsertData = Right(tExamine(i).LOTID, 3)
                .SetText 1, i + 1, vInsertData
                .BlockMode = True
                .backColor = COLOR_DISABLE
                .BlockMode = False
                .CellBorderType = 4     '�r����
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
                .CellBorderType = 4     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
            End If

            ''�@�i��
            If (i = 0) Then
                .col = 2
                .SetText 2, i + 1, tExamine(i).hinban
                sHinban = tExamine(i).hinban

                .BlockMode = True
                .backColor = vbWhite
                .BlockMode = False
                .CellBorderType = 4     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
                .Lock = False

                '' �v�e����
                .col = 7
                .SetText 7, i + 1, tExamine(i).CURRWPCS

                '�G�s��s�]���ǉ��Ή�
                .SetText 43, i + 1, tExamine(i).CURRWPCS
                .CellBorderType = 4     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16

                '�i��2�ǉ�
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
                .CellBorderType = 4     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
                .Lock = False

                '' �v�e����
                .col = 7
                .SetText 7, i + 1, tExamine(i).CURRWPCS

                '�G�s��s�]���ǉ��Ή�
                .SetText 43, i + 1, tExamine(i).CURRWPCS
                .CellBorderType = 4     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16

                '�i��2�ǉ�
                .col = 3
                .SetText 3, i + 1, Format(tExamine(i).REVNUM, "00") & tExamine(i).factory & tExamine(i).opecond
                .backColor = COLOR_DISABLE
                .BlockMode = False
            Else
                '�f�u���b�N�o
                sBuf = GetMukesaki(sHinban)
                .SetText 2, i + 1, sBuf
                sBaseMukesaki = sCmbMukeName

                .BlockMode = True
                .backColor = COLOR_DISABLE
                .BlockMode = False
            End If

            '�f�u���b�N�o
            .SetText 4, i + 1, CStr(tExamine(i).RTOP_POS)
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            ''�@�����o
            .SetText 5, i + 1, CStr(tExamine(i).RITOP_POS)
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            '�f�}�b�v�ʒu
            .SetText 6, i + 1, Trim(CStr(tExamine(i).BLOCKSEQ))
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            '�f�v�e���
            'WF��Ԕ���
            .col = 8
            .row = i + 1
            .BlockMode = True
            .backColor = COLOR_DISABLE
            .BlockMode = False

            Select Case CStr(tExamine(i).WFSTA)
                Case gsWF_STA_0   '�ʏ�
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                    '�T���v���t���O����
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
                Case gsWF_STA_1   '���L
                    .SetText 8, i + 1, gsWF_STA_NORMAL
                Case gsWF_STA_4   '����
                    '�T���v���t���O����
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

            '�f�s�ǋ敪
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
                .CellBorderType = 8     '�r����
                .CellBorderStyle = 1
                .CellBorderColor = vbBlack
                .Action = 16
            End If

            '' �敪�R���{�̐ݒ�
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
                Case "1" '����
                    .SetText 10, i + 1, gsWF_SMPL_JOINT

                    '�G�s��s�]���ǉ��Ή�
                    .SetText 38, i + 1, sBkSMPLID
                Case "0" '�ʏ�
                    If (tExamine(i).SHAFLAG <> "0") Then
                        .SetText 10, i + 1, sSmplID

                        '�G�s��s�]���ǉ��Ή�
                        .SetText 38, i + 1, sBkSMPLID
                    End If
                Case "4" '����
                    If (tExamine(i).SHAFLAG <> "0") Then
                        .SetText 10, i + 1, sSmplID
                        '�G�s��s�]���ǉ��Ή�
                        .SetText 38, i + 1, sBkSMPLID
                    End If
                Case Else
            End Select

            .col = 10
            .CellBorderType = 16    '�r���O�g
            .CellBorderStyle = 1
            .CellBorderColor = vbBlack
            .Action = 16

            '�������ڕ\��
            .col = 11    '�R�`�X���

            '�G�s��s�]���ǉ��Ή�
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
                    If .text = "1" Then    '0:�����Ȃ� ,1:�ʏ�,2:���f,3:����
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
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
'''                        .backColor = vbYellow
'''                        .ForeColor = vbYellow
                        
                        ' ��ڰ
                        .backColor = COLOR_CryJitsu
                        .ForeColor = COLOR_CryJitsu
                        '��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END
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

                ''�c���_�f�ǉ�
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
                    '�G�s��s�]���ǉ��Ή�
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

                'GD�ǉ�
                If IsNumeric(udtKensa(j).WFINDGDCW) = True Then
                    '�G�s��s�]���ǉ��Ή�
                    .col = 28
                    .col2 = 28

                    .BlockMode = True
                    .text = IIf(udtKensa(j).WFINDGDCW = "0", "", udtKensa(j).WFINDGDCW)
                    .BlockMode = False

                    '����
                    If .text = "1" And udtKensa(j).WFHSGDCW <> "1" Then
                        .backColor = vbBlack
                    '���f
                    ElseIf .text = "2" And udtKensa(j).WFHSGDCW <> "1" Then
                        .backColor = vbYellow
                        .ForeColor = vbYellow
                    '��������
                    ElseIf (.text = "1" Or .text = "2") And udtKensa(j).WFHSGDCW = "1" Then
                        .backColor = COLOR_CryJitsu
                        .ForeColor = COLOR_CryJitsu
                    End If
                End If

                '�G�s��s�]���ǉ��Ή�
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

            '' ����GD���f�Ή�
            'SXL��TOP�ʒu��BOT�ʒu�̌������т��擾����
            If (i = 0) Or (i = intUCount) Then
                udtChkHin.hinban = ""
                udtChkHin.mnorevno = 0
                udtChkHin.factory = ""
                udtChkHin.opecond = ""

                '���p���ް������݂���΂��̌��������ID�^�������FLG���
                '--GD
                If funBlkSmpDataGet(udtKensa(j).SMCRYNUMCW, udtKensa(j).TBKBNCW, udtKensa(j).INPOSCW, _
                                        udtChkHin, 1, sBlkSmplID, sCryindFlg) = FUNCTION_RETURN_SUCCESS Then
                    'TOP
                    If udtKensa(j).TBKBNCW = "T" Then
                        CpyCrySmpl.TsmplidGD = sBlkSmplID       '�����ID(GD)
                        CpyCrySmpl.TindGD = sCryindFlg          '���FLG(GD)
                    'BOT
                    ElseIf udtKensa(j).TBKBNCW = "B" Then
                        CpyCrySmpl.BsmplidGD = sBlkSmplID       '�����ID(GD)
                        CpyCrySmpl.BindGD = sCryindFlg          '���FLG(GD)
                    End If
                End If
            End If

            '�폜�{�^���̏����֎~�t���O
            If (i = 0) Or (i = intUCount) Then
                '�G�s��s�]���ǉ��Ή�
                .SetText 37, i + 1, "1"
            Else
                '�G�s��s�]���ǉ��Ή�
                .SetText 37, i + 1, "3"
            End If

            '�G�s��s�]���ǉ��Ή�
            .SetText 39, i + 1, Right(tExamine(i).LOTID, 3)

            '�\���t�B�[���h�ɕi�ԕۑ�
            '�G�s��s�]���ǉ��Ή�
            .SetText 40, i + 1, tExamine(i).hinban

            .BlockMode = False
            .MaxRows = .MaxRows + 1
        Next i

        .MaxRows = .MaxRows - 1

        '�F������
        .col = 1    '�P���
        .col2 = 1
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = COLOR_DISABLE
        .BlockMode = False

        .col = 2    '�Q���
        .col2 = 2
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = vbWhite
        .BlockMode = False

        .col = 3    '3���
        .col2 = 3
        .row = 1
        .row2 = .MaxRows
        .BlockMode = True
        .backColor = COLOR_DISABLE
        .BlockMode = False

        .col = 4    '�R�`�X���
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

        'WFϯ�ߏ�̕i�ԏ��擾
        ReDim tMapHin(0)
        m = 0
        For i = 1 To .MaxRows Step 2
            '�i�Ԏ擾
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

                        '��ۯ�ID�擾
                        .GetText 39, i, vGetBlockID
                        tMapHin(m).BLOCKID = left(txtCryNum.text, 9) & CStr(vGetBlockID)

                        '��ۯ�SEQ�擾
                        .GetText 6, i, vGetBlockSEQ_S
                        .GetText 6, i + 1, vGetBlockSEQ_E
                        tMapHin(m).BLKSEQ_S = CInt(vGetBlockSEQ_S)
                        tMapHin(m).BLKSEQ_E = CInt(vGetBlockSEQ_E)

                        '�U�������׸�
                        tMapHin(m).WARPFLG = False
                        tMapHin(m).KAKUFLG = False

                        'Add Start 2011/04/25 SMPK Miyata
                        tMapHin(m).XTALCS = txtCryNum.text      '�����ԍ�
                        .GetText 5, i, vGetIngot
                        tMapHin(m).INPOSCS_S = CInt(vGetIngot)  '�������ʒu(Start)
                        .GetText 5, i + 1, vGetIngot
                        tMapHin(m).INPOSCS_E = CInt(vGetIngot)  '�������ʒu(End)
                        'Add End   2011/04/25 SMPK Miyata

                        Exit For
                    End If
                Next j
            End If
        Next i

        '�M�������f�����ǉ�
        'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
        For i = 1 To .MaxRows Step 2
            '�i�Ԏ擾
            .GetText 2, i, vGetHinban
            If vGetHinban <> vbNullString And vGetHinban <> "Z" And vGetHinban <> "G" Then
                .col = 10
                'WF��Ԏ擾(TOP)
                .GetText 8, i, vlvGetWFstatus
                'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
                If Trim(vlvGetWFstatus) = gsWF_STA_SIJI_KEKKA Then
                    .row = i
                    .backColor = f_cmbc039_3.Label12.backColor
                End If
                'WF��Ԏ擾(BOTTOM)
                .GetText 8, i + 1, vlvGetWFstatus
                'WF��Ԃ�"����"�̏ꍇ�A�T���v��ID�̔w�i�F�𐅐F�ɂ���
                If Trim(vlvGetWFstatus) = gsWF_STA_SIJI_KEKKA Then
                    .row = i + 1
                    .backColor = f_cmbc039_3.Label12.backColor
                End If

            End If
        Next i
    End With

    '�˂炢��R�̏����������ʔ͈͂ɕύX���ĕ\������
    With sprExamine
        .GetText 5, 1, vGetIngot
        intTingot = CInt(vGetIngot)
        .GetText 5, .MaxRows, vGetIngot
        intBingot = CInt(vGetIngot)
    End With

    '�g�b�v����R�l
    Call sub_Top_Btm_TEIKOU(intTingot, dblteikoku)
    txtTopRsltR.text = CStr(Format(dblteikoku, "0.0000"))

    '�{�g������R�l
    Call sub_Top_Btm_TEIKOU(intBingot, dblteikoku)
    txtBotRsltR.text = CStr(Format(dblteikoku, "0.0000"))

    End Sub

'*******************************************************************************************
'*    �֐���        : sprExamine_ButtonClicked
'*
'*    �����T�v      : 1.�T���v���t���O�I�t�ɂ���
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
'*�@�@      �@�@      col         ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̗�ԍ�
'*�@�@      �@�@      Row()       ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̍s�ԍ�
'*�@�@      �@�@      ButtonDown  ,I  ,Integer ,�ێ��^�{�^���̏�ԁi�`�F�b�N�{�b�N�X�̏�ԁj
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sprExamine_ButtonClicked(ByVal col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    With sprExamine
        If (col <> 1) Then
            Exit Sub
        End If

        '' �T���v���t���O�I�t
        bSampFlag = False
    End With
End Sub

'*******************************************************************************************
'*    �֐���        : sprExamine_Click
'*
'*    �����T�v      : 1.sprExamine_Click����
'*                      (�������ڂ̗L����ύX)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
'*�@�@      �@�@      col         ,I  ,Long    ,��
'*�@�@      �@�@      Row         ,I  ,Long    ,�s
'*
'*    �߂�l        : �Ȃ�
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

    '�i�Ԃ�1��ǉ��������Ƃɂ���̕ύX
    With sprExamine
        If (col < 10) Then
            Exit Sub
        End If

        '�i�Ԃ�Z�̂Ƃ��͏��������Ȃ�
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

        '�G�s��s�]���ǉ��Ή�
        If col = 27 Or col = 35 Then
            .GetText 7, row, vSamp

            If vSamp = "����" Then
                Exit Sub
            End If

            '�G�s��s�]���ǉ��Ή�
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
                    udtHinban.hinban = vbNullString   '������
                    udtHinban.factory = vbNullString
                    udtHinban.mnorevno = 0
                    udtHinban.opecond = vbNullString
                    For intLoopCnt = 1 To UBound(tblsiyou)  '�����d�l�\���̂Ɣ�r���A�Y���̃t�����i�Ԏ擾
                        If tblHinbanRs(intLoopCnt).HIN.hinban = Trim(vGetHin) Then
                            udtHinban.hinban = tblHinbanRs(intLoopCnt).HIN.hinban
                            udtHinban.factory = tblHinbanRs(intLoopCnt).HIN.factory
                            udtHinban.mnorevno = tblHinbanRs(intLoopCnt).HIN.mnorevno
                            udtHinban.opecond = tblHinbanRs(intLoopCnt).HIN.opecond
                            Exit For
                        End If
                    Next
                    If udtHinban.hinban = vbNullString Then   '�Y���̕i�Ԃ��\���̂ɂȂ������ꍇ�������Ȃ�
                        Exit Sub
                    End If
                    If scmzc_getE036(udtHinban, sOT1, sOT2, sMAI1, sMAI2) = FUNCTION_RETURN_FAILURE Then
                        Exit Sub    '�G���[�̏ꍇ�������Ȃ�
                    End If

                    ''�c���_�f�������ڒǉ��ɂ��ύX
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
                    ' �K��E036�e�[�u�����`�F�b�N����悤�ɕύX�����̂ŏ����ǉ�
                    ' ���i�����j�����i�����j
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
                    '��--- 2010/01/20 SIRD�Ή� SPK habuki REP START
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
                            Call sub_Cksmp(row, col, .text)     ' ��ڰ����
                            .ForeColor = vbBlack
                            .backColor = vbBlack
                        ElseIf vKensa = "1" Then
                            vKensa = "2"
                            Call sub_Cksmp(row, col, .text)     ' ������ڰ
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
                    
                    '��--- 2010/01/20 SIRD�Ή� SPK habuki REP  END
                    .SetText col, row, vKensa
                Else
                    '����ٖ������̏ꍇ�͏����𔲂���
                    If bSampFlag = False Then Exit Sub
                    .GetText col, row, vKensa
                    '���f�^�������с@���@����
                    If vKensa = "2" Then
                        vKensa = "1"
                        Call sub_Cksmp(row, col, .text)
                        .ForeColor = vbBlack
                        .backColor = vbBlack
                    ElseIf vKensa = "1" Then
                        vKensa = "2"
                        Call sub_Cksmp(row, col, .text)
                        '�����@���@���f
                        If tblNukishi(row).WFHSGDCW <> "1" Then
                            .ForeColor = vbYellow
                            .backColor = vbYellow
                        '�����@���@��������
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
                        .GetText 44, row, vSMPLID   'UD�ʕۑ��̃T���v��ID
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
                        If intk = inth Then '�������ڂ��S�ĉ��F(���f)
                            .SetText 10, row, gsWF_SMPL_JOINT
                            .SetText 8, row, gsWF_STA_NORMAL

                            '�ʂ̍\���̂���R�s�[
                            vSMPLID = tWafk(row).SAMPLEID
                            .SetText 38, row, vSMPLID
                        Else '���f�[�^���Ƃ�ꍇ
                            .GetText 44, row, vSMPLID
                            .SetText 38, row, vSMPLID
                        End If
                    End If
                End If
            End If
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START(TOP,BOT�͌����ς݂̂��ߕύX�s��)
'''        Else
'''            '<vNukisiFlg="1">�E�E�Ecol=37
'''            .col = col: .row = row
'''            If .Lock = False Then
'''                If col = 19 Then
'''                    If .text = "1" Then
'''                        .SetText col, row, "2"  ' ������ڰ
'''                        .backColor = COLOR_CryJitsu
'''                        .ForeColor = COLOR_CryJitsu
'''                    Else
'''                        .SetText col, row, "1"  ' ��ڰ����
'''                        .backColor = vbBlack
'''                        .ForeColor = vbBlack
'''                    End If
'''                End If
'''            End If
        '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD  END
        End If
    End With
End Sub

'*******************************************************************************************
'*    �֐���        : sub_Cksmp
'*
'*    �����T�v      : 1.�������ڂ��N���b�N���ꂽ�Ƃ��ɃT���v��ID��ύX����
'*
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
'*�@�@      �@�@      R           ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̗�ԍ�
'*�@�@      �@�@      C           ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̍s�ԍ�
'*�@�@      �@�@      sHkbn        ,I  ,String  ,RC�őI�����ꂽText�̒l
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_Cksmp(R As Long, C As Long, sHkbn As String)
    If sHkbn = "2" Then '���F����
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
            ''�c���_�f�ǉ�
            Case 26
                tblNukishi(R).WFSMPLIDAOICW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESAOICW = "0"
            Case 27
                tblNukishi(R).WFSMPLIDOT1CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESOT1CW = "0"
            ' �G�s��s�]���ǉ�
            Case 35
                tblNukishi(R).WFSMPLIDOT2CW = tblNukishi(R).REPSMPLIDCW
                tblNukishi(R).WFRESOT2CW = "0"
            'GD�ǉ�
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

        'Z���ǂ����̔��f������Z�̂Ƃ��̓T���v��ID���R�s�[�����уt���O��0�ɂ���
        Call sub_Zsample(R, C, sHkbn)
    ElseIf sHkbn = "1" Then '���������F
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
            ''�c���_�f�ǉ�
            Case 26
                tblNukishi(R).WFSMPLIDAOICW = tblKns(R).WFSMPLIDAOICW
                tblNukishi(R).WFRESAOICW = tblKns(R).WFRESAOICW
            Case 27
                tblNukishi(R).WFSMPLIDOT1CW = tblKns(R).WFSMPLIDOT1CW
                tblNukishi(R).WFRESOT1CW = tblKns(R).WFRESOT1CW
            ' �G�s��s�]���ǉ�
            Case 35
                tblNukishi(R).WFSMPLIDOT2CW = tblKns(R).WFSMPLIDOT2CW
                tblNukishi(R).WFRESOT2CW = tblKns(R).WFRESOT2CW
            'GD�ǉ�
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
'*    �֐���        : sub_Zsample
'*
'*    �����T�v      : 1.��܂��͉���Z�i�Ԃ�������
'*�@�@�@�@�@�@�@�@�@�@�@Z�ɑI�����ꂽ�f�[�^�̃T���v��ID���R�s�[�����уt���O��0�ɂ���
'*�@�@�@�@�@�@�@�@�@�@�@�\���̂Ƀf�[�^���Z�b�g����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
'*�@�@      �@�@      R           ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̗�ԍ�
'*�@�@      �@�@      C           ,I  ,Long    ,�{�^�����N���b�N���ꂽ�Z���̍s�ԍ�
'*�@�@      �@�@      sHkbn        ,I  ,String  ,RC�őI�����ꂽText�̒l
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_Zsample(R As Long, C As Long, sHkbn As String)
    '�N���b�N�����i�Ԃ̏�i�Ԃ܂��͉��i�Ԃ�Z�������ꍇ
    'Z�ɑI�����ꂽ�f�[�^�̃T���v��ID���R�s�[�����уt���O��0�ɂ���
    '�\���̂Ƀf�[�^���Z�b�g����
    '3�̂Ƃ�(��i�Ԃ�Z)
    If R <= 2 Then
        Exit Sub
    End If

    If sHkbn = "2" Then '���F������
        If R Mod 2 = 0 Then
        '4�̂Ƃ�(���i��Z)
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
                    ''�c���_�f
                    Case 26
                        tblNukishi(R + 1).WFSMPLIDAOICW = tblNukishi(R).WFSMPLIDAOICW
                        tblNukishi(R + 1).WFRESAOICW = "0"
                    Case 27
                        tblNukishi(R + 1).WFSMPLIDOT1CW = ""
                        tblNukishi(R + 1).WFRESOT1CW = "0"
                    ' �G�s��s�]���ǉ�
                    Case 35
                        tblNukishi(R + 1).WFSMPLIDOT2CW = ""
                        tblNukishi(R + 1).WFRESOT2CW = "0"
                    'GD�ǉ�
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
            '3�̂Ƃ�(��i��Z)
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
                    ''�c���_�f�ǉ�
                    Case 26
                        tblNukishi(R - 1).WFSMPLIDAOICW = tblNukishi(R).WFSMPLIDAOICW
                        tblNukishi(R - 1).WFRESAOICW = "0"
                    Case 27
                        tblNukishi(R - 1).WFSMPLIDOT1CW = ""
                        tblNukishi(R - 1).WFRESOT1CW = "0"
                    ' �G�s��s�]���ǉ��Ή�
                    Case 35
                        tblNukishi(R - 1).WFSMPLIDOT2CW = ""
                        tblNukishi(R - 1).WFRESOT2CW = "0"
                    'GD�ǉ�
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
    ElseIf sHkbn = "1" Then  '���������F
        If R Mod 2 = 0 Then
        '4�̂Ƃ�(���i��Z)
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
                    ''�c���_�f�ǉ�
                    Case 26
                        tblNukishi(R + 1).WFSMPLIDAOICW = tblKns(R + 1).WFSMPLIDAOICW
                        tblNukishi(R + 1).WFRESAOICW = tblKns(R + 1).WFRESAOICW
                    Case 27
                        tblNukishi(R + 1).WFSMPLIDOT1CW = tblKns(R + 1).WFSMPLIDOT1CW
                        tblNukishi(R + 1).WFRESOT1CW = tblKns(R + 1).WFRESOT1CW
                    ' �G�s��s�]���ǉ�
                    Case 35
                        tblNukishi(R + 1).WFSMPLIDOT2CW = tblKns(R + 1).WFSMPLIDOT2CW
                        tblNukishi(R + 1).WFRESOT2CW = tblKns(R + 1).WFRESOT2CW
                    'GD�ǉ�
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
                    ''�c���_�f
                    Case 26
                        tblNukishi(R - 1).WFSMPLIDAOICW = tblKns(R - 1).WFSMPLIDAOICW
                        tblNukishi(R - 1).WFRESAOICW = tblKns(R - 1).WFRESAOICW
                    Case 27
                        tblNukishi(R - 1).WFSMPLIDOT1CW = tblKns(R - 1).WFSMPLIDOT1CW
                        tblNukishi(R - 1).WFRESOT1CW = tblKns(R - 1).WFRESOT1CW
                    ' �G�s��s�]���ǉ��Ή�
                    Case 35
                        tblNukishi(R - 1).WFSMPLIDOT2CW = tblKns(R - 1).WFSMPLIDOT2CW
                        tblNukishi(R - 1).WFRESOT2CW = tblKns(R - 1).WFRESOT2CW
                    'GD�ǉ�
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
''*    �֐���        : Combo1_Change
''*
''*    �����T�v      : 1.�g�p���Ă��Ȃ��C������H
''*
''*
''*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
''*�@�@      �@�@      �Ȃ�
''*
''*    �߂�l        : �Ȃ�
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
'                Case intConSprChg_0  '�S���w��
'                    .row = intLoopCnt
'                    .RowHidden = False
'                Case intConSprChg_1  '�Ǖi�w��
'                    '�G�s��s�]���ǉ��Ή�
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_1 Then  '�Ǖi�ȊO��������A��\��
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        .row = intLoopCnt
'                        .RowHidden = False
'                    End If
'                Case intConSprChg_2  '�T���v���w��
'                    '�G�s��s�]���ǉ��Ή�
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_2 Then  '�T���v���ȊO��������A��\��
'                        .row = intLoopCnt
'                        .RowHidden = True
'                    Else
'                        .row = intLoopCnt
'                        .RowHidden = False
'                    End If
'                Case intConSprChg_3  '�s�ǎw��
'                    '�G�s��s�]���ǉ��Ή�
'                    .GetText 38, intLoopCnt, vSprSta
'
'                    If vSprSta <> intConSprChg_3 Then  '�s�ǈȊO��������A��\��
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
'*    �֐���        : sub_F_InsertRow
'*
'*    �����T�v      : 1.�P�s�}�����s��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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

    '�ő�s���ǉ�
    sprExamine.MaxRows = sprExamine.MaxRows + 2

    intResult = sprExamine.ActiveRow Mod 2
    If intResult = 0 Then     '�����s�Ȃ��2�s�}��
        lngNewRow = sprExamine.ActiveRow
        lngNewRow2 = sprExamine.ActiveRow + 1
    Else                      '��Ȃ牺�Q�s�}��
        lngNewRow = sprExamine.ActiveRow + 1
        lngNewRow2 = sprExamine.ActiveRow + 2
    End If

    With sprExamine
        '�G�s��s�]���ǉ��Ή�
        .GetText 39, .ActiveRow, vActiveBLOCK   '�u���b�NID��col=30���擾
        .GetText 40, .ActiveRow, vOldHinban

        .row = lngNewRow
        .row2 = lngNewRow2
        .col = (-1)
        .BlockMode = True
        .Action = ActionInsertRow
        .Protect = True

        '�F�A�r���A���b�N�ݒ�
        .backColor = vbWhite

        '�G�s��s�]���ǉ��Ή�
        .SetText 40, lngNewRow, vOldHinban
        .SetText 40, lngNewRow2, vOldHinban
        .SetText 39, lngNewRow, vActiveBLOCK
        .SetText 39, lngNewRow2, vActiveBLOCK

        .col = 1                '��ԍ��̗�
        .col2 = 1
        .CellBorderType = 2
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16
        .backColor = &H80FF80

        .col = 4                '������3�Ԗڈȍ~
        .col2 = 10
        .backColor = &H80FF80
        .CellBorderType = 15
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 7                '������6�Ԗ�
        .col2 = 7
        .CellBorderType = 4 Or 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 4                '������R�Ԗڂ���̌r��
        .col2 = 10
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 4                '������3�Ԗڏ�u���b�NP��ҏW�\��
        .col2 = 4
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .backColor = vbWhite

        .col = 2                '������2�ԖڐV�i�Ԃ�ҏW�\��(��ROW�̂�)
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 3                '������3�Ԗ�
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = True
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = vbBlack
        .Action = 16

        .col = 7                'WF����
        .col2 = 7
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 7                'WF����
        .col2 = 7
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16
        vFirstFlg = 0

        '�G�s��s�]���ǉ��Ή�
        .SetText 37, lngNewRow, vFirstFlg
        .SetText 37, lngNewRow2, vFirstFlg
        .col = 9                '�s�ǋ敪
        .col2 = 9
        .row = lngNewRow
        .row2 = lngNewRow
        .Lock = False
        .CellBorderType = 4
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 9                '�s�ǋ敪
        .col2 = 9
        .row = lngNewRow2
        .row2 = lngNewRow2
        .Lock = False
        .CellBorderType = 8
        .CellBorderStyle = 1
        .CellBorderColor = &H8000000F
        .Action = 16

        .col = 4                '������3�Ԗڃu���b�NP��ҏW�s��(��ROW�̂�)
        .row = lngNewRow2
        .Lock = True

        .col = 11                '�������ڂ̓��͕s��
        .row = lngNewRow

        '�G�s��s�]���ǉ��Ή�
        .col2 = 35
        .row2 = lngNewRow2
        .Lock = True

        .col = 9                '�ǉ��s�ɋ敪�����i�敪�����鏉���\���s����̃R�s�[�j
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

        .col = 9                '������W�Ԗځi��s�j���e�L�X�g��
        .col2 = 9
        .row = lngNewRow
        .row2 = lngNewRow
        .CellType = CellTypeEdit

        .BlockMode = False
    End With

    If intResult = 0 Then     '�����s�Ȃ��2�s�}��
    Else                      '��Ȃ牺�Q�s�}��
        sprExamine.GetText 2, sprExamine.ActiveRow, vGetHinban
        sprExamine.SetText 2, sprExamine.ActiveRow, vbNullString
        sprExamine.SetText 2, sprExamine.ActiveRow + 2, vGetHinban
        sprExamine.GetText 3, sprExamine.ActiveRow, vGetHinban      '�i��2�̏�����ǉ�
        sprExamine.SetText 3, sprExamine.ActiveRow, vbNullString
        sprExamine.SetText 3, sprExamine.ActiveRow + 2, vGetHinban
    End If
End Sub

'*******************************************************************************************
'*    �֐���        : sub_SelWFmap
'*
'*    �����T�v      : 1.WFϯ�ߊǗ�ð��فiTBCMY011�j�����ް����擾
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      records()   ,O  ,typ_TBCME037 ,���o���R�[�h
'*�@�@      �@�@      sqlWhere    ,I  ,String       ,���o����(SQL��Where��:�ȗ��\)
'*�@�@      �@�@      sqlOrder    ,I  ,String       ,���o����(SQL��Order by��:�ȗ��\)
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*******************************************************************************************
Public Function sub_SelWFmap(ByVal sBlkId As String, ByVal sSXLID As String, ByRef sErrMsg As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim bBlkOn      As Boolean '��ۯ������݂��Ă�����ASXLID�̏�����AND������
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
    Else    '���s�{�^��������͕����V���O����\������i�����擾�͈͎̔w���SELECT�j
        sSXLID9 = Mid(sSXLID, 1, 9)
        sSql = sSql & " SUBSTR(MSXLID,1,9) = '" & sSXLID9 & "'"
        sSql = sSql & " AND RITOP_POS > " & SIngotP
        sSql = sSql & " AND RITOP_POS <= " & EIngotP
    End If
    sSql = sSql & " ORDER BY RITOP_POS"

    ''�f�[�^�𒊏o����
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
    '' �G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    sub_SelWFmap = FUNCTION_RETURN_FAILURE
    sErrMsg = GetMsgStr("SET47")
    rs.Close
End Function

'Del Start 2011/03/11 SMPK Miyata
''*******************************************************************************************
''*    �֐���        : fnc_SetWFmapData
''*
''*    �����T�v      : 1.�f�[�^�\������
''*
''*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
''*�@�@      �@�@      �Ȃ�
''*
''*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
''*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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
'            'WF��Ԕ���
'            Select Case gtWFmap(intLoopCnt).WFSTA
'                Case gsWF_STA_0   '�ʏ�
'                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
'                        Case gsWF_SMPL_1
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case gsWF_SMPL_2
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_OK & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case gsWF_SMPL_3
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_NG & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case Else
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_1
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = &H80FF80
'                            .BlockMode = False
'                    End Select
'                Case gsWF_STA_1   '���L
'                    .SetText 1, intLoopCnt + 1, gsWF_STA_NORMAL & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                    .SetText 30, intLoopCnt + 1, intConSprChg_1
'                    .col = 1
'                    .col2 = 32          'Warp����Ή�
'                    .row = intLoopCnt + 1
'                    .row2 = intLoopCnt + 1
'                    .BlockMode = True
'                    .backColor = &H80FF80
'                    .BlockMode = False
'                Case gsWF_STA_4   '����
'                    '�T���v���t���O����
'                    Select Case gtWFmap(intLoopCnt).SMPLEFLG
'                        Case gsWF_SMPL_4    'upd 2003/05/19 hitec)matsumoto �T���v���̌��ʈȊO�͂��ׂČ����Ɣ��f����
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_SIJI_KEKKA & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_2
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
'                            .row = intLoopCnt + 1
'                            .row2 = intLoopCnt + 1
'                            .BlockMode = True
'                            .backColor = vbYellow
'                            .BlockMode = False
'                        Case Else
'                            .SetText 1, intLoopCnt + 1, gsWF_STA_STA_K & "(" & gtWFmap(intLoopCnt).WFSTA & ")"   '2003/06/05 hitec)matsumoto  WF��ԃt���O�̕\����ǉ�
'                            .SetText 30, intLoopCnt + 1, intConSprChg_3
'                            .col = 1
'                            .col2 = 32          'Warp����Ή�
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
'            'Warp���\���ǉ�
'            If UBound(tWarpMeasG) >= intWarpPoint Then
'                '��ۯ�ID����ۯ����A�ԂŕR�t��
'                If tWarpMeasG(intWarpPoint).BLOCKID = gtWFmap(intLoopCnt).LOTID And _
'                   tWarpMeasG(intWarpPoint).WAFID = gtWFmap(intLoopCnt).BLOCKSEQ Then
'                    '���ް��������ꍇ�͕\�����Ȃ�
'                    If tWarpMeasG(intWarpPoint).EXISTFLG > 0 Then
'                        'Warp�l
'                        .SetText 31, intLoopCnt + 1, CStr(DBData2DispData_nl(tWarpMeasG(intWarpPoint).MEASDATA))
'                        '����
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
''*    �֐���        : cmbSprChg_Click
''*
''*    �����T�v      : 1.���o�����ɂ��AWFϯ�߈ꗗ����ʖ��̕\���̐؂�ւ����s��
''*
''*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
''*�@�@      �@�@      �Ȃ�
''*
''*    �߂�l        : �Ȃ�
''*
''*******************************************************************************************
'Private Sub cmbSprChg_Click()
'    Dim intLoopCnt  As Integer
'    Dim intSprSta   As Integer
'    Dim sSprSta     As String
'    Dim vSprSta     As Variant
'    Dim intRowNo    As Integer
'
'    intRowNo = 0  '�\������Ă���s�����ԍ����ӂ�Ȃ���
'     With sprWfmapView
'        .ReDraw = False
'        For intLoopCnt = 1 To .MaxRows
'            Select Case cmbSprChg.ListIndex
'                Case intConSprChg_0  '�S���w��
'                    .row = intLoopCnt
'                    .RowHidden = False
'                    intRowNo = intRowNo + 1
'                    .row = intLoopCnt
'                    .RowHidden = False
'                    .col = 0
'                    .row = intLoopCnt
'                    .text = intRowNo
'                Case intConSprChg_1  '�Ǖi�w��
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_1 Then  '�Ǖi�ȊO��������A��\��
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
'                Case intConSprChg_2  '�T���v���w��
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_2 Then  '�T���v���ȊO��������A��\��
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
'                Case intConSprChg_3  '�s�ǎw��
'                    .GetText 30, intLoopCnt, vSprSta
'                    If vSprSta <> intConSprChg_3 Then  '�s�ǈȊO��������A��\��
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
'*    �֐���        : fnc_CheckDataWfmap
'*
'*    �����T�v      : 1.�����w���ꗗ�̓��͓��e���`�F�b�N����
'*                      (�����w���ꗗ�̃f�[�^�`�F�b�N(WFϯ�ߑΉ��j)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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

    '�i�Ԃ̖����̓G���[
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

    '�u���b�N�̓��̓`�F�b�N
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
                '�G�s��s�]���ǉ��Ή�
                .GetText 37, intLoopCnt, vFlg
                If (vFlg = "1") Or (vFlg = "3") Then    '�����\���̔����s

                ElseIf intLoopCnt Mod 2 = 0 Then        '�}�����ꂽ�s�ŁA�����s�����͂̍s�ɂȂ�
                    .GetText 4, intLoopCnt, vBlockP

                    '����u���b�N�ł̐��l�͈̓`�F�b�N
                    '���݃u���b�NID
                    '�G�s��s�]���ǉ��Ή�
                    .GetText 39, intLoopCnt, vBlckID
                    sBlckID = CStr(vBlckID)

                    '��̒l���傫������
                    If intLoopCnt > 2 Then
                        '�G�s��s�]���ǉ��Ή�
                        .GetText 37, intLoopCnt - 1, vUpPos
                        If (vUpPos = "1") Or (vUpPos = "3") Or (vUpPos = "9") Then
                            '�G�s��s�]���ǉ��Ή�
                            .GetText 39, intLoopCnt - 1, vBlckID
                        Else
                            '�G�s��s�]���ǉ��Ή�
                            .GetText 39, intLoopCnt - 2, vBlckID
                        End If
                        If sBlckID = CStr(vBlckID) Then
                            '�G�s��s�]���ǉ��Ή�
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

                    '���̒l��������������
                    If intLoopCnt < .MaxRows Then
                        '�G�s��s�]���ǉ��Ή�
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

                    '�_�~�[�̃u���b�NID�E�i�ԂɊY���ް������Ă���
                    '���݈ʒu�̃u���b�NID��NULL���������̍s�ɒT���ɂ���
                    .GetText 1, intLoopCnt, vDummyBlk
                    If vDummyBlk = vbNullString Then
                        For intDummyBlkSet = intLoopCnt - 1 To 1 Step -1
                            .GetText 1, intDummyBlkSet, vDummyBlk
                            If vDummyBlk <> vbNullString Then
                                Exit For
                            End If
                        Next
                    End If

                    '�G�s��s�]���ǉ��Ή�
                    .SetText 39, intLoopCnt, vDummyBlk
                    '�i��
                    .GetText 2, intLoopCnt + 1, vDummyHin

                    '�G�s��s�]���ǉ��Ή�
                    .SetText 40, intLoopCnt, vDummyHin

                    '�G�s��s�]���ǉ��Ή�
                    .SetText 37, intLoopCnt, "2"
                    intLoopCnt = intLoopCnt + 1  '2�s���̍s�����ɍs�����߂Ɂ{�P���Ă���
                End If
            End If
        Next

        '�T���v�����㉺���L�̏ꍇ�̃G���[���b�Z�[�W---------------
        For intLoopCnt = 2 To .MaxRows Step 2
            .GetText 10, intLoopCnt, vSmpId1
            .GetText 10, intLoopCnt + 1, vSmpId2
            If (vSmpId1 = vSmpId2) And (vSmpId1 <> vbNullString) And (vSmpId2 <> vbNullString) Then    '�T���v��ID���A�㉺���L
                fnc_CheckDataWfmap = FUNCTION_RETURN_FAILURE
                lblMsg.Caption = GetMsgStr("ENSP3")

                '�G�s��s�]���ǉ��Ή�
                 .GetText 39, intLoopCnt, vBlckID
                 .GetText 38, intLoopCnt, vStrSam
                 vStrB = Right(vBlckID, 3) & "-" & Right(vStrSam, 4)
                 .SetText 10, intLoopCnt, vStrB ''�T���v��ID�\��
                 .SetText 10, intLoopCnt + 1, "���L"
                Exit Function
            End If
        Next
    End With
    fnc_CheckDataWfmap = FUNCTION_RETURN_SUCCESS
End Function

'*******************************************************************************************
'*    �֐���        : fnc_DispChangeDataWfmap
'*
'*    �����T�v      : 1.WF�}�b�v�̕\���f�[�^�ɉ�ʂł̕ύX�𔽉f����iDB�o�^�O�ł��j
'*                      (WF�}�b�v�\���f�[�^�ύX(WFϯ�ߑΉ��j)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
'*
'*
'*  Chg 2011/03/11 SMPK Miyata  ���̊֐�����WFϯ�߂�\�����ύX����B
'*                              f_cmbc039_4.sprWfmapView��f_cmbc039_4.sprWfmapView�ɕύX
'*                              �v���O�������ώG�ɂȂ�̂ŏC�������͎c���Ă��܂���B
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

    '�T���v����ʂ̕ύX�_���}�b�v��ʂ�
    With f_cmbc039_3.sprExamine

        For lngRoopCnt = 1 To lngSmplMaxRows Step 2
            .GetText 1, lngRoopCnt, vBlockId
            If (vBlockId <> "") Then
                '�T���v��_�u���b�NID�ۑ�
                sBlockId = CStr(vBlockId)
            End If
            '�T���v����ʃf�[�^�擾
            .GetText 6, lngRoopCnt, vFromMap        '�T���v��_�}�b�v�ʒu�i�J�n�j
            .GetText 6, lngRoopCnt + 1, vToMap      '�T���v��_�}�b�v�ʒu�i�I���j
            .GetText 2, lngRoopCnt, vHinban         '�T���v��_�i��

            '�}�b�v��ʌ������ύX
            For lngMapRoopCnt = 1 To lngMapMaxRows
                '�}�b�v_�u���b�NID�擾
                f_cmbc039_4.sprWfmapView.GetText 3, lngMapRoopCnt, vMapBlockID
                If sBlockId = Mid(CStr(vMapBlockID), 10, 3) Then
                    f_cmbc039_4.sprWfmapView.GetText 5, lngMapRoopCnt, vNowMap
                    '�����s
                    If CInt(vNowMap) = CInt(vFromMap) Or CInt(vNowMap) = CInt(vToMap) Then
                        If vHinban = "Z" Then
                            vStatusWfmap = "����" & "(4)"
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
                                .GetText 8, lngRoopCnt, vWFStatus1         '�T���v��_WF���
                                sWFStatusWfmap = CStr(vWFStatus1)
                                .GetText 10, lngRoopCnt, vGetSampId
                            Else
                                .GetText 8, lngRoopCnt + 1, vWFStatus1       '�T���v��_WF���
                                sWFStatusWfmap = CStr(vWFStatus1)
                                .GetText 10, lngRoopCnt + 1, vGetSampId
                            End If
                            '�ύX�͈�(�T���v������j
                            If sWFStatusWfmap = "���L" Or sWFStatusWfmap = "�ʏ�" Then
                                If vGetSampId = "���L" Then
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
                            Else            '�T���v������Œʏ�Łi�����w��WF�j
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
                    '�����s�ȊO
                    ElseIf CInt(vNowMap) > CInt(vFromMap) And CInt(vNowMap) < CInt(vToMap) Then
                        If vHinban = "Z" Then
                            vStatusWfmap = "����" & "(4)"
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
'                            '�ύX�͈�(�T���v���Ȃ��j
'                            vStatusWfmap = "�ʏ�" & "(0)"
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
'*    �֐���        : sub_LOADDISP_ADD_CHECKBOX
'*
'*    �����T�v      : 1.����SXL����i�Ԃ̃u���b�N���E�ɔ����L���̃`�F�b�N�{�b�N�X�\��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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
'*    �֐���        : fnc_ErrDispCheck
'*
'*    �����T�v      : 1.�����X�v���b�h�ɕ\�����ꂽ�f�[�^���A����ł��邩�`�F�b�N
'*                      �@���������s�ɃT���v��ID�������Ă��邩
'*                      �A������TOP�ʒu��BOTTOM�ʒu���ASXL�ʒu�ƈ�v���Ă��邩
'*
'*    �p�����[�^    : �ϐ���    ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      sMsg      ,O  ,String       ,�\�����b�Z�[�W������
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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
        '�T���v��ID���̓`�F�b�N
        For intLoopCnt = 1 To .MaxRows

            '�G�s��s�]���ǉ�
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

        '������TOP�ʒu��BOTTOM�ʒu�̃`�F�b�N
        .GetText 5, 1, vDispSIngotP
        .GetText 5, .MaxRows, vDispEIngotP

        'tuku +-1�̃Y����OK�Ƃ���B��WF���Ԃɓ��͉\�l���Q����ꍇ�̑Ή�
        fnc_ErrDispCheck = FUNCTION_RETURN_SUCCESS
    End With
End Function

'*******************************************************************************************
'*    �֐���        : sub_FurikaeMotoDataSet
'*
'*    �����T�v      : 1.�U�֌��i�Ԃ��Z�[�u���A�U�֓��e�ݒ���N���A����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
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
'*    �֐���        : sub_FurikaeKouho
'*
'*    �����T�v      : 1.�U�։\���i�ԉ�ʂ��U�֐�i�Ԃ��擾����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_FurikaeKouho()
    With sprExamine
        .row = .ActiveRow
        .col = .ActiveCol
        If .text <> "" Then
            .col = 5
            If .text = "" Then
                lblMsg.Caption = "�����ʒu������܂���"
                Exit Sub
            End If
            If sub_FurikaeKouho_Check() = False Then
                Exit Sub
            End If
        Else
            lblMsg.Caption = "�U�֌��̕i�Ԃ��w�肳��Ă��܂���"
            Exit Sub
        End If
    End With

    '' �U�։\���i�ԉ�ʌĂяo��
    f_cmzcFKKH.Show 1

    '' �U�֐�i�Ԃ�\������
    If FKKH_SakiHinban <> "" Then
        With sprExamine
            .row = .ActiveRow
            .col = 2
            .text = left$(FKKH_SakiHinban, 8)
            .col = 3
            .text = Right$(FKKH_SakiHinban, 4)

            '' �������ύX���ꂽ��
            Call sprExamine_Change(2, .row)
        End With
    End If
End Sub

'*******************************************************************************************
'*    �֐���        : sub_FurikaeKouho_Check
'*
'*    �����T�v      : 1.�i�Ԗ��I���̃`�F�b�N���s��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : Boolean�i�`�F�b�NOK:True�A�`�F�b�NNG:False�j
'*
'*******************************************************************************************
Public Function sub_FurikaeKouho_Check() As Boolean
    Dim intIchi         As Integer
    Dim lngCnt          As Long
    Dim sCkHinban       As String
    Dim sCkHinbanRev    As String

    sub_FurikaeKouho_Check = False

    With sprExamine
        '' �I��i�Ԃ̐ݒ�
        .row = .ActiveRow
        .col = 5            ' �ʒu
        If .text = "" Then
            .row = .ActiveRow - 1
        End If
        intIchi = .text

        '' �U�֌��i�ԃ`�F�b�N
        For lngCnt = 1 To UBound(MotoHinban)
            If MotoHinban(lngCnt).MOTOICHIS <= intIchi And intIchi < MotoHinban(lngCnt).MOTOICHIE Then
                sCkHinban = MotoHinban(lngCnt).MOTOHIN.hinban
                sCkHinbanRev = Format$(MotoHinban(lngCnt).MOTOHIN.mnorevno, "00") & MotoHinban(lngCnt).MOTOHIN.factory & MotoHinban(lngCnt).MOTOHIN.opecond
                Exit For
            End If
        Next

        If (Trim$(sCkHinban) = "G" Or Trim$(sCkHinban) = "Z" Or _
            Trim$(sCkHinban) = "") Then
            .col = 2           ' �i��
            .backColor = COLOR_NG
            lblMsg.Caption = "�I���̕i�Ԃ͐U�։\�i�Ԃł͂���܂���"
            Exit Function
        End If

        '' �K�v�f�[�^�ݒ�
        FKKH_Proccd = "CW760"
        FKKH_MotoHinban = sCkHinban & sCkHinbanRev
        FKKH_Crynum = txtCryNum.text
    End With

    sub_FurikaeKouho_Check = True
End Function

'*******************************************************************************************
'*    �֐���        : fnc_Furikae_Check
'*
'*    �����T�v      : 1.�U�։ۃ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : Boolean�i�`�F�b�NOK:True�A�`�F�b�NNG:False�j
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
    Dim intIchiE        As Integer      '�i�ԏI���ʒu
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
    Dim blWKchkFlg      As Boolean      'Warp/�����p�x�������L��
    Dim udtCAhin()      As typ_XSDCA    '��ۯ����i�Ԏ擾�p
    Dim sSqlWhere       As String       'WHERE��
    Dim udtHin          As tFullHinban  '�i��
    Dim udtChkHin()     As tFullHinban  '�g���������p�i��
    Dim intHinRow()     As Integer      '�i�ԍs�ʒu
    Dim intHinCnt       As Integer      '�i�Ԑ�
    Dim intHinGrp       As Integer      '�i�Ը�ٰ��(��ۯ��P��)
    Dim sBlkChk         As String       '��ۯ�ID
    Dim intOiChk        As Integer      '�Ǎ�������
    Dim intNGrow        As Integer      '�g��������NG�i�ԍs�ʒu

    intHinCnt = 0
    intHinGrp = 0
    intOiChk = 0
    ReDim udtChkHin(0)
    ReDim intHinRow(0)

    Dim nInpos      As Integer          ' �������ʒu�i��\�T���v���h�c�擾�ׁ̈j

    fnc_Furikae_Check = False

    With sprExamine
        ReDim FurikaeNaiyouWK(.MaxRows)
        m = 0       '06/01/12 ooba
        '' �U�փ`�F�b�N
        For intLp = 1 To .MaxRows - 1
            blWKchkFlg = False

            .row = intLp
            .col = 5                    ' �ʒu
            If .text = "" Then          '���ɍs��}�������ꍇ�ɂ͏����\�����Ɉʒu���ݒ肳��Ă��Ȃ��̂ŏ�̈ʒu������
                .row = intLp - 1
            End If
            intIchi = .text

            '�i�ԏI���ʒu�擾
            .row = intLp + 1
            intIchiE = .text
            .row = intLp
            intOiChk = 0      '�Ǎ��������׸ޏ�����

            .col = 2           ' �i��
            udtSakiHin.hinban = .text

            Call GetLastHinban(udtSakiHin.hinban, udtSakiHin) '�t�����i�Ԃ��擾

            .col = 3        ' �i�ԃ��r�W����
            .text = Format$(udtSakiHin.mnorevno, "00") & udtSakiHin.factory & udtSakiHin.opecond

            If Trim(udtSakiHin.hinban) <> "" And _
               Trim(udtSakiHin.hinban) <> "G" And _
               Trim(udtSakiHin.hinban) <> "Z" Then

                m = m + 1
                tMapHinG = tMapHin(m)
                blWKchkFlg = True
            End If

            ''�i�ԑg���������Ή�
            '��Ǎ�������
            If intIchi <> tExamine(0).RITOP_POS Then intOiChk = intOiChk + 1

            '���Ǎ�������
            If intIchiE <> tExamine(UBound(tExamine)).RITOP_POS Then intOiChk = intOiChk + 1

            'G/Z�i�ȊO
            If Trim(udtSakiHin.hinban) <> "" And Trim(udtSakiHin.hinban) <> "G" And _
               Trim(udtSakiHin.hinban) <> "Z" Then
                intHinCnt = intHinCnt + 1
                ReDim Preserve udtChkHin(intHinCnt)
                ReDim Preserve intHinRow(intHinCnt)

                '�i�Ծ��
                udtChkHin(intHinCnt) = udtSakiHin
                intHinRow(intHinCnt) = intLp

                '��ۯ�ID�\���L�Ȃ����ۯ�
                    .col = 1:   sBlkChk = .text
                    If Trim(sBlkChk) <> "" Then intHinGrp = intHinGrp + 1

                '��ۯ��P��No���
                If intHinGrp < 10 Then
                    udtChkHin(intHinCnt).Hinkubun = Chr$(intHinGrp + vbKey0)
                Else
                    udtChkHin(intHinCnt).Hinkubun = Chr$(intHinGrp - 10 + vbKeyA)
                End If
            End If

            '�G�s��s�]���ǉ��Ή�
            .col = 39
            .row = intLp
            sVgetlotid = .text

            ' �T���v��ID�̎擾
            .col = 5
            nInpos = CInt(.text)
            Call fnc_Furikae_Set_RepID(sRepidT, sRepidB, nInpos)

            '' n�u���b�N1SXL�p�ɃN���A����
            .row = intLp
            .col = 10
            If Trim(.text) = "" Then sRepidT = ""
            .row = intLp + 1
            .col = 10
            If Trim(.text) = "" Then sRepidB = ""

            .row = intLp


            '' �U�֗L���`�F�b�N
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

            '' �U�։ۃ`�F�b�N
            If intUmuFlg = 1 Then
                '' �U�֓��e�ݒ�f�[�^
                FurikaeNaiyouWK(intLp).FURIUMU = intUmuFlg
                FurikaeNaiyouWK(intLp).ICHI = tblNukishi(intLp).INPOSCW         '2003/10/27 SystemBrain
                FurikaeNaiyouWK(intLp).MOTOHIN = MotoHinban(lngCnt).MOTOHIN
                FurikaeNaiyouWK(intLp).SAKIHIN = udtSakiHin
                FurikaeNaiyouWK(intLp).TREPID = sRepidT
                FurikaeNaiyouWK(intLp).BREPID = sRepidB

                '' �ʒu�̍Đݒ�
                If tblNukishi(intLp).INPOSCW = -1 Then
                    FurikaeNaiyouWK(intLp).ICHI = intIchi
                End If

                '' �����̓��̔z��̈ʒu�ƻ����ID��������
                For lngCnt3 = 1 To TokuCntWK - 1
                    For lngCnt4 = 1 To UBound(FurikaeNaiyouWK)
                        If TokusaiBangou(lngCnt3).SAKIHIN.hinban = FurikaeNaiyouWK(lngCnt4).SAKIHIN.hinban And _
                           TokusaiBangou(lngCnt3).SAKIHIN.mnorevno = FurikaeNaiyouWK(lngCnt4).SAKIHIN.mnorevno And _
                           TokusaiBangou(lngCnt3).SAKIHIN.factory = FurikaeNaiyouWK(lngCnt4).SAKIHIN.factory And _
                           TokusaiBangou(lngCnt3).SAKIHIN.opecond = FurikaeNaiyouWK(lngCnt4).SAKIHIN.opecond Then
                            TokusaiBangou(lngCnt3).ICHI = FurikaeNaiyouWK(lngCnt4).ICHI
                            TokusaiBangou(lngCnt3).TREPID = FurikaeNaiyouWK(lngCnt4).TREPID   '��\�T���v��ID
                            TokusaiBangou(lngCnt3).BREPID = FurikaeNaiyouWK(lngCnt4).BREPID   '��\�T���v��ID
                        End If
                    Next lngCnt4
                Next lngCnt3

                '' ���̔ԍ����̓`�F�b�N
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
                        ' ���̔ԍ��f�[�^����œ��̔ԍ��Ȃ��̏ꍇ
                        If TokusaiBangou(lngCnt2).BANGOU = "" Then
                            intTokuFlg = 2
                        End If
                        Exit For
                    End If
                Next

                '' ���̔ԍ����͂̏ꍇ�̓`�F�b�N�Ȃ�
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
                            '' ��̓��̔ԍ��f�[�^���ݒ�
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
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
'                            intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
'                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2)
                            intRet = funChkFurikaeShiyou("CW761", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                        Else
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
'                            intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
'                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2)
                            intRet = funChkFurikaeShiyou("CW762", tblSXL.SXLID, MotoHinban(lngCnt).MOTOHIN, _
                                udtSakiHin, intErrCode, intErrMsg, typ_b, typ_CType, intSmpGetFlg, sSmpllID1, "", 0, 0, 2, , , , intIchi, intIchiE)
'2012/01/12 Update E_N_D DCS)Shoryuji
                        End If

                        If Trim(udtSakiHin.hinban) <> "" And _
                           Trim(udtSakiHin.hinban) <> "G" And _
                           Trim(udtSakiHin.hinban) <> "Z" Then

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                        End If

                        ''�����i�ԑΉ�
                        If intRet = 0 Then
                            sSqlWhere = "where crynumca = '" & Mid(tblSXL.CRYNUM, 1, 9) & Trim(sVgetlotid) & "' "
                            sSqlWhere = sSqlWhere & "and livkca = '0' "

                            '��ۯ����i�Ԏ擾
                            If DBDRV_GetXSDCA(udtCAhin(), sSqlWhere) = FUNCTION_RETURN_FAILURE Then
                                lblMsg.Caption = GetMsgStr("EGET2", "XSDCA")
                                Exit Function
                            End If

                            '' 1-4(Cs)�`�F�b�N
                            For lngCnt4 = 1 To UBound(udtCAhin)
                                udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                udtHin.opecond = udtCAhin(lngCnt4).OPECA
                                'Cs�̎d�l����(1-4)�ɂ��Ă���ۯ����S�i�Ԃōs��
                                If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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

                            '' 1-4(EPD)�`�F�b�N
                            If intRet = 0 Then
                                For lngCnt4 = 1 To UBound(udtCAhin)
                                    udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                    udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                    udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                    udtHin.opecond = udtCAhin(lngCnt4).OPECA
                                    'EPD�̎d�l����(1-4)�ɂ��Ă���ۯ����S�i�Ԃōs��
                                    If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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

                            '' 1-4(LT)�`�F�b�N
                            If intRet = 0 Then
                                For lngCnt4 = 1 To UBound(udtCAhin)
                                    udtHin.hinban = udtCAhin(lngCnt4).HINBCA
                                    udtHin.mnorevno = udtCAhin(lngCnt4).REVNUMCA
                                    udtHin.factory = udtCAhin(lngCnt4).FACTORYCA
                                    udtHin.opecond = udtCAhin(lngCnt4).OPECA

                                    'LT�̎d�l����(1-4)�ɂ��Ă���ۯ����S�i�Ԃōs��
                                    If intOiChk <> 2 Then
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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
'2012/01/12 Update START DCS)Shoryuji �����{�g�������o�K���ǉ��Ή�
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

                        If intRet = 1 Then  '�U��NG�ƐU��OK
                            '' ���̔ԍ��f�[�^
                            TokuCntWK = TokuCntWK + 1
                            ReDim Preserve TokusaiBangou(TokuCntWK)
                            TokusaiBangou(TokuCntWK).ICHI = FurikaeNaiyouWK(intLp).ICHI
                            TokusaiBangou(TokuCntWK).MOTOHIN = FurikaeNaiyouWK(intLp).MOTOHIN
                            TokusaiBangou(TokuCntWK).SAKIHIN = FurikaeNaiyouWK(intLp).SAKIHIN
                            TokusaiBangou(TokuCntWK).BANGOU = ""
                            TokusaiBangou(TokuCntWK).RIYUU = ""
                            TokusaiBangou(TokuCntWK).ERRRIYUU = left$(intErrMsg, 25)
                            TokusaiBangou(TokuCntWK).TREPID = FurikaeNaiyouWK(intLp).TREPID   '��\�T���v��ID
                            TokusaiBangou(TokuCntWK).BREPID = FurikaeNaiyouWK(intLp).BREPID   '��\�T���v��ID

                            .col = 2           ' �i��
                            .backColor = COLOR_NG
                            lblMsg.Caption = intErrMsg
                            If bTokuKengenFlag = True Then      ' ���̌�������̏ꍇ
                                cmdF(5).Enabled = True
                                cmdF(5).backColor = vbRed
                            End If
                            Exit Function
                        ElseIf intRet < 0 Then '�U�փG���[
                            lblMsg.Caption = intErrMsg
                            Exit Function
                        ElseIf intRet = 0 Then
                            .col = 2
                            .backColor = vbWhite
                        End If
                    Else
                        '�U�������ΏۊO��Warp/�����p�x����
                        For intLp2 = 1 To 2
                            If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                                intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                          tMapHinG.HIN, tMapHinG.HIN, _
                                                          intErrCode, intErrMsg, _
                                                          typ_b, typ_CType, 0)

                                tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                                tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                            End If
                        Next intLp2
                    End If

                '' ���̔ԍ����͂̏ꍇ�͐ݒ�
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
                            '' ��̓��̔ԍ��f�[�^����œ��̔ԍ�����̏ꍇ
                            If TokusaiBangou(lngCnt3).BANGOU <> "" Then
                                TokusaiBangou(lngCnt2).BANGOU = TokusaiBangou(lngCnt3).BANGOU
                                intTokuFlg = 4
                                Exit For
                            End If
                        End If
                    Next

                    '' ���̔ԍ��f�[�^���쐬�����ɃG���[
                    If intTokuFlg = 2 Then
                        .col = 2           ' �i��
                        .backColor = COLOR_NG
                        lblMsg.Caption = "�U�֔ԍ��̓��͂�����܂���"
                        If bTokuKengenFlag = True Then      ' ���̌�������̏ꍇ
                            cmdF(5).Enabled = True
                            cmdF(5).backColor = vbRed
                        End If
                        Exit Function
                    End If

                    '�U�������ΏۊO��Warp/�����p�x����
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                        End If
                    Next intLp2
                Else
                    '�U�������ΏۊO��Warp/�����p�x����
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                        End If
                    Next intLp2
                End If
            Else
                '�U�������ΏۊO��Warp/�����p�x����
                If blWKchkFlg Then
                    For intLp2 = 1 To 2
                        If Not (tMapHinG.WARPFLG And tMapHinG.KAKUFLG) Then
                            intRet = funChkFurikaeShiyou("CW763", tblSXL.SXLID, _
                                                      tMapHinG.HIN, tMapHinG.HIN, _
                                                      intErrCode, intErrMsg, _
                                                      typ_b, typ_CType, 0)

                            tMapHin(m).WARPFLG = tMapHinG.WARPFLG   'Warp�U�������׸޾��
                            tMapHin(m).KAKUFLG = tMapHinG.KAKUFLG   '�����p�x�U�������׸޾��
                            '����NG
                            If intRet = 1 Then
                                If Not tMapHinG.KAKUFLG Then
                                    lblMsg.Caption = "Warp����G���[�@�i�ԐU�ւ��s���Ă��������B"
                                Else
                                    lblMsg.Caption = "�����p�x����G���[�@�i�ԐU�ւ��s���Ă��������B"
                                End If
                                Exit Function
                            '�U�������װ
                            ElseIf intRet < 0 Then
                                lblMsg.Caption = intErrMsg
                                Exit Function
                            End If
                        End If
                    Next intLp2
                End If
            End If
            intLp = intLp + 1 '1�s���Ƃ̍s
        Next intLp

        '' Z�i�Ԃ݂̂̏ꍇ�͕i�ԑg�����`�F�b�N�͕s�v
        If UBound(udtChkHin) > 0 Then
            ''�i�ԑg���������Ή�
            '�Ǎ���
            If intOiChk = 2 Then
                intRet = funChkKumiHinban("CW762", tblSXL.CRYNUM, _
                                        udtChkHin(), intHinRow(), intNGrow, intErrCode, intErrMsg)
            '�U��
            Else
                intRet = funChkKumiHinban("CW761", tblSXL.CRYNUM, _
                                        udtChkHin(), intHinRow(), intNGrow, intErrCode, intErrMsg)
            End If
            If intRet = 1 Then
                .row = intNGrow
                .col = 2
                .backColor = COLOR_NG
                '�i�Ԃ̑g�������s���ł�(%s)
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
'*    �֐���        : fnc_Furikae_Set_RepID
'*
'*    �����T�v      : 1.�U�։ۃ`�F�b�N
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      sRepidT   ,O  ,String       ,��\�T���v��ID(Top)
'*�@�@      �@�@      sRepidB   ,O  ,String       ,��\�T���v��ID(Bot)
'*�@�@      �@�@      intInpos      ,O  ,Integer      ,�������ʒu�i��\�T���v���h�c�擾�ׁ̈j
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub fnc_Furikae_Set_RepID(sRepidT As String, sRepidB As String, intInpos As Integer)
    Dim i As Integer

    ' ���������Ă���
    sRepidT = ""
    sRepidB = ""
    For i = 0 To UBound(tKensa()) - 1 Step 2
        ' �������ʒu���㉺�T���v���͈͓̔��ɂ��邩�H
        If tKensa(i).INPOSCW <= intInpos And tKensa(i + 1).INPOSCW >= intInpos Then
            sRepidT = tKensa(i).REPSMPLIDCW
            sRepidB = tKensa(i + 1).REPSMPLIDCW
            Exit For
        End If
    Next
End Sub

'*******************************************************************************************
'*    �֐���        : sub_TokusaiInput
'*
'*    �����T�v      : 1.���̔ԍ����͉�ʂɂē��̔ԍ�����͂���
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@�@ ,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_TokusaiInput()
    '' �U�֌��i�ԂƐU�֐�i�Ԃ̐ݒ�
    With TokusaiBangou(TokuCntWK)
        TBN_MotoHinban = .MOTOHIN.hinban & Format$(.MOTOHIN.mnorevno, "00") & .MOTOHIN.factory & .MOTOHIN.opecond
        TBN_SakiHinban = .SAKIHIN.hinban & Format$(.SAKIHIN.mnorevno, "00") & .SAKIHIN.factory & .SAKIHIN.opecond
        TBN_Bangou = .BANGOU
        TBN_Riyuu = .RIYUU
    End With

    '' ���̔ԍ����͉�ʌĂяo��
    f_cmzcTBN.Show 1

    '' ���̔ԍ���ݒ肷��
    If TBN_Bangou <> "" Then
        TokusaiBangou(TokuCntWK).BANGOU = TBN_Bangou
        TokusaiBangou(TokuCntWK).RIYUU = TBN_Riyuu
    End If
End Sub

'*******************************************************************************************
'*    �֐���        : fnc_RirekiKanriDB_Touroku
'*
'*    �����T�v      : 1.�����Ǘ��c�a�̓o�^���s��
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      p_ErrMsg    ,O  ,string    ,ERR���b�Z�[�W
'*
'*    �߂�l        : Boolean (True:OK False:NG)
'*
'*******************************************************************************************
Public Function fnc_RirekiKanriDB_Touroku(sP_ErrMsg As String) As Boolean
    Dim cnt             As Long
    Dim cnt2            As Long
    Dim intLp           As Integer
    Dim intKCNTC3       As Integer
    Dim udtDataXSDC3()  As typ_XSDC3    '�H�����єz��
    Dim sWhere          As String       'SQL������
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
        '' �U�֗����̓o�^�f�[�^�ݒ�
        If FurikaeNaiyou(intLp).FURIUMU = 1 Then
            cnt = cnt + 1
            ReDim Preserve FurikaeRireki(cnt)

            '�u���b�NID
            sprExamine.col = 39

            sprExamine.row = intLp
            sVgetlotid = sprExamine.text

            '' �H���A�Ԏ擾
            intKCNTC3 = GetKCNTC3(Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid), "CW760")
            If intKCNTC3 = 0 Then
                sP_ErrMsg = "�H���A�Ԃ��擾�ł��܂���ł���"
                Exit Function
            End If

            'WHERE����
            sWhere = "WHERE CRYNUMC3 = '" & Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid) & "' "
            sWhere = sWhere & "AND INPOSC3 = " & FurikaeNaiyou(intLp).ICHI & " "    '�������J�n�ʒu
            sWhere = sWhere & "AND KCNTC3  = " & intKCNTC3 & " "

            '' �H������(XSDC3)�f�[�^�擾
            sDBName = "XSDC3"
            If DBDRV_GetXSDC3(udtDataXSDC3, sWhere) = FUNCTION_RETURN_FAILURE Then
                sP_ErrMsg = GetMsgStr("EGET ", vbNullString, sDBName)
                Exit Function
            End If

            intINPOSC3 = -1
            If UBound(udtDataXSDC3) = 0 Then
                '' �������J�n�ʒu
                intINPOSC3 = GetINPOSC3(Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid), FurikaeNaiyou(intLp).ICHI, intKCNTC3)
                If intINPOSC3 = -1 Then
                    sP_ErrMsg = "�������J�n�ʒu���擾�ł��܂���ł���"
                    Exit Function
                End If

                'WHERE����
                sWhere = "WHERE CRYNUMC3 = '" & Mid(tblSXL.SXLID, 1, 9) & Trim(sVgetlotid) & "' "
                sWhere = sWhere & "AND INPOSC3 = " & intINPOSC3 & " "
                sWhere = sWhere & "AND KCNTC3  = " & intKCNTC3 & " "

                '' �H������(XSDC3)�f�[�^�擾
                sDBName = "XSDC3"
                If DBDRV_GetXSDC3(udtDataXSDC3, sWhere) = FUNCTION_RETURN_FAILURE Then
                    sP_ErrMsg = GetMsgStr("EGET ", vbNullString, sDBName)
                    Exit Function
                End If

                If UBound(udtDataXSDC3) = 0 Then
                    sP_ErrMsg = "�H������(XSDC3)���擾�ł��܂���ł���"
                    Exit Function
                End If
            End If

            '' ���̔ԍ��擾
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
                .CRYNUMCE = udtDataXSDC3(1).CRYNUMC3                    ' �u���b�NID�E�����ԍ���XSDC3���
                .INPOSCE = FurikaeNaiyou(intLp).ICHI                    ' �������J�n�ʒu
                If intINPOSC3 <> -1 Then
                    .INPOSCE = intINPOSC3                               ' �������J�n�ʒu
                End If
                .KCNTCE = udtDataXSDC3(1).KCNTC3                        ' �H���A��
                .HINBCE = FurikaeNaiyou(intLp).SAKIHIN.hinban           ' �U�֐�i��
                .REVNUMCE = FurikaeNaiyou(intLp).SAKIHIN.mnorevno       ' ���i�ԍ������ԍ�(�U�֐�)
                .FACTORYCE = FurikaeNaiyou(intLp).SAKIHIN.factory       ' �H��(�U�֐�)
                .OPECE = FurikaeNaiyou(intLp).SAKIHIN.opecond           ' ���Ə���(�U�֐�)
                .MOTHINCE = FurikaeNaiyou(intLp).MOTOHIN.hinban         ' �U�֌��i��
                .MREVNUMCE = FurikaeNaiyou(intLp).MOTOHIN.mnorevno      ' ���i�ԍ������ԍ�(�U�֌�)
                .MFACTORYCE = FurikaeNaiyou(intLp).MOTOHIN.factory      ' �H��(�U�֌�)
                .MOPECE = FurikaeNaiyou(intLp).MOTOHIN.opecond          ' ���Ə���(�U�֌�)
                .SXLIDCE = udtDataXSDC3(1).SXLIDC3                      ' SXLID
                .WKKTCE = udtDataXSDC3(1).WKKTC3                        ' �H��
                .KNKTCE = udtDataXSDC3(1).KNKTC3                        ' �Ǘ��H��
                .REPSMPLIDTCE = FurikaeNaiyou(intLp).TREPID             ' ��\�T���v��ID(TOP)
                .REPSMPLIDBCE = FurikaeNaiyou(intLp).BREPID             ' ��\�T���v��ID(BOT)
                If intTokuFlg = 0 Then
                    .TOKNUMCE = ""                                      ' ���̔ԍ�
                    .TOKCAUSECE = ""                                    ' ���̗��R
                    .ERRCAUSECE = ""                                    ' �G���[���R
                Else
                    .TOKNUMCE = TokusaiBangou(cnt2).BANGOU              ' ���̔ԍ�
                    .TOKCAUSECE = TokusaiBangou(cnt2).RIYUU             ' ���̗��R
                    .ERRCAUSECE = TokusaiBangou(cnt2).ERRRIYUU          ' �G���[���R
                End If
                .TOKCODECE = ""                                         ' ���̗��R�R�[�h
                .HULCE = udtDataXSDC3(1).TOLC3                          ' �U�֒���
                .HUWCE = udtDataXSDC3(1).TOWC3                          ' �U�֏d��
                .HUMCE = udtDataXSDC3(1).TOMC3                          ' �U�֖��� ��XSDC3���
                .TSTAFFCE = txtStaffID.text                             ' �o�^�Ј�ID
                .KSTAFFCE = txtStaffID.text                             ' �X�V�Ј�ID
                .SNDKCE = "0"                                           ' ���M�t���O (0:�����M)
                .SNDDAYCE = ""                                          ' ���M���t (�u�����N)
            End With
        End If
    Next

    ' n�u���b�N1SXL�̃f�[�^�̐U�֗����͂P��
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
        '' �U�֗����̓o�^
        If CreateXSDCE(FurikaeRireki(intLp), sErrMsg) = False Then
            sP_ErrMsg = sErrMsg
            Exit Function
        End If
        '���̓��͒ʒm���[�����M  2011/07/04 Kameda
        If intTokuFlg = 1 Then
            Call SendMailMain(FurikaeRireki(intLp))
        End If
    Next

    fnc_RirekiKanriDB_Touroku = True
End Function

'*******************************************************************************************
'*    �֐���        : fnc_FurikaeRireki_Join
'*
'*    �����T�v      : 1.���i�Ԃ����̗�������R�s�[
'*�@�@�@�@�@�@�@�@�@�@2.���̗����Ƃ̍��Z
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      intIdx�@�@     ,I  ,Integer   ,�U�֗����f�[�^�Y����
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Public Sub fnc_FurikaeRireki_Join(intIdx As Integer)
    ' ���i�Ԃ����̗�������R�s�[
    FurikaeRireki(intIdx).REPSMPLIDBCE = FurikaeRireki(intIdx + 1).REPSMPLIDBCE
    ' ���Z
    FurikaeRireki(intIdx).HULCE = CStr(CInt(FurikaeRireki(intIdx).HULCE) + CInt(FurikaeRireki(intIdx + 1).HULCE))    ' �U�֒���
    FurikaeRireki(intIdx).HUWCE = CStr(CLng(FurikaeRireki(intIdx).HUWCE) + CLng(FurikaeRireki(intIdx + 1).HUWCE))    ' �U�֏d��
    FurikaeRireki(intIdx).HUMCE = CStr(CInt(FurikaeRireki(intIdx).HUMCE) + CInt(FurikaeRireki(intIdx + 1).HUMCE))    ' �U�֖���
End Sub

'*******************************************************************************************
'*    �֐���        : fnc_FurikaeRireki_Move
'*
'*    �����T�v      : 1.�����f�[�^���P���l�߂�
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      intIdx�@�@     ,I  ,Integer   ,�U�֗����f�[�^�Y����
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Public Sub fnc_FurikaeRireki_Move(intIdx As Integer)
    ' �����f�[�^���P���l�߂�
    Dim intLp As Integer

    For intLp = intIdx To UBound(FurikaeRireki) - 1
        FurikaeRireki(intLp) = FurikaeRireki(intLp + 1)
    Next
End Sub

'*******************************************************************************************
'*    �֐���        : sub_ClearTokusai
'*
'*    �����T�v      : 1.���̏��N���A
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_ClearTokusai()

    '' ���̃o�b�t�@�N���A
    TokuCntWK = TokuCnt
    ReDim Preserve TokusaiBangou(TokuCntWK)
    '' ���̃{�^���N���A
    cmdF(5).Enabled = False
    cmdF(5).backColor = &H8000000F
    '���b�Z�[�W�G���A�N���A
    lblMsg.Caption = ""
End Sub

'*******************************************************************************************
'*    �֐���        : sub_DispSumple_Hanei_Ep_1
'*
'*    �����T�v      : 1.�T���v�����f(�G�s)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      i  �@�@     ,I  ,Integer   ,Spread�̍s�Ɏg�p���Ă���Y��
'*�@�@      �@�@      intNukisiRow�@,I  ,Integer   ,�����w���e�[�u���ʒu
'*�@�@      �@�@      intSmpkbn   ,I  ,Integer   ,�T���v���敪����\�T���v�������f����
'*
'*    �߂�l        : �Ȃ�
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
'*    �֐���        : sub_DispSumple_Hanei_Ep_2
'*
'*    �����T�v      : 1.�T���v�����f(�G�s)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      i  �@�@     ,I  ,Integer   ,Spread�̍s�Ɏg�p���Ă���Y��
'*�@�@      �@�@      intNukisiRow�@,I  ,Integer   ,�����w���e�[�u���ʒu
'*�@�@      �@�@      skensa1   ,I  ,Integer   ,�����p
'*�@�@      �@�@      skensa2   ,I  ,Integer   ,�����p
'*�@�@      �@�@      intZkbn�@   ,I  ,Integer   ,Z�敪
'*
'*    �߂�l        : �Ȃ�
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
                     '�R�s�[
                     tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                     tblNukishi(i + 1).EPRESB1CW = tblNukishi(i).EPRESB1CW
                 End If
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
                     '�R�s�[
                     tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                     tblNukishi(i + 1).EPRESB2CW = tblNukishi(i).EPRESB2CW
                 End If
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
                     '�R�s�[
                     tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                     tblNukishi(i + 1).EPRESB3CW = tblNukishi(i).EPRESB3CW
                 End If
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
             ElseIf intZkbn = 0 Then 'Z����Ȃ��Ƃ�
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
'*    �֐���        : sub_DispSumple_Hanei_Ep_3
'*
'*    �����T�v      : 1.�T���v�����f(�G�s)
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      i  �@�@     ,I  ,Integer   ,Spread�̍s�Ɏg�p���Ă���Y��
'*�@�@      �@�@      intNukisiRow�@,I  ,Integer   ,�����w���e�[�u���ʒu
'*�@�@      �@�@      skensa1   ,I  ,Integer   ,�����p
'*�@�@      �@�@      skensa2   ,I  ,Integer   ,�����p
'*�@�@      �@�@      intZkbn�@   ,I  ,Integer   ,Z�敪
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_DispSumple_Hanei_Ep_3(i As Integer, intNukisiRow As Integer, skensa1() As String, skensa2() As String, intZkbn As Integer)
    With sprExamine
        '' BMD1E
        .col = 29
        .backColor = vbWhite
        If .text <> "2" And .text <> "" Then
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDB1CW = "1" Then
                    tblNukishi(i + 1).EPINDB1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB1CW = tblNukishi(i).EPSMPLIDB1CW
                    tblNukishi(i + 1).EPRESB1CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB1CW = tblNukishi(i + 1).EPSMPLIDB1CW
                    tblNukishi(i).EPRESB1CW = tblNukishi(i + 1).EPRESB1CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDB2CW = "1" Then
                    tblNukishi(i + 1).EPINDB2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB2CW = tblNukishi(i).EPSMPLIDB2CW
                    tblNukishi(i + 1).EPRESB2CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB2CW = tblNukishi(i + 1).EPSMPLIDB2CW
                    tblNukishi(i).EPRESB2CW = tblNukishi(i + 1).EPRESB2CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDB3CW = "1" Then
                    tblNukishi(i + 1).EPINDB3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDB3CW = tblNukishi(i).EPSMPLIDB3CW
                    tblNukishi(i + 1).EPRESB3CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDB3CW = tblNukishi(i + 1).EPSMPLIDB3CW
                    tblNukishi(i).EPRESB3CW = tblNukishi(i + 1).EPRESB3CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDL1CW = "1" Then
                    tblNukishi(i + 1).EPINDL1CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL1CW = tblNukishi(i).EPSMPLIDL1CW
                    tblNukishi(i + 1).EPRESL1CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL1CW = tblNukishi(i + 1).EPSMPLIDL1CW
                    tblNukishi(i).EPRESL1CW = tblNukishi(i + 1).EPRESL1CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDL2CW = "1" Then
                    tblNukishi(i + 1).EPINDL2CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL2CW = tblNukishi(i).EPSMPLIDL2CW
                    tblNukishi(i + 1).EPRESL2CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL2CW = tblNukishi(i + 1).EPSMPLIDL2CW
                    tblNukishi(i).EPRESL2CW = tblNukishi(i + 1).EPRESL2CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
            If intZkbn = 3 Or intZkbn = 1 Then '�㑤Z
                If tblNukishi(i).EPINDL3CW = "1" Then
                    tblNukishi(i + 1).EPINDL3CW = "2"
                    tblNukishi(i + 1).EPSMPLIDL3CW = tblNukishi(i).EPSMPLIDL3CW
                    tblNukishi(i + 1).EPRESL3CW = "1"
                Else
                    tblNukishi(i).EPSMPLIDL3CW = tblNukishi(i + 1).EPSMPLIDL3CW
                    tblNukishi(i).EPRESL3CW = tblNukishi(i + 1).EPRESL3CW
                End If
            ElseIf intZkbn = 0 Then 'Z�ł͂Ȃ�
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
'*    �֐���        : fnc_ChkMukesaki
'*
'*    �����T�v      : 1.�I�����ꂽ�i�Ԃɑ΂��������`�F�b�N����
'*
'*    �p�����[�^    : �ϐ���      ,IO ,�^      �@,����
'*�@�@      �@�@      �Ȃ�
'*
'*    �߂�l        : ����I������FUNCTION_RETURN_SUCCESS(0),
'*                    �G���[�I������ FUNCTION_RETURN_FAILURE(-1)
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
'*    �֐���        : sub_cmbc039_3_ChangeHinSpec
'*
'*    �����T�v      : 1.WF�d�l�̃G�s�d�l�̕\���ؑ�
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@�@intCategory�@ ,I  ,Integer         ,�\���J�e�S��(0:WF�d�l,1:�G�s�d�l)
'*
'*    �߂�l        : �Ȃ�
'*
'*******************************************************************************************
Private Sub sub_cmbc039_3_ChangeHinSpec(Optional intCategory As Integer = 0)
    Dim i       As Long

    On Error Resume Next

    With f_cmbc039_3.sprSpec

        .ReDraw = False

        Select Case intCategory
        Case 0          ' WF�d�l�f�[�^�̕\��
            ' Rs(6���)�`AO(21���)
            For i = 6 To 21 Step 1
                .ColWidth(i) = 2.75
            Next i

            ' OT1(22���)
            i = 22:   .ColWidth(i) = 3

            ' GD(23���)
            i = 23:   .ColWidth(i) = 2.75

            ' B1E(24���)�`OT2(30���)
            For i = 24 To 30 Step 1
                .ColWidth(i) = 0      ' ��\��
            Next i
        Case 1          ' �G�s�d�l�f�[�^�̕\��
            ' Rs(6���)�`GD(23���)
            For i = 6 To 23 Step 1
                .ColWidth(i) = 0      ' ��\��
            Next i

            ' B1E(24���)�`OT2(30���)
            For i = 24 To 30 Step 1
                .ColWidth(i) = 3
            Next i
        Case Else

        End Select

        .ReDraw = True
    End With
End Sub

'***********************************************************************************
'*    �֐���        : fnc_GetMukesaki_XSDCB
'*
'*    �����T�v      : 1.���������iSXL�j��SXLID�ɑ΂�������\��
'*
'*    �p�����[�^    : �ϐ���        ,IO ,�^              ,����
'*�@�@�@�@�@�@�@�@�@�@sSXLid �@�@�@ ,I  ,String          ,SXL ID
'*
'*    �߂�l        : String(����)
'*
'***********************************************************************************
Private Function fnc_GetMukesaki_XSDCB(sSXLID As String) As FUNCTION_RETURN
    Dim sSql        As String
    Dim rs          As OraDynaset
    Dim lngRecCnt   As Long      '���R�[�h��
    Dim i           As Long

    '�G���[�n���h���̐ݒ�
    On Error GoTo proc_err
    gErr.Push "s_cmbc039.bas -- Function Getstaffauthority"

    fnc_GetMukesaki_XSDCB = FUNCTION_RETURN_FAILURE

    sSql = "Select PLANTCATCB "
    sSql = sSql & "from XSDCB "
    sSql = sSql & "where SXLIDCB = '" & Trim(sSXLID) & "' "

    ''�f�[�^�𒊏o����
    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''���o���ʂ��i�[����
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
    '�I��
    gErr.Pop
    Exit Function

proc_err:
    '�G���[�n���h��
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    Resume proc_exit
End Function

'�T�v      :�������ڂ̃X�v���b�h�̏�Ԃ���T���v���L���J�E���g����
'�@2008/03/21 aoyagi
Private Function sub_DispSample_SCnt01(IRow As Integer) As Integer
Dim strKval     As String
Dim i           As Integer '��
Dim iCnt        As Integer '����

    iCnt = 0
    
    '�X�v���b�h�̌�������(""�F��=vbWhite�A1�F��=vbBlack�A2�F���F=vbYellow)
    '                   1or2�F��ڰ=COLOR_CryJitsu
    With sprExamine
    
        .row = IRow
        
        ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
        For i = 11 To 26
            .col = i
            If .backColor = vbBlack Then
                iCnt = iCnt + 1
            End If
            If .backColor = vbYellow Then
                iCnt = iCnt + 1
            End If
            '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD START
            If i = 19 Then
                If .backColor = COLOR_CryJitsu Then
                    iCnt = iCnt + 1
                End If
            End If
            '��--- 2010/01/20 SIRD�Ή� SPK habuki ADD END
        Next i

        ''�c���_�f�������ڒǉ��ɂ��ύX�@03/12/09 ooba
        '--- 2006/08/15 Cng �G�s��s�]���ǉ��Ή� SMP)kondoh
        For i = 27 To 27
            .col = i
            If .backColor = vbBlack Then
                iCnt = iCnt + 1
            End If
            If .backColor = vbYellow Then
                iCnt = iCnt + 1
            End If
        Next i

        'GD�����w���\���ǉ��@05/02/10 ooba
        .col = 28
        If .backColor = vbBlack Then
            iCnt = iCnt + 1
        End If
        If .backColor = vbYellow Then
            iCnt = iCnt + 1
        End If
        
        ''BM1E�`OF3E
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
        If .text = 1 Then       '�����\���s�̓`�F�b�N�OOK�Ƃ���
            iCnt = 9999
        End If
        
        .col = 8
        If .text = "����" Then  '�����̓`�F�b�N�OOK�Ƃ���
            iCnt = 9999
        End If
        
        .col = 10
        If Trim(.text) = "" Then    '�T���v��ID���̓`�F�b�N�OOK�Ƃ���@08/08/01 ooba
            iCnt = 9999
        End If
        
    End With

    '������߂�
    sub_DispSample_SCnt01 = iCnt
    
End Function

'�T�v      :�������ڂ̃X�v���b�h�̏�Ԃ���L��SXL�����J�E���g����
'�@2008/03/21 aoyagi
Private Function sub_DispSample_SCnt02(IRow As Integer) As Integer
Dim i           As Integer '��
Dim iCnt        As Integer '����
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
    
    '������߂�
    sub_DispSample_SCnt02 = iCnt
    
    End With

End Function

'�T�v      :�������ڂ̃X�v���b�h�̏�Ԃ��瓯��u���b�N�s�����J�E���g����
'�@2008/03/21 aoyagi
Private Function sub_DispSample_SCnt03(IRow As Integer) As Integer
Dim i           As Integer '��
Dim iCnt        As Integer '����
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
    
    '������߂�
    sub_DispSample_SCnt03 = iCnt
    
    End With

End Function




