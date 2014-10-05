VERSION 5.00
Begin VB.Form f_cmmc001a 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "f_cmmc001a"
   ClientHeight    =   8205
   ClientLeft      =   1875
   ClientTop       =   2820
   ClientWidth     =   11880
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   547
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   792
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtRsBot2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   7770
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   15
      Text            =   "9999.9999"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6420
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   14
      Text            =   "9999.9999"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5070
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   13
      Text            =   "9999.9999"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   12
      Text            =   "9999.9999"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2370
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   11
      Text            =   "9999.9999"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   4
      Left            =   7770
      Locked          =   -1  'True
      TabIndex        =   143
      Text            =   "9999.9999"
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   3
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   142
      Text            =   "9999.9999"
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   5070
      Locked          =   -1  'True
      TabIndex        =   141
      Text            =   "9999.9999"
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   140
      Text            =   "9999.9999"
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtRsBot1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   139
      Text            =   "9999.9999"
      Top             =   3600
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   4
      Left            =   7770
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   10
      Text            =   "9999.9999"
      Top             =   3285
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   3
      Left            =   6420
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   9
      Text            =   "9999.9999"
      Top             =   3285
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   2
      Left            =   5070
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   8
      Text            =   "9999.9999"
      Top             =   3285
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   7
      Text            =   "9999.9999"
      Top             =   3285
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop2 
      Alignment       =   1  '右揃え
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   0
      Left            =   2370
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   6
      Text            =   "9999.9999"
      Top             =   3285
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   4
      Left            =   7770
      Locked          =   -1  'True
      TabIndex        =   138
      Text            =   "9999.9999"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   3
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   137
      Text            =   "9999.9999"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   5070
      Locked          =   -1  'True
      TabIndex        =   136
      Text            =   "9999.9999"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   135
      Text            =   "9999.9999"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtRsTop1 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   2370
      Locked          =   -1  'True
      TabIndex        =   134
      Text            =   "9999.9999"
      Top             =   3000
      Width           =   1200
   End
   Begin VB.TextBox txtSuccLenBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   128
      Text            =   "9,999"
      Top             =   6660
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   126
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   124
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   122
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccWtBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   8460
      Locked          =   -1  'True
      TabIndex        =   120
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtSuccWtBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   7380
      Locked          =   -1  'True
      TabIndex        =   118
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtSuccWtBot 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   6300
      Locked          =   -1  'True
      TabIndex        =   116
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtSuccLenTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   4284
      Locked          =   -1  'True
      TabIndex        =   109
      Text            =   "9,999"
      Top             =   6660
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   4284
      Locked          =   -1  'True
      TabIndex        =   107
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   3204
      Locked          =   -1  'True
      TabIndex        =   105
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccPosTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   2124
      Locked          =   -1  'True
      TabIndex        =   103
      Text            =   "9,999"
      Top             =   6408
      Width           =   660
   End
   Begin VB.TextBox txtSuccWtTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   4284
      Locked          =   -1  'True
      TabIndex        =   101
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtSuccWtTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   3204
      Locked          =   -1  'True
      TabIndex        =   99
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtSuccWtTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   2124
      Locked          =   -1  'True
      TabIndex        =   97
      Text            =   "999,999"
      Top             =   6156
      Width           =   768
   End
   Begin VB.TextBox txtEfehsTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   4284
      Locked          =   -1  'True
      TabIndex        =   96
      Text            =   "99.99999"
      Top             =   5904
      Width           =   1020
   End
   Begin VB.TextBox txtEfehsTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   3204
      Locked          =   -1  'True
      TabIndex        =   95
      Text            =   "99.99999"
      Top             =   5904
      Width           =   1020
   End
   Begin VB.TextBox txtEfehsTop 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   2124
      Locked          =   -1  'True
      TabIndex        =   94
      Text            =   "99.99999"
      Top             =   5904
      Width           =   1020
   End
   Begin VB.TextBox txtRpRs 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6336
      MaxLength       =   11
      TabIndex        =   20
      Text            =   "9999.999"
      Top             =   4950
      Width           =   1236
   End
   Begin VB.TextBox txtRpRs 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   6336
      Locked          =   -1  'True
      TabIndex        =   85
      Text            =   "9999.99999"
      Top             =   4680
      Width           =   1236
   End
   Begin VB.TextBox txtSpRsMax 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   3708
      MaxLength       =   11
      TabIndex        =   19
      Text            =   "9999.999"
      Top             =   4932
      Width           =   1236
   End
   Begin VB.TextBox txtSpRsMin 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1944
      MaxLength       =   11
      TabIndex        =   18
      Text            =   "9999.999"
      Top             =   4932
      Width           =   1236
   End
   Begin VB.TextBox txtSpRsMax 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   3708
      Locked          =   -1  'True
      TabIndex        =   82
      Text            =   "9999.99999"
      Top             =   4680
      Width           =   1236
   End
   Begin VB.TextBox txtSpRsMin 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   1944
      Locked          =   -1  'True
      TabIndex        =   80
      Text            =   "9999.99999"
      Top             =   4680
      Width           =   1236
   End
   Begin VB.TextBox txtHighStInMng 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   2  'ｵﾌ
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   17
      Text            =   "99"
      Top             =   4215
      Width           =   444
   End
   Begin VB.TextBox txtLowStInMng 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   2  'ｵﾌ
      Left            =   2268
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "99"
      Top             =   4212
      Width           =   444
   End
   Begin VB.TextBox txtRsRRG 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   3
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   74
      Text            =   "999.9"
      Top             =   3885
      Width           =   840
   End
   Begin VB.TextBox txtRsRRG 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   2
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   73
      Text            =   "999.9"
      Top             =   3600
      Width           =   840
   End
   Begin VB.TextBox txtRsRRG 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   1
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   72
      Text            =   "999.9"
      Top             =   3285
      Width           =   840
   End
   Begin VB.TextBox txtRsRRG 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   70
      Text            =   "999.9"
      Top             =   3000
      Width           =   840
   End
   Begin VB.TextBox txtTopcutWt 
      Alignment       =   1  '右揃え
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1041
         SubFormatType   =   0
      EndProperty
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   9756
      MaxLength       =   10
      TabIndex        =   5
      Text            =   "999,999"
      Top             =   2412
      Width           =   1020
   End
   Begin VB.TextBox txtCutAfterLen 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   6948
      MaxLength       =   4
      TabIndex        =   4
      Text            =   "9,999"
      Top             =   2412
      Width           =   1020
   End
   Begin VB.TextBox txtChargeWt 
      Alignment       =   1  '右揃え
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Index           =   1
      Left            =   1440
      MaxLength       =   10
      TabIndex        =   3
      Text            =   "9,999,999"
      Top             =   2412
      Width           =   990
   End
   Begin VB.TextBox txtTopcutWt 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   9756
      Locked          =   -1  'True
      TabIndex        =   58
      Text            =   "999,999"
      Top             =   2124
      Width           =   1020
   End
   Begin VB.TextBox txtCutAfterLen 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   6948
      Locked          =   -1  'True
      TabIndex        =   55
      Text            =   "9,999"
      Top             =   2124
      Width           =   1020
   End
   Begin VB.TextBox txtCutAfterWt 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   4356
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "999,999"
      Top             =   2124
      Width           =   1020
   End
   Begin VB.TextBox txtChargeWt 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Index           =   0
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "9,999,999"
      Top             =   2124
      Width           =   990
   End
   Begin VB.TextBox txtSuiRsHinban 
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Left            =   5985
      MaxLength       =   12
      TabIndex        =   2
      Text            =   "XXXXXXXXXXXX"
      Top             =   1692
      Width           =   1515
   End
   Begin VB.TextBox txtDateTopRou 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   1764
      Locked          =   -1  'True
      TabIndex        =   47
      Text            =   "1999/12/31"
      Top             =   1692
      Width           =   2040
   End
   Begin VB.TextBox txtRRG 
      Alignment       =   1  '右揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   9300
      Locked          =   -1  'True
      TabIndex        =   44
      Text            =   "999.9"
      Top             =   1332
      Width           =   1020
   End
   Begin VB.TextBox txtRpHinban 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   5976
      Locked          =   -1  'True
      TabIndex        =   42
      Text            =   "XXXXXXXXXXXX"
      Top             =   1332
      Width           =   1515
   End
   Begin VB.TextBox txtHoui 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   3276
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "999"
      Top             =   1332
      Width           =   516
   End
   Begin VB.TextBox txtType 
      Alignment       =   2  '中央揃え
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   1152
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "XX"
      Top             =   1332
      Width           =   516
   End
   Begin VB.TextBox txtCryNum 
      BackColor       =   &H00FFFFFF&
      Height          =   264
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1230
      MaxLength       =   12
      TabIndex        =   1
      Text            =   "XXXXXXXXXXXX"
      Top             =   900
      Width           =   1500
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
      Left            =   0
      TabIndex        =   34
      Top             =   7110
      Width           =   11895
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11]　　＊＊＊"
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
         Left            =   9900
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]　　＊＊＊"
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
         Left            =   8940
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
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
         Left            =   6900
         TabIndex        =   28
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
         Left            =   5940
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F６]　　＊＊＊"
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
         Left            =   4980
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F５]　　＊＊＊"
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
         Left            =   4020
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
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
         Left            =   2940
         TabIndex        =   24
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
         Left            =   7980
         TabIndex        =   29
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
         Left            =   60
         TabIndex        =   21
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
         Left            =   1020
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F３]　　ｷｬﾝｾﾙ"
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
         Left            =   1980
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]　　実行"
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
         Left            =   10860
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
   End
   Begin VB.Frame FraTitle 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      Begin VB.Label lblTime 
         Height          =   255
         Left            =   10350
         TabIndex        =   133
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblMsg 
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4695
         TabIndex        =   132
         Top             =   255
         Width           =   5550
      End
      Begin VB.Label lblTitle 
         Caption         =   "抵抗偏析計算処理"
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
         TabIndex        =   33
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Label Label1 
      Caption         =   "５"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   8265
      TabIndex        =   131
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "４"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   6915
      TabIndex        =   130
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   59
      Left            =   6984
      TabIndex        =   129
      Top             =   6696
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   63
      Left            =   9144
      TabIndex        =   127
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   61
      Left            =   8064
      TabIndex        =   125
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   58
      Left            =   6984
      TabIndex        =   123
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   62
      Left            =   9252
      TabIndex        =   121
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   60
      Left            =   8172
      TabIndex        =   119
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   57
      Left            =   7092
      TabIndex        =   117
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "BOT"
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
      Index           =   39
      Left            =   7455
      TabIndex        =   115
      Top             =   5430
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   417
      X2              =   632
      Y1              =   369
      Y2              =   369
   End
   Begin VB.Label Label1 
      Caption         =   "Ｃ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   45
      Left            =   8748
      TabIndex        =   114
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Ｂ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   44
      Left            =   7668
      TabIndex        =   113
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Ａ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   43
      Left            =   6588
      TabIndex        =   112
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "TOP"
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
      Index           =   38
      Left            =   3345
      TabIndex        =   111
      Top             =   5430
      Width           =   585
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   139
      X2              =   352
      Y1              =   369
      Y2              =   370
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   56
      Left            =   4968
      TabIndex        =   110
      Top             =   6696
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   55
      Left            =   4968
      TabIndex        =   108
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   53
      Left            =   3888
      TabIndex        =   106
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   51
      Left            =   2808
      TabIndex        =   104
      Top             =   6444
      Width           =   372
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   54
      Left            =   5076
      TabIndex        =   102
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   52
      Left            =   3996
      TabIndex        =   100
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   50
      Left            =   2916
      TabIndex        =   98
      Top             =   6192
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "Ｃ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   42
      Left            =   4644
      TabIndex        =   93
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Ｂ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   41
      Left            =   3564
      TabIndex        =   92
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "Ａ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   40
      Left            =   2484
      TabIndex        =   91
      Top             =   5688
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "合格長さ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   49
      Left            =   828
      TabIndex        =   90
      Top             =   6696
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "合格位置"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   48
      Left            =   828
      TabIndex        =   89
      Top             =   6444
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "合格重量"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   47
      Left            =   828
      TabIndex        =   88
      Top             =   6192
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "偏析値"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   46
      Left            =   828
      TabIndex        =   87
      Top             =   5940
      Width           =   1164
   End
   Begin VB.Label Label1 
      Caption         =   "＜計算結果＞"
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
      Index           =   37
      Left            =   660
      TabIndex        =   86
      Top             =   5430
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "ねらい抵抗"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   35
      Left            =   5184
      TabIndex        =   84
      Top             =   4716
      Width           =   1056
   End
   Begin VB.Label Label1 
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   36
      Left            =   3348
      TabIndex        =   83
      Top             =   5004
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "－"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   34
      Left            =   3348
      TabIndex        =   81
      Top             =   4752
      Width           =   228
   End
   Begin VB.Label Label1 
      Caption         =   "規格抵抗"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   33
      Left            =   612
      TabIndex        =   79
      Top             =   4716
      Width           =   1272
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   32
      Left            =   5508
      TabIndex        =   78
      Top             =   4248
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "上限規格内側管理"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   31
      Left            =   3384
      TabIndex        =   77
      Top             =   4248
      Width           =   1704
   End
   Begin VB.Label Label1 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   30
      Left            =   2736
      TabIndex        =   76
      Top             =   4248
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "下限規格内側管理"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Index           =   29
      Left            =   612
      TabIndex        =   75
      Top             =   4248
      Width           =   1704
   End
   Begin VB.Label Label1 
      Caption         =   "ＲＲＧ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   9510
      TabIndex        =   71
      Top             =   2820
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "３"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   5565
      TabIndex        =   69
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "２"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   4200
      TabIndex        =   68
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "１"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   2820
      TabIndex        =   67
      Top             =   2820
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "(BOT)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   28
      Left            =   1605
      TabIndex        =   66
      Top             =   3660
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "(TOP)"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   27
      Left            =   1605
      TabIndex        =   65
      Top             =   3060
      Width           =   585
   End
   Begin VB.Label Label1 
      Caption         =   "測定抵抗"
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
      Index           =   26
      Left            =   570
      TabIndex        =   64
      Top             =   3060
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   19
      Left            =   10800
      TabIndex        =   63
      Top             =   2484
      Width           =   264
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   16
      Left            =   7956
      TabIndex        =   62
      Top             =   2484
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   11
      Left            =   2484
      TabIndex        =   61
      Top             =   2484
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "（チャージ量）"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   240
      TabIndex        =   60
      Top             =   2445
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   18
      Left            =   10800
      TabIndex        =   59
      Top             =   2196
      Width           =   264
   End
   Begin VB.Label Label1 
      Caption         =   "トップカット重量"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   17
      Left            =   8340
      TabIndex        =   57
      Top             =   2175
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      Index           =   15
      Left            =   7956
      TabIndex        =   56
      Top             =   2196
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "本切断後長さ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   14
      Left            =   5760
      TabIndex        =   54
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Index           =   13
      Left            =   5400
      TabIndex        =   53
      Top             =   2196
      Width           =   336
   End
   Begin VB.Label Label1 
      Caption         =   "本切断後重量"
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
      Index           =   12
      Left            =   3090
      TabIndex        =   51
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   10
      Left            =   2484
      TabIndex        =   50
      Top             =   2196
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "ルツボ内量"
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
      Index           =   8
      Left            =   330
      TabIndex        =   48
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "推定抵抗品番"
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
      Index           =   6
      Left            =   4755
      TabIndex        =   46
      Top             =   1725
      Width           =   1230
   End
   Begin VB.Label Label1 
      Caption         =   "測定日(TOP ρ)"
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
      Index           =   7
      Left            =   330
      TabIndex        =   45
      Top             =   1725
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "ＲＲＧ"
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
      Index           =   5
      Left            =   8760
      TabIndex        =   43
      Top             =   1365
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "ねらい品番"
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
      Index           =   4
      Left            =   4980
      TabIndex        =   41
      Top             =   1365
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "＜　　　　　＞"
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
      Index           =   3
      Left            =   3000
      TabIndex        =   39
      Top             =   1365
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "結晶軸"
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
      Index           =   2
      Left            =   2235
      TabIndex        =   38
      Top             =   1365
      Width           =   750
   End
   Begin VB.Label Label1 
      Caption         =   "タイプ"
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
      Index           =   1
      Left            =   330
      TabIndex        =   36
      Top             =   1365
      Width           =   810
   End
   Begin VB.Label Label1 
      Caption         =   "結晶番号"
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
      Index           =   0
      Left            =   330
      TabIndex        =   35
      Top             =   930
      Width           =   870
   End
End
Attribute VB_Name = "f_cmmc001a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ResTemp(3, 4) As Double
Dim ResTemp1(1, 2) As Double

'概要      :計算処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:dRsTop()　　　,O  ,Double         　,トップ側の測定抵抗値配列
'      　　:dRsBot()　　　,O  ,Double         　,ボトム側の測定抵抗値配列
'          :errMsg        ,O  ,String           ,エラー箇所
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function CalculateProcess(dRsTop() As Double, dRsBot() As Double, errMsg$) As FUNCTION_RETURN
Dim Index       As Integer
Dim dDM         As Double
Dim tHenInf     As type_Coefficient
Dim dHenseki(2) As Double
Dim tCalcPosInf As type_ResPosCal
Dim dPassPosTop(2) As Double
Dim dPassPosBot(2) As Double
Dim dPassWtTop(2)  As Double
Dim dPassWtBot(2)  As Double
Dim dPassLenTop    As Double
Dim dPassLenBot    As Double
Dim dLen        As Double
Dim dWt         As Double

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function CalculateProcess"

    CalculateProcess = FUNCTION_RETURN_FAILURE
1:
    '' 本切断後重量の計算と表示
    txtCutAfterWt.Text = MakeParamWeight(GetCutAfterWt(Val(txtCutAfterLen(1).Text)))
2:
    '' RRGの計算と表示
    ShowRrgParam
3:
    '' ねらい抵抗の計算と表示
    ShowRpRsParam
4:
    With g_PlupEndRslt
        dDM = (.DM1 + .DM2 + .DM3) / 3#
    End With
5:
    '' 偏析値の計算
    For Index = 0 To 2
        tHenInf.DUNMENSEKI = AreaOfCircle(dDM)
        If UBound(g_tblRs) > 0 Then
            tHenInf.TOPSMPLPOS = g_tblRs(1).POSITION
            tHenInf.BOTSMPLPOS = g_tblRs(UBound(g_tblRs)).POSITION
        Else
            tHenInf.TOPSMPLPOS = 0
            tHenInf.BOTSMPLPOS = Val(txtCutAfterLen(1).Text)
        End If
        tHenInf.CHARGEWEIGHT = Val(txtChargeWt(1).Text)
        tHenInf.TOPWEIGHT = g_PlupEndRslt.WGHTTOP + Val(txtTopcutWt(1).Text)
        tHenInf.TOPRES = dRsTop(Index)
        tHenInf.BOTRES = dRsBot(Index)
        dHenseki(Index) = CoefficientCalculation(tHenInf)
        If dHenseki(Index) = -9999 Then
            '偏析係数計算エラー
            errMsg = "偏析係数"
            GoTo proc_exit
        End If
    Next Index
6:
    '' 合格位置の計算（TOP）
    For Index = 0 To 2
        tCalcPosInf.COEFFICIENT = dHenseki(Index)
        tCalcPosInf.DUNMENSEKI = AreaOfCircle(dDM)
        tCalcPosInf.CHARGEWEIGHT = Val(txtChargeWt(1).Text)
        tCalcPosInf.TOPWEIGHT = g_PlupEndRslt.WGHTTOP + Val(txtTopcutWt(1).Text)
        If UBound(g_tblRs) > 0 Then
            tCalcPosInf.TOPSMPLPOS = g_tblRs(1).POSITION
        Else
            tCalcPosInf.TOPSMPLPOS = 0
        End If
        tCalcPosInf.TOPRES = dRsTop(Index)
        tCalcPosInf.target = Val(txtSpRsMax(1).Text) * (1 - (Val(txtHighStInMng.Text) / 100))
        dPassPosTop(Index) = PosCalculation(tCalcPosInf)
        If dPassPosTop(Index) = -9999 Then
            '合格位置(TOP)推定計算エラー
            errMsg = "合格位置(TOP)推定"
            GoTo proc_exit
        End If
    Next Index
7:
    '' 合格位置の計算（BOT）
    For Index = 0 To 2
        tCalcPosInf.COEFFICIENT = dHenseki(Index)
        tCalcPosInf.DUNMENSEKI = AreaOfCircle(dDM)
        tCalcPosInf.CHARGEWEIGHT = Val(txtChargeWt(1).Text)
        tCalcPosInf.TOPWEIGHT = g_PlupEndRslt.WGHTTOP + Val(txtTopcutWt(1).Text)
        If UBound(g_tblRs) > 0 Then
            tCalcPosInf.TOPSMPLPOS = g_tblRs(1).POSITION
        Else
            tCalcPosInf.TOPSMPLPOS = 0
        End If
        tCalcPosInf.TOPRES = dRsTop(Index)
        tCalcPosInf.target = Val(txtSpRsMin(1).Text) * (1 + (Val(txtLowStInMng.Text) / 100))
        dPassPosBot(Index) = PosCalculation(tCalcPosInf)
        If dPassPosBot(Index) = -9999 Then
            '合格位置(BOT)推定計算エラー
            errMsg = "合格位置(BOT)推定"
            GoTo proc_exit
        End If
    Next Index
8:
    '' 合格長さの決定(TOP)
    dPassLenTop = GetMax(dPassPosTop)
    If dPassLenTop < 0 Then dPassLenTop = 0
9:
    '' 合格長さの決定(BOT)
    dPassLenBot = GetMin(dPassPosBot)
    If dPassLenBot < 0 Then dPassLenBot = 0
10:
    '-*-*- 人見　０除算エラーの対処　20010804
    dLen = Val(txtCutAfterLen(1).Text)
    If dLen = 0 Then GoTo proc_exit
    dWt = GetCutAfterWt(dLen)
11:
    '' 合格重量の計算（TOP）
    For Index = 0 To 2
        dPassWtTop(Index) = dWt * (dPassPosTop(Index) / dLen)
    Next Index
12:
    '' 合格重量の計算（BOT）
    For Index = 0 To 2
        dPassWtBot(Index) = dWt * (dPassPosBot(Index) / dLen)
    Next Index
13:
    '' 計算値の表示
    For Index = 0 To 2
        txtEfehsTop(Index).Text = MakeParamCoefficient(dHenseki(Index))
    Next Index
    For Index = 0 To 2
        txtSuccPosTop(Index).Text = MakeParamLength(dPassPosTop(Index))
    Next Index
    For Index = 0 To 2
        txtSuccPosBot(Index).Text = MakeParamLength(dPassPosBot(Index))
    Next Index
    For Index = 0 To 2
        txtSuccWtTop(Index).Text = MakeParamWeight(dPassWtTop(Index))
    Next Index
    For Index = 0 To 2
        txtSuccWtBot(Index).Text = MakeParamWeight(dPassWtBot(Index))
    Next Index
    txtSuccLenTop.Text = MakeParamLength(dPassLenTop)
    txtSuccLenBot.Text = MakeParamLength(dPassLenBot)
14:
    CalculateProcess = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :実行処理を行う
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function ExecutionProcess() As FUNCTION_RETURN
Dim iRet As Integer
Dim dRsTop(2) As Double
Dim dRsBot(2) As Double
Dim errMsg As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function ExecutionProcess"

    ExecutionProcess = FUNCTION_RETURN_FAILURE

    '' 入力項目のチェック
    Call CheckInputText(txtChargeWt(1), 7, 0)
    Call CheckInputText(txtCutAfterLen(1), 4, 0)
    Call CheckInputText(txtTopcutWt(1), 6, 0)
    Call CheckInputText(txtLowStInMng, 2, 0)
    Call CheckInputText(txtHighStInMng, 2, 0)
    Call CheckInputText(txtSpRsMin(1), 4, 4)
    Call CheckInputText(txtSpRsMax(1), 4, 4)
    Call CheckInputText(txtRpRs(1), 4, 4)

    '' 測定点を決定（３点or５点測定）して、測定値を取得する
    iRet = CheckMeasPoint(dRsTop, dRsBot)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        GoTo proc_exit
    End If

    '' 計算を行う
    iRet = CalculateProcess(dRsTop, dRsBot, errMsg)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        lblMsg.Caption = GetMsgStr("ECLC1", errMsg)
        GoTo proc_exit
    End If

    '' 処理正常終了
    ExecutionProcess = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :テキスト入力のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名     ,IO ,型               ,説明
'      　　:tbox　　　 ,IO ,TextBox        　,テキストボックス
'      　　:uplen　　　,I  ,Integer        　,整数部桁数
'      　　:lwlen　　　,I  ,Integer        　,小数点以下桁数
'      　　:戻り値     ,O  ,FUNCTION_RETURN　,チェックの成否
'説明      :
Private Function CheckInputText(tbox As TextBox, ByVal uplen As Integer, ByVal lwlen As Integer) As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function CheckInputText"

    Select Case ChkNumber(tbox.Text, uplen, lwlen)
    Case CHK_NG
        tbox.BackColor = COLOR_NG
        lblMsg.Caption = GetMsgStr("EINPM")
        CheckInputText = FUNCTION_RETURN_FAILURE
    Case CHK_NULL
        tbox.BackColor = COLOR_NG
        lblMsg.Caption = GetMsgStr("EINIM")
        CheckInputText = FUNCTION_RETURN_FAILURE
    Case Else
        tbox.BackColor = COLOR_OK
        lblMsg.Caption = ""
        CheckInputText = FUNCTION_RETURN_SUCCESS
    End Select

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :測定点を決定（３点or５点測定）して、測定値を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:dRsTop()　　　,O  ,Double         　,トップ側の測定抵抗値配列
'      　　:dRsBot()　　　,O  ,Double         　,ボトム側の測定抵抗値配列
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function CheckMeasPoint(dRsTop() As Double, dRsBot() As Double) As FUNCTION_RETURN
Dim Index  As Integer
Dim bNonRs As Boolean
Dim bNotIn As Boolean

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function CheckMeasPoint"

    CheckMeasPoint = FUNCTION_RETURN_FAILURE

    bNonRs = False
    bNotIn = False

    '' トップ側測定抵抗、ボトム側測定抵抗を調べて測定点の決定を行う
    For Index = 0 To 4
        If Val(txtRsTop2(Index).Text) <= 0 Then
            bNonRs = True
        End If
        If Val(txtRsBot2(Index).Text) <= 0 Then
            bNonRs = True
        End If
    Next

    If bNonRs <> True Then
        '' ５点測定（１点目、４点目、５点目が計算対象）
        '' 測定抵抗値を取得
        dRsTop(0) = Val(txtRsTop2(0).Text)
        dRsTop(1) = Val(txtRsTop2(3).Text)
        dRsTop(2) = Val(txtRsTop2(4).Text)
        dRsBot(0) = Val(txtRsBot2(0).Text)
        dRsBot(1) = Val(txtRsBot2(3).Text)
        dRsBot(2) = Val(txtRsBot2(4).Text)
        For Index = 0 To 4
            '' 測定値入力項目を有効にする
            CtrlEnabled txtRsTop2(Index), CTRL_ENABLE
            CtrlEnabled txtRsBot2(Index), CTRL_ENABLE
        Next Index
    Else
        '' ３点測定（１点目、２点目、３点目が計算対象）
        For Index = 0 To 2
            If Val(txtRsTop2(Index).Text) <= 0 Then
                CtrlEnabled txtRsTop2(Index), CTRL_WARNING: bNotIn = True
            Else
                CtrlEnabled txtRsTop2(Index), CTRL_ENABLE
            End If
            If Val(txtRsBot2(Index).Text) <= 0 Then
                CtrlEnabled txtRsBot2(Index), CTRL_WARNING: bNotIn = True
            Else
                CtrlEnabled txtRsBot2(Index), CTRL_ENABLE
            End If
            '' 測定抵抗値を取得
            dRsTop(Index) = Val(txtRsTop2(Index).Text)
            dRsBot(Index) = Val(txtRsBot2(Index).Text)
        Next Index
        '' 未入力項目があった場合
        If bNotIn = True Then
            lblMsg.Caption = GetMsgStr(MSG_NOTINPUT_ERROR)
            GoTo proc_exit
        End If
    End If

    CheckMeasPoint = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :ファンクションボタンクリック処理
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
'　　      :Index        ,I  ,Integer　,コントロール配列の添字
'説明      :ファンクションボタンがクリックされたら、各処理に分岐する
'履歴      :
Private Sub cmdF_Click(Index As Integer)
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub cmdF_Click"

    Select Case Index
'    Case 1              '' Ｆ１キー（メインメニュー）
'       unload me
    Case 2              '' Ｆ２キー（サブメニュー）
        '' サブメニューに戻る
        CloseFormProc f_cmhc001a, Me
        GoTo proc_exit
    Case 3              '' Ｆ３キー（キャンセル）
        '' フォーム初期化処理
        cmdF(Index).Enabled = False
        InitForm
        cmdF(Index).Enabled = True
    Case 12             '' Ｆ１２キー（実行）
        cmdF(Index).Enabled = False
        BeginProcess    '' プロセス開始
        '' 実行処理を行う
        If ExecutionProcess <> FUNCTION_RETURN_SUCCESS Then
            EndProcess  '' プロセス終了
        End If
        EndProcess      '' プロセス終了
        '' 処理完了時間セット
        SetPresentTime lblTime
        cmdF(Index).Enabled = True
    End Select

    cmdF(Index).Enabled = True

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    cmdF(Index).Enabled = True
    Resume proc_exit
End Sub

'概要      :キーボード押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
'　　      :KeyCode      ,I  ,Integer　,キーコード
'　　      :Shift        ,I  ,Integer　,Shiftキーの状態
'説明      :キーが押されたら、各処理に分岐する
'履歴      :
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
Dim Index As Integer
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmhc001a.frm -- Sub Form_KeyUp"

    Select Case KeyCode
      Case vbKeyF1
        Index = 1
      Case vbKeyF2
        Index = 2
      Case vbKeyF3
        Index = 3
      Case vbKeyF4
        Index = 4
      Case vbKeyF5
        Index = 5
      Case vbKeyF6
        Index = 6
      Case vbKeyF7
        Index = 7
      Case vbKeyF8
        Index = 8
      Case vbKeyF9
        Index = 9
      Case vbKeyF10
        Index = 10
      Case vbKeyF11
        Index = 11
      Case vbKeyF12
        Index = 12
      Case Else
        GoTo proc_exit
    End Select
    If cmdF(Index).Visible And cmdF(Index).Enabled Then
        cmdF_Click Index
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :フォームのロード
'説明      :
Private Sub Form_Load()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub Form_Load"
    
    ReDim g_tblRs(0)

    '' フォーム位置セット
    CenterForm Me
    '' 処理時間セット
    SetPresentTime lblTime

    '' フォームを表示する。
    Me.Show
    '' フォーム初期化処理
    InitForm


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :フォームのコントロールの状態を設定する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型                 ,説明
'          :state         ,I   ,enm_CtrlStateKind   ,入力可能項目の状態
'          :bClear        ,I   ,Boolean             ,入力可能項目クリア指定(True:クリアする False:クリアしない)
'          :[stateDisp]   ,I   ,enm_CtrlStateKind   ,表示項目の状態
'          :[bDispClear]  ,I   ,Boolean             ,表示項目クリア指定(True:クリアする False:クリアしない)
'説明      :
Private Sub SetFormCtrlEnabled(state As enm_CtrlStateKind, bClear As Boolean, Optional stateDisp As enm_CtrlStateKind = CTRL_DISABLE, Optional bDispClear As Boolean = False)
    Dim Index As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub SetFormCtrlEnabled"

    lblMsg.Caption = ""
    CtrlEnabled txtType, stateDisp, bDispClear
    CtrlEnabled txtHoui, stateDisp, bDispClear
    CtrlEnabled txtRpHinban, stateDisp, bDispClear
    CtrlEnabled txtRRG, stateDisp, bDispClear
    CtrlEnabled txtDateTopRou, stateDisp, bDispClear
    CtrlEnabled txtSuiRsHinban, state, bClear
    
    CtrlEnabled txtChargeWt(0), stateDisp, bDispClear
    CtrlEnabled txtChargeWt(1), state, bClear
    CtrlEnabled txtCutAfterWt, stateDisp, bDispClear
    CtrlEnabled txtCutAfterLen(0), stateDisp, bDispClear
    CtrlEnabled txtCutAfterLen(1), state, bClear
    CtrlEnabled txtTopcutWt(0), stateDisp, bDispClear
    CtrlEnabled txtTopcutWt(1), state, bClear
    
    For Index = 0 To 4
        CtrlEnabled txtRsTop1(Index), stateDisp, bDispClear
        CtrlEnabled txtRsTop2(Index), state, bClear
        CtrlEnabled txtRsBot1(Index), stateDisp, bDispClear
        CtrlEnabled txtRsBot2(Index), state, bClear
    Next Index
    
    For Index = 0 To 3
        CtrlEnabled txtRsRRG(Index), stateDisp, bDispClear
    Next Index
    
    CtrlEnabled txtLowStInMng, state, bClear
    CtrlEnabled txtHighStInMng, state, bClear
    
    CtrlEnabled txtSpRsMin(0), stateDisp, bDispClear
    CtrlEnabled txtSpRsMin(1), state, bClear
    CtrlEnabled txtSpRsMax(0), stateDisp, bDispClear
    CtrlEnabled txtSpRsMax(1), state, bClear
    CtrlEnabled txtRpRs(0), stateDisp, bDispClear
    CtrlEnabled txtRpRs(1), state, bClear
    
    For Index = 0 To 2
        CtrlEnabled txtEfehsTop(Index), stateDisp, bDispClear
        CtrlEnabled txtSuccWtTop(Index), stateDisp, bDispClear
        CtrlEnabled txtSuccPosTop(Index), stateDisp, bDispClear
    Next Index
    
    CtrlEnabled txtSuccLenTop, stateDisp, bDispClear
    
    For Index = 0 To 2
        CtrlEnabled txtSuccWtBot(Index), stateDisp, bDispClear
        CtrlEnabled txtSuccPosBot(Index), stateDisp, bDispClear
    Next Index
    
    CtrlEnabled txtSuccLenBot, stateDisp, bDispClear


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :フォーム初期化処理
'説明      :
Private Sub InitForm()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub InitForm"

    '' フォーカスセット
    txtCryNum.SetFocus

    '' 各コントロール初期化
    CtrlEnabled txtCryNum, CTRL_ENABLE, True
    SetFormCtrlEnabled CTRL_DISABLE, True, CTRL_DISABLE, True
    cmdF(12).Enabled = False

    '' 初期表示パラメータ初期化処理
    InitDisplay


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :初期表示パラメータ初期化処理
'説明      :
Private Sub InitDisplay()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub InitDisplay"

    '' 下限規格内側管理
    txtLowStInMng.Text = "0"
    '' 上限規格内側管理
    txtHighStInMng.Text = "0"


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub Form_Unload"

    '' メニューに戻る
    f_cmhc001a.Show


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :結晶番号入力チェック処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtCryNum_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iRet    As Integer
Dim tHinInf As tFullHinban
Dim tblRpHinban As typ_TBCME037
Dim strCryNum As String


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtCryNum_KeyDown"

    If KeyCode = vbKeyReturn Then
        '' 表示項目になっている場合、処理終了
        If txtCryNum.Locked = True Then GoTo proc_exit
        '' 結晶番号を取得する
        strCryNum = Trim(txtCryNum.Text)
        '' 結晶番号より、ねらい品番を取得する
        iRet = GetRpHinban(tblRpHinban, strCryNum)
        If iRet <> FUNCTION_RETURN_SUCCESS Then
            lblMsg.Caption = GetMsgStr(MSG_NOTFOUND_CRYNUM)
            GoTo proc_exit
        End If
        '' 結晶情報、および、製品仕様SXLデータ、抵抗実績情報等を表示する
        iRet = ShowRsInfo(tblRpHinban)
        If iRet <> FUNCTION_RETURN_SUCCESS Then
            '' このエラー表示は、ShowRsInfo()内で行う
            GoTo proc_exit
        End If
        '' フォームの項目を入力可能にする
        SetFormCtrlEnabled CTRL_ENABLE, False
        '' 結晶番号入力エリアを入力不可にする
        CtrlEnabled txtCryNum, CTRL_DISABLE
        '' フォーカスセット
        txtChargeWt(1).SetFocus
        '' ボタン有効
        cmdF(12).Enabled = True
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :結晶情報、および、製品仕様SXLデータ、抵抗実績情報等を表示する
'ﾊﾟﾗﾒｰﾀ　　:変数名           ,IO ,型               ,説明
'      　　:tblRpHinban　　　,I  ,typ_TBCME037   　,結晶情報テーブル
'      　　:戻り値           ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function ShowRsInfo(tblRpHinban As typ_TBCME037) As FUNCTION_RETURN
Dim iRet    As Integer
Dim tHinInf As tFullHinban
Dim dDM     As Double       '直径
Dim dWTop   As Double       'Top重量
Dim dWTCut  As Double       'TopCut重量
Dim lCharge As Long         'チャージ量


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function ShowRsInfo"

    ShowRsInfo = FUNCTION_RETURN_FAILURE
    
    '' ねらい品番より、製品仕様SXLデータ1を取得する
    tHinInf.HINBAN = tblRpHinban.RPHINBAN
    tHinInf.mnorevno = tblRpHinban.RPREVNUM
    tHinInf.factory = tblRpHinban.RPFACT
    tHinInf.opecond = tblRpHinban.RPOPCOND
    iRet = GetSPSXLData1(g_PrSpSXLData1, tHinInf)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        lblMsg.Caption = GetMsgStr(MSG_GETERROR_DBDATA, "E018")
        GoTo proc_exit
    End If
    
    '' 引上げ終了実績を取得する
    iRet = GetPlupEndRslt(g_PlupEndRslt, tblRpHinban.CRYNUM)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        lblMsg.Caption = GetMsgStr(MSG_GETERROR_DBDATA, "H004")
        GoTo proc_exit
    End If
    
    ''偏析計算用パラメータを取得して、引上げ終了実績の値を差し替える
    If GetCoeffParams(Trim$(txtCryNum.Text), lCharge, dWTop, dWTCut, dDM) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = GetMsgStr(MSG_NOTFOUND_CRYNUM)
        GoTo proc_exit
    End If
    With g_PlupEndRslt
        .CHARGE = lCharge
    End With
    
    '' 比抵抗実績を取得する
    ReDim g_tblRs(0)
    iRet = GetResultsRs(g_tblRs, tblRpHinban.CRYNUM)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
'        lblMsg.Caption = GetMsgStr("PNORS")
'        GoTo proc_exit
        '' 抵抗実績がなくても処理継続
    End If
        

    '' 取得データの情報の表示を行う
    iRet = ShowRsInfoParam(g_tblRs)
    If iRet <> FUNCTION_RETURN_SUCCESS Then
        lblMsg.Caption = GetMsgStr(MSG_DISPLAY_ERROR)
        GoTo proc_exit
    End If

    ShowRsInfo = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶情報、および、製品仕様SXLデータ、抵抗実績情報等を表示する（詳細）
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:tblRs()　　　,I  ,typ_TBCMJ002   　,比抵抗実績テーブル配列(1～)
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function ShowRsInfoParam(tblRs() As typ_TBCMJ002) As FUNCTION_RETURN
Dim Index As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function ShowRsInfoParam"

    ShowRsInfoParam = FUNCTION_RETURN_FAILURE

    txtType.Text = g_PrSpSXLData1.HSXTYPE
    txtHoui.Text = g_PrSpSXLData1.HSXCDIR
    txtRpHinban.Text = g_PrSpSXLData1.HINBAN + Format(g_PrSpSXLData1.mnorevno, "00") + g_PrSpSXLData1.factory + g_PrSpSXLData1.opecond
    txtRRG.Text = MakeParamRs(g_PrSpSXLData1.HSXRMBNP)
    If UBound(tblRs) > 0 Then txtDateTopRou.Text = tblRs(1).REGDATE
    txtSuiRsHinban.Text = ""

    txtChargeWt(0).Text = MakeParamWeight(g_PlupEndRslt.CHARGE)
    txtChargeWt(1).Text = txtChargeWt(0).Text
    txtCutAfterWt.Text = MakeParamWeight(GetCutAfterWt(g_PlupEndRslt.LENGFREE))

    txtCutAfterLen(0).Text = g_PlupEndRslt.LENGFREE
    txtCutAfterLen(1).Text = txtCutAfterLen(0).Text
    txtTopcutWt(0).Text = MakeParamWeight(g_PlupEndRslt.WGTOPCUT)
    txtTopcutWt(1).Text = txtTopcutWt(0).Text

    For Index = 1 To UBound(tblRs)
        If tblRs(Index).MEAS1 >= 0 Then
            txtRsTop1(0).Text = MakeParamRs(tblRs(Index).MEAS1, "")
            txtRsTop1(1).Text = MakeParamRs(tblRs(Index).MEAS2, "")
            txtRsTop1(2).Text = MakeParamRs(tblRs(Index).MEAS3, "")
            txtRsTop1(3).Text = MakeParamRs(tblRs(Index).MEAS4, "")
            txtRsTop1(4).Text = MakeParamRs(tblRs(Index).MEAS5, "")
            ResTemp(0, 0) = tblRs(Index).MEAS1
            ResTemp(0, 1) = tblRs(Index).MEAS2
            ResTemp(0, 2) = tblRs(Index).MEAS3
            ResTemp(0, 3) = tblRs(Index).MEAS4
            ResTemp(0, 4) = tblRs(Index).MEAS5
            ResTemp(1, 0) = tblRs(Index).MEAS1
            ResTemp(1, 1) = tblRs(Index).MEAS2
            ResTemp(1, 2) = tblRs(Index).MEAS3
            ResTemp(1, 3) = tblRs(Index).MEAS4
            ResTemp(1, 4) = tblRs(Index).MEAS5
            Exit For
        End If
    Next Index
    For Index = UBound(tblRs) To 1 Step -1
        If tblRs(Index).MEAS1 >= 0 Then
            txtRsBot1(0).Text = MakeParamRs(tblRs(Index).MEAS1, "")
            txtRsBot1(1).Text = MakeParamRs(tblRs(Index).MEAS2, "")
            txtRsBot1(2).Text = MakeParamRs(tblRs(Index).MEAS3, "")
            txtRsBot1(3).Text = MakeParamRs(tblRs(Index).MEAS4, "")
            txtRsBot1(4).Text = MakeParamRs(tblRs(Index).MEAS5, "")
            ResTemp(2, 0) = tblRs(Index).MEAS1
            ResTemp(2, 1) = tblRs(Index).MEAS2
            ResTemp(2, 2) = tblRs(Index).MEAS3
            ResTemp(2, 3) = tblRs(Index).MEAS4
            ResTemp(2, 4) = tblRs(Index).MEAS5
            ResTemp(3, 0) = tblRs(Index).MEAS1
            ResTemp(3, 1) = tblRs(Index).MEAS2
            ResTemp(3, 2) = tblRs(Index).MEAS3
            ResTemp(3, 3) = tblRs(Index).MEAS4
            ResTemp(3, 4) = tblRs(Index).MEAS5
            Exit For
        End If
    Next Index

    For Index = 0 To 4
        txtRsTop2(Index).Text = txtRsTop1(Index).Text
        txtRsBot2(Index).Text = txtRsBot1(Index).Text
    
        ResValSet Index
    Next Index

    txtSpRsMin(0).Text = MakeParamRs(g_PrSpSXLData1.HSXRMIN, "")
    txtSpRsMin(1).Text = txtSpRsMin(0).Text
    txtSpRsMax(0).Text = MakeParamRs(g_PrSpSXLData1.HSXRMAX, "")
    txtSpRsMax(1).Text = txtSpRsMax(0).Text
    ResTemp1(0, 0) = g_PrSpSXLData1.HSXRMIN
    ResTemp1(1, 0) = g_PrSpSXLData1.HSXRMIN
    ResTemp1(0, 1) = g_PrSpSXLData1.HSXRMAX
    ResTemp1(1, 1) = g_PrSpSXLData1.HSXRMAX
    '' RRG計算
    ShowRrgParam
    '' ねらい抵抗計算
    ShowRpRsParam
    ResValSet1

    ShowRsInfoParam = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :RRGを計算して表示する
'説明      :
Private Sub ShowRrgParam()
Dim Index    As Integer
Dim dParam() As Double
Dim dGet     As Double

    ReDim dParam(4)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub ShowRrgParam"

    For Index = 0 To 4
        If txtRsTop1(Index).Text = "" Then Exit For
            dParam(Index) = Val(txtRsTop1(Index).Text)
    Next Index
    If Index > 0 Then
        ReDim Preserve dParam(Index - 1)
        txtRsRRG(0).Text = MakeParamRs(GetRG(dParam))
    End If

    ReDim dParam(4)
    For Index = 0 To 4
        If txtRsTop2(Index).Text = "" Then Exit For
            dParam(Index) = Val(txtRsTop2(Index).Text)
    Next Index
    If Index > 0 Then
        ReDim Preserve dParam(Index - 1)
        txtRsRRG(1).Text = MakeParamRs(GetRG(dParam))
    End If

    ReDim dParam(4)
    For Index = 0 To 4
        If txtRsBot1(Index).Text = "" Then Exit For
            dParam(Index) = Val(txtRsBot1(Index).Text)
    Next Index
    If Index > 0 Then
        ReDim Preserve dParam(Index - 1)
        txtRsRRG(2).Text = MakeParamRs(GetRG(dParam))
    End If

    ReDim dParam(4)
    For Index = 0 To 4
        If txtRsBot2(Index).Text = "" Then Exit For
            dParam(Index) = Val(txtRsBot2(Index).Text)
    Next Index
    If Index > 0 Then
        ReDim Preserve dParam(Index - 1)
        txtRsRRG(3).Text = MakeParamRs(GetRG(dParam))
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :ねらい抵抗を計算して表示する
'説明      :
Private Sub ShowRpRsParam()
Dim dParam

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub ShowRpRsParam"

    dParam = ResTemp1(0, 1) 'Val(txtSpRsMax(0).Text)
    txtRpRs(0).Text = MakeParamRs((dParam - (dParam * Val(txtHighStInMng.Text) / 100#)) * 0.97, "")
    ResTemp1(0, 2) = (dParam - (dParam * Val(txtHighStInMng.Text) / 100#)) * 0.97
    
    dParam = ResTemp1(1, 1) 'Val(txtSpRsMax(1).Text)
    txtRpRs(1).Text = MakeParamRs((dParam - (dParam * Val(txtHighStInMng.Text) / 100#)) * 0.97, "")
    ResTemp1(1, 2) = (dParam - (dParam * Val(txtHighStInMng.Text) / 100#)) * 0.97


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :ねらい抵抗より規格抵抗値Maxを計算して表示する
'説明      :
Private Sub ShowSpRsMaxParam()
Dim dParam

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub ShowSpRsMaxParam"

    dParam = ResTemp1(1, 2) 'Val(txtRpRs(1).Text)
    txtSpRsMax(1).Text = MakeParamRs((100 * dParam) / (0.97 * (100 - Val(txtHighStInMng.Text))), "")
    ResTemp1(1, 1) = (100 * dParam) / (0.97 * (100 - Val(txtHighStInMng.Text)))


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub


'概要      :数値の丸めを行う（偏析値）
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型      ,説明
'      　　:param 　　　,I  ,Double　,数値
'      　　:戻り値      ,O  ,Double　,数値
'説明      :
Private Function MakeParamCoefficient(ByVal param As Double) As Double
Dim strParam As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function MakeParamCoefficient"

    strParam = Format(param, "0.00000")
    MakeParamCoefficient = Val(strParam)


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :数値の丸めを行う（位置、長さ）
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型      ,説明
'      　　:param 　　　,I  ,Double　,数値
'      　　:戻り値      ,O  ,Double　,数値
'説明      :
Private Function MakeParamLength(ByVal param As Double) As Double
Dim strParam As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function MakeParamLength"

    strParam = Format(param, "0")
    MakeParamLength = Val(strParam)


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :数値の丸めを行う（重量）
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型      ,説明
'      　　:param 　　　,I  ,Double　,数値
'      　　:戻り値      ,O  ,Double　,数値
'説明      :
Private Function MakeParamWeight(ByVal param As Double) As Double
Dim strParam As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function MakeParamWeight"

    strParam = Format(param, "0")
    MakeParamWeight = Val(strParam)


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :抵抗値文字列を作成する
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型      ,説明
'      　　:param 　　　,I  ,Double　,数値
'      　　:戻り値      ,O  ,String　,作成文字列
'説明      :
Private Function MakeParamRs(ByVal param As Double, Optional formatstr As String = "0.0000") As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function MakeParamRs"

    If param < 0 Then
        MakeParamRs = ""
        GoTo proc_exit
    End If

    If formatstr = "" Then
        MakeParamRs = toRsStr(param)
    Else
        MakeParamRs = Format(param, formatstr)
    End If

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :本切断後重量を計算する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Function GetCutAfterWt(ByVal dCutAfterLen As Double) As Double
Dim dDM As Double

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function GetCutAfterWt"

    With g_PlupEndRslt
        dDM = (.DM1 + .DM2 + .DM3) / 3#
        GetCutAfterWt = WeightOfCylinder(dDM, dCutAfterLen)
    End With

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :本切断後長さ入力値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtCutAfterLen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtCutAfterLen_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtCutAfterLen(Index).Locked <> True) Then
        txtCutAfterLen(Index).Text = MakeParamLength(Val(txtCutAfterLen(Index).Text))
        '' 本切断後重量の計算
        txtCutAfterWt.Text = MakeParamWeight(GetCutAfterWt(Val(txtCutAfterLen(Index).Text)))
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :トップ側抵抗測定値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'      　　:Index 　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtCutAfterLen_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtCutAfterLen_LostFocus"

    Call txtCutAfterLen_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :トップ側抵抗測定値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtRsTop2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRsTop2_KeyDown"

    If (KeyCode = vbKeyReturn) And (txtRsTop2(Index).Locked <> True) Then
        ResTemp(1, Index) = CDbl(Val(txtRsTop2(Index).Text))
        txtRsTop2(Index).Text = MakeParamRs(Val(txtRsTop2(Index).Text), "")
        
        ResValSet Index
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :ボトム側抵抗測定値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtRsBot2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRsBot2_KeyDown"

    If (KeyCode = vbKeyReturn) And (txtRsBot2(Index).Locked <> True) Then
        ResTemp(3, Index) = CDbl(Val(txtRsBot2(Index).Text))
        txtRsBot2(Index).Text = MakeParamRs(Val(txtRsBot2(Index).Text), "")
        
        ResValSet Index
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :トップ側抵抗測定値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtRsTop2_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRsTop2_LostFocus"

    If txtRsTop2(Index).Text <> "" Then
        Call txtRsTop2_KeyDown(Index, vbKeyReturn, 0)
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :ボトム側抵抗測定値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtRsBot2_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRsBot2_LostFocus"

    If txtRsBot2(Index).Text <> "" Then
        Call txtRsBot2_KeyDown(Index, vbKeyReturn, 0)
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :規格抵抗MIN値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtSpRsMin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtSpRsMin_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtSpRsMin(Index).Locked <> True) Then
        ResTemp1(Index, 0) = Val(txtSpRsMin(Index).Text)
        txtSpRsMin(Index).Text = MakeParamRs(Val(txtSpRsMin(Index).Text), "")
        ShowRpRsParam
        ResValSet1
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :規格抵抗MIN値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtSpRsMin_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtSpRsMin_LostFocus"

    Call txtSpRsMin_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :規格抵抗MAX値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtSpRsMax_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtSpRsMax_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtSpRsMax(Index).Locked <> True) Then
        ResTemp1(Index, 1) = Val(txtSpRsMax(Index).Text)
        txtSpRsMax(Index).Text = MakeParamRs(Val(txtSpRsMax(Index).Text), "")
        ShowRpRsParam
        ResValSet1
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :規格抵抗MAX値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtSpRsMax_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtSpRsMax_LostFocus"

    Call txtSpRsMax_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :ねらい抵抗値キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtRpRs_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRpRs_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtRpRs(Index).Locked <> True) Then
        ResTemp1(Index, 2) = Val(txtRpRs(Index).Text)
        txtRpRs(Index).Text = MakeParamRs(Val(txtRpRs(Index).Text), "")
        ShowSpRsMaxParam
        ResValSet1
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :ねらい抵抗値フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtRpRs_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtRpRs_LostFocus"

    Call txtRpRs_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :チャージ量キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtChargeWt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtChargeWt_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtChargeWt(Index).Locked <> True) Then
        txtChargeWt(Index).Text = Val(Format(txtChargeWt(Index).Text, "0"))
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :チャージ量フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'      　　:Index 　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtChargeWt_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtChargeWt_LostFocus"

    Call txtChargeWt_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :トップカット重量キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtTopcutWt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtTopcutWt_KeyDown"

    If (KeyCode = vbKeyReturn) And (Index = 1) And (txtTopcutWt(Index).Locked <> True) Then
        txtTopcutWt(Index).Text = Val(Format(txtTopcutWt(Index).Text, "0"))
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :トップカット重量フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'      　　:Index 　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtTopcutWt_LostFocus(Index As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtTopcutWt_LostFocus"

    Call txtTopcutWt_KeyDown(Index, vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :下限規格内側管理キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtLowStInMng_KeyDown(KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtLowStInMng_KeyDown"

    If (KeyCode = vbKeyReturn) And (txtLowStInMng.Locked <> True) Then
        txtLowStInMng.Text = Val(Format(txtLowStInMng.Text, "0"))
        ShowRpRsParam
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :下限規格内側管理フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'      　　:Index 　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtLowStInMng_LostFocus()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtLowStInMng_LostFocus"

    Call txtLowStInMng_KeyDown(vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :上限値規格内側管理キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'      　　:KeyCode　　　,I  ,Integer　,キーコード
'      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtHighStInMng_KeyDown(KeyCode As Integer, Shift As Integer)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtHighStInMng_KeyDown"

    If (KeyCode = vbKeyReturn) And (txtHighStInMng.Locked <> True) Then
        txtHighStInMng.Text = Val(Format(txtHighStInMng.Text, "0"))
        ShowRpRsParam
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :上限値規格内側管理フォーカス外れ処理
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'      　　:Index 　　　,I  ,Integer　,コントロールインデックス番号
'説明      :
Private Sub txtHighStInMng_LostFocus()

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtHighStInMng_LostFocus"

    Call txtHighStInMng_KeyDown(vbKeyReturn, 0)


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :推定抵抗品番キー押下処理
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'      　　:Index  　　　,I  ,Integer　,コントロールインデックス番号
'　　      :KeyCode　　　,I  ,Integer　,キーコード
'　　      :Shift  　　　,I  ,Integer　,Shiftキーの状態
'説明      :
Private Sub txtSuiRsHinban_KeyDown(KeyCode As Integer, Shift As Integer)
Dim iRet        As Integer
Dim strHinban   As String
Dim tHinInf     As tFullHinban
Dim tblSXL1     As typ_TBCME018

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Sub txtSuiRsHinban_KeyDown"

    If KeyCode = vbKeyReturn Then
        lblMsg.Caption = vbNullString
        
        '' 表示項目になっている場合、処理終了
        If (txtSuiRsHinban.Locked = True) Or (Trim(txtSuiRsHinban.Text) = "") Then GoTo proc_exit
        '' 推定抵抗品番を取得する
        strHinban = txtSuiRsHinban.Text
        If Len(strHinban) <> 12 Then
            lblMsg.Caption = GetMsgStr(MSG_HINBAN_ERROR)
            GoTo proc_exit
        End If
        '' 製品仕様SXLデータの取得
        tHinInf.HINBAN = Mid(strHinban, 1, 8)
        tHinInf.mnorevno = Val(Mid(strHinban, 9, 2))
        tHinInf.factory = Mid(strHinban, 11, 1)
        tHinInf.opecond = Mid(strHinban, 12, 1)
        iRet = GetSPSXLData1(tblSXL1, tHinInf)
        If iRet <> FUNCTION_RETURN_SUCCESS Then
            lblMsg.Caption = GetMsgStr(MSG_NOTFOUND_HINBAN_ERR)
            GoTo proc_exit
        End If
        '' 取得製品仕様SXLデータ値のセット
        iRet = SetSuiRsHinbanSXLData(tblSXL1)
        If iRet <> FUNCTION_RETURN_SUCCESS Then
            GoTo proc_exit
        End If
    End If


proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Sub

'概要      :結晶情報、および、製品仕様SXLデータ、抵抗実績情報等を表示する（詳細）
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:tblSXL1　　　,I  ,typ_TBCME018   　,製品仕様ＳＸＬデータ1テーブル
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,処理成功：FUNCTION_RETURN_SUCCESS　処理失敗：FUNCTION_RETURN_FAILURE
'説明      :
Private Function SetSuiRsHinbanSXLData(tblSXL1 As typ_TBCME018) As FUNCTION_RETURN

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmmc001a.frm -- Function SetSuiRsHinbanSXLData"

    SetSuiRsHinbanSXLData = FUNCTION_RETURN_FAILURE

'    txtType.Text = tblSXL1.HSXTYPE
'    txtHoui.Text = tblSXL1.HSXCDIR
'    txtRRG.Text = MakeParamRs(tblSXL1.HSXRMBNP)

    'txtSpRsMin(0).Text = MakeParamRs(tblSXL1.HSXRMIN)
    txtSpRsMin(1).Text = MakeParamRs(tblSXL1.HSXRMIN, "")
    'txtSpRsMax(0).Text = MakeParamRs(tblSXL1.HSXRMAX)
    txtSpRsMax(1).Text = MakeParamRs(tblSXL1.HSXRMAX, "")

    ShowRpRsParam

    SetSuiRsHinbanSXLData = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function

''概要      :推定抵抗品番キーアップ処理
''ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
''      　　:KeyCode　　　,I  ,Integer　,キーコード
''      　　:Shift  　　　,I  ,Integer　,Shiftキーの状態
''説明      :
'Private Sub txtSuiRsHinban_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim iRet      As Integer
'Dim strHinban As String
'Dim tHinInf   As tFullHinban
'Dim tblSXL1   As typ_TBCME018
'
'    'エラーハンドラの設定
'    On Error GoTo proc_err
'    gErr.Push "f_cmmc001a.frm -- Sub txtSuiRsHinban_KeyUp"
'
'    If Len(txtSuiRsHinban.Text) = 0 Then
'        '' 表示項目になっている場合、処理終了
'        If txtSuiRsHinban.Locked = True Then GoTo proc_exit
'        '' ねらい品番を取得する
'        strHinban = txtRpHinban.Text
'        '' 製品仕様SXLデータの取得
'        tHinInf.HINBAN = Mid(strHinban, 1, 8)
'        tHinInf.mnorevno = Val(Mid(strHinban, 9, 2))
'        tHinInf.factory = Mid(strHinban, 11, 1)
'        tHinInf.opecond = Mid(strHinban, 12, 1)
'        iRet = GetSPSXLData1(tblSXL1, tHinInf)
'        If iRet <> FUNCTION_RETURN_SUCCESS Then
'            lblMsg.Caption = GetMsgStr(MSG_GETERROR_DBDATA)
'            GoTo proc_exit
'        End If
'        '' 取得製品仕様SXLデータ値のセット
'        iRet = SetSuiRsHinbanSXLData(tblSXL1)
'        If iRet <> FUNCTION_RETURN_SUCCESS Then
'            GoTo proc_exit
'        End If
'    End If
'
'
'proc_exit:
'    '終了
'    gErr.Pop
'    Exit Sub
'
'proc_err:
'    'エラーハンドラ
'    gErr.HandleError
'    Resume proc_exit
'End Sub

Public Sub ResValSet(dpos As Integer)
    Dim MaxLowerCol As Integer '2002/01/15 S.Sano
    Dim c0 As Integer
    Dim lowCol As Integer

    MaxLowerCol = 0
    lowCol = GetLowerCol(ResTemp(0, dpos), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp(1, dpos), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp(2, dpos), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp(3, dpos), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol

    txtRsTop1(dpos).Text = toRsStrByPlace(ResTemp(0, dpos), MaxLowerCol)
    txtRsTop2(dpos).Text = toRsStrByPlace(ResTemp(1, dpos), MaxLowerCol)
    txtRsBot1(dpos).Text = toRsStrByPlace(ResTemp(2, dpos), MaxLowerCol)
    txtRsBot2(dpos).Text = toRsStrByPlace(ResTemp(3, dpos), MaxLowerCol)
End Sub

Public Sub ResValSet1()
    Dim MaxLowerCol As Integer '2002/01/15 S.Sano
    Dim c0 As Integer
    Dim lowCol As Integer

    MaxLowerCol = 0
    lowCol = GetLowerCol(ResTemp1(0, 0), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp1(1, 0), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    
    txtSpRsMin(0).Text = toRsStrByPlace(ResTemp1(0, 0), MaxLowerCol)
    txtSpRsMin(1).Text = toRsStrByPlace(ResTemp1(1, 0), MaxLowerCol)

    MaxLowerCol = 0
    lowCol = GetLowerCol(ResTemp1(0, 1), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp1(1, 1), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    
    txtSpRsMax(0).Text = toRsStrByPlace(ResTemp1(0, 1), MaxLowerCol)
    txtSpRsMax(1).Text = toRsStrByPlace(ResTemp1(1, 1), MaxLowerCol)

    MaxLowerCol = 0
    lowCol = GetLowerCol(ResTemp1(0, 2), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    lowCol = GetLowerCol(ResTemp1(1, 2), True)
    If lowCol > MaxLowerCol Then MaxLowerCol = lowCol
    
    txtRpRs(0).Text = toRsStrByPlace(ResTemp1(0, 2), MaxLowerCol)
    txtRpRs(1).Text = toRsStrByPlace(ResTemp1(1, 2), MaxLowerCol)
End Sub

