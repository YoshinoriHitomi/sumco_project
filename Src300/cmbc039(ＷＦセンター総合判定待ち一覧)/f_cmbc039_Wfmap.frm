VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_4 
   BorderStyle     =   1  '固定(実線)
   Caption         =   " f_cmbc039_4"
   ClientHeight    =   8205
   ClientLeft      =   1575
   ClientTop       =   1680
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
   Begin VB.ComboBox cmbSprChg 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9.75
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "f_cmbc039_Wfmap.frx":0000
      Left            =   1665
      List            =   "f_cmbc039_Wfmap.frx":0010
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   11
      Top             =   1035
      Width           =   1425
   End
   Begin VB.TextBox txtSxlId 
      BackColor       =   &H0080FF80&
      Height          =   264
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1050
      Width           =   1335
   End
   Begin VB.CommandButton cmdF 
      Caption         =   "閉じる"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   7620
      Width           =   1050
   End
   Begin FPSpread.vaSpread sprExamine 
      Height          =   6000
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   11535
      _Version        =   196608
      _ExtentX        =   20346
      _ExtentY        =   10583
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      BackColorStyle  =   1
      ColsFrozen      =   6
      MaxCols         =   30
      MaxRows         =   1
      ProcessTab      =   -1  'True
      RetainSelBlock  =   0   'False
      ShadowColor     =   12632256
      ShadowDark      =   8421504
      ShadowText      =   0
      SpreadDesigner  =   "f_cmbc039_Wfmap.frx":0030
      UserResize      =   0
      VisibleCols     =   12
   End
   Begin VB.Frame fraHead 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   -30
      Width           =   11925
      Begin VB.Label lblTime 
         Height          =   150
         Left            =   10300
         TabIndex        =   3
         Top             =   300
         Width           =   1450
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
         Left            =   3810
         TabIndex        =   2
         Top             =   240
         Width           =   7890
      End
      Begin VB.Label lblTitle 
         Caption         =   "WFマップ状態表示"
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
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   4575
      End
   End
   Begin VB.Label Label5 
      Caption         =   "SXLID"
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
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label6 
      Caption         =   "抽出条件"
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
      Left            =   1710
      TabIndex        =   9
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label3 
      Caption         =   "抜試WF"
      Height          =   255
      Left            =   7650
      TabIndex        =   7
      Top             =   7740
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "欠落WF"
      Height          =   255
      Left            =   5340
      TabIndex        =   6
      Top             =   7740
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "通常WF"
      Height          =   255
      Left            =   3180
      TabIndex        =   5
      Top             =   7740
      Width           =   705
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  '不透明
      BorderStyle     =   3  '点線
      Height          =   285
      Left            =   6450
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  '不透明
      BorderStyle     =   3  '点線
      Height          =   285
      Left            =   4140
      Top             =   7680
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000FF00&
      BackStyle       =   1  '不透明
      BorderStyle     =   3  '点線
      Height          =   285
      Left            =   1980
      Top             =   7680
      Width           =   1155
   End
End
Attribute VB_Name = "f_cmbc039_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'概要      :抽出条件コンボボックス切替処理
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　,IO ,型       ,説明
'　　      :なし
'説明      :抽出条件により、WFﾏｯﾌﾟ一覧を種別毎の表示の切り替えを行う
'履歴      :
Private Sub cmbSprChg_Click()

    Dim iLoopCnt    As Integer
    Dim iSprSta     As Integer
    Dim sSprSta     As String
    Dim vSprSta     As Variant

    With sprExamine
        .ReDraw = False
        For iLoopCnt = 1 To .MaxRows
            Select Case cmbSprChg.ListIndex
                Case mSprChg_0  '全件指定
                    .row = iLoopCnt
                    .RowHidden = False
                Case mSprChg_1  '良品指定
                    .GetText 30, iLoopCnt, vSprSta
                    If vSprSta <> mSprChg_1 Then  '良品以外だったら、非表示
                        .row = iLoopCnt
                        .RowHidden = True
                    Else
                        .row = iLoopCnt
                        .RowHidden = False
                    End If
                Case mSprChg_2  'サンプル指定
                    .GetText 30, iLoopCnt, vSprSta
                    If vSprSta <> mSprChg_2 Then  'サンプル以外だったら、非表示
                        .row = iLoopCnt
                        .RowHidden = True
                    Else
                        .row = iLoopCnt
                        .RowHidden = False
                    End If
                Case mSprChg_3  '不良指定
                    .GetText 30, iLoopCnt, vSprSta
                    If vSprSta <> mSprChg_3 Then  '不良以外だったら、非表示
                        .row = iLoopCnt
                        .RowHidden = True
                    Else
                        .row = iLoopCnt
                        .RowHidden = False
                    End If
            End Select
        Next
        .ReDraw = True
    End With

End Sub

Private Sub cmdF_Click(Index As Integer)

    Me.Visible = False
    Unload Me

End Sub

Private Sub Form_Activate()

    cmbSprChg.ListIndex = 0
    With sprExamine
        .col = 30
        .ColHidden = True
    End With

End Sub


Private Sub Form_Load()

    ' 現在日時の表示
    SetPresentTime lblTime
    

End Sub
