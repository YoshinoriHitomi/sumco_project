VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_1 
   Appearance      =   0  'ﾌﾗｯﾄ
   BorderStyle     =   1  '固定(実線)
   Caption         =   "f_cmbc039_1(CW750) - 300mm結晶操業システム"
   ClientHeight    =   10875
   ClientLeft      =   825
   ClientTop       =   1155
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   1018
   StartUpPosition =   2  '画面の中央
   Begin VB.OptionButton optHold 
      BackColor       =   &H8000000A&
      Caption         =   "全表示"
      Height          =   375
      Index           =   2
      Left            =   9420
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   33
      Top             =   840
      Width           =   855
   End
   Begin VB.OptionButton optHold 
      BackColor       =   &H8000000A&
      Caption         =   "流動停止"
      Height          =   375
      Index           =   1
      Left            =   8460
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   32
      Top             =   840
      Width           =   975
   End
   Begin VB.OptionButton optHold 
      BackColor       =   &H8000000A&
      Caption         =   "流動可能"
      Height          =   375
      Index           =   0
      Left            =   7500
      Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
      TabIndex        =   31
      Top             =   840
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CheckBox chkY4Disp 
      Caption         =   "流動停止項目表示"
      Height          =   180
      Left            =   13200
      TabIndex        =   30
      Top             =   960
      Value           =   1  'ﾁｪｯｸ
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CheckBox chk_Alldata 
      Caption         =   "仕掛全表示"
      Enabled         =   0   'False
      Height          =   180
      Left            =   11940
      TabIndex        =   29
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox cmbMukesaki 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   336
      ItemData        =   "f_cmbc039_1.frx":0000
      Left            =   4464
      List            =   "f_cmbc039_1.frx":0002
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   28
      Top             =   864
      Width           =   1356
   End
   Begin VB.TextBox txtSxlId 
      Height          =   285
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1770
      MaxLength       =   13
      TabIndex        =   17
      Top             =   888
      Width           =   1515
   End
   Begin FPSpread.vaSpread spdWait 
      Height          =   7605
      Left            =   180
      TabIndex        =   15
      Top             =   1290
      Width           =   14895
      _Version        =   196608
      _ExtentX        =   26273
      _ExtentY        =   13414
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      MaxCols         =   15
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "f_cmbc039_1.frx":0004
      UserResize      =   0
      VisibleCols     =   11
      VisibleRows     =   1
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
      TabIndex        =   14
      Top             =   9540
      Width           =   15195
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
         Left            =   12680
         TabIndex        =   11
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
         Left            =   11448
         TabIndex        =   10
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
         Caption         =   "[F６]　  ０枚ﾛｯﾄ"
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
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F５]　　＊＊＊"
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
         Left            =   4056
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F９]　　抽出"
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
         Left            =   10230
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
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
         Left            =   2824
         TabIndex        =   3
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
         Left            =   13920
         TabIndex        =   12
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
      Width           =   15225
      Begin VB.Label lblvers 
         BackColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   13740
         TabIndex        =   20
         Top             =   480
         Width           =   1365
      End
      Begin VB.Label lblTime 
         BackColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   13740
         TabIndex        =   19
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblMsg 
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
         Left            =   4665
         TabIndex        =   18
         Top             =   240
         Width           =   8670
      End
      Begin VB.Label lblTitle 
         Caption         =   "WFセンター総合判定待ち一覧"
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
         TabIndex        =   13
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "向先"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   13.5
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   276
      Left            =   3792
      TabIndex        =   27
      Top             =   888
      Width           =   576
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "待ち"
      Height          =   180
      Left            =   1920
      TabIndex        =   26
      Top             =   9180
      Width           =   330
   End
   Begin VB.Label lblMachi 
      BorderStyle     =   1  '実線
      Height          =   135
      Left            =   1080
      TabIndex        =   25
      Top             =   9180
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "流動停止品"
      Height          =   180
      Left            =   6240
      TabIndex        =   24
      Top             =   9180
      Width           =   900
   End
   Begin VB.Label lblHoldColor 
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   1  '実線
      Height          =   135
      Left            =   5400
      TabIndex        =   23
      Top             =   9180
      Width           =   735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "正常"
      Height          =   180
      Left            =   4080
      TabIndex        =   22
      Top             =   9180
      Width           =   360
   End
   Begin VB.Label lblGoodColor 
      BackColor       =   &H0080FF80&
      BorderStyle     =   1  '実線
      Height          =   135
      Left            =   3240
      TabIndex        =   21
      Top             =   9180
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "仮ＳＸＬ－ＩＤ"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Left            =   600
      TabIndex        =   16
      Top             =   888
      Width           =   1128
   End
End
Attribute VB_Name = "f_cmbc039_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MlngSpread_Col As Long
Dim MlngSpread_Row As Long
Dim MblClickFlg As Boolean
Private objIE()     As Object   'add 09/03/16 SETkimizuka

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
    '' 処理分岐
    Select Case intIndex
        Case 1          '' Ｆ１キー（メインメニュー）
            '' プログラム終了処理
             GotoMainMenu
        Case 2          '' Ｆ２キー（サブメニュー）
            '' サブメニューに戻る
            GotoSubMenu
        Case 3          '' Ｆ３キー（キャンセル）
            '' 表示画面クリア
            sub_InitDisp_f_cmbc039_1

            'フォーカスセット（担当者）
            CtrlEnabled txtSxlId, CTRL_ENABLE, True       '担当者コード
            f_cmbc039_1.txtSxlId.SetFocus

            '待ち一覧表示
            Call sub_PutSpdMatiData(typ_ww())

            'コマンドボタンの状態変更
            sub_FunctionKeySet
            
        '>>>>> ０枚ロット一覧表示追加 SETsw Marushita 2012/09/07 Start
        Case 6   '' Ｆ6キー
            Dim lLp1 As Long
            
            '初期日付範囲セット
            If cmbMukesaki.text <> "" Then
                For lLp1 = 0 To UBound(s_MukesakiBase)
                    If Trim(cmbMukesaki.text) = s_MukesakiBase(lLp1).sMukeName Then
                        gsMukeCd = s_MukesakiBase(lLp1).sMukeCode
                        Exit For
                    End If
                Next lLp1
            Else
                lblMsg.Caption = "向先を選択してください"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
            '0枚ロット一覧情報取得
            f_cmbc039_1.lblMsg.Caption = GetMsgStr(PWAIT)
            DoEvents
            f_cmbc039_6.Show vbModal
        '<<<<< ０枚ロット一覧表示追加 SETsw Marushita 2012/09/07 End
        Case 9   '' Ｆ9キー
            BeginProcess '' プロセス開始
            '' 表示画面クリア
            sub_InitDisp_f_cmbc039_1

            'フォーカスセット（担当者）
            CtrlEnabled txtSxlId, CTRL_ENABLE, True       '担当者コード
            f_cmbc039_1.txtSxlId.SetFocus

            f_cmbc039_1.Enabled = False
            '' 表示画面クリア
            sub_InitDisp_f_cmbc039_1

            'コマンドボタンの状態変更
            sub_FunctionKeySet

            '待ち一覧情報取得
            f_cmbc039_1.lblMsg.Caption = GetMsgStr(PWAIT)
            DoEvents

            '画面情報を情報構造体に設定する
            If WfWaitSetAllData(typ_ww()) <> True Then
                f_cmbc039_1.lblMsg.Caption = ""
                f_cmbc039_1.Enabled = True
                EndProcess '' プロセス終了
                Exit Sub
            End If

            '待ち一覧表示
            sub_PutSpdMatiData typ_ww()

            f_cmbc039_1.Enabled = True

            EndProcess '' プロセス終了
        Case 12   '' Ｆ12キー
            If iMode = 0 Then

                Dim lLp As Long

                Screen.MousePointer = vbHourglass
                DoEvents

                If cmbMukesaki.text <> "" Then
                    For lLp = 0 To UBound(s_MukesakiBase)
                        If Trim(cmbMukesaki.text) = s_MukesakiBase(lLp).sMukeName Then
                            sCmbMukesaki = s_MukesakiBase(lLp).sMukeCode
                        End If
                    Next lLp
                Else
                    lblMsg.Caption = "向先を選択してください"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If

                '待ち一覧
                sub_RunWfWait

                Screen.MousePointer = vbDefault
            Else
                If VerChk(f_cmbc039_1) = False Then Exit Sub

                BeginProcess '' プロセス開始
                '' 実行処理を行う
                If fnc_ExecutionProcess = FUNCTION_RETURN_FAILURE Then
                    EndProcess '' プロセス終了
                    Exit Sub
                End If
                EndProcess '' プロセス終了

                '画面情報をＷＦセンター総合判定に渡す
                f_cmbc039_2.Sub_SetParamData

                Me.Visible = False
                f_cmbc039_2.Show
            End If
    End Select
End Sub

'*******************************************************************************
'*    関数名        : Form_Activate
'*
'*    処理概要      : 1.選択されたSXLIDを削除し待ち一覧を表示する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Activate()
    '' 画面表示メッセージクリア
    f_cmbc039_1.lblMsg.Caption = ""

    If intModoru = 2 Then  '実行時
        f_cmbc039_1.txtSxlId = ""
        f_cmbc039_1.txtSxlId.SetFocus

        '選択されたSXLIDを削除し待ち一覧を表示する
        With spdWait
            .row = .ActiveRow
            .Action = ActionDeleteRow
            .MaxRows = .MaxRows - 1
        End With
    End If
End Sub

'*******************************************************************************
'*    関数名        : Form_KeyDown
'*
'*    処理概要      : 1.キーが押されたら、各処理に分岐する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*　　               KeyCode      ,I  ,Integer　,キーコード
'*         　　      Shift        ,I  ,Integer　,Shiftキーの状態
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
End Sub

'*******************************************************************************
'*    関数名        : Form_Load
'*
'*    処理概要      : 1.Form_Load処理
'*                    2.Initial処理
'*                    3.向先の取得
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*　　               KeyCode      ,I  ,Integer　,キーコード
'*         　　      Shift        ,I  ,Integer　,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Load()
    'フォームを表示
    f_cmbc039_1.Show
    
    ' バージョン情報の表示
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision

    'フォーカスセット（担当者）
    f_cmbc039_1.txtSxlId.SetFocus

    'プログラム起動時の初期化処理
    InitExe
    If VerChk(f_cmbc039_1) = False Then
       lblvers.backColor = COLOR_GRAY
       lblTime.backColor = COLOR_GRAY
        Exit Sub
    End If
    ' 現在日時の表示
    '' 処理時間セット
    SetPresentTime lblTime

    '' フォーム位置セット
    CenterForm Me

    iMode = 0   '(一覧未表示)
    
    '向先コード取得
    If GetMukeCode = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = "向先取得エラー(KODA9)"
    End If

    Dim iLp  As Integer
    Dim sMukesaki As String

    If sCmbMukesaki <> "" Then
        For iLp = 0 To 2
            sMukesaki = cmbMukesaki.List(iLp)

            If sCmbMukeName = sMukesaki Then
                cmbMukesaki.Enabled = False
                cmbMukesaki.text = cmbMukesaki.List(iLp)
                Exit For
            End If
        Next iLp
    End If
    
    ReDim objIE(0)  'add SETkimizuka
    
    '初期状態はキャンセル・表示更新ボタンを押下不可に 2010/06/16 SETsw kubota
    cmdF(3).Enabled = False
    cmdF(9).Enabled = False
    
End Sub

'*******************************************************************************
'*    関数名        : txtSxlId_KeyDown
'*
'*    処理概要      : 1.Enterキーが押されたら、SXL-IDをチェックする
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*　　               KeyCode      ,I  ,Integer　,キーコード
'*         　　      Shift        ,I  ,Integer　,Shiftキーの状態
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtSxlId_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intC0       As Integer
    Dim c1          As Integer
    Dim blChecFlag  As Boolean
    Dim vTmpSXLID   As Variant

    ''画面表示メッセージクリア
    lblMsg.Caption = ""

    If KeyCode = vbKeyReturn Then
        BeginProcess '' プロセス開始
        ''ブロックIDの長さを確認
        If ChkTextBox(txtSxlId, CHK_STRING, 13, 13) = FUNCTION_RETURN_SUCCESS Then
            blChecFlag = False
            For intC0 = 1 To MaxLine
                spdWait.GetText 3, intC0, vTmpSXLID
                If vTmpSXLID = Trim(txtSxlId.text) Then
                    spdWait_ButtonClicked 1, intC0, 0
                    sub_Spread_Change spdWait, intC0, MlngSpread_Row
                    blChecFlag = True
                    Exit For
                End If
            Next
            If blChecFlag = False Then
                lblMsg.Caption = GetMsgStr(ESXL0)
            End If
        Else
            ''ブロックID異常
            lblMsg.Caption = GetMsgStr(ESXL1)
        End If
        EndProcess '' プロセス終了
    End If
End Sub

'*******************************************************************************
'*    関数名        : txtSxlID_Change
'*
'*    処理概要      : 1.チェックボックスのチェックを外す
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*         　　       なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub txtSxlID_Change()
    If MblClickFlg = False And MlngSpread_Row <> 0 Then
        spdWait.SetInteger 1, MlngSpread_Row, 0
    End If
End Sub

'*******************************************************************************
'*    関数名        : spdWait_ButtonClicked
'*
'*    処理概要      : 1.選択されたSxlIDを表示
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*         　　       なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub spdWait_ButtonClicked(ByVal col As Long, ByVal row As Long, ByVal ButtonDown As Integer)
    Dim FuncAns     As FUNCTION_RETURN
    Dim iCnt        As Integer
    Dim sSplit()    As String
    Dim sVal        As String
    Dim vNo         As Variant

    lblMsg.Caption = ""
    f_cmzc003a.Hide
 
    If ButtonDown = 1 And col = 1 Then  'upd SETkimizuka 09/03/16
'    If ButtonDown = 1 Then
        If MlngSpread_Col <> 0 And MlngSpread_Row <> row Then
            spdWait.row = MlngSpread_Row
            spdWait.col = MlngSpread_Col
            spdWait.Value = False
        End If

        MlngSpread_Col = 1
        MlngSpread_Row = row
        MblClickFlg = True

        spdWait.row = row
        spdWait.col = 2
        sCmbMukeName = spdWait.text

        spdWait.col = 2 + 1
        txtSxlId.text = spdWait.text
        MblClickFlg = False
    'add 停止項目追加 SETkimizuka 09/03/16 Start
    ElseIf col = 15 And row > 0 Then
        spdWait.GetText 3, row, vNo
        For iCnt = 1 To UBound(typ_ww)
            If typ_ww(iCnt).SXLIDCA = CStr(vNo) Then
                sSplit = Split(typ_ww(iCnt).PRINTNO, Chr(9))
                Exit For
            End If
        Next
        Call sub_CloseIE
        
        If UBound(sSplit) > 0 Then
            ReDim objIE(UBound(sSplit))
            For iCnt = 0 To UBound(sSplit) - 1
                sVal = GetSWSUrl(Mid(sSplit(iCnt), 1, 10), Mid(sSplit(iCnt), 11))
                Call SetIEOption(objIE(iCnt), Me)
                objIE(iCnt).Navigate sVal
                objIE(iCnt).Visible = True
            Next
        End If
    'add 停止項目追加 SETkimizuka 09/03/16 End
    Else
'        txtSxlId.Text = ""
    End If

    'コマンドボタンの状態変更
    sub_FunctionKeySet
End Sub
'*******************************************************************************
'*    関数名        : CloseIE
'*
'*    処理概要      : 先行評価依頼票(IE)を閉じる
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*         　　       なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_CloseIE()
    On Error Resume Next
    Dim iCnt As Integer
    If UBound(objIE) > 0 Then
        For iCnt = 0 To UBound(objIE) - 1
            objIE(iCnt).Quit
        Next
    End If
End Sub
'**************************************************************************************
'*    関数名        : sub_Spread_Change
'*
'*    処理概要      : 1.チェックボックスが付いていたところのチェックを外す
'*                    2.新たにテキストに記載されたSxlIDを探し、その行にチェックをつける
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*         　　       spread　　　,I  ,vaSpread ,一覧のSpread
'*         　　       Row　   　　,I  ,Long     ,チェックをつける先の行
'*         　　       Rowb  　　　,I  ,Long     ,チェックが付いていたところの行
'*
'*    戻り値        : なし
'*
'**************************************************************************************
Public Sub sub_Spread_Change(spread As vaSpread, ByVal row As Long, ByVal Rowb As Long)

    With spread
        If Rowb <> 0 Then
            .row = Rowb
            .col = 1
            .Value = False
        End If
        .row = row
        .col = 1
        If .Lock = False Then
            .Value = True
        End If
    End With
End Sub

'**************************************************************************************
'*    関数名        : sub_FunctionKeySet
'*
'*    処理概要      : 1.WFの状態によりコマンドボタンのEnableをTrue／Falseを設定する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'**************************************************************************************
Public Sub sub_FunctionKeySet()
    Dim lngCol          As Long
    Dim lngC0           As Long
    Dim lngTemp         As Long
    Dim blCheckFlag(0)  As Boolean
    Dim blAns           As Boolean
    Dim sFlagString     As String

    lngCol = 1

    For lngC0 = 1 To MaxLine
        blAns = spdWait.GetInteger(lngCol, lngC0, lngTemp)
        blCheckFlag(0) = ((lngTemp = 1) Or (lngTemp = -1))
        If blCheckFlag(0) Then Exit For
    Next

    If blCheckFlag(0) Then
        '>>>>> 2012/09/07 SETsw Marushita ０枚ロット一覧表示対応
        'sFlagString = "111000001001"
        sFlagString = "111001001001"
        '<<<<< 2012/09/07 SETsw Marushita ０枚ロット一覧表示対応
    Else
        '>>>>> 2012/09/07 SETsw Marushita ０枚ロット一覧表示対応
        'sFlagString = "111000001000"
        sFlagString = "111001001000"
        '<<<<< 2012/09/07 SETsw Marushita ０枚ロット一覧表示対応
    End If

    For lngC0 = 1 To 12
        cmdF(lngC0).Enabled = (Mid(sFlagString, lngC0, 1) = "1")
    Next
End Sub

'**************************************************************************************
'*    関数名        : Fnc_ExecutionProcess
'*
'*    処理概要      : 1.入力画面においての入力された値を登録する
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : 正常終了時はFUNCTION_RETURN_SUCCESS(0),
'*                    エラー終了時は FUNCTION_RETURN_FAILURE(-1)
'*
'**************************************************************************************
Private Function fnc_ExecutionProcess() As FUNCTION_RETURN
    Dim udtAllTypes As typ_AllTypes       '全情報構造体
    Dim intC0       As Integer
    Dim blCheckFlag As Boolean

    '' パラメータ初期化
    fnc_ExecutionProcess = FUNCTION_RETURN_FAILURE

    '' パラメータ判定処理を行う
    If ChkTextBox(txtSxlId, CHK_STRING, 13, 13) = FUNCTION_RETURN_SUCCESS Then
        blCheckFlag = False
        For intC0 = 1 To MaxLine
            If typ_ww(intC0).SXLIDCA = Trim(txtSxlId.text) Then
                blCheckFlag = True
                lblMsg.Caption = GetMsgStr(ESXL0)
                Exit For
            End If
        Next
        If Not blCheckFlag Then
            Exit Function
        End If
    Else
        lblMsg.Caption = GetMsgStr(ESXL1)
        Exit Function
    End If

    typ_AType = udtAllTypes
    SelectSxlID039 = Trim(txtSxlId.text)

    For intC0 = 1 To MaxLine
        If RTrim(typ_ww(intC0).SXLIDCA) = RTrim(SelectSxlID039) Then
            typ_Param001b = typ_ww(intC0)
            sKanrenFlg = typ_ww(intC0).KANREN      '関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
        End If
    Next

    iCntJ015upd = 0             'TBCMJ015-UPDATEﾚｺｰﾄﾞ数の初期化
    ReDim typ_J015_WFGDUpd(0)   'TBCMJ015-UPDATE用GD実績の初期化

    '' 処理正常終了
    fnc_ExecutionProcess = FUNCTION_RETURN_SUCCESS
End Function

'**************************************************************************************
'*    関数名        : spdWait_click
'*
'*    処理概要      : 1.列タイトルを押下した場合、押下された列を元にソート
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    col         ,I  ,Long     ,Spreadの列
'*                    Row         ,I  ,Long     ,Spreadの行
'*
'*    戻り値        : なし
'*
'**************************************************************************************
Private Sub spdWait_click(ByVal col As Long, ByVal row As Long)
    On Error GoTo Err:
    'スプレッドシートの表示を更新しない
    spdWait.ReDraw = False

    Select Case row
        'P1 列タイトルを押下した場合、押下された列を元にソート
        Case 0
            Call sprSort(spdWait, col)
    End Select

    'スプレッドシートの表示を更新する
    spdWait.ReDraw = True

    Exit Sub
Err:
    MsgBox "sprsort err(clik)"
End Sub

'*******************************************************************************
'*    関数名        : sub_RunWfWait
'*
'*    処理概要      : 1.待ち一覧の情報取得
'*                    2.待ち一覧の情報表示
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_RunWfWait()
    f_cmbc039_1.Enabled = False

    '待ち一覧情報取得
    BeginProcess '' プロセス開始
    f_cmbc039_1.lblMsg.Caption = GetMsgStr(PWAIT)
    DoEvents

    '画面情報を情報構造体に設定する
    If WfWaitSetAllData(typ_ww()) <> True Then
        f_cmbc039_1.Enabled = True
        f_cmbc039_1.lblMsg.Caption = "データがありません。"
        EndProcess '' プロセス終了
        Exit Sub
    End If

    '待ち一覧表示
    Call sub_PutSpdMatiData(typ_ww())

    'ボタン制御変更 2010/06/16 SETsw kubota
    cmdF(3).Enabled = True
    cmdF(9).Enabled = True
    cmdF(12).Enabled = False

    f_cmbc039_1.cmbMukesaki.Enabled = False
    iMode = 1

    EndProcess '' プロセス終了

    f_cmbc039_1.Enabled = True
End Sub

'*******************************************************************************
'*    関数名        : sub_InitDisp_f_cmbc039_1
'*
'*    処理概要      : 1.WF待ち一覧画面初期化
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_InitDisp_f_cmbc039_1()
    'スプレッドコントロールの初期化処理
    SpCtrlInit f_cmbc039_1.spdWait, 0
    SpCtrlInit f_cmbc039_1.spdWait, 22
End Sub

'*******************************************************************************
'*    関数名        : sub_PutSpdMatiData
'*
'*    処理概要      : 1.画面待ち一覧情報データ表示(情報構造体を画面に表示する)
'*
'*    パラメータ    : 変数名        ,IO ,型                            ,説明
'*                    udt_ww         I   ,DBDRV_scmzc_fcmlc001b_SXL039 ,SXL管理
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub sub_PutSpdMatiData(udt_ww() As DBDRV_scmzc_fcmlc001b_SXL039)
    Dim i As Integer, j As Integer       ' ﾙｰﾌﾟ ｶｳﾝﾀ
    Dim intRow As Integer
    Dim smpId(2) As String
    Dim strRecord As String
    Dim bRc As Boolean

    MaxLine = UBound(udt_ww())
'    DoEvents
'    SpCtrlInit f_cmbc039_1.spdWait, UBound(udt_ww())
'    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 2, 1, 11, UBound(udt_ww()), CTRL_DISABLE

    With f_cmbc039_1
        .spdWait.ReDraw = False
        intRow = 0

        For i = 1 To UBound(udt_ww())
            If udt_ww(i).MAICB Then
'Add Start 2012/07/17 Y.Hitomi　'opHold(0):流動可能　(1):流動不可　(2):全表示
                If (optHold(0).Value = True And WFJudgExecOkFlag(i) And udt_ww(i).STOP <> "1") Or _
                (optHold(1).Value = True And udt_ww(i).STOP = "1") Or _
                optHold(2).Value = True Then
'Add End 2012/07/17 Y.Hitomi
'Del Start 2012/07/17 Y.Hitomi
''                '表示方法変更(仕掛全表示ﾁｪｯｸﾎﾞｯｸｽがﾁｪｯｸ無の場合は実行可能SXLのみ)　08/10/29 ooba
''                ' 既存ホールドを流動停止に置き換え upd 09/04/16 Start
''                If chk_Alldata.Value = 1 Or (chk_Alldata.Value <> 1 And _
''                        WFJudgExecOkFlag(i) And udt_ww(i).STOP <> "1") Then
'''                If chk_Alldata.Value = 1 Or (chk_Alldata.Value <> 1 And _
'''                        WFJudgExecOkFlag(i) And udt_ww(i).HOLDBCB <> "1" And udt_ww(i).WFHOLDFLGCB <> "1") Then
''                ' 既存ホールドを流動停止に置き換え upd 09/04/16 End
'Del End 2012/07/17 Y.Hitomi
                    intRow = intRow + 1
    
                    strRecord = ""
    
                    If UBound(udt_ww(i).WFSMP) >= 1 Then
                        smpId(1) = Trim(udt_ww(i).WFSMP(1).REPSMPLIDCW)
                    Else
                        smpId(1) = vbNullString
                    End If
                    If UBound(udt_ww(i).WFSMP) >= 2 Then
                        'Chg Start 2011/04/28 SMPK Miyata
                        'smpId(2) = Trim(udt_ww(i).WFSMP(2).REPSMPLIDCW)
                        smpId(2) = Trim(udt_ww(i).WFSMP(UBound(udt_ww(i).WFSMP)).REPSMPLIDCW)
                        'Chg End   2011/04/28 SMPK Miyata

                    Else
                        smpId(2) = vbNullString
                    End If
                    
                    .spdWait.MaxRows = intRow
                    .spdWait.row = intRow
    
                    '向先
    '                .spdWait.col = 2
    '                .spdWait.Value = udt_ww(i).PLANTCAT
                    strRecord = Chr$(9) & strRecord & udt_ww(i).PLANTCAT & Chr$(9)
    
                    '仮SXL-ID
    '                .spdWait.col = 2 + 1
    '                .spdWait.Value = udt_ww(i).SXLIDCA
                    strRecord = strRecord & udt_ww(i).SXLIDCA & Chr$(9)
    
                    '品番
    '                .spdWait.col = 3 + 1
    '                .spdWait.Value = udt_ww(i).HINBCA
                    strRecord = strRecord & udt_ww(i).HINBCA & Chr$(9)
    
                    '長さ
    '                .spdWait.col = 4 + 1
    '                .spdWait.Value = udt_ww(i).GNLCA
                    strRecord = strRecord & udt_ww(i).GNLCA & Chr$(9)
    
                    '枚数
    '                .spdWait.col = 5 + 1
    '                .spdWait.Value = udt_ww(i).MAICB
                    strRecord = strRecord & udt_ww(i).MAICB & Chr$(9)
    
                    'WF払出日
    '                .spdWait.col = 6 + 1
    '                .spdWait.Value = Format(udt_ww(i).TDAYCB, "mm/dd")
                    strRecord = strRecord & Format(udt_ww(i).TDAYCB, "mm/dd") & Chr$(9)
    
                    'サンプルID(上側)
    '                .spdWait.col = 7 + 1
    '                .spdWait.Value = smpId(1)
                    strRecord = strRecord & smpId(1) & Chr$(9)
    
                    'サンプルID(下側)
    '                .spdWait.col = 8 + 1
    '                .spdWait.Value = smpId(2)
                    strRecord = strRecord & smpId(2) & Chr$(9)
    
                    '最終受信日(上側)
    '                .spdWait.col = 9 + 1
                    If Not (smpId(1) = "" Or _
                            left(smpId(1), 1) = vbNullChar) Then
                        If UBound(udt_ww(i).WFSMP) >= 1 Then
    '                        .spdWait.Value = Format(udt_ww(i).WFSMP(1).KDAYCW, "mm/dd")
                            strRecord = strRecord & Format(udt_ww(i).WFSMP(1).KDAYCW, "mm/dd") & Chr$(9)
    
                        End If
                    End If
    
                    '最終受信日(下側)
    '                .spdWait.col = 10 + 1
                    If Not (smpId(2) = "" Or _
                            left(smpId(2), 1) = vbNullChar) Then
                        If UBound(udt_ww(i).WFSMP) >= 2 Then
    '                        .spdWait.Value = Format(udt_ww(i).WFSMP(2).KDAYCW, "mm/dd")
                            'Chg Start 2011/04/28 SMPK Miyata
                            'strRecord = strRecord & Format(udt_ww(i).WFSMP(2).KDAYCW, "mm/dd") & Chr$(9)
                            strRecord = strRecord & Format(udt_ww(i).WFSMP(UBound(udt_ww(i).WFSMP)).KDAYCW, "mm/dd") & Chr$(9)
                            'Chg End   2011/04/28 SMPK Miyata
                        End If
                    End If

                    '欠(上側)
    '                .spdWait.col = 11 + 1
    
                    If udt_ww(i).KETURAKU = True Then
                        strRecord = strRecord & "有" & Chr$(13) & Chr$(10)
    '                    .spdWait.Value = "有"
                    Else
                        strRecord = strRecord & "無" & Chr$(13) & Chr$(10)
    '                    .spdWait.Value = "無"
                    End If
                    
                    'データ表示
                    bRc = gFnc_SS_RecordSet(.spdWait, intRow, strRecord, udt_ww, i)
                    
                    'ｴﾗｰが発生した場合
                    If bRc = False Then
                        'ｴﾗｰ処理
                        .lblMsg.Caption = "表示エラー"
                        Exit Sub
                    End If
                    
    '                If Not WFJudgExecOkFlag(i) Then
    '' 2007/10/17 SET miyatake Add Start
    ''                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, intRow, 11, intRow, CTRL_DISABLE_GRAY
    '                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, intRow, 12, intRow, CTRL_DISABLE_GRAY
    '' 2007/10/17 SET miyatake Add End
    '                End If
    '
    '                'ホールドロット（0=通常，1=流動停止）
    '                If udt_ww(i).HOLDBCB = "1" Or udt_ww(i).WFHOLDFLGCB = "1" Then
    '                    'ﾎｰﾙﾄﾞ区分またはﾎｰﾙﾄﾞ区分(WF)が「1」のﾛｯﾄは選択不可とする
    '' 2007/10/17 SET miyatake Add Start
    ''                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, intRow, 11, intRow, CTRL_DISABLE_RED
    '                    SpCtrlBlockEnabled f_cmbc039_1.spdWait, 1, intRow, 12, intRow, CTRL_DISABLE_RED
    '' 2007/10/17 SET miyatake Add End
    ''''                End If
    '                End If
                End If
            End If
        Next 'Loop
        .spdWait.MaxRows = intRow
        .spdWait.ReDraw = True
    End With
End Sub
