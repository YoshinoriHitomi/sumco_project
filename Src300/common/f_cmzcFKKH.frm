VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmzcFKKH 
   BorderStyle     =   1  '固定(実線)
   Caption         =   "振替可能候補品番"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdKettei 
      Caption         =   "決定"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdKouho 
      Caption         =   "候補品番表示"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
   Begin FPSpread.vaSpread sprHinban 
      Height          =   2295
      Left            =   840
      TabIndex        =   1
      Top             =   2760
      Width           =   2295
      _Version        =   196608
      _ExtentX        =   4048
      _ExtentY        =   4048
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      MaxCols         =   1
      ScrollBars      =   2
      SelectBlockOptions=   2
      ShadowColor     =   14215660
      ShadowDark      =   10070188
      ShadowText      =   0
      SpreadDesigner  =   "f_cmzcFKKH.frx":0000
      VisibleCols     =   1
      VisibleRows     =   500
   End
   Begin VB.TextBox txtMotoHinban 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "INS0017A00Y1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Caption         =   "↓"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "振替元品番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "振替可能候補品番"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   15.75
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "f_cmzcFKKH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'                                     2003/09/01
'======================================================
' 振替可能候補品番
' 概要    : 振替元品番より振替可能候補品番を一覧表示し、
'           振替先品番として決定する。
' 参照    :
'======================================================

'概要      :Form Load時処理
'説明      :初期表示
'履歴      :2003/09/01  高田 作成
Private Sub Form_Load()
    '' 振替元品番を設定する
    txtMotoHinban.text = FKKH_MotoHinban
    
    '' 初期設定
    '''  品番
    With sprHinban
        .MaxRows = 0
        .col = -1
        .row = -1
        .Lock = True
        .RowHeight(-1) = 12.27
    End With
    '''  決定ボタン
    cmdKettei.Enabled = False

    '' 表示位置
    Me.Move 9000, 3540
End Sub

'概要      :候補品番表示ボタン押下時処理
'説明      :振替可能な品番を一覧に表示する
'履歴      :2003/09/01  高田 作成
Private Sub cmdKouho_Click()
    Dim RET As Integer
    Dim ErrCode As Integer
    Dim ErrMsg As String
    
    ' マウスポインタを砂時計に変更
    Screen.MousePointer = vbHourglass

    '' 振替候補品番取得(仕様チェック)共通関数
    RET = fncGetKouhoHinbanShiyou(FKKH_Proccd, FKKH_Crynum, FKKH_MotoHinban, KouhoHinban(), ErrCode, ErrMsg)
    
    If RET <> 0 Then
        Screen.MousePointer = vbDefault
        Call MsgBox(ErrMsg, vbOKOnly, "振替可能候補品番")
        Exit Sub
    End If
    
    '' 振替候補品番を一覧に表示する
    Call FurikaeKouhoSet

    ' マウスポインタを元に戻す
    Screen.MousePointer = vbDefault
    
    '' 決定ボタン
    cmdKettei.Enabled = True

End Sub

'概要      :一覧表示
'説明      :振替候補品番を一覧に表示する
'履歴      :2003/09/01  高田 作成
Private Sub FurikaeKouhoSet()
    Dim tblCnt As Long
    Dim cnt As Long
    
    With sprHinban
        .ReDraw = False
        .MaxRows = 0
        
        tblCnt = UBound(KouhoHinban)
        .MaxRows = tblCnt + 1
                
        For cnt = 0 To tblCnt
            .row = cnt + 1
            
            '振替候補品番
            .col = 1
            .text = KouhoHinban(cnt).GETHINBAN
        Next
        .ReDraw = True
    End With
End Sub

'概要      :決定ボタン押下時処理
'説明      :振替先品番として決定し、呼出元画面に戻る
'履歴      :2003/09/01  高田 作成
Private Sub cmdKettei_Click()
    '' 振替先品番を設定する
    With sprHinban
        .row = .ActiveRow
        .col = 1
        FKKH_SakiHinban = .text
    End With
    
    Unload Me
End Sub

'概要      :キャンセルボタン押下時処理
'説明      :呼出元画面に戻る
'履歴      :2003/09/01  高田 作成
Private Sub cmdCancel_Click()
    '' 振替先品番をクリアする
    FKKH_SakiHinban = ""
    
    Unload Me
End Sub

' @(f)
'
' 機能      : スプレッドシートクリック
'
' 返り値    : なし
'
' 引き数    :
'
' 機能説明  : イベント関数
'
' 備考      : スプレッドシートのソート処理  2008/05/28 追加:Kameda
'
Private Sub sprHinban_click(ByVal col As Long, ByVal row As Long)
    
    'スプレッドシートの表示を更新しない
    sprHinban.ReDraw = False
    Select Case row
        'P1 列タイトルを押下した場合、押下された列を元にソート
        Case 0
            'Call sprSort(sprHinban, col)
            With sprHinban
                .BlockMode = True                               '  セルブロックを有効
                .col = 1                                        '  列を設定
                .col2 = .MaxCols                                '  最終列を設定
                .row = 1                                        '  行を設定
                .row2 = .MaxRows                                '  最終行を設定
                .SortBy = SortByRow                             '  行単位に並び替え
                .SortKey(1) = col                               '  並び替えのキーを設定
                
                If .SortKey(1) = col And .SortKeyOrder(1) = SortKeyOrderAscending Then
                    .SortKeyOrder(1) = SortKeyOrderDescending   '  降順に並び替えを設定
                Else
                    .SortKeyOrder(1) = SortKeyOrderAscending    '  昇順に並び替えを設定
                End If
                
                .Action = ActionSort                            '  並び替えを実行
                .BlockMode = False                              '  セルブロックを無効
            End With
    End Select

    'スプレッドシートの表示を更新する
    sprHinban.ReDraw = True

End Sub


