VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Begin VB.Form f_cmbc039_6 
   BorderStyle     =   1  '固定(実線)
   Caption         =   " f_cmbc039_6"
   ClientHeight    =   10875
   ClientLeft      =   1575
   ClientTop       =   1680
   ClientWidth     =   15270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   725
   ScaleMode       =   3  'ﾋﾟｸｾﾙ
   ScaleWidth      =   1018
   StartUpPosition =   2  '画面の中央
   Begin VB.TextBox txtDateT 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   4200
      MaxLength       =   8
      TabIndex        =   1
      Text            =   "20120808"
      Top             =   960
      Width           =   1140
   End
   Begin VB.TextBox txtDateF 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2400
      MaxLength       =   8
      TabIndex        =   0
      Text            =   "20120808"
      Top             =   960
      Width           =   1140
   End
   Begin VB.Frame fraHead 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15225
      Begin VB.Label lblvers 
         Height          =   195
         Left            =   13680
         TabIndex        =   24
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lblTime 
         Height          =   150
         Left            =   13680
         TabIndex        =   5
         Top             =   240
         Width           =   1455
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
         Left            =   3810
         TabIndex        =   4
         Top             =   240
         Width           =   7050
      End
      Begin VB.Label lblTitle 
         Caption         =   "０枚ロット一覧表示"
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
         TabIndex        =   3
         Top             =   270
         Width           =   3495
      End
   End
   Begin VB.Frame fraF 
      Height          =   1095
      Left            =   30
      TabIndex        =   7
      Top             =   9540
      Width           =   15195
      Begin VB.CommandButton cmdF 
         Caption         =   "[F12]　　　 閉じる"
         Height          =   735
         Index           =   12
         Left            =   13920
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F３]　　　＊＊＊"
         Height          =   735
         Index           =   3
         Left            =   2824
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F２]　　　＊＊＊"
         Height          =   735
         Index           =   2
         Left            =   1592
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F１]　　　＊＊＊"
         Height          =   735
         Index           =   0
         Left            =   360
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F９]　　　抽出"
         Height          =   735
         Index           =   9
         Left            =   10216
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   240
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F４]　　　＊＊＊"
         Height          =   735
         Index           =   4
         Left            =   4056
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F５]　　　＊＊＊"
         Height          =   735
         Index           =   5
         Left            =   5288
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F６]　　　＊＊＊"
         Height          =   735
         Index           =   6
         Left            =   6520
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F７]　　　＊＊＊"
         Height          =   735
         Index           =   7
         Left            =   7752
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F８]　　　＊＊＊"
         Height          =   735
         Index           =   8
         Left            =   8984
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F10]　　　＊＊＊"
         Height          =   735
         Index           =   10
         Left            =   11448
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton cmdF 
         Caption         =   "[F11]　　　＊＊＊"
         Height          =   735
         Index           =   11
         Left            =   12680
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin FPSpread.vaSpread spdZeroView 
      Height          =   7725
      Left            =   240
      TabIndex        =   20
      Top             =   1680
      Width           =   14655
      _Version        =   196608
      _ExtentX        =   25850
      _ExtentY        =   13626
      _StockProps     =   64
      BackColorStyle  =   1
      ColHeaderDisplay=   0
      ColsFrozen      =   4
      MaxCols         =   15
      MaxRows         =   1
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "f_cmbc039_6.frx":0000
      UserResize      =   0
      VisibleCols     =   15
      VisibleRows     =   1
   End
   Begin VB.Label Label3 
      Caption         =   "(YYYYMMDD)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4140
      TabIndex        =   23
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label2 
      Caption         =   "(YYYYMMDD)"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2340
      TabIndex        =   22
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "WF払出日："
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label5 
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   375
   End
End
Attribute VB_Name = "f_cmbc039_6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===============================================================================
' ０枚ロット一覧表示画面
' 2012/09/07 SETsw Marushita
' 概要    :　サンプル消費で0枚ロットとなったデータを一覧表示する(10枚以下ロット流動対応)
'===============================================================================

'*******************************************************************************
'*    関数名        : DispSpdZeroView
'*
'*    処理概要      : 1.抽出条件により、0件ロット一覧を取得し表示する
'*
'*    パラメータ    : 変数名        ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub DispSpdZeroView()
    Dim intLoopCnt  As Integer
    Dim sSprSta     As String
    Dim intRowNo    As Integer
    Dim sDateF      As String
    Dim sDateT      As String
    Dim sMukesaki   As String
    Dim smpId(2)    As String
    Dim sErrMsg     As String
    
    intRowNo = 0
    
    'スプレッドコントロールの初期化処理
    SpCtrlInit f_cmbc039_6.spdZeroView, 0

    '画面から抽出条件を取得する
    If Trim(txtDateF.text) = "" Then
    Else
        '入力日付チェック
        If DateCheck(Trim(txtDateF.text), 0) = False Then
            lblMsg.Caption = "正しい日付を入力してください。"
            txtDateF.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txtDateT.text) = "" Then
    Else
        '入力日付チェック
        If DateCheck(Trim(txtDateT.text), 0) = False Then
            lblMsg.Caption = "正しい日付を入力してください。"
            txtDateT.SetFocus
            Exit Sub
        End If
        '入力日付大小チェック
        If Trim(txtDateF.text) > Trim(txtDateT.text) Then
            lblMsg.Caption = "正しい日付範囲を入力してください。"
            txtDateF.SetFocus
            Exit Sub
        End If
    End If
    
    sDateF = txtDateF.text
    'TO日付指定時は時間を付加
    If Trim(txtDateT.text) = "" Then
        sDateT = txtDateT.text
    Else
        sDateT = txtDateT.text & " 23:59:59"
    End If
    
    '0枚ロット一覧情報取得
    lblMsg.Caption = GetMsgStr(PWAIT)
    DoEvents
    
    '抽出条件を指定して対象データを取得する。
    If DBDRV_fcmbc039_6_Disp(gsMukeCd, sDateF, sDateT, typ_zero(), sErrMsg) = FUNCTION_RETURN_FAILURE Then
        lblMsg.Caption = "対象データ取得エラーです。"
        Exit Sub
    End If
        
    If UBound(typ_zero) = 0 Then
        lblMsg.Caption = "対象データがありません。"
        Exit Sub
    Else
        lblMsg.Caption = ""
    End If
    
    '対象データを一覧表示する
    With spdZeroView
        .ReDraw = False
        For intLoopCnt = 1 To UBound(typ_zero)
            intRowNo = intRowNo + 1
    
            sSprSta = ""

            If UBound(typ_zero(intLoopCnt).WFSMP) >= 1 Then
                smpId(1) = Trim(typ_zero(intLoopCnt).WFSMP(1).REPSMPLIDCW)
            Else
                smpId(1) = vbNullString
            End If
            If UBound(typ_zero(intLoopCnt).WFSMP) >= 2 Then
                smpId(2) = Trim(typ_zero(intLoopCnt).WFSMP(UBound(typ_zero(intLoopCnt).WFSMP)).REPSMPLIDCW)
            Else
                smpId(2) = vbNullString
            End If
                    
            .MaxRows = intRowNo
            .row = intRowNo
    
            '向先
            .col = 1
            .SetText 1, intRowNo, typ_zero(intLoopCnt).PLANTCAT
    
            '仮SXL-ID
            .col = 2
            .SetText 2, intRowNo, typ_zero(intLoopCnt).SXLIDCA
    
            '品番
            .col = 3
            .SetText 3, intRowNo, typ_zero(intLoopCnt).HINBCA
            
            '長さ
            .col = 4
            .SetText 4, intRowNo, typ_zero(intLoopCnt).GNLCA

            '枚数
            .col = 5
            .SetText 5, intRowNo, typ_zero(intLoopCnt).MAICB
    
            'WF払出日
            .col = 6
            .SetText 6, intRowNo, Format(typ_zero(intLoopCnt).TDAYCB, "yyyy/mm/dd")
    
            'サンプルID(上側)
            .col = 7
            .SetText 7, intRowNo, smpId(1)
    
            'サンプルID(下側)
            .col = 8
            .SetText 8, intRowNo, smpId(2)
    
            '最終受信日(上側)
            If Not (smpId(1) = "" Or _
                left(smpId(1), 1) = vbNullChar) Then
                If UBound(typ_zero(intLoopCnt).WFSMP) >= 1 Then
                    .col = 9
                    .SetText 9, intRowNo, Format(typ_zero(intLoopCnt).WFSMP(1).KDAYCW, "yyyy/mm/dd")
                End If
            End If
    
            '最終受信日(下側)
            If Not (smpId(2) = "" Or _
                left(smpId(2), 1) = vbNullChar) Then
                If UBound(typ_zero(intLoopCnt).WFSMP) >= 2 Then
                    .col = 10
                    .SetText 10, intRowNo, Format(typ_zero(intLoopCnt).WFSMP(UBound(typ_zero(intLoopCnt).WFSMP)).KDAYCW, "yyyy/mm/dd")
                End If
            End If

            '現在工程(TEST)
            .col = 11
            .SetText 11, intRowNo, typ_zero(intLoopCnt).NOWPROC

'            '欠(上側)
'            If typ_zero(intLoopCnt).KETURAKU = True Then
'                sSprSta = sSprSta & "有" & Chr$(13) & Chr$(10)
'            Else
'                sSprSta = sSprSta & "無" & Chr$(13) & Chr$(10)
'            End If
                    
            'データ表示
            '1行データセット
            '.Clip = sSprSta
            
'            bRc = gFnc_SS_RecordSet(.spdWait, intRow, strRecord, udt_ww, i)
'            'ｴﾗｰが発生した場合
'            If bRc = False Then
'                'ｴﾗｰ処理
'                .lblMsg.Caption = "表示エラー"
'                Exit Sub
'            End If
            
        Next
        .MaxRows = intRowNo
        .ReDraw = True
        
        If .MaxRows > 0 Then
            .col = 1
            .row = 1
            .Action = ActionActiveCell
        End If
        
    End With
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
Private Sub cmdF_Click(Index As Integer)
    '' 処理分岐
    Select Case Index
    Case 9        '' Ｆ9キー（抽出）
        Call DispSpdZeroView
    Case 12       '' Ｆ12キー（実行）
        Me.Visible = False
        Unload Me
    End Select
End Sub

'*******************************************************************************
'*    関数名        : Form_Activate
'*
'*    処理概要      : 1.Form_Activate処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Activate()
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
'*    処理概要      : 1.Form_Load処理
'*
'*    パラメータ    : 変数名      ,IO ,型       ,説明
'*                    なし
'*
'*    戻り値        : なし
'*
'*******************************************************************************
Private Sub Form_Load()
    ' 現在日時の表示
    SetPresentTime lblTime
    
    ' バージョン情報の表示
    lblvers.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    
    ' 抽出日付の初期値セット
    txtDateF.text = Format(DateAdd("m", -1, Date), "yyyymmdd")
    txtDateT.text = Format(Date, "yyyymmdd")

    ' ０枚ロット一覧画面の表示
    Call DispSpdZeroView

End Sub

