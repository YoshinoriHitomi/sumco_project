VERSION 5.00
Begin VB.Form f_cmzcTBN 
   Caption         =   "振替番号入力"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows の既定値
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   3120
      MaxLength       =   5
      TabIndex        =   14
      Top             =   5350
      Width           =   1695
   End
   Begin VB.TextBox txtRiyuu 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      IMEMode         =   4  '全角ひらがな
      Left            =   720
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "f_cmzcTBN.frx":0000
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox txtSakiHinban 
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
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "TNS0013D00Y2"
      Top             =   1320
      Width           =   1815
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
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "INS0017A00Y1"
      Top             =   1320
      Width           =   1815
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
      Left            =   3360
      TabIndex        =   3
      Top             =   5880
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ＯＫ"
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
      Left            =   1560
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.TextBox txtBangou 
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   2040
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "TS-0000001"
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  '中央揃え
      Caption         =   "パスワード"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1560
      TabIndex        =   13
      Top             =   5400
      Width           =   1395
   End
   Begin VB.Label lblMsg 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   4800
      Width           =   5295
   End
   Begin VB.Label Label8 
      Alignment       =   2  '中央揃え
      Caption         =   "振替理由"
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
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   1395
   End
   Begin VB.Label Label7 
      Alignment       =   2  '中央揃え
      Caption         =   "振替先品番"
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
      Left            =   3960
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  '中央揃え
      Caption         =   "→"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   24
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   8
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
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
      Left            =   840
      TabIndex        =   7
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  '中央揃え
      Caption         =   "振替番号"
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
      Left            =   600
      TabIndex        =   5
      Top             =   2160
      Width           =   1400
   End
   Begin VB.Label Label1 
      Alignment       =   2  '中央揃え
      Caption         =   "振替番号入力"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "f_cmzcTBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'                                     2003/09/01
'======================================================
' 特採番号入力
' 概要    : 特採番号と特採理由の入力を行う。
' 参照    :
'======================================================

'概要      :Form Load時処理
'説明      :初期表示
'履歴      :2003/09/01  高田 作成
Private Sub Form_Load()
    '' 振替元品番を設定する
    txtMotoHinban.text = TBN_MotoHinban
    '' 振替先品番を設定する
    txtSakiHinban.text = TBN_SakiHinban
    
    ''  特採番号と特採理由
    txtBangou.text = TBN_Bangou
    txtRiyuu.text = TBN_Riyuu

    '' 表示位置
    Me.Move 9000, 3540
End Sub

'概要      :特採番号 LostFocus時処理
'説明      :大文字変換
'履歴      :2003/09/01  高田 作成
Private Sub txtBangou_LostFocus()
    '' 大文字変換
    txtBangou.text = StrConv(txtBangou.text, vbUpperCase)
    lblMsg.Caption = ""
End Sub

'概要      :ＯＫボタン押下時処理
'説明      :入力チェックを行う
'           特採番号と特採理由を保持し、呼出元画面に戻る
'履歴      :2003/09/01  高田 作成
Private Sub cmdOK_Click()
    '' 入力チェック
    If Trim$(txtBangou.text) = "" Then
'        lblMsg.Caption = GetMsgStr("EINIM")
        lblMsg.Caption = "振替番号が未入力です。"
        txtBangou.SetFocus
        TBN_Msg = ""
        Exit Sub
    End If

    '2011/07/01 add Kameda
    If Trim$(txtPass.text) = "" Then
        lblMsg.Caption = "パスワード未入力です。"
        txtPass.SetFocus
        Exit Sub
    End If
    If ChkPass = False Then
        lblMsg.Caption = "パスワードが違います。"
        txtPass.SetFocus
        Exit Sub
    End If
    
  
    
    '' 特採番号と特採理由を設定する
    TBN_Bangou = txtBangou.text
    TBN_Riyuu = txtRiyuu.text
    
    TBN_Msg = "振替入力しました"
    
    '' メール送信
    
    Unload Me
End Sub

'概要      :キャンセルボタン押下時処理
'説明      :呼出元画面に戻る
'履歴      :2003/09/01  高田 作成
Private Sub cmdCancel_Click()
    '' 特採番号と特採理由をクリアする
    TBN_Bangou = ""
    TBN_Riyuu = ""
    TBN_Msg = ""
    Unload Me
End Sub
'概要      :パスワード照合処理
'説明      :KODA9.X.55.'TOKSAI'.KCODEA９
'履歴      :2011/07/01  Kameda
Private Function ChkPass() As Boolean
    Dim sSql        As String
    Dim objOraDyn   As Object
    Dim iCount As Integer
    
        ChkPass = False
    
        sSql = ""
        sSql = sSql & "select   count(*) COUNT " & vbLf
        sSql = sSql & "from     KODA9" & vbLf
        sSql = sSql & "where    SYSCA9 = 'X'" & vbLf
        sSql = sSql & "and      SHUCA9 = '55'" & vbLf
        sSql = sSql & "and      CODEA9 = 'TOKSAI'" & vbLf
        sSql = sSql & "and      KCODEA9 = '" & Trim(txtPass.text) & "'" & vbLf
    
        'SQL文実行
        If DynSet2(objOraDyn, sSql) = False Then
            ''取得失敗
            Call MsgOut(100, sSql, ERR_DISP_LOG, "KODA9")
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        ''ﾃﾞｰﾀなし
        If objOraDyn.EOF Then
            Call MsgOut(55, "管理ｺｰﾄﾞﾃｰﾌﾞﾙ", ERR_DISP)
            Set objOraDyn = Nothing
            Exit Function
        End If
    
        iCount = objOraDyn.Fields("COUNT").Value
        
        If iCount = 0 Then
            Set objOraDyn = Nothing
            Exit Function
        End If
        
        '開放
        If Not objOraDyn Is Nothing Then objOraDyn.Close: Set objOraDyn = Nothing
        
        ChkPass = True

End Function
