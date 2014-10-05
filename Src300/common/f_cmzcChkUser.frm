VERSION 5.00
Begin VB.Form f_cmzcChkUser 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "ﾛｸﾞｵﾝ"
   ClientHeight    =   1245
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735.587
   ScaleMode       =   0  'ﾕｰｻﾞｰ
   ScaleWidth      =   3549.215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ｷｬﾝｾﾙ"
      Height          =   390
      Left            =   2040
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   1290
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "ﾊﾟｽﾜｰﾄﾞ(&P):"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
End
Attribute VB_Name = "f_cmzcChkUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InpPass As String

Private Sub cmdCancel_Click()
    InpPass = ""
    txtPassword.Text = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    InpPass = txtPassword.Text
    txtPassword.Text = ""
    Me.Hide
End Sub

Private Sub Form_Load()
    InpPass = ""
End Sub

Public Function GetInpPass() As String
    GetInpPass = InpPass
End Function

'概要      :権限管理マスターと社員マスターより権限チェックをおこなう。
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :StaffID       ,I  ,String           ,社員ID
'          :ProcID        ,I  ,String           ,工程ID
'          :戻り値        ,O  ,FUNCTION_RETURN  ,正常・異常
'説明      :見つからない場合はVbNullStringを返す
'履歴      :2001/08/16 人見
'          :2009/08 Sumco 秋月

Private Function ChkUser(ByVal StaffID As String, ByVal ProcID As String) As FUNCTION_RETURN
    
    Dim dbIsMine As Boolean
    Dim rs As OraDynaset
    Dim sql As String
    Dim Pass As String
    Dim Pwchk As String
    Dim InpPwd As String

    On Error GoTo proc_err
    gErr.Push "s_cmzcChkUser.bas -- Function ChkUser"

    ChkUser = FUNCTION_RETURN_SUCCESS
    
    If OraDB Is Nothing Then
        dbIsMine = True
        OraDBOpen
    End If
    
    '' 引数で指定された社員IDと工程コードで検索を行う
    sql = " select T1.PASSWD as PASS, T4.PWCHECK as PWCHK from TBCMB001 T1, TBCMB004 T4 "
    sql = sql & " Where T1.EXECODE = T4.AUTHCODE "
    sql = sql & " and T1.STAFFID = '" & StaffID & "' "
    sql = sql & " and T4.TRANID = '" & ProcID & "' "
    
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)
    
    ''みつからなかったらエラー
    If rs.RecordCount = 0 Then
        ChkUser = FUNCTION_RETURN_FAILURE
    
    Else ''見つかったら、パスワードチェックをおこなう
        Pass = rs("PASS")
        Pwchk = rs("PWCHK")
        
        ' パスワードチェックあり
        If Trim(Pwchk) = "1" Then
            'Load f_cmzcChkUser
            f_cmzcChkUser.Show 1
            InpPwd = f_cmzcChkUser.GetInpPass()
            Unload f_cmzcChkUser
            If InpPwd <> Pass Then
                ChkUser = FUNCTION_RETURN_FAILURE
            End If
        End If
    End If
    rs.Close
    
    If dbIsMine Then
        OraDBClose
    End If
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    gErr.HandleError
    ChkUser = FUNCTION_RETURN_FAILURE
    Resume proc_exit
End Function

'概要      :権限チェック関数を呼ぶ前のフォーム名ー＞工程コード変換関数
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :StaffID        ,I  ,String          ,社員ID
'          :ProcID          ,I  ,String         ,工程ID
'          :戻り値        ,O  ,String           ,正常・異常
'説明      :見つからない場合は""空文字を返す
'履歴      :2001/08/16 人見
'          :2002/07/29 筑　（新フォーム名対応）
'          :2009/08 SUMCO 秋月　(Ｘ線測定実績入力:f_cmbc053_1)追加

Private Function CnvFrm2Prc(ByVal FrmN As String) As String

    On Error GoTo proc_err
    gErr.Push "s_cmzcChkUser.bas -- Function CnvFrm2Prc"

    Select Case (FrmN)
        Case "f_cmac001_1":  CnvFrm2Prc = vbNullString                   'メインメニュー画面
        Case "f_cmac001_2":  CnvFrm2Prc = vbNullString                   '仕様受入メニュー
        Case "f_cmac001_3":  CnvFrm2Prc = vbNullString                   '原料メニュー
        Case "f_cmac001_4":  CnvFrm2Prc = vbNullString                   '結晶引上メニュー
        Case "f_cmac001_5":  CnvFrm2Prc = vbNullString                   '結晶加工メニュー
        Case "f_cmac001_6":  CnvFrm2Prc = vbNullString                   '結晶検査メニュー
        Case "f_cmac001_7":  CnvFrm2Prc = vbNullString                   '結晶払出メニュー
        Case "f_cmac001_8":  CnvFrm2Prc = vbNullString                   'WF加工メニュー
        Case "f_cmac001_9":  CnvFrm2Prc = vbNullString                   'その他メニュー
        Case "f_cmac001_10":  CnvFrm2Prc = vbNullString                   'バーコードラベル再発行メニュー
        
        
        Case "f_cmbc001_1":  CnvFrm2Prc = PROCD_SIYOU_UKEIRE             '仕様受入仕掛一覧
        Case "f_cmbc001_2":  CnvFrm2Prc = PROCD_SIYOU_UKEIRE             '製品仕様受入
        Case "f_cmbc003_1":  CnvFrm2Prc = PROCD_SIYOU_NYUURYOKU          '製品仕様入力
        
        
        Case "f_cmbc004_1":  CnvFrm2Prc = PROCD_TAKESSYOU_UKEIRE         '多結晶受入棚入
        Case "f_cmbc005_1":  CnvFrm2Prc = PROCD_RIMERUTO_UKEIRE          'リメルト受入・切断
        Case "f_cmbc006_1":  CnvFrm2Prc = PROCD_RIMERUTO_HARAIDASI       'リメルト洗浄・払出
        Case "f_cmbc007_1":  CnvFrm2Prc = PROCD_GENRYOU_ZAIKO_SYUUSEI    '原料在庫修正
        Case "f_cmbc008_1":  CnvFrm2Prc = PROCD_KAKUAGE                  'クリスタルカタログ検索格上
        
        
        Case "f_cmbc008_3":  CnvFrm2Prc = vbNullString                   '号機一覧
        Case "f_cmbc009_2":  CnvFrm2Prc = vbNullString                   '指示一覧
        Case "f_cmbc009_3":  CnvFrm2Prc = PROCD_HIKIAGE_SIJI             '指示内容入力
        Case "f_cmbc009_4":  CnvFrm2Prc = vbNullString                   '原料一覧
        Case "f_cmbc009_5":  CnvFrm2Prc = vbNullString                   'PG-ID一覧
        Case "f_cmbc009_6":  CnvFrm2Prc = vbNullString                   'PG - ID詳細
        Case "f_cmbc009_7":  CnvFrm2Prc = vbNullString                   '結晶ブロック組合せ
        Case "f_cmbc013_1":  CnvFrm2Prc = PROCD_HIKIAGE_TOUNYUU          '引上投入実績初期画面
        Case "f_cmbc013_2":  CnvFrm2Prc = PROCD_HIKIAGE_TOUNYUU          '引上投入実績
        Case "f_cmbc013_3":  CnvFrm2Prc = vbNullString                   '原料投入実績一覧
        Case "f_cmbc014_1":  CnvFrm2Prc = PROCD_HIKIAGE_SYUURYOU         '引上終了実績入力
        Case "f_cmbc015_1":  CnvFrm2Prc = PROCD_TEIKO_HENSEKI_KEISSAN    '抵抗偏析計算
        
        
        Case "f_cmbc016_1":  CnvFrm2Prc = PROCD_KAKOU_HARAIDASI          '結晶加工払出
        Case "f_cmbc017_1":  CnvFrm2Prc = PROCD_KENNSAKU_KAKOU           '結晶研削加工実績入力
        Case "f_cmbc018_1":  CnvFrm2Prc = vbNullString                   '切断待ち一覧
        Case "f_cmbc018_2":  CnvFrm2Prc = PROCD_SETUDAN                  '切断
        Case "f_cmbc019_1":  CnvFrm2Prc = PROCD_KESSYOU_HOLD             '結晶ホールド
        Case "f_cmbc020_1":  CnvFrm2Prc = PROCD_KESSYOU_HOLD_KAIJO       '結晶ホールド解除
        
        
        Case "f_cmbc021_1":  CnvFrm2Prc = PROCD_FTIR                     'FTIR(Oi,Cs)実績入力
        Case "f_cmbc022_1":  CnvFrm2Prc = PROCD_GFA                      'GFA(Oi)実績入力
        Case "f_cmbc023_1":  CnvFrm2Prc = PROCD_TEIKOU                   '抵抗実績入力
        Case "f_cmbc024_1":  CnvFrm2Prc = PROCD_BMD                      'BMD実績入力
        Case "f_cmbc025_1":  CnvFrm2Prc = PROCD_OSF                      'OSF実績入力
        Case "f_cmbc026_1":  CnvFrm2Prc = PROCD_GD                       'GD実績入力
        Case "f_cmbc027_1":  CnvFrm2Prc = PROCD_LIFETIME                 'ライフタイム実績入力
        Case "f_cmbc028_1":  CnvFrm2Prc = PROCD_EPD                      'EPD実績入力
        Case "f_cmbc029_1":  CnvFrm2Prc = PROCD_GFA_KOUSEIJOHO           'GFA校正情報設定
        '2009/08 SUMCO Akizuki
        Case "f_cmbc053_1":  CnvFrm2Prc = PROCD_X                        'X線実績入力

                
'' 09/01/28 FAE)akiyama start
        Case "f_cmbc030_1":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '待ち一覧(総合判定）
'' 09/01/28 FAE)akiyama start
        Case "f_cmbc030_2":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '総合判定
        Case "f_cmbc031_1":  CnvFrm2Prc = PROCD_KESSYOU_SOUGOUHANTEI     '総合判定 - 再分割
        
        Case "f_cmbc032_1":  CnvFrm2Prc = vbNullString                   '待ち一覧（結晶最終払出）
        Case "f_cmbc032_2":  CnvFrm2Prc = PROCD_KESSYOU_SAISYUU_HARAIDASI ' 結晶最終払出入力
        Case "f_cmbc033_1":  CnvFrm2Prc = vbNullString                   '待ち一覧（抜試指示）
        Case "f_cmbc033_2":  CnvFrm2Prc = PROCD_NUKISI_SIJI              '抜試指示入力
        Case "f_cmbc034_1":  CnvFrm2Prc = PROCD_WFC_HARAIDASI            'WFセンター払出
        Case "f_cmbc035_1":  CnvFrm2Prc = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU ' 結晶情報変更
        Case "f_cmbc036_1":  CnvFrm2Prc = vbNullString                   '欠落ブロック一覧
        Case "f_cmbc036_2":  CnvFrm2Prc = PROCD_NUKISI_HENKOU            '抜試指示変更入力
        Case "f_cmbc036_3":  CnvFrm2Prc = PROCD_NUKISI_HENKOU            '関連ﾌﾞﾛｯｸ管理　08/01/28 ooba
        Case "f_cmbc037_1":  CnvFrm2Prc = PROCD_KOUNYU_TAN_KESSYOU       '購入単結晶受入
        Case "f_cmbc038_1":  CnvFrm2Prc = PROCD_BLOCK                    'ブロックラベル発行 4/16  Yam
        
        
        Case "f_cmbc039_1":  CnvFrm2Prc = vbNullString                   'WF総合判定 - 判定待ち一覧
        Case "f_cmbc039_2":  CnvFrm2Prc = PROCD_WFC_SOUGOUHANTEI         'WF総合判定 - 総合判定
        Case "f_cmbc039_3":  CnvFrm2Prc = PROCD_WFC_SOUGOUHANTEI         'WF総合判定 - 再抜試指示
        Case "f_cmbc040_1":  CnvFrm2Prc = PROCD_SXL_KAKUTEI              'シングル確定

''Add Start 2011/04/06 SMPK Nakamura
        Case "f_cmbc055_1":  CnvFrm2Prc = PROCD_FRS_KOUSEIJOHO           'FRS校正情報設定
''Add End 2011/04/06 SMPK Nakamura
        
        Case "f_cmcc001_1":  CnvFrm2Prc = PROCD_PGID_MNT                 'PG-ID一覧
        Case "f_cmcc001_2":  CnvFrm2Prc = PROCD_PGID_MNT                 'PG-ID詳細
        Case "f_cmcc002_1":  CnvFrm2Prc = PROCD_SEISAKUJOUKEN_MNT        '製作条件メンテナンス
        
        Case "f_cmcc003_1":  CnvFrm2Prc = PROCD_SYAIN_MNT                '社員マスタメンテナンス'02/3/30 Yam
        Case "f_cmcc004_1":  CnvFrm2Prc = PROCD_KENGEN_MNT               '権限マスタメンテナンス'02/3/30 Yam
        Case "f_cmcc005_1":  CnvFrm2Prc = PROCD_CODE_MNT                 'コードマスタメンテナンス'02/3/30 Yam
        Case "f_cmcc006_1":  CnvFrm2Prc = PROCD_PRINTER_MNT              'ラベルプリンタマスタメンテナンス'02/3/30 Yam
        ' 特採用を追加  2003/09/12 SystemBrain ===================> START
        Case "TOKUSAI":      CnvFrm2Prc = PROCD_TOKUSAI_KENGEN           '特採権限
        ' 特採用を追加  2003/09/12 SystemBrain ===================> END
        Case Else:          CnvFrm2Prc = vbNullString
    End Select
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    gErr.HandleError
    CnvFrm2Prc = ""
    Resume proc_exit
End Function


'概要      :社員IDと画面名から、実行権限を得る
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型        ,説明
'          :frmName       ,   ,String    ,
'          :staffID       ,   ,String    ,
'          :戻り値        ,O  ,FUNCTION_RETURN,
'説明      :
'履歴      :2001/8/17 作成  野村
Public Function CanExec(ByVal frmName$, ByVal StaffID$) As Boolean
    If ChkUser(StaffID, CnvFrm2Prc(frmName)) = FUNCTION_RETURN_SUCCESS Then
        CanExec = True
        XSDC3_StaffID = StaffID
    Else
        CanExec = False
    End If
End Function
