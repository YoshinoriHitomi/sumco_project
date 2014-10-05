VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "c_cmzcLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'===============================================================================
'クラス  : c_cmzcLabel
'解説    : ラベルプリンタ印刷要求テーブル（TBCMC001）登録用データを
'          ラベル発行要求テーブル（KODZ6）に登録するためのクラス
'作成者  : SETsw kakeida
'===============================================================================

'プロパティの値を格納する変数
'ラベルプリンタ印刷要求テーブル（TBCMC001）項目
Private m_strQueDate As Variant             'キュー日付
Private m_strReqKind As Variant             '印刷要求区分
Private m_strPrintKind As Variant           '印刷種類
Private m_strEndFlg As Variant              '完了区分
Private m_strStatus As Variant              '終了ステータス
Private m_strBlockIdUmu As Variant          'ブロックID有無区分
Private m_strProcCode As Variant            '工程コード
Private m_strEtcPrKind As Variant           'その他ラベル種類
Private m_strCryNum As Variant              '結晶番号

Private m_numIngotPos As Variant            '結晶内位置
Private m_numSmplNo As Variant              'サンプル№
Private m_strMtrlNum As Variant             '原料番号
Private m_strSmtrlNum As Variant            '精製原料番号
Private m_strBlockId As Variant             'ブロックID

Private m_strHinban As Variant              '品番
Private m_numRevNum As Variant              '製品番号改訂番号
Private m_strFactory As Variant             '工場
Private m_strOpeCond As Variant             '操業条件
Private m_strMachine As Variant             '要求マシン名

Private m_strStaffId As Variant             '要求担当者ID
Private m_strRegDate As Variant             '登録日付
Private m_strUpdDate As Variant             '更新日付
Private m_strCryIndRs As Variant            '結晶検査指示（Rs）
Private m_strCryIndOi As Variant            '結晶検査指示（Oi）
Private m_strCryIndB1 As Variant            '結晶検査指示（B1）
Private m_strCryIndB2 As Variant            '結晶検査指示（B2）
Private m_strCryIndB3 As Variant            '結晶検査指示（B3）
Private m_strCryIndL1 As Variant            '結晶検査指示（L1）
Private m_strCryIndL2 As Variant            '結晶検査指示（L2）
Private m_strCryIndL3 As Variant            '結晶検査指示（L3）
Private m_strCryIndL4 As Variant            '結晶検査指示（L4）
Private m_strCryIndCs As Variant            '結晶検査指示（Cs）
Private m_strCryIndGD As Variant            '結晶検査指示（GD）
Private m_strCryIndT As Variant             '結晶検査指示（T）
Private m_strCryIndEP As Variant            '結晶検査指示（EPD）
Private m_strCryIndX As Variant             '結晶検査指示（X）      '2009/07/31追加 SETsw kubota
Private m_strCryIndCO3 As Variant           '結晶検査指示（CO3）    '2010/12/02追加 SETsw kubota
Private m_strCryIndC As Variant             '結晶検査指示（C）      '2010/12/02追加 SETsw kubota
Private m_strCryIndCJ As Variant            '結晶検査指示（CJ）     '2010/12/02追加 SETsw kubota
Private m_strCryIndCJ2 As Variant           '結晶検査指示（CJ2）    '2010/12/02追加 SETsw kubota

Private m_strPrKey01 As Variant             '帳票キーデータ1

'ラベル発行要求テーブル（KODZ6）項目
Private m_strLabelId As Variant             '依頼帳票ラベル名（新ラベル種別）
Private m_strPrCdSub As Variant             'ラベル連番
Private m_strEFlgz6 As Variant              '終了フラグ
Private m_strStatz6 As Variant              '状態ステータス
Private m_strKariHon As Variant             '仮本区分

'新旧プロセス区分
Private m_procKubun As Boolean              'Trueは新プロセス

'ラベル実行プロセスフラグ
Private Const LBL_NEW_PROC = True           '新プロセス

'===============================================================================
'概要      :初期化処理
'解説      :各変数を初期化する
'===============================================================================
Private Sub Class_Initialize()

    m_strQueDate = "":        m_strReqKind = ""
    m_strPrintKind = "":      m_strEndFlg = ""
    m_strStatus = "":         m_strBlockIdUmu = ""
    m_strProcCode = "":       m_strEtcPrKind = ""
    m_strCryNum = "":         m_numIngotPos = 0
    m_numSmplNo = 0:          m_strMtrlNum = ""
    m_strSmtrlNum = "":       m_strBlockId = ""
    m_strHinban = "":         m_numRevNum = 0
    m_strFactory = "":        m_strOpeCond = ""
    m_strMachine = "":        m_strStaffId = ""
    m_strRegDate = "":        m_strUpdDate = ""
    m_strCryIndRs = "":       m_strCryIndOi = ""
    m_strCryIndB1 = "":       m_strCryIndB2 = ""
    m_strCryIndB3 = "":       m_strCryIndL1 = ""
    m_strCryIndL2 = "":       m_strCryIndL3 = ""
    m_strCryIndL4 = "":       m_strCryIndCs = ""
    m_strCryIndGD = "":       m_strCryIndT = ""
    m_strCryIndEP = "":       m_strPrKey01 = ""
    m_strCryIndX = ""
    m_strCryIndCO3 = ""
    m_strCryIndC = ""
    m_strCryIndCJ = ""
    m_strCryIndCJ2 = ""
    
    m_strLabelId = "":        m_strPrCdSub = "1"
    m_strEFlgz6 = "0":        m_strStatz6 = "0000"
    m_strKariHon = "0"
    
    m_procKubun = False    '新旧プロセス区分を旧に初期化
    
End Sub

'===============================================================================
'関数名     :新旧プロセス判定処理（ファサード）
'機能       :その他ラベル種類（旧ラベル種別）に応じて各メソッドを呼ぶ。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :なし
'===============================================================================
Public Sub Label_Process_Check()
      
  'その他ラベル種類
  If m_strEtcPrKind = "00" Then          'サンプルラベル
      Call SampleLabel_Check
  ElseIf m_strEtcPrKind = "15" Then      '払出ブロックラベル
      Call HaraidashiLabel_Check
  Else
      Call Set_LabelId                   'ラベルＩＤの決定
      Call Judge_Process                 '新旧プロセス判定
  End If
    
End Sub

'===============================================================================
'関数名     :ラベル発行要求テーブル登録処理（ファサード）
'機能       :その他ラベル種類（旧ラベル種別）に応じて各メソッドを呼ぶ。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :成功 True、エラー False
'===============================================================================
Public Function Label_Facade() As Boolean
    
  Dim iRet As Boolean
    
  'その他ラベル種類
  If m_strEtcPrKind = "00" Then          'サンプルラベル
      iRet = SampleLabel()
  ElseIf m_strEtcPrKind = "15" Then      '払出ブロックラベル
      iRet = HaraidashiLabel()
  Else
      Call Set_LabelId                   'ラベルＩＤの決定
      iRet = KODZ6_Insert()              'ＫＯＤＺ６に登録
  End If
  
  '戻り値
  Label_Facade = iRet
    
End Function

'===============================================================================
'関数名     :新旧プロセス判定処理（サンプルラベル）
'機能       :サンプルラベルに応じて新ラベル種別をセットし、旧プロセスか
'            新プロセスか判定する。
'            判定したサンプルが全て新規の場合のみ新プロセスとし、それ以外は
'            旧プロセスとする。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :なし
'===============================================================================
Private Sub SampleLabel_Check()

'旧プロセス(TBCMC001書込)の判断を廃止 2010/12/14 SETsw kubota
'    Dim iRet As Boolean
'
'    'Oi・Cs
'    If m_strCryIndOi = "1" Or m_strCryIndCs = "1" Then
'        m_strLabelId = "L300"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
'
'    '比抵抗
'    If m_strCryIndRs = "1" Then
'        m_strLabelId = "L301"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
'
'    'EPD
'    If m_strCryIndEP = "1" Then
'        m_strLabelId = "L302"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
'
'    'OSF・BMD
'    If m_strCryIndB1 = "1" Or m_strCryIndB2 = "1" Or m_strCryIndB3 = "1" Or m_strCryIndL1 = "1" Or _
'                              m_strCryIndL2 = "1" Or m_strCryIndL3 = "1" Or m_strCryIndL4 = "1" Then
'        m_strLabelId = "L303"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
'
'    'ライフタイム
'    If m_strCryIndT = "1" Then
'        m_strLabelId = "L304"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
'
'    'GD
'    If m_strCryIndGD = "1" Then
'        m_strLabelId = "L305"
'        Call Judge_Process
'        If m_procKubun = False Then
'            Exit Sub
'        End If
'    End If
    
    '新プロセスの場合、新旧プロセス区分をTrueにする
    m_procKubun = True

End Sub

'===============================================================================
'関数名     :サンプルラベル生成
'機能       :ラベル種類がサンプルラベルの場合の処理。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function SampleLabel() As Boolean

    Dim iRet As Boolean

    iRet = True

    'Oi・Cs
    If m_strCryIndOi = "1" Or m_strCryIndCs = "1" Then
        m_strLabelId = "L300"
        iRet = KODZ6_Insert()
    End If

    '比抵抗
    If (iRet) Then
        If m_strCryIndRs = "1" Then
            m_strLabelId = "L301"
            iRet = KODZ6_Insert()
        End If
    End If
    
    'EPD
    If (iRet) Then
        If m_strCryIndEP = "1" Then
            m_strLabelId = "L302"
            '全サンプルラベル出力プロセスにEPDラベル出力を追加 2011/03/24 SETsw kubota
            'iRet = KODZ6_Insert()
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If

'>>>>> 全サンプルラベル出力プロセスを使用 2010/12/02 SETsw kubota -------------
'    'OSF・BMD
'    If (iRet) Then
'        If m_strCryIndB1 = "1" Or m_strCryIndB2 = "1" Or m_strCryIndB3 = "1" Or m_strCryIndL1 = "1" Or _
'                                  m_strCryIndL2 = "1" Or m_strCryIndL3 = "1" Or m_strCryIndL4 = "1" Then
'            m_strLabelId = "L303"
'            iRet = KODZ6_Insert()
'        End If
'    End If
    
    'OSF・BMD
    If (iRet) Then
        If m_strCryIndB1 = "1" Then
            m_strLabelId = "L413"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndB2 = "1" Then
            m_strLabelId = "L414"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndB3 = "1" Then
            m_strLabelId = "L415"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndL1 = "1" Then
            m_strLabelId = "L405"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndL2 = "1" Then
            m_strLabelId = "L406"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndL3 = "1" Then
            m_strLabelId = "L407"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    If (iRet) Then
        If m_strCryIndL4 = "1" Then
            m_strLabelId = "L408"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
'<<<<< 全サンプルラベル出力プロセスを使用 2010/12/02 SETsw kubota -------------

    'ライフタイム
    If (iRet) Then
        If m_strCryIndT = "1" Then
            m_strLabelId = "L304"
            iRet = KODZ6_Insert()
        End If
    End If

    'GD
    If (iRet) Then
        If m_strCryIndGD = "1" Then
            m_strLabelId = "L305"
            iRet = KODZ6_Insert()
        End If
    End If

    'X線測定    2009/07/31追加 SETsw kubota
    If (iRet) Then
        If m_strCryIndX = "1" Then
            m_strLabelId = "L325"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numIngotPos, m_numSmplNo)
        End If
    End If

    'CO3測定    2010/12/02追加 SETsw kubota
    If (iRet) Then
        If m_strCryIndCO3 = "1" Then
            m_strLabelId = "L409"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    
    'C測定      2010/12/02追加 SETsw kubota
    If (iRet) Then
        If m_strCryIndC = "1" Then
            m_strLabelId = "L410"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    
    'CJ測定     2010/12/02追加 SETsw kubota
    If (iRet) Then
        If m_strCryIndCJ = "1" Then
            m_strLabelId = "L411"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If
    
    'CJ2測定    2010/12/02追加 SETsw kubota
    If (iRet) Then
        If m_strCryIndCJ2 = "1" Then
            m_strLabelId = "L412"
            iRet = KODZ6_Insert_2(m_strCryNum, m_numSmplNo)
        End If
    End If

    '戻り値
    SampleLabel = iRet

End Function

'===============================================================================
'関数名     :新旧プロセス判定処理（払出ブロックラベル）
'機能       :新ラベル種別をセットし、旧プロセスか新プロセスか判定する。
'            ２レコード分（ラベル連番が'1'と'2'）判定し、どちらも新規の場合のみ
'            新プロセスとし、それ以外は旧プロセスとする。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :なし
'===============================================================================
Private Sub HaraidashiLabel_Check()

    m_strLabelId = "L323"
    m_strPrCdSub = "1"
    
    Call Judge_Process
    
    If m_procKubun = False Then
        Exit Sub
    End If

    m_strLabelId = "L323"
    m_strPrCdSub = "2"
    
    Call Judge_Process

End Sub

'===============================================================================
'関数名     :払出ブロックラベル生成
'機能       :払出ブロックラベルは２レコード同時に作成する。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function HaraidashiLabel() As Boolean

    Dim iRet As Boolean

    iRet = True
    
    m_strLabelId = "L323"
    m_strPrCdSub = "1"
    
    iRet = KODZ6_Insert()

    If (iRet) Then
        m_strLabelId = "L323"
        m_strPrCdSub = "2"
        iRet = KODZ6_Insert()
    End If
    
    '戻り値
    HaraidashiLabel = iRet

End Function

'===============================================================================
'関数名     :ラベル発行要求テーブル登録処理
'機能       :新規プロセス時、ラベルプリンタ印刷要求テーブル登録用データを
'            ラベル発行要求テーブルに登録する。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function KODZ6_Insert() As Boolean
    Dim sSql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcLabel.cls -- Function KODZ6_Insert"

    '新規プロセス時、KODZ6にINSERTする
    sSql = ""
    sSql = "INSERT INTO KODZ6 ("
    sSql = sSql & " IRDAYZ6"
    sSql = sSql & ", WSIDZ6"
    sSql = sSql & ", PRCODEZ6"
    sSql = sSql & ", PRCNTZ6"
    sSql = sSql & ", PRIDZ6"
    sSql = sSql & ", PRKEY01Z6"
    sSql = sSql & ", PRKEY02Z6"
    sSql = sSql & ", PRKEY03Z6"
    sSql = sSql & ", PRKEY04Z6"
    sSql = sSql & ", PRKEY05Z6"
    sSql = sSql & ", PRCDSUBZ6"
    sSql = sSql & ", ENDFLGZ6"
    sSql = sSql & ", STATUSZ6"
    sSql = sSql & ", KARIHONZ6"
    sSql = sSql & " ) VALUES ("
    sSql = sSql & "TO_DATE('" & m_strQueDate & "','yyyy/mm/dd hh24:mi:ss')"  '依頼日時
    sSql = sSql & ", '" & m_strMachine & "'"                                 '依頼パソコン名
    sSql = sSql & ", '" & m_strLabelId & "'"                                 '依頼帳票ラベル名
    sSql = sSql & ", 1"                                                      '出力枚数
    sSql = sSql & ", ' '"                                                    '印字先情報（プリンタ名）
    
    '桁数チェック
    sSql = sSql & ", '" & Str_Check(m_strQueDate, 14) & Str_Check(m_strReqKind, 1) _
                        & Str_Check(m_strPrintKind, 4) & Str_Check(m_strEndFlg, 1) _
                        & Str_Check(m_strStatus, 4) & Str_Check(m_strBlockIdUmu, 1) _
                        & Str_Check(m_strProcCode, 5) & Str_Check(m_strEtcPrKind, 2) _
                        & Str_Check(m_strCryNum, 12) & "'"                   '帳票キーデータ1
                        
    sSql = sSql & ", '" & Num_Check(m_numIngotPos, 8) & Num_Check(m_numSmplNo, 8) _
                        & Str_Check(m_strMtrlNum, 10) & Str_Check(m_strSmtrlNum, 10) _
                        & Str_Check(m_strBlockId, 12) & "'"                  '帳票キーデータ2
                        
    sSql = sSql & ", '" & Str_Check(m_strHinban, 8) & Num_Check(m_numRevNum, 8) _
                        & Str_Check(m_strFactory, 1) & Str_Check(m_strOpeCond, 1) _
                        & Str_Check(m_strMachine, 30) & "'"                  '帳票キーデータ3
                        
    sSql = sSql & ", '" & Str_Check(m_strStaffId, 8) & Str_Check(m_strRegDate, 14) _
                        & Str_Check(m_strUpdDate, 14) & Str_Check(m_strCryIndRs, 1) _
                        & Str_Check(m_strCryIndOi, 1) & Str_Check(m_strCryIndB1, 1) _
                        & Str_Check(m_strCryIndB2, 1) & Str_Check(m_strCryIndB3, 1) _
                        & Str_Check(m_strCryIndL1, 1) & Str_Check(m_strCryIndL2, 1) _
                        & Str_Check(m_strCryIndL3, 1) & Str_Check(m_strCryIndL4, 1) _
                        & Str_Check(m_strCryIndCs, 1) & Str_Check(m_strCryIndGD, 1) _
                        & Str_Check(m_strCryIndT, 1) & Str_Check(m_strCryIndEP, 1) _
                        & "'"                                                '帳票キーデータ4
    
    sSql = sSql & ", '" & m_strPrKey01 & "'"                                 '帳票キーデータ5
    
    sSql = sSql & ", '" & m_strPrCdSub & "'"                                 'ラベル連番
    sSql = sSql & ", '" & m_strEFlgz6 & "'"                                  '終了フラグ
    sSql = sSql & ", '" & m_strStatz6 & "'"                                  '状態ステータス
    sSql = sSql & ", '" & m_strKariHon & "'"                                 '仮本区分
    sSql = sSql & " )"

    If OraDB.ExecuteSQL(sSql) < 1 Then
        'ロールバック
        OraDB.Rollback
        KODZ6_Insert = False
        GoTo proc_exit
    End If

    KODZ6_Insert = True
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    KODZ6_Insert = False
    Resume proc_exit
End Function

'===============================================================================
'関数名     :ラベル発行要求テーブル登録処理(新規ラベル用)
'機能       :ラベル発行要求テーブルに登録する。(TBCMC001形式への変換なし)
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2009/08/11  1.00    SETsw kubota
'更新
'-------------------------------------------------------------------------------
'引数　     :strPrKey01   I   帳票キーデータ1登録値
'           :strPrKey02   I   帳票キーデータ2登録値
'           :strPrKey03   I   帳票キーデータ3登録値
'           :strPrKey04   I   帳票キーデータ4登録値
'           :strPrKey05   I   帳票キーデータ5登録値
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function KODZ6_Insert_2(Optional ByVal strPrKey01 As String _
                              , Optional ByVal strPrKey02 As String _
                              , Optional ByVal strPrKey03 As String _
                              , Optional ByVal strPrKey04 As String _
                              , Optional ByVal strPrKey05 As String _
                              ) As Boolean
    Dim sSql As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "c_cmzcLabel.cls -- Function KODZ6_Insert_2"

    '新規プロセス時、KODZ6にINSERTする
    sSql = ""
    sSql = "INSERT INTO KODZ6 ("
    sSql = sSql & " IRDAYZ6"
    sSql = sSql & ", WSIDZ6"
    sSql = sSql & ", PRCODEZ6"
    sSql = sSql & ", PRCNTZ6"
    sSql = sSql & ", PRIDZ6"
    sSql = sSql & ", PRKEY01Z6"
    sSql = sSql & ", PRKEY02Z6"
    sSql = sSql & ", PRKEY03Z6"
    sSql = sSql & ", PRKEY04Z6"
    sSql = sSql & ", PRKEY05Z6"
    sSql = sSql & ", PRCDSUBZ6"
    sSql = sSql & ", ENDFLGZ6"
    sSql = sSql & ", STATUSZ6"
    sSql = sSql & ", KARIHONZ6"
    sSql = sSql & " ) VALUES ("
    sSql = sSql & "TO_DATE('" & m_strQueDate & "','yyyy/mm/dd hh24:mi:ss')"  '依頼日時
    sSql = sSql & ", '" & m_strMachine & "'"        '依頼パソコン名
    sSql = sSql & ", '" & m_strLabelId & "'"        '依頼帳票ラベル名
    sSql = sSql & ", 1"                             '出力枚数
    sSql = sSql & ", ' '"                           '印字先情報（プリンタ名）
    sSql = sSql & ", '" & strPrKey01 & "'"          '帳票キーデータ1
    sSql = sSql & ", '" & strPrKey02 & "'"          '帳票キーデータ2
    sSql = sSql & ", '" & strPrKey03 & "'"          '帳票キーデータ3
    sSql = sSql & ", '" & strPrKey04 & "'"          '帳票キーデータ4
    sSql = sSql & ", '" & strPrKey05 & "'"          '帳票キーデータ5
    sSql = sSql & ", '" & m_strPrCdSub & "'"        'ラベル連番
    sSql = sSql & ", '" & m_strEFlgz6 & "'"         '終了フラグ
    sSql = sSql & ", '" & m_strStatz6 & "'"         '状態ステータス
    sSql = sSql & ", '" & m_strKariHon & "'"        '仮本区分
    sSql = sSql & " )"

    If OraDB.ExecuteSQL(sSql) < 1 Then
        'ロールバック
        OraDB.Rollback
        KODZ6_Insert_2 = False
        GoTo proc_exit
    End If

    KODZ6_Insert_2 = True
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sSql
    gErr.HandleError
    KODZ6_Insert_2 = False
    Resume proc_exit
End Function

'===============================================================================
'関数名     :ラベル実行プロセス取得処理
'機能       :ラベル属性情報テーブルからラベル実行プロセスを取得する。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :なし
'===============================================================================
Public Sub Judge_Process()
    Dim rs As OraDynaset
    Dim sSql As String

    On Error GoTo proc_err
    gErr.Push "c_cmzcLabel.cls -- Function Judge_Process"
    
    'ラベル属性情報テーブル(XODZ1)からラベル実行プロセス(PRSKBNZ1)を取得
    sSql = ""
    sSql = sSql & "SELECT"
    sSql = sSql & " PRSKBNZ1"                  'ラベル実行プロセス
    sSql = sSql & " FROM"
    sSql = sSql & " XODZ1"
    sSql = sSql & " WHERE LBLKEYZ1 = '" & m_strLabelId & "'"
    sSql = sSql & " AND LBLSUBZ1 = '" & m_strPrCdSub & "'"
    
    Set rs = OraDB.CreateDynaset(sSql, ORADYN_DEFAULT)
    
    '新旧プロセス区分を初期化
    m_procKubun = False
    
    If rs.RecordCount > 0 Then
        '取得したラベル実行プロセスが'1'の場合
        If Trim(rs(0).Value) = "1" Then
            '新旧プロセス区分をTrue(新プロセス)にする
            m_procKubun = True
        End If
    End If
    
    rs.Close
    
proc_exit:
    '終了
    gErr.Pop
    Exit Sub

proc_err:
    'エラーハンドラ
    Resume proc_exit
End Sub

'===============================================================================
'関数名     :桁数チェック処理（文字列）
'機能       :文字列データの文字数をチェックし、指定された桁数に合わせる。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :strData   IO   ラベルプリンタ印刷要求テーブル登録用文字列データ
'           :keta      I    各データの指定桁数
'戻り値     :strData or 桁数分のスペース
'===============================================================================
Private Function Str_Check(strData As Variant, keta As Integer) As Variant
    
    If IsNull(strData) Or strData = "" Then
        'データがNULL値もしくは空文字列の場合、指定桁数分スペースを入れる
        Str_Check = Space(keta)
    Else
        If Len(strData) = keta Then
            'データの文字数と指定桁数が同じ場合、そのまま返す
            Str_Check = strData
        ElseIf Len(strData) > keta Then
            'データの文字数が指定桁数より大きい場合、強制的に指定桁数で区切る
            Str_Check = Mid(strData, 1, keta)
        Else
            'データの文字数が指定桁数より小さい場合、不足分スペースを追加する
            Str_Check = strData & Space(keta - Len(strData))
        End If
    End If
        
End Function

'===============================================================================
'関数名     :桁数チェック処理（数値）
'機能       :数値データの桁数をチェックし、指定された桁数に合わせて文字列変換する。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :numData   IO   ラベルプリンタ印刷要求テーブル登録用数値データ
'           :keta      I    各データの指定桁数
'戻り値     :文字列変換されたnumData or 桁数分の'0'
'===============================================================================
Private Function Num_Check(numData As Variant, keta As Integer) As Variant
    
    If IsNull(numData) Then
        'データがNULL値の場合、指定桁数分'0'を入れる
        Num_Check = String(keta, "0")
    Else
        If Len(numData) = keta Then
            'データの桁数と指定桁数が同じ場合、文字列変換してそのまま返す
            Num_Check = Trim(Str(numData))
        ElseIf Len(numData) < keta Then
            'データの桁数が指定桁数より小さい場合、不足分先頭に'0'を追加する
            Num_Check = String(keta - Len(numData), "0") & Trim(Str(numData))
        End If
    End If
        
End Function

'===============================================================================
'関数名     :新ラベル種別格納処理（サンプル、払出ブロックラベル以外）
'機能       :旧ラベル種別に対応する新ラベル種別をプロパティにセットする。
'-------------------------------------------------------------------------------
'       日付        版      担当者           コメント
'作成   2008/11/05  1.00    SETsw kakeida
'更新
'-------------------------------------------------------------------------------
'引数　     :なし
'戻り値     :なし
'===============================================================================
Private Sub Set_LabelId()
        
    Select Case m_strEtcPrKind
        Case "01"                                          '01:インゴットラベル
            m_strLabelId = "L306"
        Case "02"                                          '02:トップ・ボトムラベル
            m_strLabelId = "L307"
        Case "03"                                          '03:新原料ラベル
            m_strLabelId = "L308"
        Case "04"                                          '04:精製原料ラベル(分類前)
            m_strLabelId = "L309"
        Case "05"                                          '05:精製原料ラベル(洗浄後)
            m_strLabelId = "L310"
        Case "06"                                          '06:クリスタルカタログラベル
            m_strLabelId = "L311"
        Case "07"                                          '07:ブロックラベル
            m_strLabelId = "L312"
        Case "08"                                          '08:ブロックラベル２
            m_strLabelId = "L313"
        Case "10"                                          '10:精製原料ラベル(分類前)
            m_strLabelId = "L314"
        Case "11"                                          '11:精製原料ラベル(洗浄前)
            m_strLabelId = "L315"
        Case "12"                                          '12:精製原料ラベル(再洗浄)
            m_strLabelId = "L316"
        Case "13"                                          '13:精製原料ラベル
            m_strLabelId = "L317"
        Case "14"                                          '14:電極材ラベル
            m_strLabelId = "L318"
'Add Start 2011/04/15 SMPK Nakamura FRSシステム化対応
        Case "16"                                          '16:FRS測定ラベル
            m_strLabelId = "L416"
'Add End 2011/04/15 SMPK Nakamura FRSシステム化対応
        Case "81"                                          '81:精製原料ラベル(発生工程)
            m_strLabelId = "L319"
        Case "82"                                          '82:精製原料ラベル(切断受入)
            m_strLabelId = "L320"
        Case "84"                                          '84:精製原料ラベル(切断払出以降共通)
            m_strLabelId = "L321"
        Case "83"                                          '83:ドーパント準備ラベル
            m_strLabelId = "L322"
        Case "85"                                          '85:ルツボラベル
            m_strLabelId = "L324"
    End Select
    
End Sub

'===============================================================================
'概要      :各ラベルプリンタ印刷要求テーブル登録用データを
'           プロパティとして設定＆取得
'説明      :文字列としての"NULL"はセットしてはいけない
'===============================================================================
Public Property Get StrQueDate() As Variant
    '戻り値 ： キュー日付

    StrQueDate = m_strQueDate
    
End Property

Public Property Let StrQueDate(ByVal strValue As Variant)
    'strValue : キュー日付
    
    m_strQueDate = strValue
    
End Property

Public Property Get StrReqKind() As Variant
    '戻り値 ： 印刷要求区分

    StrReqKind = m_strReqKind
    
End Property

Public Property Let StrReqKind(ByVal strValue As Variant)
    'strValue : 印刷要求区分
    
    m_strReqKind = strValue
    
End Property

Public Property Get StrPrintKind() As Variant
    '戻り値 ： 印刷種類

    StrPrintKind = m_strPrintKind
    
End Property

Public Property Let StrPrintKind(ByVal strValue As Variant)
    'strValue : 印刷種類
    
    m_strPrintKind = strValue
    
End Property

Public Property Get StrEndFlg() As Variant
    '戻り値 ： 完了区分

    StrEndFlg = m_strEndFlg
    
End Property

Public Property Let StrEndFlg(ByVal strValue As Variant)
    'strValue : 完了区分
    
    m_strEndFlg = strValue
    
End Property

Public Property Get StrStatus() As Variant
    '戻り値 ： 終了ステータス

    StrStatus = m_strStatus
    
End Property

Public Property Let StrStatus(ByVal strValue As Variant)
    'strValue : 終了ステータス
    
    m_strStatus = strValue
    
End Property

Public Property Get StrBlockIdUmu() As Variant
    '戻り値 ： ブロックID有無区分

    StrBlockIdUmu = m_strBlockIdUmu
    
End Property

Public Property Let StrBlockIdUmu(ByVal strValue As Variant)
    'strValue : ブロックID有無区分
    
    m_strBlockIdUmu = strValue
    
End Property

Public Property Get StrProcCode() As Variant
    '戻り値 ： 工程コード

    StrProcCode = m_strProcCode
    
End Property

Public Property Let StrProcCode(ByVal strValue As Variant)
    'strValue : 工程コード
    
    m_strProcCode = strValue
    
End Property

Public Property Get StrEtcPrKind() As Variant
    '戻り値 ： その他ラベル種類

    StrEtcPrKind = m_strEtcPrKind
    
End Property

Public Property Let StrEtcPrKind(ByVal strValue As Variant)
    'strValue : その他ラベル種類
    
    m_strEtcPrKind = strValue
    
End Property

Public Property Get StrCryNum() As Variant
    '戻り値 ： 結晶番号

    StrCryNum = m_strCryNum
    
End Property

Public Property Let StrCryNum(ByVal strValue As Variant)
    'strValue : 結晶番号
    
    m_strCryNum = strValue
    
End Property

Public Property Get NumIngotPos() As Variant
    '戻り値 ： 結晶内位置

    NumIngotPos = m_numIngotPos
    
End Property

Public Property Let NumIngotPos(ByVal strValue As Variant)
    'strValue : 結晶内位置
    
    m_numIngotPos = strValue
    
End Property

Public Property Get NumSmplNo() As Variant
    '戻り値 ： サンプル№

    NumSmplNo = m_numSmplNo
    
End Property

Public Property Let NumSmplNo(ByVal strValue As Variant)
    'strValue : サンプル№
    
    m_numSmplNo = strValue
    
End Property

Public Property Get StrMtrlNum() As Variant
    '戻り値 ： 原料番号

    StrMtrlNum = m_strMtrlNum
    
End Property

Public Property Let StrMtrlNum(ByVal strValue As Variant)
    'strValue : 原料番号
    
    m_strMtrlNum = strValue
    
End Property

Public Property Get StrSmtrlNum() As Variant
    '戻り値 ：精製原料番号

    StrSmtrlNum = m_strSmtrlNum
    
End Property

Public Property Let StrSmtrlNum(ByVal strValue As Variant)
    'strValue : 精製原料番号
    
    m_strSmtrlNum = strValue
    
End Property

Public Property Get StrBlockId() As Variant
    '戻り値 ：ブロックID

    StrBlockId = m_strBlockId
    
End Property

Public Property Let StrBlockId(ByVal strValue As Variant)
    'strValue : ブロックID
    
    m_strBlockId = strValue
    
End Property

Public Property Get StrHinban() As Variant
    '戻り値 ：品番

    StrHinban = m_strHinban
    
End Property

Public Property Let StrHinban(ByVal strValue As Variant)
    'strValue : 品番
    
    m_strHinban = strValue
    
End Property

Public Property Get NumRevNum() As Variant
    '戻り値 ：製品番号改訂番号

    NumRevNum = m_numRevNum
    
End Property

Public Property Let NumRevNum(ByVal strValue As Variant)
    'strValue : 製品番号改訂番号
    
    m_numRevNum = strValue
    
End Property

Public Property Get StrFactory() As Variant
    '戻り値 ：工場

    StrFactory = m_strFactory
    
End Property

Public Property Let StrFactory(ByVal strValue As Variant)
    'strValue : 工場
    
    m_strFactory = strValue
    
End Property

Public Property Get StrOpeCond() As Variant
    '戻り値 ：操業条件

    StrOpeCond = m_strOpeCond
    
End Property

Public Property Let StrOpeCond(ByVal strValue As Variant)
    'strValue : 操業条件
    
    m_strOpeCond = strValue
    
End Property

Public Property Get StrMachine() As Variant
    '戻り値 ：要求マシン名

    StrMachine = m_strMachine
    
End Property

Public Property Let StrMachine(ByVal strValue As Variant)
    'strValue : 要求マシン名
    
    m_strMachine = strValue
    
End Property

Public Property Get StrStaffId() As Variant
    '戻り値 ：要求担当者ID

    StrStaffId = m_strStaffId
    
End Property

Public Property Let StrStaffId(ByVal strValue As Variant)
    'strValue : 要求担当者ID
    
    m_strStaffId = strValue
    
End Property

Public Property Get StrRegDate() As Variant
    '戻り値 ：登録日付

    StrRegDate = m_strRegDate
    
End Property

Public Property Let StrRegDate(ByVal strValue As Variant)
    'strValue : 登録日付
    
    m_strRegDate = strValue
    
End Property

Public Property Get StrUpdDate() As Variant
    '戻り値 ：更新日付

    StrUpdDate = m_strUpdDate
    
End Property

Public Property Let StrUpdDate(ByVal strValue As Variant)
    'strValue : 更新日付
    
    m_strUpdDate = strValue
    
End Property

Public Property Get StrCryIndRs() As Variant
    '戻り値 ：結晶検査指示（Rs）

    StrCryIndRs = m_strCryIndRs
    
End Property

Public Property Let StrCryIndRs(ByVal strValue As Variant)
    'strValue : 結晶検査指示（Rs）
    
    m_strCryIndRs = strValue
    
End Property

Public Property Get StrCryIndOi() As Variant
    '戻り値 ：結晶検査指示（Oi）

    StrCryIndOi = m_strCryIndOi
    
End Property

Public Property Let StrCryIndOi(ByVal strValue As Variant)
    'strValue : 結晶検査指示（Oi）
    
    m_strCryIndOi = strValue
    
End Property

Public Property Get StrCryIndB1() As Variant
    '戻り値 ：結晶検査指示（B1）

    StrCryIndB1 = m_strCryIndB1
    
End Property

Public Property Let StrCryIndB1(ByVal strValue As Variant)
    'strValue : 結晶検査指示（B1）
    
    m_strCryIndB1 = strValue
    
End Property

Public Property Get StrCryIndB2() As Variant
    '戻り値 ：結晶検査指示（B2）

    StrCryIndB2 = m_strCryIndB2
    
End Property

Public Property Let StrCryIndB2(ByVal strValue As Variant)
    'strValue : 結晶検査指示（B2）
    
    m_strCryIndB2 = strValue
    
End Property

Public Property Get StrCryIndB3() As Variant
    '戻り値 ：結晶検査指示（B3）

    StrCryIndB3 = m_strCryIndB3
    
End Property

Public Property Let StrCryIndB3(ByVal strValue As Variant)
    'strValue : 結晶検査指示（B3）
    
    m_strCryIndB3 = strValue
    
End Property

Public Property Get StrCryIndL1() As Variant
    '戻り値 ：結晶検査指示（L1）

    StrCryIndL1 = m_strCryIndL1
    
End Property

Public Property Let StrCryIndL1(ByVal strValue As Variant)
    'strValue : 結晶検査指示（L1）
    
    m_strCryIndL1 = strValue
    
End Property

Public Property Get StrCryIndL2() As Variant
    '戻り値 ：結晶検査指示（L2）

    StrCryIndL2 = m_strCryIndL2
    
End Property

Public Property Let StrCryIndL2(ByVal strValue As Variant)
    'strValue : 結晶検査指示（L2）
    
    m_strCryIndL2 = strValue
    
End Property

Public Property Get StrCryIndL3() As Variant
    '戻り値 ：結晶検査指示（L3）

    StrCryIndL3 = m_strCryIndL3
    
End Property

Public Property Let StrCryIndL3(ByVal strValue As Variant)
    'strValue : 結晶検査指示（L3）
    
    m_strCryIndL3 = strValue
    
End Property

Public Property Get StrCryIndL4() As Variant
    '戻り値 ：結晶検査指示（L4）

    StrCryIndL4 = m_strCryIndL4
    
End Property

Public Property Let StrCryIndL4(ByVal strValue As Variant)
    'strValue : 結晶検査指示（L4）
    
    m_strCryIndL4 = strValue
    
End Property

Public Property Get StrCryIndCs() As Variant
    '戻り値 ：結晶検査指示（Cs）

    StrCryIndCs = m_strCryIndCs
    
End Property

Public Property Let StrCryIndCs(ByVal strValue As Variant)
    'strValue : 結晶検査指示（Cs）
    
    m_strCryIndCs = strValue
    
End Property

Public Property Get StrCryIndGD() As Variant
    '戻り値 ：結晶検査指示（GD）

    StrCryIndGD = m_strCryIndGD
    
End Property

Public Property Let StrCryIndGD(ByVal strValue As Variant)
    'strValue : 結晶検査指示（GD）
    
    m_strCryIndGD = strValue
    
End Property

Public Property Get StrCryIndT() As Variant
    '戻り値 ：結晶検査指示（T）

    StrCryIndT = m_strCryIndT
    
End Property

Public Property Let StrCryIndT(ByVal strValue As Variant)
    'strValue : 結晶検査指示（T）
    
    m_strCryIndT = strValue
    
End Property

Public Property Get StrCryIndEP() As Variant
    '戻り値 ：結晶検査指示（EPD）

    StrCryIndEP = m_strCryIndEP
    
End Property

Public Property Let StrCryIndEP(ByVal strValue As Variant)
    'strValue : 結晶検査指示（EPD）
    
    m_strCryIndEP = strValue
    
End Property

'>>>>> 2009/07/31追加 SETsw kubota ------------
Public Property Get StrCryIndX() As Variant
    '戻り値 ：結晶検査指示（X）

    StrCryIndX = m_strCryIndX
    
End Property

Public Property Let StrCryIndX(ByVal strValue As Variant)
    'strValue : 結晶検査指示（X）
    
    m_strCryIndX = strValue
    
End Property
'<<<<< 2009/07/31追加 SETsw kubota ------------

'>>>>> 2010/12/02追加 SETsw kubota ------------
'結晶検査指示(CO3)
Public Property Get StrCryIndCO3() As Variant
    StrCryIndCO3 = m_strCryIndCO3
End Property
Public Property Let StrCryIndCO3(ByVal strValue As Variant)
    m_strCryIndCO3 = strValue
End Property
'結晶検査指示(C)
Public Property Get StrCryIndC() As Variant
    StrCryIndC = m_strCryIndC
End Property
Public Property Let StrCryIndC(ByVal strValue As Variant)
    m_strCryIndC = strValue
End Property
'結晶検査指示(CJ)
Public Property Get StrCryIndCJ() As Variant
    StrCryIndCJ = m_strCryIndCJ
End Property
Public Property Let StrCryIndCJ(ByVal strValue As Variant)
    m_strCryIndCJ = strValue
End Property
'結晶検査指示(CJ2)
Public Property Get StrCryIndCJ2() As Variant
    StrCryIndCJ2 = m_strCryIndCJ2
End Property
Public Property Let StrCryIndCJ2(ByVal strValue As Variant)
    m_strCryIndCJ2 = strValue
End Property
'<<<<< 2010/12/02追加 SETsw kubota ------------

Public Property Get strPrKey01() As Variant
    '戻り値 ：帳票キーデータ1

    strPrKey01 = m_strPrKey01
    
End Property

Public Property Let strPrKey01(ByVal strValue As Variant)
    'strValue : 帳票キーデータ1
    
    m_strPrKey01 = strValue
    
End Property

Public Property Get StrLabelId() As Variant
    '戻り値 ：依頼帳票ラベル名

    StrLabelId = m_strLabelId
    
End Property

Public Property Let StrLabelId(ByVal strValue As Variant)
    'strValue : 依頼帳票ラベル名
    
    m_strLabelId = strValue
    
End Property

Public Property Get StrPrCdSub() As Variant
    '戻り値 ：ラベル連番

    StrPrCdSub = m_strPrCdSub
    
End Property

Public Property Let StrPrCdSub(ByVal strValue As Variant)
    'strValue : ラベル連番
    
    m_strPrCdSub = strValue
    
End Property

Public Property Get StrEFlgz6() As Variant
    '戻り値 ：終了フラグ
    
    StrEFlgz6 = m_strEFlgz6
    
End Property

Public Property Let StrEFlgz6(ByVal strValue As Variant)
    'strValue : 終了フラグ
    
    m_strEFlgz6 = strValue
    
End Property

Public Property Get StrStatz6() As Variant
    '戻り値 ：状態ステータス

    StrStatz6 = m_strStatz6
    
End Property

Public Property Let StrStatz6(ByVal strValue As Variant)
    'strValue : 状態ステータス
    
    m_strStatz6 = strValue
    
End Property

Public Property Get StrKariHon() As Variant
    '戻り値 ：仮本区分
    
    StrKariHon = m_strKariHon
    
End Property

Public Property Let StrKariHon(ByVal strValue As Variant)
    'strValue : 仮本区分
    
    m_strKariHon = strValue
    
End Property

Public Property Get ProcKubun() As Variant
    '戻り値 ：新旧プロセス区分
    
    ProcKubun = m_procKubun
    
End Property

Public Property Let ProcKubun(ByVal strValue As Variant)
    'strValue : 新旧プロセス区分
    
    m_procKubun = strValue
    
End Property
