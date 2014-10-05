Attribute VB_Name = "s_cmzclabel"
'**************************************************
'   結晶システム／バーコードラベル機能
'
'   プログラム名    : ラベル発行関数
'   ファイル名　    : s_cmzclabel.bas
'   作成者　　　    : JCE
'   作成日　　　    : 2001/08/03
'
'**************************************************

Option Explicit

'===============================================================================
'   リテラル値定義
'===============================================================================
'ORACLE Object for OLE 定数
Private Const ORADYN_DEFAULT = &H0&          'ダイナセットの初期設定パラメータ

'ラベル種類
Private Const cLBL_INGOT   As String = "01"  'インゴットラベル
Private Const cLBL_TOPBTM  As String = "02"  'トップ・ボトムラベル"
Private Const cLBL_NWMTRL  As String = "03"  '新原料ラベル"
Private Const cLBL_SBMTRL  As String = "04"  '精製原料ラベル（洗浄前）"
Private Const cLBL_SAMTRL  As String = "05"  '精製原料ラベル（洗浄後）"
Private Const cLBL_CRYCAT  As String = "06"  'クリスタルカタログラベル"
Private Const cLBL_BLOCK   As String = "07"  'ブロックラベル"
Private Const cLBL_KARI    As String = "15"  '準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ"
'Add Start 2011/04/15 SMPK Nakamura FRSシステム化対応
Private Const cLBL_FRS     As String = "16"  'FRS測定ラベル
'Add End 2011/04/15 SMPK Nakamura FRSシステム化対応

'===============================================================================
'   変数定義
'===============================================================================
'ラベルプリンタ要求テーブル項目
Private mdtmQueDate As String               'キュー日付
Private mstrReqKind As String               '印刷要求区分
Private mstrPrintKind As String             '印刷種類
Private mstrEndFlg As String                '完了区分
Private mstrStatus As String                '終了ステータス
Private mstrBlockIDUmu As String            'ブロックＩＤ有無区分
Private mstrProcCode As String              '工程コード
Private mstrEtcPrKind As String             'その他ラベル種類
Private mstrCryNum As String                '結晶番号
Private mintIngotPos As Integer             '結晶内位置
Private mintSmplNo As Long                  'サンプルＮｏ． Integer→Long 6桁対応 2007/05/28 SETsw kubota
Private mstrMtrlNum As String               '原料番号
Private mstrSmtrlNum As String              '精製原料番号
Private mstrBlockID As String               'ブロックＩＤ
Private mstrHinban As String                '品番
Private mintRevNum As Integer               '製品番号改訂番号
Private mstrFactry As String                '工場
Private mstrOpecond As String               '操業条件
Private mstrCryindrs As String              '結晶検査指示（Rs）
Private mstrCryIndoi As String              '結晶検査指示（Oi）
Private mstrCryIndb1 As String              '結晶検査指示（B1）
Private mstrCryIndb2 As String              '結晶検査指示（B2）
Private mstrCryIndb3 As String              '結晶検査指示（B3）
Private mstrCryIndl1 As String              '結晶検査指示（L1）
Private mstrCryIndl2 As String              '結晶検査指示（L2）
Private mstrCryIndl3 As String              '結晶検査指示（L3）
Private mstrCryIndl4 As String              '結晶検査指示（L4）
Private mstrCryIndcs As String              '結晶検査指示（Cs）
Private mstrCryIndgd As String              '結晶検査指示（GD）
Private mstrCryIndt As String               '結晶検査指示（T）
Private mstrCryIndep As String              '結晶検査指示（EPD）
Private mstrCryIndx As String               '結晶検査指示（X）  '2009/07/31追加 SETsw kubota
Private mstrCryIndco3 As String             '結晶検査指示（CO3）'2010/12/14追加 SETsw kubota
Private mstrCryIndc As String               '結晶検査指示（C）  '2010/12/14追加 SETsw kubota
Private mstrCryIndcj As String              '結晶検査指示（CJ） '2010/12/14追加 SETsw kubota
Private mstrCryIndcjlt As String            '結晶検査指示（CJLT） 'Add 2011/02/02 SMPK Miyata
Private mstrCryIndcj2 As String             '結晶検査指示（CJ2）'2010/12/14追加 SETsw kubota
Private mstrStaffID As String               '要求担当者ＩＤ
Private mstrMachine As String               '要求マシン名
Private mdtmRegDate As Date                 '登録日付
Private mdtmUpdDate As Date                 '更新日付

'===============================================================================
'   Windows API 定義
'===============================================================================
' コンピュータ名を取得
Private Declare Function GetComputerNameA Lib "kernel32" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'*****************************************************
' 関数名　 : サンプルラベル出力関数
' 目的説明 : 工程管理画面からサンプルラベルを出力する。
'
' 引数　　 :
'     strCryNum(i)   : 結晶番号
'     intIngotPos(i) : サンプル位置
'     intSmplNo(i)   : サンプルNo   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'     strProcCode(i) : 工程コード
'     strStaffID(i)  : 社員ID
'     strBlockID(i)  : ブロックID  2009/09/16 SETsw kubota
'
' 戻り値　 :  0 : 正常終了
' 　　　　　 -1 : 異常終了
'*****************************************************
Public Function samlabel(StrCryNum As String, _
                          intIngotPos As Integer, _
                          intSmplNo As Long, _
                          StrProcCode As String, _
                          StrStaffId As String _
                        , Optional ByVal StrBlockId As String = "" _
                          ) As Integer

    Dim objRec    As Object
    Dim strMsg    As String
    
    samlabel = -1
    
    strMsg = "結晶番号=" & StrCryNum & Chr(13) & _
             "サンプル位置=" & CStr(intIngotPos) & Chr(13) & _
             "サンプルNo=" & CStr(intSmplNo) & Chr(13) & _
             "工程コード=" & StrProcCode & Chr(13) & _
             "社員ID=" & StrStaffId

'    MsgBox (strMsg)
    
    ' 引数設定内容確認
    If StrCryNum = "" Then Exit Function
    If IsNull(intIngotPos) Then Exit Function
    If IsNull(intSmplNo) Then Exit Function
    If StrProcCode = "" Then Exit Function
    If StrStaffId = "" Then Exit Function
    If StrBlockId = "" Then Exit Function
    
    '結晶情報管理テーブル情報取得
    'If Not Select_VECME017(StrCryNum, intIngotPos, intSmplNo, objRec) Then Exit Function
    If Not Select_VECME017(StrBlockId, intIngotPos, intSmplNo, objRec) Then Exit Function

' 追加  2003/10/08 SystemBrain ===================> START
    If objRec.RecordCount = 0 Then
        samlabel = 0
        Exit Function
    End If
' 追加  2003/10/08 SystemBrain ===================> START

    'ラベルプリンタ要求情報設定
    mdtmQueDate = Format(Now, "yyyymmddhhmmss")         'キュー日付
    mstrReqKind = "0"                                   '印刷要求区分　 0:工程内からの出力
    mstrPrintKind = "0"                                 '印刷種類 　　　0:サンプルラベル
    mstrEndFlg = "0"                                    '完了区分　　　 0:印刷待ち
    mstrStatus = "0"                                    '終了ステータス 0:正常
    mstrBlockIDUmu = "0"                                'ブロックID有無区分(サンプルラベルでは未使用の為'0'を設定)
    mstrProcCode = Trim(StrProcCode)                    '工程コード
    mstrEtcPrKind = "00"                                'その他ラベル種類(サンプルラベルでは未使用の為'00'を設定)
    mstrCryNum = Trim(StrCryNum)                        '結晶番号
    mintIngotPos = intIngotPos                          '結晶内位置
    mintSmplNo = intSmplNo                              'サンプルNo
    mstrMtrlNum = "0"                                   '原料番号(サンプルラベルでは未使用の為'0'を設定)
    mstrSmtrlNum = "0"                                  '精製原料番号(サンプルラベルでは未使用の為'0'を設定)
    mstrBlockID = "0"                                   'ブロックID(サンプルラベルでは未使用の為'0'を設定)

    If Not objRec.EOF Then
' 修正  2003/09/26 SystemBrain ===================> START
'        mstrHinban = Trim(objRec.Fields!hinban)         '品番
'        mintRevNum = objRec.Fields!REVNUM               '製品番号改訂番号
'        mstrFactry = Trim(objRec.Fields!factory)        '工場
'        mstrOpecond = Trim(objRec.Fields!opecond)       '操業条件
'        mstrCryindrs = Trim(objRec.Fields!CRYINDRS)     '結晶検査指示（Rs）
'        mstrCryIndoi = Trim(objRec.Fields!CRYINDOI)     '結晶検査指示（Oi）
'        mstrCryIndb1 = Trim(objRec.Fields!CRYINDB1)     '結晶検査指示（B1）
'        mstrCryIndb2 = Trim(objRec.Fields!CRYINDB2)     '結晶検査指示（B2）
'        mstrCryIndb3 = Trim(objRec.Fields!CRYINDB3)     '結晶検査指示（B3）
'        mstrCryIndl1 = Trim(objRec.Fields!CRYINDL1)     '結晶検査指示（L1）
'        mstrCryIndl2 = Trim(objRec.Fields!CRYINDL2)     '結晶検査指示（L2）
'        mstrCryIndl3 = Trim(objRec.Fields!CRYINDL3)     '結晶検査指示（L3）
'        mstrCryIndl4 = Trim(objRec.Fields!CRYINDL4)     '結晶検査指示（L4）
'        mstrCryIndcs = Trim(objRec.Fields!CRYINDCS)     '結晶検査指示（Cs）
'        mstrCryIndgd = Trim(objRec.Fields!CRYINDGD)     '結晶検査指示（GD）
'        mstrCryIndt = Trim(objRec.Fields!CRYINDT)       '結晶検査指示（T）
'        mstrCryIndep = Trim(objRec.Fields!CRYINDEP)     '結晶検査指示（EPD）
        mstrHinban = Trim(objRec.Fields!HINBCS)         '品番
        mintRevNum = objRec.Fields!REVNUMCS             '製品番号改訂番号
        mstrFactry = Trim(objRec.Fields!FACTORYCS)      '工場
        mstrOpecond = Trim(objRec.Fields!OPECS)         '操業条件
        mstrCryindrs = Trim(objRec.Fields!CRYINDRSCS)   '結晶検査指示（Rs）
        mstrCryIndoi = Trim(objRec.Fields!CRYINDOICS)   '結晶検査指示（Oi）
        mstrCryIndb1 = Trim(objRec.Fields!CRYINDB1CS)   '結晶検査指示（B1）
        mstrCryIndb2 = Trim(objRec.Fields!CRYINDB2CS)   '結晶検査指示（B2）
        mstrCryIndb3 = Trim(objRec.Fields!CRYINDB3CS)   '結晶検査指示（B3）
        mstrCryIndl1 = Trim(objRec.Fields!CRYINDL1CS)   '結晶検査指示（L1）
        mstrCryIndl2 = Trim(objRec.Fields!CRYINDL2CS)   '結晶検査指示（L2）
        mstrCryIndl3 = Trim(objRec.Fields!CRYINDL3CS)   '結晶検査指示（L3）
        mstrCryIndl4 = Trim(objRec.Fields!CRYINDL4CS)   '結晶検査指示（L4）
        mstrCryIndcs = Trim(objRec.Fields!CRYINDCSCS)   '結晶検査指示（Cs）
        mstrCryIndgd = Trim(objRec.Fields!CRYINDGDCS)   '結晶検査指示（GD）
        mstrCryIndt = Trim(objRec.Fields!CRYINDTCS)     '結晶検査指示（T）
        mstrCryIndep = Trim(objRec.Fields!CRYINDEPCS)   '結晶検査指示（EPD）
        mstrCryIndx = Trim(objRec.Fields!CRYINDXCS)     '結晶検査指示（X）  '2009/07/31追加 SETsw kubota
' 修正  2003/09/26 SystemBrain ===================> END
        'Add Start 2011/02/02 SMPK Miyata
        mstrCryIndc = Trim(objRec.Fields!CRYINDCCS)         '結晶検査指示（C）
        mstrCryIndcj = Trim(objRec.Fields!CRYINDCJCS)       '結晶検査指示（CJ）
        mstrCryIndcjlt = Trim(objRec.Fields!CRYINDCJLTCS)   '結晶検査指示（CJLT）
        mstrCryIndcj2 = Trim(objRec.Fields!CRYINDCJ2CS)     '結晶検査指示（CJ2）
        'Add End   2011/02/02 SMPK Miyata

        ''C,CJ,CJ2追加対応 2010/12/14 SETsw kubota
        mstrCryIndco3 = Trim(objRec.Fields!CRYINDL4CS)  '結晶検査指示（CO3）
        'Del Start 2011/02/02 SMPK Miyata
        'mstrCryIndc = "0"
        'mstrCryIndcj = "0"
        'mstrCryIndcj2 = "0"
        'If mstrCryIndco3 = "1" Then
        '    '結晶検査指示（C）
        '    If objRec.Fields!HSXCHS = "H" _
        '    Or objRec.Fields!HSXCHS = "S" Then
        '        mstrCryIndc = "1"
        '    End If
        '    '結晶検査指示（CJ）
        '    If objRec.Fields!HSXCJHS = "H" _
        '    Or objRec.Fields!HSXCJHS = "S" Then
        '        mstrCryIndcj = "1"
        '    End If
        '    '結晶検査指示（CJ2）
        '    If objRec.Fields!HSXCJ2HS = "H" _
        '    Or objRec.Fields!HSXCJ2HS = "S" Then
        '        mstrCryIndcj2 = "1"
        '    End If
        'End If
        'Del End   2011/02/02 SMPK Miyata
    
    End If
    objRec.Close

    mstrStaffID = StrStaffId                            '要求担当者ID
    mstrMachine = StrConv(GetComputerName, vbUpperCase) '要求マシン名
    
    'ラベルプリンタ要求テーブル情報追加
    If Not Insert_TBCMC001() Then Exit Function
    
    samlabel = 0
    
End Function


'*****************************************************
' 関数名　 : その他ラベル出力関数
' 目的説明 : 工程管理画面からその他ラベルを出力する。
'
' 引数　　 :
'     strEtcPrKind(i) : その他帳票区分
'                       01: インゴットラベル
'                       02:トップ・ボトムラベル
'                       03:新原料ラベル
'                       04:精製原料ラベル(分類前)
'                       05:精製原料ラベル(洗浄後)
'                       06:クリスタルカタログラベル
'                       07:ブロックラベル
'準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ発行処理追加依頼　yakimura 2002.12.12 start
'                       15:準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ
'準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ発行処理追加依頼　yakimura 2002.12.12 start
'     strKey(i)       : キー項目
'                       (結晶番号／原料No.／ブロックID／精製原料No.)
'     strBlockID(i)   : ブロックID有無区分
'     strProcCode(i)  : 工程コード
'     strStaffID(i)   : 社員ID
'
' 戻り値　 :  0 : 正常終了
' 　　　　　 -1 : 異常終了
'*****************************************************
Public Function etclabel(StrEtcPrKind As String, _
                          strKey As String, _
                          StrBlockId As String, _
                          StrProcCode As String, _
                          StrStaffId As String) As Integer
    Dim strMsg    As String
    
    etclabel = -1
    
    strMsg = "その他帳票区分=" & StrEtcPrKind & Chr(13) & _
             "キー項目=" & strKey & Chr(13) & _
             "ブロックID有無区分=" & StrBlockId & Chr(13) & _
             "工程コード=" & StrProcCode & Chr(13) & _
             "社員ID=" & StrStaffId

'    MsgBox (strMsg)

    ' 引数確認
    If StrEtcPrKind = "" Then Exit Function
    If strKey = "" Then Exit Function
    If StrBlockId = "" Then Exit Function
    If StrProcCode = "" Then Exit Function
    If StrStaffId = "" Then Exit Function
    
    'ラベルプリンタ要求情報設定
    mdtmQueDate = Format(Now, "yyyymmddhhmmss")         'キュー日付
    mstrReqKind = "0"                                   '印刷要求区分　 0:工程内からの出力
    mstrPrintKind = "1"                                 '印刷種類 　　　1:その他ラベル
    mstrEndFlg = "0"                                    '完了区分　　　 0:印刷待ち
    mstrStatus = "0"                                    '終了ステータス 0:正常
    mstrProcCode = Trim(StrProcCode)                    '工程コード
    mstrEtcPrKind = Trim(StrEtcPrKind)                  'その他ラベル種類
                                                        '  01:インゴットラベル
                                                        '  02:トップ・ボトムラベル
                                                        '  03:新原料ラベル
                                                        '  04:精製原料ラベル(分類前)
                                                        '  05:精製原料ラベル(洗浄後)
                                                        '  06:クリスタルカタログラベル
                                                        '  07:ブロックラベル
    
    'インゴットラベル, トップ・ボトムラベル
    If mstrEtcPrKind = cLBL_INGOT Or mstrEtcPrKind = cLBL_TOPBTM Then
        mstrCryNum = Trim(strKey)                       '結晶番号
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = "0"                               'ブロックID(未使用の為'0'を設定)
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
    '新原料ラベル
    ElseIf mstrEtcPrKind = cLBL_NWMTRL Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = Trim(strKey)                      '原料番号
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = "0"                               'ブロックID(未使用の為'0'を設定)
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
    '精製原料ラベル(分類前), クリスタルカタログラベル
    ElseIf mstrEtcPrKind = cLBL_SBMTRL Or mstrEtcPrKind = cLBL_CRYCAT Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = Trim(strKey)                      'ブロックID
        mstrBlockIDUmu = Trim(StrBlockId)               'ブロックID有無区分  0:ブロックID有り 1:ブロックID無し
    '精製原料ラベル(洗浄後)
    ElseIf mstrEtcPrKind = cLBL_SAMTRL Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = Trim(strKey)                     '精製原料番号
        mstrBlockID = "0"                               'ブロックID(未使用の為'0'を設定)
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
    'ブロックラベル
    ElseIf mstrEtcPrKind = cLBL_BLOCK Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = Trim(strKey)                      'ブロックID
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
'Add Start 2011/04/15 SMPK Nakamura FRSシステム化対応
    'FRS測定ラベル
    ElseIf mstrEtcPrKind = cLBL_FRS Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = Trim(strKey)                      'ブロックID
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
'Add End 2011/04/15 SMPK Nakamura FRSシステム化対応
    End If
    
    mintIngotPos = 0                                    '結晶内位置(その他ラベルでは未使用の為'0'を設定)
    mintSmplNo = 0                                      'サンプルNo(その他ラベルでは未使用の為'0'を設定)
    mstrHinban = "0"                                    '品番
    mintRevNum = 0                                      '製品番号改訂番号
    mstrFactry = "0"                                    '工場
    mstrOpecond = "0"                                   '操業条件
    mstrCryindrs = "0"                                  '結晶検査指示（Rs）
    mstrCryIndoi = "0"                                  '結晶検査指示（Oi）
    mstrCryIndb1 = "0"                                  '結晶検査指示（B1）
    mstrCryIndb2 = "0"                                  '結晶検査指示（B2）
    mstrCryIndb3 = "0"                                  '結晶検査指示（B3）
    mstrCryIndl1 = "0"                                  '結晶検査指示（L1）
    mstrCryIndl2 = "0"                                  '結晶検査指示（L2）
    mstrCryIndl3 = "0"                                  '結晶検査指示（L3）
    mstrCryIndl4 = "0"                                  '結晶検査指示（L4）
    mstrCryIndcs = "0"                                  '結晶検査指示（Cs）
    mstrCryIndgd = "0"                                  '結晶検査指示（GD）
    mstrCryIndt = "0"                                   '結晶検査指示（T）
    mstrCryIndep = "0"                                  '結晶検査指示（EPD）
    mstrCryIndx = "0"                                   '結晶検査指示（X）
    mstrCryIndco3 = "0"                                 '結晶検査指示（CO3）
    mstrCryIndc = "0"                                   '結晶検査指示（C）
    mstrCryIndcj = "0"                                  '結晶検査指示（CJ）
    mstrCryIndcj2 = "0"                                 '結晶検査指示（CJ2）
    
    mstrStaffID = StrStaffId                            '要求担当者ID
    mstrMachine = StrConv(GetComputerName, vbUpperCase) '要求マシン名
    
'準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ発行処理追加依頼　yakimura 2002.12.12 start
    'ブロックラベル　，　準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ
    If mstrEtcPrKind = cLBL_KARI Then
        mstrCryNum = "0"                                '結晶番号(未使用の為'0'を設定)
        mstrMtrlNum = "0"                               '原料番号(未使用の為'0'を設定)
        mstrSmtrlNum = "0"                              '精製原料番号(未使用の為'0'を設定)
        mstrBlockID = Trim(StrBlockId)                  'ブロックID
        mstrHinban = Trim(strKey)                       '品番
        mstrBlockIDUmu = "0"                            'ブロックID有無区分(未使用の為'0'を設定)
    End If
'準備済ﾌﾞﾛｯｸﾗﾍﾞﾙ発行処理追加依頼　yakimura 2002.12.12 end
    
    'ラベルプリンタ要求テーブル情報追加
    If Not Insert_TBCMC001() Then Exit Function
    
    etclabel = 0

End Function

'*****************************************************
' 関数名　 : 結晶サンプル情報管理テーブル情報取得関数
' 目的説明 : 結晶サンプル情報管理テーブルから情報を取得する。
'
' 引数　　 :
'     strCryNum(i)   : 結晶番号
'     intIngotPos(i) : サンプル位置
'     intSmplNo(i)   : サンプルNo   Integer→Long サンプル№6桁対応 2007/05/28 SETsw kubota
'     objRec(o)      : 検索結果
'
' 戻り値　 : True  : 正常終了
' 　　　     False : 異常終了
'*****************************************************
Private Function Select_VECME017(StrCryNum As String, _
                                intIngotPos As Integer, _
                                intSmplNo As Long, _
                                objRec As Object) As Boolean
    Dim strSQL     As String
    Dim strErrMsg  As String   'エラー時のメッセージ

    Select_VECME017 = False

    'SQL作成
' 修正  2003/09/26 SystemBrain ===================> START
'    strSQL = "SELECT * FROM VECME017 "
'    strSQL = strSQL & "WHERE CRYNUM = '" & strCryNum & "' "
'    strSQL = strSQL & "AND INGOTPOS = " & intIngotPos & " "
'    strSQL = strSQL & "AND SMPLNO = " & intSmplNo & " "
'    strSQL = strSQL & "AND KTKBN = '0'"
    
    strSQL = "SELECT"
    strSQL = strSQL & " XTALCS"           ' 0:結晶番号
    strSQL = strSQL & ",INPOSCS"          ' 1:結晶内位置
    strSQL = strSQL & ",REPSMPLIDCS"      ' 2:代表サンプルID
    strSQL = strSQL & ",HINBCS"           ' 3:品番
    strSQL = strSQL & ",REVNUMCS"         ' 4:製品番号改訂番号
    strSQL = strSQL & ",FACTORYCS"        ' 5:工場
    strSQL = strSQL & ",OPECS"            ' 6:操業条件
    strSQL = strSQL & ",CRYINDRSCS"       ' 7:状態FLG(Rs)
    strSQL = strSQL & ",CRYINDOICS"       ' 8:状態FLG(Oi)
    strSQL = strSQL & ",CRYINDCSCS"       ' 9:状態FLG(Cs)
    strSQL = strSQL & ",CRYINDB1CS"       '10:状態FLG(B1)
    strSQL = strSQL & ",CRYINDB2CS"       '11:状態FLG(B2)
    strSQL = strSQL & ",CRYINDB3CS"       '12:状態FLG(B3)
    strSQL = strSQL & ",CRYINDL1CS"       '13:状態FLG(L1)
    strSQL = strSQL & ",CRYINDL2CS"       '14:状態FLG(L2)
    strSQL = strSQL & ",CRYINDL3CS"       '15:状態FLG(L3)
    strSQL = strSQL & ",CRYINDL4CS"       '16:状態FLG(L4)
    strSQL = strSQL & ",CRYINDGDCS"       '17:状態FLG(GD)
    strSQL = strSQL & ",CRYINDTCS"        '18:状態FLG(T)
    strSQL = strSQL & ",CRYINDEPCS"       '19:状態FLG(EPD)
    strSQL = strSQL & ",NVL(CRYINDXCS,'0') CRYINDXCS"     '20:状態FLG(X線)
    'Add Start 2011/02/02 SMPK Miyata
    strSQL = strSQL & ",NVL(CRYINDCCS,'0') CRYINDCCS"       '21:状態FLG(C)
    strSQL = strSQL & ",NVL(CRYINDCJCS,'0') CRYINDCJCS"     '22:状態FLG(CJ)
    strSQL = strSQL & ",NVL(CRYINDCJLTCS,'0') CRYINDCJLTCS" '23:状態FLG(CJLT)
    strSQL = strSQL & ",NVL(CRYINDCJ2CS,'0') CRYINDCJ2CS"   '24:状態FLG(CJ2)
    'Add End   2011/02/02 SMPK Miyata

    'Del Start 2011/02/02 SMPK Miyata
    ''C,CJ,CJ2対応 2010/12/14 SETsw kubota
    'strSQL = strSQL & ",NVL(E020.HSXCHS , ' ') HSXCHS"          '品SXL/C保証方法＿処
    'strSQL = strSQL & ",NVL(E020.HSXCJHS , ' ') HSXCJHS"        '品SXL/CJ保証方法＿処
    'strSQL = strSQL & ",NVL(E020.HSXCJ2HS , ' ') HSXCJ2HS"      '品SXL/CJ2保証方法＿処
    'Del End   2011/02/02 SMPK Miyata
    
    strSQL = strSQL & "  FROM XSDCS "
    strSQL = strSQL & "     , TBCME020 E020 "      '2010/12/14 SETsw kubota
    
''' 09/03/02 FAE)akiyama start
''    strSQL = strSQL & "WHERE XTALCS = '" & StrCryNum & "' "
'    strSQL = strSQL & "WHERE CRYNUMCS LIKE '" & Left(StrCryNum, 9) & "%' "
''' 09/03/02 FAE)akiyama end
    strSQL = strSQL & " WHERE CRYNUMCS = '" & StrCryNum & "' "
    strSQL = strSQL & " AND INPOSCS = " & intIngotPos & " "
    strSQL = strSQL & " AND REPSMPLIDCS = " & intSmplNo & " "
    strSQL = strSQL & " AND (CRYINDRSCS = '1'"
    strSQL = strSQL & " OR CRYINDOICS = '1'"
    strSQL = strSQL & " OR CRYINDCSCS = '1'"
    strSQL = strSQL & " OR CRYINDB1CS = '1'"
    strSQL = strSQL & " OR CRYINDB2CS = '1'"
    strSQL = strSQL & " OR CRYINDB3CS = '1'"
    strSQL = strSQL & " OR CRYINDL1CS = '1'"
    strSQL = strSQL & " OR CRYINDL2CS = '1'"
    strSQL = strSQL & " OR CRYINDL3CS = '1'"
    strSQL = strSQL & " OR CRYINDL4CS = '1'"
    strSQL = strSQL & " OR CRYINDGDCS = '1'"
    strSQL = strSQL & " OR CRYINDTCS = '1'"
    strSQL = strSQL & " OR CRYINDEPCS = '1'"
    strSQL = strSQL & " OR CRYINDXCS = '1'"
    'Add Start 2011/02/02 SMPK Miyata
    strSQL = strSQL & " OR CRYINDCCS = '1'"
    strSQL = strSQL & " OR CRYINDCJCS = '1'"
    strSQL = strSQL & " OR CRYINDCJLTCS = '1'"
    strSQL = strSQL & " OR CRYINDCJ2CS = '1'"
    'Add End   2011/02/02 SMPK Miyata
    strSQL = strSQL & " )"
'    strSQL = strSQL & "AND BLKKTFLAGCS = '0'"      'ﾌﾞﾛｯｸ確定ﾌﾗｸﾞは呼び元で判断する 2009/09/16 SETsw kubota
' 修正  2003/09/26 SystemBrain ===================> END
    
    'C,CJ,CJ2対応 2010/12/14 SETsw kubota
    strSQL = strSQL & " AND HINBCS = E020.HINBAN(+)"
    strSQL = strSQL & " AND REVNUMCS = E020.MNOREVNO(+)"
    strSQL = strSQL & " AND FACTORYCS = E020.FACTORY(+)"
    strSQL = strSQL & " AND OPECS = E020.OPECOND(+)"

Debug.Print strSQL
    'DBを検索する
    'SQL発行
    If Not Fun_Ora_Select((strSQL), objRec, True) Then
        Exit Function
    End If
    
' 削除(該当ﾃﾞｰﾀなしでもｴﾗｰとしないで無処理とする) 2003/10/08 SystemBrain ===================> START
'''''    If objRec.RecordCount <= 0 Then
'''''        '該当データなし
'''''        strErrMsg = GetMsgStr("ECRY7")
'''''        MsgBox strErrMsg, vbCritical
'''''        Exit Function
'''''    End If
' 削除(該当ﾃﾞｰﾀなしでもｴﾗｰとしないで無処理とする) 2003/10/08 SystemBrain ===================> END
    
    Select_VECME017 = True

End Function

'*****************************************************
' 関数名　 : ラベルプリンタ要求テーブル情報追加関数
' 目的説明 : ラベルプリンタ要求テーブルに情報を追加する。
'
' 引数　　 : なし
'
' 戻り値　 : True  : 正常終了
' 　　　     False : 異常終了
'*****************************************************
Private Function Insert_TBCMC001() As Boolean
    Dim strSQL     As String
    Dim strSQL2    As String
    Dim strErrMsg  As String   'エラー時のメッセージ
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    Dim labelCls    As c_cmzcLabel      'クラスモジュール用変数
    Set labelCls = New c_cmzcLabel      'クラスオブジェクト生成
    '<<<<< -------------------------------------------------------END
    
    Insert_TBCMC001 = False

    'ロールバック
    Call Fun_Ora_Rollback(False)

    'トランザクション開始
    Fun_Ora_BeginTransaction
    
    strSQL = "INSERT INTO TBCMC001 (": strSQL2 = " VALUES (":
    strSQL = strSQL & " QUEDATE":      strSQL2 = strSQL2 & " TO_DATE('" & mdtmQueDate & "','YYYYMMDDHH24MISS')"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrQueDate = mdtmQueDate                        'キュー日付をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REQKIND":      strSQL2 = strSQL2 & ",'" & mstrReqKind & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrReqKind = mstrReqKind                        '印刷要求区分をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",PRINTKIND":    strSQL2 = strSQL2 & ",'" & mstrPrintKind & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrPrintKind = mstrPrintKind                    '印刷種類をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",ENDFLG":       strSQL2 = strSQL2 & ",'" & mstrEndFlg & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrEndFlg = mstrEndFlg                          '完了区分をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",STATUS":       strSQL2 = strSQL2 & ",'" & mstrStatus & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrStatus = mstrStatus                          '終了ステータスをプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",BLOCKIDUMU":   strSQL2 = strSQL2 & ",'" & mstrBlockIDUmu & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrBlockIdUmu = mstrBlockIDUmu                   'ブロックID有無区分をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",PROCCODE":     strSQL2 = strSQL2 & ",'" & mstrProcCode & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrProcCode = mstrProcCode                       '工程コードをプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",ETCPRKIND":    strSQL2 = strSQL2 & ",'" & mstrEtcPrKind & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrEtcPrKind = mstrEtcPrKind                     'その他ラベル種類をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYNUM":       strSQL2 = strSQL2 & ",'" & mstrCryNum & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryNum = mstrCryNum                           '結晶番号をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",INGOTPOS":     strSQL2 = strSQL2 & "," & mintIngotPos
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.NumIngotPos = mintIngotPos                        '結晶内位置をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",SMPLNO":       strSQL2 = strSQL2 & "," & mintSmplNo
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.NumSmplNo = mintSmplNo                            'サンプル№をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",MTRLNUM":      strSQL2 = strSQL2 & ",'" & mstrMtrlNum & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrMtrlNum = mstrMtrlNum                          '原料番号をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",SMTRLNUM":     strSQL2 = strSQL2 & ",'" & mstrSmtrlNum & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrSmtrlNum = mstrSmtrlNum                        '精製原料番号をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",BLOCKID":      strSQL2 = strSQL2 & ",'" & mstrBlockID & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrBlockId = mstrBlockID                          'ブロックIDをプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",HINBAN":       strSQL2 = strSQL2 & ",'" & mstrHinban & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrHinban = mstrHinban                            '品番をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REVNUM":       strSQL2 = strSQL2 & "," & mintRevNum
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.NumRevNum = mintRevNum                            '製品番号改訂番号をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",FACTORY":      strSQL2 = strSQL2 & ",'" & mstrFactry & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrFactory = mstrFactry                           '工場をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",OPECOND":      strSQL2 = strSQL2 & ",'" & mstrOpecond & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrOpeCond = mstrOpecond                          '操業条件をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
' 修正  2003/09/26 SystemBrain ===================> START
    strSQL = strSQL & ",CRYINDRS":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryindrs = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndRs = IIf(mstrCryindrs = "1", "1", "0")   '結晶検査指示(Rs)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDOI":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndoi = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndOi = IIf(mstrCryIndoi = "1", "1", "0")   '結晶検査指示(Oi)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB1":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb1 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB1 = IIf(mstrCryIndb1 = "1", "1", "0")   '結晶検査指示(B1)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB2":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb2 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB2 = IIf(mstrCryIndb2 = "1", "1", "0")   '結晶検査指示(B2)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDB3":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndb3 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndB3 = IIf(mstrCryIndb3 = "1", "1", "0")   '結晶検査指示(B3)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL1":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl1 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL1 = IIf(mstrCryIndl1 = "1", "1", "0")   '結晶検査指示(L1)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL2":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl2 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL2 = IIf(mstrCryIndl2 = "1", "1", "0")   '結晶検査指示(L2)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL3":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl3 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndL3 = IIf(mstrCryIndl3 = "1", "1", "0")   '結晶検査指示(L3)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDL4":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndl4 = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
'L4をCO3に変更 2010/12/14 SETsw kubota
'    labelCls.StrCryIndL4 = IIf(mstrCryIndl4 = "1", "1", "0")   '結晶検査指示(L4)をプロパティにセット
    labelCls.StrCryIndL4 = "0"
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDCS":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndcs = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndCs = IIf(mstrCryIndcs = "1", "1", "0")   '結晶検査指示(Cs)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDGD":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndgd = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndGD = IIf(mstrCryIndgd = "1", "1", "0")   '結晶検査指示(GD)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDT":      strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndt = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndT = IIf(mstrCryIndt = "1", "1", "0")     '結晶検査指示(T)をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",CRYINDEP":     strSQL2 = strSQL2 & ",'" & IIf(mstrCryIndep = "1", "1", "0") & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrCryIndEP = IIf(mstrCryIndep = "1", "1", "0")   '結晶検査指示(EPD)をプロパティにセット
    '<<<<< -------------------------------------------------------END
' 修正  2003/09/26 SystemBrain ===================> END

    strSQL = strSQL & ",STAFFID":      strSQL2 = strSQL2 & ",'" & mstrStaffID & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrStaffId = mstrStaffID                          '要求担当者IDをプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",MACHINE":      strSQL2 = strSQL2 & ",'" & mstrMachine & "'"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrMachine = mstrMachine                          '要求マシン名をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",REGDATE":      strSQL2 = strSQL2 & ",SYSDATE"
    
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    If Not GetSysdate() Then GoTo proc_exit  'システム日付を取得
    
    labelCls.StrRegDate = Format(gsSysdate, "yyyymmddhhmmss")  '登録日付をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ",UPDDATE":      strSQL2 = strSQL2 & ",SYSDATE"
    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    labelCls.StrUpdDate = Format(gsSysdate, "yyyymmddhhmmss")  '更新日付をプロパティにセット
    '<<<<< -------------------------------------------------------END
    
    strSQL = strSQL & ")":             strSQL2 = strSQL2 & ")"

    '結晶検査指示(X)    2009/07/31追加 SETsw kubota
    labelCls.StrCryIndX = IIf(mstrCryIndx = "1", "1", "0")

    '結晶検査指示(C,CJ,CJ2)     2010/12/14追加 SETsw kubota
    labelCls.StrCryIndCO3 = IIf(mstrCryIndco3 = "1", "1", "0")  '結晶検査指示(CO3)
    labelCls.StrCryIndC = IIf(mstrCryIndc = "1", "1", "0")      '結晶検査指示(C)
    labelCls.StrCryIndCJ = IIf(mstrCryIndcj = "1", "1", "0")    '結晶検査指示(CJ)
    labelCls.StrCryIndCJ2 = IIf(mstrCryIndcj2 = "1", "1", "0")  '結晶検査指示(CJ2)

    '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
    '新旧プロセスのチェック
    labelCls.Label_Process_Check
    
    '新プロセスの場合
    If labelCls.ProcKubun = True Then
       'KODZ6に登録
        If labelCls.Label_Facade = False Then
            GoTo proc_exit
        End If
    Else
        '旧プロセスの場合、既存の処理(TBCMC001に登録)
        If Not Fun_Ora_Execute(strSQL & strSQL2) Then
            'ロールバック
            If Not Fun_Ora_Rollback(True) Then GoTo proc_exit
            GoTo proc_exit
        End If
    End If
    
    If Not Fun_Ora_Commit() Then GoTo proc_exit
    '<<<<< -------------------------------------------------------END
        
'    If Not Fun_Ora_Execute(strSQL & strSQL2) Then
'        'ロールバック
'        If Not Fun_Ora_Rollback(True) Then Exit Function
'        Exit Function
'    End If
'
'    If Not Fun_Ora_Commit() Then Exit Function
        
    '1秒間待ち
    Sleep (1000)

    Insert_TBCMC001 = True
    
'>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
proc_exit:
    Set labelCls = Nothing              'クラスオブジェクト解放
'<<<<< -------------------------------------------------------END

End Function

'Oracle関連
'===============================================================================
'関数名     :SELECT文を実行する Fun_Ora_Select
'機能       :接続先に対して、SELECT文を実行します。
'-------------------------------------------------------------------------------
'       日付        版      担当者      コメント
'作成   2001/08/01  1.00    JCE
'更新
'-------------------------------------------------------------------------------
'引数　     :sPrmSql     （発行ＳＱＬ）
'　　　     :bPrmOutMsg  （メッセージ出力フラグ(デフォルト:True){True:出力する,False:出力しない}）
'戻り値     :成功 True、エラー False
'　　　     :sPrmRslt    （ＤＢ取得内容）
'===============================================================================
Private Function Fun_Ora_Select(sPrmSql As String, sPrmRslt As OraDynaset, _
                            Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     'マウスポインター格納
    Dim strErrMsg As String   'エラー時のメッセージ

    Fun_Ora_Select = False
    
    'マウスポインタ状態の保存
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    '結果セットの作成
    Set sPrmRslt = OraDB.DBCreateDynaset(sPrmSql, ORADYN_DEFAULT)

    Fun_Ora_Select = True
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元

Exit Function

'ＤＢ関連エラー（処理中断）
ErrSub:
    'エラーメッセージ生成
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    'エラーメッセージ出力（指定時のみ出力）
    If bPrmOutMsg Then
        MsgBox "エラーが発生しました" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元
End Function

'===============================================================================
'関数名     :直接SQL発行処理 Fun_Ora_Execute
'機能       :接続先に対して、SELECT以外のSQLを実行します。
'-------------------------------------------------------------------------------
'       日付        版      担当者      コメント
'作成   2001/08/01  1.00    JCE
'更新
'-------------------------------------------------------------------------------
'引数　     :sPrmSql     （発行ＳＱＬ）
'　　　     :bPrmOutMsg  （メッセージ出力フラグ(デフォルト:True){True:出力する,False:出力しない}）
'戻り値     :成功 True、エラー False
'　　　     :lPrmCnt     （処理件数）
'===============================================================================
Private Function Fun_Ora_Execute(sPrmSql As String, Optional lPrmCnt As Long = 0, _
                            Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     'マウスポインター格納
    Dim strErrMsg As String   'エラー時のメッセージ
    
    Fun_Ora_Execute = False
    
    'マウスポインタ状態の保存
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub

    lPrmCnt = 0     '更新件数を初期化

    'ＳＱＬの発行
    sPrmSql = Fun_Empty_To_Null(sPrmSql)   'SQL文中の「,'',」や「,,」を、「,Null,」に置き換えます
    lPrmCnt = OraDB.DbExecuteSQL(sPrmSql)

    Fun_Ora_Execute = True
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元

Exit Function

'ＤＢ関連エラー（処理中断）
ErrSub:
    'エラーメッセージ生成
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    'エラーメッセージ出力（指定時のみ出力）
    If bPrmOutMsg Then
        MsgBox "エラーが発生しました" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元
End Function

'===============================================================================
'関数名     :トランザクション開始処理 Fun_Ora_BeginTransaction
'更新
'-------------------------------------------------------------------------------
'       日付        版      担当者      コメント
'作成   2001/08/01  1.00    JCE
'更新
'-------------------------------------------------------------------------------
'引数　     :bPrmOutMsg  （メッセージ出力フラグ(デフォルト:True){True:出力する,False:出力しない}）
'
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function Fun_Ora_BeginTransaction(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     'マウスポインター格納
    Dim strErrMsg As String   'エラー時のメッセージ

    Fun_Ora_BeginTransaction = False
    
    'マウスポインタ状態の保存
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub

    'ＳＱＬの発行
    Call OraSess.DbBeginTrans

    Fun_Ora_BeginTransaction = True
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元

Exit Function

'ＤＢ関連エラー（処理中断）
ErrSub:
    'エラーメッセージ生成
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    'エラーメッセージ出力（指定時のみ出力）
    If bPrmOutMsg Then
        MsgBox "エラーが発生しました" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元
End Function

'===============================================================================
'関数名     :トランザクション正常終了処理 Fun_Ora_Commit
'機能       :lPrmConNoのトランザクションをコミットします。
'-------------------------------------------------------------------------------
'       日付        版      担当者      コメント
'作成   2001/08/01  1.00    JCE
'更新
'-------------------------------------------------------------------------------
'引数　     :bPrmOutMsg  （メッセージ出力フラグ(デフォルト:True){True:出力する,False:出力しない}）
'
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function Fun_Ora_Commit(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     'マウスポインター格納
    Dim strErrMsg As String   'エラー時のメッセージ

    Fun_Ora_Commit = False
    
    'マウスポインタ状態の保存
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    'コミットの発行
    Call OraSess.DbCommitTrans

    Fun_Ora_Commit = True
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元

Exit Function

'ＤＢ関連エラー（処理中断）
ErrSub:
    'エラーメッセージ生成
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    'エラーメッセージ出力（指定時のみ出力）
    If bPrmOutMsg Then
        MsgBox "エラーが発生しました" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元
End Function

'===============================================================================
'関数名     :トランザクション異常終了処理 Fun_Ora_Rollback
'機能       :lPrmConNoのトランザクションをロールバックします。
'-------------------------------------------------------------------------------
'       日付        版      担当者      コメント
'作成   2001/08/01  1.00    JCE
'更新
'-------------------------------------------------------------------------------
'引数　     :bPrmOutMsg  （メッセージ出力フラグ(デフォルト:True){True:出力する,False:出力しない}）
'
'戻り値     :成功 True、エラー False
'===============================================================================
Private Function Fun_Ora_Rollback(Optional bPrmOutMsg As Boolean = True) As Boolean

    Dim lSv_Mouse As Long     'マウスポインター格納
    Dim strErrMsg As String   'エラー時のメッセージ
    
    Fun_Ora_Rollback = False
    
    'マウスポインタ状態の保存
    lSv_Mouse = Screen.MousePointer
    Screen.MousePointer = vbHourglass

    On Error GoTo ErrSub
    
    'ロールバックの発行
    Call OraSess.DbRollback

    Fun_Ora_Rollback = True
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元

Exit Function

'ＤＢ関連エラー（処理中断）
ErrSub:
    'エラーメッセージ生成
    strErrMsg = ""
    strErrMsg = strErrMsg & "ErrorCode:" & Err.Number
    strErrMsg = strErrMsg & "  Message:" & Err.Description
    'エラーメッセージ出力（指定時のみ出力）
    If bPrmOutMsg Then
        MsgBox "エラーが発生しました" & vbCrLf & strErrMsg, vbCritical
    End If
    Screen.MousePointer = lSv_Mouse     'マウスポインタ状態の復元
End Function

'関数名     :Null置換 Fun_Empty_To_Null
'機能       :カンマ区切りの各項目で、数値項目、文字項目が空であるものを、
'            'null'に置き換えます。
Private Function Fun_Empty_To_Null(sPrmStr As String) As String
    Dim lCnt As Long
    Dim lWkPos As Long
    Dim sWkStr As String
    
    '空文字
    Do Until InStr(sPrmStr, "''") = 0
        lWkPos = InStr(sPrmStr, "''")
        sWkStr = Left$(sPrmStr, lWkPos - 1) & "null" & Mid$(sPrmStr, lWkPos + 2)
        sPrmStr = sWkStr
    Loop
    
    '空数値
    Do Until InStr(sPrmStr, ",,") = 0
        lWkPos = InStr(sPrmStr, ",,")
        sWkStr = Left$(sPrmStr, lWkPos - 1) & ",null," & Mid$(sPrmStr, lWkPos + 2)
        sPrmStr = sWkStr
    Loop

    Fun_Empty_To_Null = sPrmStr

End Function

'コンピュータ名を取得
Private Function GetComputerName() As String
    'コンピュータ名を取得
    Dim sBuffer As String * 255
    
    'バッファのクリア
    sBuffer = Space(255)
    
    'APIコール
    If GetComputerNameA(sBuffer, 255) <> 0 Then
        '取得できた場合
        GetComputerName = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    Else
        '取得できなかった場合
        GetComputerName = "ERROR"
    End If
    
End Function

'>>>>>>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -----------------START
'>>>>>>>>>> mdlCommon.basのIns_TBCMC001_New関数をs_cmzclabel.basに移動 ------
    ' @(f)
    '
    ' 機能    : ラベル発行用ﾓｼﾞｭｰﾙ
    '
    ' 返り値  : True:成功 False:失敗
    '
    ' 引き数  : after:次工程コード
    '           cyoku:直区分
    '
    ' 機能説明:　精製原料候補発生時、精製原料ラベルを発行する
    '
    ' 備考    :
    '       引数：
    '           sProcCode   工程ｺｰﾄﾞ
    '           sEtcPrKind  その他ﾗﾍﾞﾙ種類
    '           sStaffID    要求担当者
    '           sPrKey01    帳票ｷｰﾃﾞｰﾀ1
    '           sSysdate    ｷｭｰ日付
    '           sRegDate    登録日付　※登録日付はPKの為、一回の処理で複数件登録する場合、
    '                                   呼び出し元で1秒ずらす等の制御が必要
    '       使用ﾌﾟﾛｸﾞﾗﾑ：
    '           cmbc008     ｸﾘｽﾀﾙｶﾀﾛｸﾞ検索格上げ
    '           cmbc030     結晶総合判定
    '           cmbc018     切断・ｻﾝﾌﾟﾙ指示照会
    '
    Public Function Ins_TBCMC001_New(sProcCode As String, sEtcPrKind As String, sStaffID As String, sPrKey01 As String, sSysDate As String) As Boolean
        Dim sSql      As String       'SQL文格納
        Dim iRet    As Integer        'データ追加数
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        Dim labelCls    As c_cmzcLabel      'クラスモジュール用変数
        Set labelCls = New c_cmzcLabel      'クラスオブジェクト生成
        '<<<<< -------------------------------------------------------END
                
    'エラーハンドラ
    On Error GoTo ErrHand

        '戻り値設定
        Ins_TBCMC001_New = False
            
        'ｺﾝﾋﾟｭｰﾀ名設定
        gsCompName = GetCompName
        
        '登録用ｸｴﾘｰ設定
        sSql = ""
        sSql = sSql & "insert into tbcmc001(" & vbLf    ''
        sSql = sSql & "                 quedate                     " & vbLf    ''キュー日付
        sSql = sSql & "                 ,reqkind                    " & vbLf    ''印刷要求区分
        sSql = sSql & "                 ,printkind                  " & vbLf    ''印刷種類
        sSql = sSql & "                 ,endflg                     " & vbLf    ''完了区分
        sSql = sSql & "                 ,status                     " & vbLf    ''終了ステータス
        sSql = sSql & "                 ,blockidumu                 " & vbLf    ''ブロックID有無区分
        sSql = sSql & "                 ,proccode                   " & vbLf    ''工程コード
        sSql = sSql & "                 ,etcprkind                  " & vbLf    ''その他ラベル種類
        sSql = sSql & "                 ,crynum                     " & vbLf    ''結晶番号
        sSql = sSql & "                 ,ingotpos                   " & vbLf    ''結晶内位置
        sSql = sSql & "                 ,smplno                     " & vbLf    ''サンプルNo
        sSql = sSql & "                 ,mtrlnum                    " & vbLf    ''原料番号
        sSql = sSql & "                 ,smtrlnum                   " & vbLf    ''精製原料番号
        sSql = sSql & "                 ,blockid                    " & vbLf    ''ブロックID
        sSql = sSql & "                 ,hinban                     " & vbLf    ''品番
        sSql = sSql & "                 ,revnum                     " & vbLf    ''製品番号改定番号
        sSql = sSql & "                 ,factory                    " & vbLf    ''工場
        sSql = sSql & "                 ,opecond                    " & vbLf    ''操業条件
        sSql = sSql & "                 ,cryindrs                   " & vbLf    ''結晶検査指示(Rs)
        sSql = sSql & "                 ,cryindoi                   " & vbLf    ''結晶検査指示(Oi)
        sSql = sSql & "                 ,cryindb1                   " & vbLf    ''結晶検査指示(B1)
        sSql = sSql & "                 ,cryindb2                   " & vbLf    ''結晶検査指示(B2)
        sSql = sSql & "                 ,cryindb3                   " & vbLf    ''結晶検査指示(B3)
        sSql = sSql & "                 ,cryindl1                   " & vbLf    ''結晶検査指示(L1)
        sSql = sSql & "                 ,cryindl2                   " & vbLf    ''結晶検査指示(L2)
        sSql = sSql & "                 ,cryindl3                   " & vbLf    ''結晶検査指示(L3)
        sSql = sSql & "                 ,cryindl4                   " & vbLf    ''結晶検査指示(L4)
        sSql = sSql & "                 ,cryindcs                   " & vbLf    ''結晶検査指示(Cs)
        sSql = sSql & "                 ,cryindgd                   " & vbLf    ''結晶検査指示(Gd)
        sSql = sSql & "                 ,cryindt                    " & vbLf    ''結晶検査指示(T)
        sSql = sSql & "                 ,cryindep                   " & vbLf    ''結晶検査指示(EPD)
        sSql = sSql & "                 ,staffid                    " & vbLf    ''要求担当者
        sSql = sSql & "                 ,machine                    " & vbLf    ''要求マシン名
        sSql = sSql & "                 ,regdate                    " & vbLf    ''登録日付
        sSql = sSql & "                 ,upddate                    " & vbLf    ''更新日付
        sSql = sSql & "                 ,prkey01                     " & vbLf    ''帳票キーデータ１
        sSql = sSql & "     )                                       " & vbLf
        sSql = sSql & "values(                                      " & vbLf
        sSql = sSql & "                 to_date('" & sSysDate & "','yyyy/mm/dd hh24:mi:ss')       " & vbLf    ''キュー日付
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrQueDate = Format(sSysDate, "yyyymmddhhmmss")  'キュー日付をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''印刷要求区分
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrReqKind = "0"                                 '印刷要求区分をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'1'                        " & vbLf    ''印刷種類
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrPrintKind = "1"                               '印刷種類をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''完了区分
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrEndFlg = "0"                                  '完了区分をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''終了ステータス
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrStatus = "0"                                  '終了ステータスをプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'0'                        " & vbLf    ''ブロックID有無区分
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrBlockIdUmu = "0"                              'ブロックID有無区分をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sProcCode & "'        " & vbLf    ''工程コード
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrProcCode = sProcCode                          '工程コードをプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sEtcPrKind & "'       " & vbLf    ''その他ラベル種類
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrEtcPrKind = sEtcPrKind                        'その他ラベル種類をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶番号
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryNum = Null                                 '結晶番号をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶内位置
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.NumIngotPos = Null                               '結晶内位置をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''サンプルNo
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.NumSmplNo = Null                                 'サンプル№をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''原料番号
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrMtrlNum = Null                                '原料番号をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''精製原料番号
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrSmtrlNum = Null                               '精製原料番号をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''ブロックID
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrBlockId = Null                                'ブロックIDをプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''品番
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrHinban = Null                                 '品番をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''製品番号改定番号
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.NumRevNum = Null                                 '製品番号改訂番号をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''工場
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrFactory = Null                                '工場をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''操業条件
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrOpeCond = Null                                '操業条件をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Rs)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndRs = Null                               '結晶検査指示(Rs)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Oi)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndOi = Null                               '結晶検査指示(Oi)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B1)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB1 = Null                               '結晶検査指示(B1)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B2)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB2 = Null                               '結晶検査指示(B2)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(B3)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndB3 = Null                               '結晶検査指示(B3)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L1)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL1 = Null                               '結晶検査指示(L1)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L2)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL2 = Null                               '結晶検査指示(L2)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L3)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL3 = Null                               '結晶検査指示(L3)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(L4)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndL4 = Null                               '結晶検査指示(L4)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Cs)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndCs = Null                               '結晶検査指示(Cs)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Gd)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndGD = Null                               '結晶検査指示(GD)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(T)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndT = Null                                '結晶検査指示(T)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,null                       " & vbLf    ''結晶検査指示(Epd)
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrCryIndEP = Null                               '結晶検査指示(EPD)をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sStaffID & "'         " & vbLf    ''要求担当者名
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrStaffId = sStaffID                            '要求担当者IDをプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & gsCompName & "'       " & vbLf    ''
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrMachine = gsCompName                          '要求マシン名をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''登録日付
        
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        If Not GetSysdate() Then GoTo proc_exit  'システム日付を取得
        
        labelCls.StrRegDate = Format(gsSysdate, "yyyymmddhhmmss") '登録日付をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,SYSDATE                    " & vbLf    ''更新日付
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.StrUpdDate = Format(gsSysdate, "yyyymmddhhmmss") '更新日付をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                 ,'" & sPrKey01 & "'         " & vbLf    ''帳票キーデータ１
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        labelCls.strPrKey01 = sPrKey01                            '帳票キーデータ1をプロパティにセット
        '<<<<< -------------------------------------------------------END
        
        sSql = sSql & "                             )               " & vbLf

        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        '新旧プロセスのチェック
        labelCls.Label_Process_Check
        
        '新プロセスの場合
        If labelCls.ProcKubun = True Then
            'KODZ6に登録
            If labelCls.Label_Facade = False Then
                Call MsgOut(100, sSql, ERR_DISP_LOG, "KODZ6")
                GoTo proc_exit
            End If
        Else
            '旧プロセスの場合､既存の処理 (TBCMC001に登録)
            iRet = SqlExec2(sSql)
            If iRet < 0 Then
                Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
                GoTo proc_exit
            End If
        End If
        '<<<<< -------------------------------------------------------END
        
        '実行
'        iRet = SqlExec2(sSql)
'
'        If iRet < 0 Then
'            Call MsgOut(100, sSql, ERR_DISP_LOG, "TBCMC001")
'            Exit Function
'        End If
        
        '戻り値設定
        Ins_TBCMC001_New = True
        
'>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
proc_exit:
        Set labelCls = Nothing              'クラスオブジェクト解放
        Exit Function
'<<<<< -------------------------------------------------------END
           
    'エラー時
ErrHand:
        ''ｴﾗｰ
        Call MsgOut(100, "", ERR_DISP_LOG, "TBCMC001")
        '>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -------START
        Resume proc_exit
        '<<<<< -------------------------------------------------------END
    End Function
'>>>>>>>>>> ラベル印刷方式統合対応 2008/11/07 SETsw kakeida -----------------END
