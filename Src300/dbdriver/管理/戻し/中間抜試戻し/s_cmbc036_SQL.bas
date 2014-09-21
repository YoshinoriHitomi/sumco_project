Attribute VB_Name = "s_cmbc036_SQL"
    Option Explicit

' 抜試変更指示

' ブロック情報
Public Type typ_BlkInf3
    BLOCKID     As String * 12          ' ブロックID
    LENGTH      As Integer              ' 長さ
    REALLEN     As Integer              ' 実長さ
    NOWPROC     As String * 5           ' 現在工程
    DELFLG      As String * 1           ' 削除区分
    COF         As type_Coefficient     ' 偏析係数計算
    '--- 野村追加：Disp()では初期値としてブロックの０～実長さを設定
    TOPPOS              As Integer          ' ブロックの最初のウェハ位置
    BOTPOS              As Integer          ' ブロックの最後のウェハ位置
End Type

' 欠落ウェハー情報
Public Type typ_LackWaf
    BLOCKID             As String * 12      ' ブロックID
    WAFERNO             As Integer          ' ウェハー連番
    WAFERTO             As Integer          ' ウェハー連番(to)
    TOP_POS             As Double           ' ウェハー開始位置'2002/02/27 S.Sano
    TAIL_POS            As Double           ' ウェハー終了位置'2002/02/27 S.Sano
    REJCAT              As String * 1       ' 欠落理由
    ALLSCRAP            As String * 1       ' 全数スクラップ
End Type

'2002/09/11 ADD hitec)N.MATSUMOTO Start
Private HSXCTCEN        As Double           ' 品ＳＸ結晶面傾縦中心
Private HSXCYCEN        As Double           ' 品ＳＸ結晶面傾横中心
'WF枚数計算用のパラメータ
Private SEEDDEG         As Integer          ' SEED傾き
Private Loss0           As Integer          ' 傾き差0度のときの傾きロス
Private Loss4           As Integer          ' 傾き差4度のときの傾きロス
Private Mlt4            As Double           ' 傾き差4度の時の係数
Private Pitch           As Double           ' ワイヤソーメインローラピッチ

'ブロック管理
Public Type typ_cmkc001f_Block
    'E040 ブロック管理
    INGOTPOS            As Integer          ' 結晶内開始位置
    LENGTH              As Integer          ' 長さ
    REALLEN             As Integer          ' 実長さ
    KRPROCCD            As String * 5       ' 現在管理工程
    NOWPROC             As String * 5       ' 現在工程
    LPKRPROCCD          As String * 5       ' 最終通過管理工程
    LASTPASS            As String * 5       ' 最終通過工程
    DELCLS              As String * 1       ' 削除区分
    RSTATCLS            As String * 1       ' 流動状態区分
    LSTATCLS            As String * 1       ' 最終状態区分 */
    'E037 結晶情報管理
    SEED                As String           'SEED
End Type

'仕様取得用
Public Type typ_cmkc001f_Disp
    '品番管理
    hinban              As String * 8       ' 品番
    INGOTPOS            As Integer          ' 結晶内開始位置
    REVNUM              As Integer          ' 製品番号改訂番号
    factory             As String * 1       ' 工場
    opecond             As String * 1       ' 操業条件
    LENGTH              As Integer          ' 長さ
    '製品仕様SXLデータ
    HSXD1CEN            As Double           ' 品ＳＸ直径１中心
    HSXRMIN             As Double           ' 品ＳＸ比抵抗下限
    HSXRMAX             As Double           ' 品ＳＸ比抵抗上限
    HSXRMBNP            As Double           ' 品ＳＸ比抵抗面内分布
    HSXRHWYS            As String * 1       ' 品ＳＸ比抵抗保証方法＿処
    HSXONMIN            As Double           ' 品ＳＸ酸素濃度下限
    HSXONMAX            As Double           ' 品ＳＸ酸素濃度上限
    HSXONMBP            As Double           ' 品ＳＸ酸素濃度面内分布
    HSXONHWS            As String * 1       ' 品ＳＸ酸素濃度保証方法＿処
    HSXCNMIN            As Double           ' 品ＳＸ炭素濃度下限
    HSXCNMAX            As Double           ' 品ＳＸ炭素濃度上限
    HSXCNHWS            As String * 1       ' 品ＳＸ炭素濃度保証方法＿処
    HSXTMMAX            As Double           ' 品ＳＸ転位密度上限         項目追加，修正対応 2003.05.20 yakimura
    HSXBMnAN(1 To 3)    As Double           ' 品ＳＸＢＭＤn 平均下限
    HSXBMnAX(1 To 3)    As Double           ' 品ＳＸＢＭＤn 平均上限
    HSXBMnHS(1 To 3)    As String * 1       ' 品ＳＸＢＭＤn 保証方法＿処
    HSXOFnAX(1 To 4)    As Double           ' 品ＳＸＯＳＦn平均上限
    HSXOFnMX(1 To 4)    As Double           ' 品ＳＸＯＳＦn上限
    HSXOFnHS(1 To 4)    As String * 1       ' 品ＳＸＯＳＦn 保証方法＿処
    HSXDENMX            As Integer          ' 品ＳＸＤｅｎ上限
    HSXDENMN            As Integer          ' 品ＳＸＤｅｎ下限
    HSXDENHS            As String * 1       ' 品ＳＸＤｅｎ保証方法＿処
    HSXDVDMX            As Integer          ' 品ＳＸＤＶＤ２上限
    HSXDVDMN            As Integer          ' 品ＳＸＤＶＤ２下限
    HSXDVDHS            As String * 1       ' 品ＳＸＤＶＤ２保証方法＿処
    HSXLDLMX            As Integer          ' 品ＳＸＬ／ＤＬ上限
    HSXLDLMN            As Integer          ' 品ＳＸＬ／ＤＬ下限
    HSXLDLHS            As String * 1       ' 品ＳＸＬ／ＤＬ保証方法＿処
    HSXLTMIN            As Integer          ' 品ＳＸＬタイム下限
    HSXLTMAX            As Integer          ' 品ＳＸＬタイム上限
    HSXLTHWS            As String * 1       ' 品ＳＸＬタイム保証方法＿処
    HSXDPDIR            As String * 2       ' 品ＳＸ溝位置方位
    HSXDPDRC            As String * 1       ' 品ＳＸ溝位置方向
    HSXDWMIN            As Double           ' 品ＳＸ溝巾下限
    HSXDWMAX            As Double           ' 品ＳＸ溝巾上限
    HSXDDMIN            As Double           ' 品ＳＸ溝深下限
    HSXDDMAX            As Double           ' 品ＳＸ溝深上限
    HSXD1MIN            As Double           ' 品ＳＸ直径１下限
    HSXD1MAX            As Double           ' 品ＳＸ直径１上限
    HSXCTCEN            As Double           ' 品ＳＸ結晶面傾縦中心
    HSXCYCEN            As Double           ' 品ＳＸ結晶面傾横中心
    EPDUP               As Integer          ' 結晶内側管理 EPD　上限
End Type

'=================================
'2003/02/28 ADD HITEC)okazaki start

Public Type type_DBDRV_Nukisi
    LOTID               As String * 12      ' ブロックID
    SXLID               As String * 13      ' SXLID
    MinMax              As Integer          ' 0:MIN 1:MAX
    BLOCKSEQ            As String * 3       ' ブロック内連番
    WFSTA               As String * 1       ' WF状態
    hinban              As String * 8       ' 品番
    RTOP_POS            As Double           ' 論理ブロック内位置
    RITOP_POS           As Double           ' 論理結晶内位置
    SMPLEID             As String * 16      ' 抜試位置
    SHAFLAG             As String * 1       ' サンプルフラグ
    INDTM               As Date
    BASKETID            As String * 6
    SLOTNO              As Integer
    CURRWPCS            As Integer
    EXISTFLG            As String * 1
    TOP_POS             As Integer
    REJCAT              As String * 1
    TXID                As String * 6
    REGDATE             As Date
    SUMMITSENDFLAG      As String * 1
    SENDFLAG            As String * 1
    SENDDATE            As Date
    HREJCODE            As String * 4
    UPDPROC             As String * 5
    UPDDATE             As Date
    REVNUM              As Integer
    factory             As String * 1
    opecond             As String * 1
    KANKBN              As String * 1
    NREJCODE            As String * 6
    RINGOTPOS           As Double
End Type

Public Type type_DBDRV_LOTSXL
    LOTID               As String * 12      ' ブロックID
    SXLID               As String * 13      ' SXLID
    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    INGOTPOS            As Integer          '結晶内位置
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
End Type
'2003/02/28 Add HITEC)okazaki end

'2003/02/28 Hitec)okazaki add start
Public tExamine()       As type_DBDRV_Nukisi    '画面表示時
Public tGetExamine()    As type_DBDRV_Nukisi    '実行時
                                                'ウェハーセンター入庫情報テーブル
Public tKeturaku()      As typ_TBCMY012
'表示条件
Public Const mSprChg_0  As Integer = 0          '全体
Public Const mSprChg_1  As Integer = 1          '良品
Public Const mSprChg_2  As Integer = 2          'サンプル
Public Const mSprChg_3  As Integer = 3          '不良

Public tSXLID() As type_DBDRV_LOTSXL
'2003/02/28 Hitec)okazaki add end

'add start 2003/03/15 hitec)matsumoto ---------------

''WFﾏｯﾌﾟ管理ﾃｰﾌﾞﾙ構造体
'Public Type typeWFmap
'    LOTID       As String       'ブロックID
'    BLOCKSEQ    As Integer      'ブロック内連番
'    INDTM       As Variant      'ＷＦセンター入庫日時
'    BASKETID    As String       'バスケットID
'    SLOTNO      As Integer      'スロットNO
'    CURRWPCS    As Integer      'ＷＦ枚数
'    EXISTFLG    As String       '存在フラグ
'    TOP_POS     As Integer      'ブロックのＴＯＰからの位置
'    REJCAT      As String       '欠落理由
'    TXID        As String       'トランザクションID
'    REGDATE     As Variant      '登録日付
'    SUMMITSENDFLG   As String   'SUMIT送信フラグ
'    SENDFLG     As String       '送信フラグ
'    SENDDATE    As Variant      '送信日時
'    WFSTA       As String       'WF状態
'    HREJCODE    As String       '不良理由コード
'    UPDPROC     As String       '更新工程
'    UPDDATE     As Variant      '更新日時
'    SXLID       As String       'SXLID
'    hinban      As String       '品番
'    REVNUM      As Integer      '製品番号改訂番号
'    factory     As String       '工場
'    opecond     As String       '操業条件
'    KANKBN      As String       '完了区分
'    SMPLEID     As String       '抜試位置
'    NREJCODE    As String       '抜試返答理由コード
'    SMPLEFLG    As String       'サンプルフラグ
'    RTOP_POS    As Double       '論理ブロック内位置
'    RITOP_POS   As Double       '論理結晶内位置
'End Type

'Public gtWFmap() As typeWFmap
Public bWfmapView As Boolean

''WFﾏｯﾌﾟ対応画面データ格納構造体
'Public Type typeSprWFmap
'    LOTID       As Variant      'ブロックID
'    hinban      As Variant      '品番
'    REVNUM      As Variant      '製品番号改訂番号
'    factory     As Variant      '工場
'    opecond     As Variant      '操業条件
'    HINUP       As tFullHinban  ' 上品番
'    HINDN       As tFullHinban  ' 下品番
'    blockp      As Variant      'ブロックP
'''''BLOCKP_T    As Variant      'ブロックP（上）
'''''BLOCKP_B    As Variant      'ブロックP（下）
'    KESSYOUP    As Variant      '結晶P
'''''KESSYOUP_T  As Variant      '結晶P（上）
'''''KESSYOUP_B  As Variant      '結晶P（下）
'    BLOCKSEQ    As Integer      'マップ位置
'''''BLOCKSEQ_T  As Integer      'マップ位置（上）
'''''BLOCKSEQ_B  As Integer      'マップ位置（下）
'    wfnum       As Integer      'ＷＦ枚数
'    WFSTA       As Variant      'WF状態
'''''WFSTA_T     As Variant      'WF状態（上）
'''''WFSTA_B     As Variant      'WF状態（下）
'    REJCODE     As Integer      '不良区分
'    SAMPLEID    As Variant      'サンプルID
'''''SAMPLEID_T  As Variant      'サンプルID（上）
'''''SAMPLEID_B  As Variant      'サンプルID（下）
'    WFSMP_Rs    As Variant      '検査項目（Rs）
'    WFSMP_Oi    As Variant      '検査項目（Oi）
'    WFSMP_B1    As Variant      '検査項目（B1）
'    WFSMP_B2    As Variant      '検査項目（B2）
'    WFSMP_B3    As Variant      '検査項目（B3）
'    WFSMP_L1    As Variant      '検査項目（L1）
'    WFSMP_L2    As Variant      '検査項目（L2）
'    WFSMP_L3    As Variant      '検査項目（L3）
'    WFSMP_L4    As Variant      '検査項目（L4）
'    WFSMP_DS    As Variant      '検査項目（DS）
'    WFSMP_DZ    As Variant      '検査項目（DZ）
'    WFSMP_SP    As Variant      '検査項目（SP）
'    WFSMP_D1    As Variant      '検査項目（D1）
'    WFSMP_D2    As Variant      '検査項目（D2）
'    WFSMP_D3    As Variant      '検査項目（D3）
'    SHAFLAG     As Variant      'サンプルフラグ
'''''SHAFLAG_T   As Variant      'サンプルフラグ（上）
'''''SHAFLAG_B   As Variant      'サンプルフラグ（下）
'    ADD_FLG     As String       '0：既存抜試行，1：追加抜試行
'End Type
'Public gtSprWfMap() As typeSprWFmap

''WF状態
'Public Const gsWF_STA_0  As String = "0"      '通常
'Public Const gsWF_STA_1  As String = "1"      '共有
''Public Const gsWF_STA_2 As String = "2"      '指示待ち
''Public Const gsWF_STA_3 As String = "3"      '指示OK
'Public Const gsWF_STA_4  As String = "4"      '欠落
''Public Const gsWF_STA_5 As String = "5"      '結果
'
''サンプルフラグ
'Public Const gsWF_SMPL_0 As String = "0"      '欠落
'Public Const gsWF_SMPL_1 As String = "1"      '指示待ち
'Public Const gsWF_SMPL_2 As String = "2"      '指示OK
'Public Const gsWF_SMPL_3 As String = "3"      '指示NG
'Public Const gsWF_SMPL_4 As String = "4"      '結果
'
''WF状態（画面表示）
'Public Const gsWF_STA_NORMAL      As String = "通常"       '通常
'Public Const gsWF_STA_STA_K       As String = "欠落"       '欠落
'Public Const gsWF_STA_SIJI        As String = "指示待ち"   '指示待ち
'Public Const gsWF_STA_SIJI_OK     As String = "指示OK"     '指示OK
'Public Const gsWF_STA_SIJI_NG     As String = "指示NG"     '指示NG
'Public Const gsWF_STA_SIJI_KEKKA  As String = "結果"       '結果
'サンプルフラグ（画面表示）
'Public Const gsWF_SMPL_JOINT    As String = "共有"  '共有

'add end 2003/03/15 hitec)matsumoto ---------------------

'add 2003/03/25 hitec)matsumoto ｸﾞﾛｰﾊﾞﾙ関数として使いたいので、f_cmbc039_3.frmより移動----------------
Public SIngotP      As Integer              ' インゴット上側位置
Public EIngotP      As Integer              ' インゴット下側位置
'add 2003/03/25 hitec)matsumoto ------------------------------
Public tmpSXLMng()  As typ_TBCME042

Public sWrpLOTID()      As String           'ﾌﾞﾛｯｸID(Warp実績紐付け用)　05/12/26 ooba
Public iWrpBLOCKSEQ()   As Integer          'ﾌﾞﾛｯｸ内連番(Warp実績紐付け用)　05/12/26 ooba
Public pWafSmp_wk()     As typ_XSDCW        'ｻﾝﾌﾟﾙ管理(初期ﾃﾞｰﾀ退避用)　08/02/04 ooba

Public CngSmpID_UD()    As String



'2002/09/11 ADD hitec)N.MATSUMOTO End


'概要      :抜試変更指示用 ブロックＩＤ入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sBlockID　　　,I  ,String         　,ブロックID
'      　　:bKounyu 　　　,I  ,Boolean        　,購入単結晶フラグ
'　　      :pCryInf 　　　,O  ,typ_TBCME037   　,結晶情報
'　　      :pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'　　      :pHinMng 　　　,O  ,typ_TBCME041   　,品番管理
'      　　:pSXLMng 　　　,O  ,typ_TBCME042   　,SXL管理
'      　　:pWafSmp 　　　,O  ,typ_XSDCW   　   ,新サンプル管理（SXL）
'　　      :pBlkInf 　　　,O  ,typ_BlkInf3    　,ブロック情報
'　　      :pHinSpec　　　,O  ,typ_HinSpec    　,製品仕様
'　　      :pLackWaf　　　,O  ,typ_LackWaf    　,欠落ウェハー情報
'　　      :pBlkID  　　　,O  ,String         　,払出単位ブロックID
'      　　:dNeraiRes 　　,O  ,Double         　,ねらい品番の比抵抗上限値（P+の判断用）
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/10 小林 作成
Public Function DBDRV_scmzc_fcmkc001k_Disp(ByVal sBlockId As String, bKounyu As Boolean, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pHinMng() As typ_TBCME041, pSXLMng() As typ_TBCME042, _
                                           pWafSmp() As typ_XSDCW, pBlkInf() As typ_BlkInf3, _
                                           pHinSpec() As typ_HinSpec, pLackWaf() As typ_LackWaf, _
                                           pBlkID() As String, dNeraiRes As Double, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim sDbName     As String
    Dim sCryNum     As String
    Dim sHin        As String
    Dim sBLK        As String
    Dim dMenseki    As Double
    Dim dTopWght    As Double
    Dim dCharge     As Double
    Dim dMeas(4)    As Double
    Dim bFlag       As Boolean
    Dim recCnt      As Long
    Dim i           As Long
    Dim j           As Long
    Dim k           As Long
    Dim REJCAT      As String

    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    Dim tmpHinMng() As typ_TBCME041     '品番情報更新用
'    Dim ltXSDCA()   As typ_XSDCA        'XSDCWデータ補完用データ
    Dim tmpWafSmp() As typ_XSDCW        'XSDCWデータ補完用データ
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp"
    sErrMsg = ""

    '' ブロック管理の取得
    sDbName = "E040"
    sCryNum = Left(sBlockId, 9) & "000"
    sql = "select INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, LSTATCLS"
    sql = sql & " from  TBCME040"
    sql = sql & " where CRYNUM   ='" & sCryNum & "'"
    sql = sql & "   and INGOTPOS>= 0"
    sql = sql & "   and LENGTH  >  0"
    sql = sql & " order by INGOTPOS"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    bFlag = False
    ReDim pBlkInf(recCnt)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.TOPSMPLPOS = rs("INGOTPOS")
            .LENGTH = rs("LENGTH")
            .REALLEN = rs("REALLEN")
            .BLOCKID = rs("BLOCKID")
            .NOWPROC = rs("NOWPROC")
            .COF.BOTSMPLPOS = .COF.TOPSMPLPOS + .LENGTH
            .DELFLG = "0"
            .TOPPOS = 0
            .BOTPOS = .REALLEN
            If .BLOCKID = sBlockId Then
                '' 工程チェック
                If rs("LSTATCLS") <> "W" Then
                    sErrMsg = GetMsgStr("EPRC2")
                    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                bFlag = True
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close

    '' ブロックID存在チェック
    If bFlag = False Then
        sErrMsg = GetMsgStr("EBLK0")
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)


    '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
    sDbName = "E039"
    '2004.09.08 Y.K 紐付け変更
'    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' and LENGTH>0 order by INGOTPOS"
    sql = " where substr(CRYNUM,1,9)='" & Left(sCryNum, 7) & "0" & Mid(sCryNum, 9, 1) & "' and LENGTH>0 order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' 品番管理の取得(s_cmzcTBCME041_SQL.bas が必要)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' SXL管理の取得(s_cmzcTBCME042_SQL.bas が必要)
    sDbName = "E042"
    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    sql = " where XTALCB='" & sCryNum & "' order by INPOSCB"
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pSXLMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    '' WFサンプル管理の取得(s_cmzcTBCME044_SQL.bas が必要)
    '　　↑新サンプル管理に変更-------2003/09/18
    ' ↓生死区分を見る必要有り
'   sDbName = "E044"
    sDbName = "XSDCW"
    sql = " where XTALCW='" & sCryNum & "'" _
        & "   and LIVKCW='0'" _
        & " order by INPOSCW"
'    If DBDRV_GetTBCME044(pWafSmp(), sql) = FUNCTION_RETURN_FAILURE Then
    'XSDCBの情報をﾍﾞｰｽに取得　08/02/04 ooba
    If DBDRV_GetXSDCW(sCryNum, pWafSmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 関連ﾌﾞﾛｯｸ情報取得　08/10/28 ooba
    If sKanrenFlg = "1" Then
        sDbName = "Y023"
        sql = "SELECT "
        sql = sql & "BLOCKID, "
        sql = sql & "PROCCAT "
        sql = sql & "FROM TBCMY023 "
        sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
        sql = sql & "AND TRANCNT = ( "
        sql = sql & "    SELECT "
        sql = sql & "    MAX(TRANCNT) "
        sql = sql & "    FROM TBCMY023 "
        sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
        sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
        sql = sql & ") "
        sql = sql & "ORDER BY BLOCKID "
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
        recCnt = rs.RecordCount
        If recCnt <= 0 Then
            rs.Close
            sErrMsg = GetMsgStr("EGET2", sDbName)
            DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        '関連ﾌﾞﾛｯｸでない場合
        If rs.Fields("PROCCAT") = "D" Then
            ReDim pBlkID(1)
            pBlkID(1) = sBlockId
        Else
            ReDim pBlkID(recCnt)
            'ﾌﾞﾛｯｸIDｾｯﾄ
            For i = 1 To recCnt
                pBlkID(i) = rs("BLOCKID")
                rs.MoveNext
            Next i
        End If
        rs.Close
    '関連ﾌﾞﾛｯｸでない場合
    Else
        ReDim pBlkID(1)
        pBlkID(1) = sBlockId
    End If
    
'''    '' ブロック新規情報の取得
'''    sDbName = "Y001"
'''    sql = "select SBLOCKID from TBCMY001 where BLOCKID='" & sBlockID & "'"
'''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''    If rs.RecordCount <= 0 Then
'''        rs.Close
'''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
'''        GoTo proc_exit
'''    End If
'''    sBLK = rs("SBLOCKID")
'''    rs.Close
'''
'''
'''    sql = "select BLOCKID from TBCMY001"
'''    sql = sql & " where SBLOCKID='" & sBLK & "'"
'''    sql = sql & " order by SBLOCKID, BLOCKORDER"
'''    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'''    recCnt = rs.RecordCount
'''    If recCnt <= 0 Then
'''        rs.Close
'''        sErrMsg = GetMsgStr("EGET2", sDbName)
'''        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
'''        GoTo proc_exit
'''    End If
'''
'''    ReDim pBlkID(recCnt)
'''    For i = 1 To recCnt
'''        pBlkID(i) = rs("BLOCKID")
'''        rs.MoveNext
'''    Next i
'''    rs.Close

    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    ''品番管理の取得(XSDCA,XSDCB:指定ブロックのみ)
    sDbName = "E041update"
    If DBDRV_GetTBCME041_Clone(tmpHinMng(), sCryNum, pBlkID) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '結晶全体の品番データにブロック指定のデータを合成する
    Call s_cmbc036_2_F_SynHinban(pHinMng, tmpHinMng)

    '' XSDCW補完用データをXSDCAから作成し
    '' 取得したXSDCAデータを元にXSDCWの取得データを補完する
    sDbName = "XSDCA"
    ReDim tmpWafSmp(UBound(pWafSmp))
    ReDim pWafSmp_wk(UBound(pWafSmp))       '08/02/04 ooba
    For i = 0 To UBound(pWafSmp)
        tmpWafSmp(i) = pWafSmp(i)
        pWafSmp_wk(i) = pWafSmp(i)          '初期ﾃﾞｰﾀ退避　08/02/04 ooba
    Next i

    '↓追加 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応
    ReDim tSXLID(0)
    '↑追加 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応

    If DBDRV_GetXSDCWUpdate(tmpWafSmp(), sCryNum, pBlkID()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    '補完なしの場合、tmpWafSmpは初期化されるので、元のpWafSmpをそのまま使用する
    If UBound(tmpWafSmp) <> 0 Then
        ReDim pWafSmp(UBound(tmpWafSmp))
        For i = 0 To UBound(pWafSmp)
            pWafSmp(i) = tmpWafSmp(i)
        Next i
    End If

    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川

    '' 引上げ終了実績の取得
    '2001/07/23 S.Sano Start
    If Not bKounyu Then
    '2001/07/23 S.Sano End
        sDbName = "H004"
        sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE from TBCMH004 where CRYNUM='" & sCryNum & "'"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            sErrMsg = GetMsgStr("EGET2", sDbName)
            DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        dMenseki = AreaOfCircle(rs("DM"))
        dTopWght = rs("WGHTTOP")
        dCharge = rs("CHARGE")
        rs.Close
    '2001/07/23 S.Sano Start
    End If
    '2001/07/23 S.Sano End


    '' 結晶抵抗実績の取得
    sDbName = "J002"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            .COF.DUNMENSEKI = dMenseki      ' 断面積
            .COF.CHARGEWEIGHT = dCharge     ' チャージ量
            .COF.TOPWEIGHT = dTopWght       ' トップ重量

            '' トップ側比抵抗中央値の取得
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and POSITION= " & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.TOPRES = JudgCenter(dMeas())
            Else
                .COF.TOPRES = -9999
            End If
            rs.Close

            '' ボトム側比抵抗中央値の取得
            sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and POSITION= " & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM  ='" & sCryNum & "'"
                sql = sql & "   and POSITION= " & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & "   and TRANCNT = ANY(select MAX(TRANCNT) from TBCMJ002 where CRYNUM='" & sCryNum & "' and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
                Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            End If
            If rs.RecordCount > 0 Then
                dMeas(0) = rs("MEAS1")
                dMeas(1) = rs("MEAS2")
                dMeas(2) = rs("MEAS3")
                dMeas(3) = rs("MEAS4")
                dMeas(4) = rs("MEAS5")
                .COF.BOTRES = JudgCenter(dMeas())
            Else
                .COF.BOTRES = -9999
            End If
            rs.Close
        End With
    Next i


    '' 製品仕様の取得
    sDbName = "VE004"
    recCnt = UBound(pHinMng)
    ReDim pHinSpec(recCnt)
    k = 0
    For i = 1 To recCnt
        With pHinMng(i)
            sHin = RTrim$(.hinban)
            If sHin <> "" And sHin <> "G" And sHin <> "Z" Then
                For j = 1 To k
                    If pHinSpec(j).hin.hinban = .hinban Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).INGOTPOS = .INGOTPOS
                    pHinSpec(k).hin.hinban = .hinban
                    pHinSpec(k).hin.mnorevno = .REVNUM
                    pHinSpec(k).hin.factory = .factory
                    pHinSpec(k).hin.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH

                    ''残存酸素仕様チェック　03/12/11 ooba START ==============================>
                    iChkAoi = ChkAoiSiyou(pHinSpec(k).hin)
                    If iChkAoi < 0 Then
                        sErrMsg = "残存酸素(AOi)仕様エラー"
                        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    ''残存酸素仕様チェック　03/12/11 ooba END ================================>

                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET") & sDbName
                        DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)



ReDim pLackWaf(0)   '2003/05/05 hitec)okazaki
    '' 欠落ウェハー情報の取得
#If False Then
    sDbName = "VW002"
    sql = "select distinct BLOCKID, WAFERNO, TOP_POS, TAIL_POS"
    sql = sql & " from VECMW002 where CRYNUM='" & sCryNum & "' order by BLOCKID, WAFERNO"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pLackWaf(recCnt)
    For i = 1 To recCnt
        With pLackWaf(i)
            .BLOCKID = rs("BLOCKID")    ' ブロックID
            .WAFERNO = rs("WAFERNO")    ' ウェハー連番
            .TOP_POS = rs("TOP_POS")    ' ウェハー開始位置
            .TAIL_POS = rs("TAIL_POS")  ' ウェハー終了位置
        End With
        rs.MoveNext
    Next i
    rs.Close
#Else
''    sDbName = "VW004"
''    sql = "select distinct LOTID as BLOCKID, REJCAT, REJWFFROM as WAFERNO, REJWFTO as WAFERTO, REJFROM as TOP_POS, REJTO as TAIL_POS, ALLSCRAP"
''    sql = sql & " from VECMW004"
''    sql = sql & " where (LOTID like '" & Left$(sCryNum, 9) & "%') and (REJCAT<>'C')"
''    sql = sql & " order by LOTID, WAFERNO "

    'ﾋﾞｭｰ参照停止　06/02/06 ooba START ====================================================>
    sDbName = "Y012"
    sql = "select distinct LOTID as BLOCKID, REJCAT, REJWFFROM as WAFERNO, REJWFTO as WAFERTO, REJFROM as TOP_POS, REJTO as TAIL_POS, ALLSCRAP from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  case when (XXX.REJFROM<=B.WFFROM) then 0 else XXX.REJFROM end as REJFROM,"
    sql = sql & "  case when (XXX.REJTO>=B.WFTO) then C.LENGTH else XXX.REJTO end as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      'ﾌﾞﾛｯｸ状態での一部欠量対応 09/02/27 ooba
    sql = sql & " and lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  (select LOTID, min(TOP_POS)/10.0 as WFFROM, max(TOP_POS)/10.0 as WFTO from TBCMY011 "
    sql = sql & " where lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " group by LOTID) B,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=B.LOTID)"
    sql = sql & "  and (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='N')"
    sql = sql & " union all"
    sql = sql & " select distinct"
    sql = sql & "  C.CRYNUM,"
    sql = sql & "  XXX.LOTID,"
    sql = sql & "  REJCAT,"
    sql = sql & "  ALLSCRAP,"
    sql = sql & "  0 as REJFROM,"
    sql = sql & "  C.LENGTH as REJTO,"
    sql = sql & "  REJWFFROM,"
    sql = sql & "  REJWFTO"
    sql = sql & " from "
    sql = sql & "("
    sql = sql & "select "
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    0 as REJFROM,"
    sql = sql & "    LENGTH as REJTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCME040 B"
    sql = sql & "  where (A.LOTID=B.BLOCKID)"
    sql = sql & "    and (A.ALLSCRAP='Y')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    LOTID,"
    sql = sql & "    REJCAT,"
    sql = sql & "    ALLSCRAP,"
    sql = sql & "    LENFROM,"
    sql = sql & "    LENTO,"
    sql = sql & "    0 as REJWFFROM,"
    sql = sql & "    0 as REJWFTO"
    sql = sql & "  from TBCMY012"
'    sql = sql & "  where (REJCAT='A') and (ALLSCRAP='N')"
    sql = sql & "  where (REJCAT in ('A','E')) and (ALLSCRAP='N')"      'ﾌﾞﾛｯｸ状態での一部欠量対応 09/02/27 ooba
    sql = sql & " and lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    A.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    A.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ as REJWFTO"
    sql = sql & "  from TBCMY012 A"
    sql = sql & "  where (A.REJCAT='B') and (ALLSCRAP='N')"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " union all"
    sql = sql & "  select distinct"
    sql = sql & "    A.LOTID,"
    sql = sql & "    A.REJCAT,"
    sql = sql & "    A.ALLSCRAP,"
    sql = sql & "    B.TOP_POS/10.0 as REJFROM,"
    sql = sql & "    C.TOP_POS/10.0 as REJTO,"
    sql = sql & "    A.BLOCKSEQ as REJWFFROM,"
    sql = sql & "    A.BLOCKSEQ+A.REJPCS-1 as REJWFTO"
    sql = sql & "  from TBCMY012 A,"
    sql = sql & "    TBCMY011 B,"
    sql = sql & "    TBCMY011 C"
    sql = sql & "  where (A.REJCAT='C')"
    sql = sql & "    and (A.LOTID=B.LOTID) and (A.BLOCKSEQ=B.BLOCKSEQ)"
    sql = sql & "    and (A.LOTID=C.LOTID) and (A.BLOCKSEQ+A.REJPCS-1=C.BLOCKSEQ)"
    sql = sql & " and a.lotid like '" & Left$(sCryNum, 9) & "%'"
    sql = sql & " order by LOTID,REJFROM"
    sql = sql & ") XXX,"
    sql = sql & "  TBCME040 C"
    sql = sql & " where (XXX.LOTID=C.BLOCKID)"
    sql = sql & "  and (XXX.ALLSCRAP='Y')"
    sql = sql & ")"
    sql = sql & " where (REJCAT<>'C')"
    sql = sql & " order by LOTID, WAFERNO "
    'ﾋﾞｭｰ参照停止　06/02/06 ooba END ======================================================>

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pLackWaf(recCnt)
    k = 0
    For i = 1 To recCnt

'2002/08/22
'        If (REJCAT = rs("REJCAT")) _
'          And rs("ALLSCRAP") = "N" _
'          And (pLackWaf(k).BLOCKID = rs("BLOCKID")) _
'          And (pLackWaf(k).WAFERTO + 1 = rs("WAFERNO")) Then
        If (REJCAT = rs("REJCAT")) _
          And rs("ALLSCRAP") = "N" _
          And (pLackWaf(k).ALLSCRAP = rs("ALLSCRAP")) _
          And (pLackWaf(k).BLOCKID = rs("BLOCKID")) _
          And (pLackWaf(k).WAFERTO + 1 = rs("WAFERNO")) Then

            With pLackWaf(k)
                .WAFERTO = rs("WAFERTO")    ' ウェハー連番(to)
                .TAIL_POS = rs("TAIL_POS")  ' ウェハー終了位置
            End With
        Else
            k = k + 1
            With pLackWaf(k)
                .BLOCKID = rs("BLOCKID")    ' ブロックID
                .WAFERNO = rs("WAFERNO")    ' ウェハー連番
                .WAFERTO = rs("WAFERTO")    ' ウェハー連番(to)
                .TOP_POS = rs("TOP_POS")    ' ウェハー開始位置
                .TAIL_POS = rs("TAIL_POS")  ' ウェハー終了位置
                .ALLSCRAP = rs("ALLSCRAP")  ' 全数スクラップ
                .REJCAT = rs("REJCAT")      ' 欠落理由
            End With
        End If
        REJCAT = rs("REJCAT")
        rs.MoveNext
    Next i
    rs.Close
    ReDim Preserve pLackWaf(k)
#End If


    '' ねらい品番の比抵抗上限値を取得
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM  ='" & Left$(sBlockId, 9) & "000')"
    sql = sql & "   and (E37.RPHINBAN=E18.HINBAN)  and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & "   and (E37.RPFACT  =E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      'ここまではこないはず
    End If
    rs.Close


    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDbName)
    DBDRV_scmzc_fcmkc001k_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :関連ﾌﾞﾛｯｸ情報取得
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sBlockID       ,I  ,String         　,ﾌﾞﾛｯｸID
'      　　:tKanrenDisp()  ,I  ,typ_KanrenDisp   ,関連ﾌﾞﾛｯｸ一覧
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :08/01/23 ooba
Public Function DBDRV_scmzc_fcmkc001k_Disp2(sBlockId As String, _
                                            tKanrenDisp() As typ_KanrenDisp) As FUNCTION_RETURN

    Dim i, j        As Integer
    Dim iBlkCnt     As Integer      'ﾌﾞﾛｯｸ数
    Dim iHinCnt     As Integer      '品番数
    Dim skanblock() As String       '関連ﾌﾞﾛｯｸ　08/10/28 ooba
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp2"
    
    DBDRV_scmzc_fcmkc001k_Disp2 = FUNCTION_RETURN_FAILURE
    
    '関連ﾌﾞﾛｯｸ紐切紐付ﾃｰﾌﾞﾙより関連ﾌﾞﾛｯｸ取得　08/10/28 ooba
    sql = "SELECT "
    sql = sql & "BLOCKID, "
    sql = sql & "PROCCAT "
    sql = sql & "FROM TBCMY023 "
    sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TRANCNT = ( "
    sql = sql & "    SELECT "
    sql = sql & "    MAX(TRANCNT) "
    sql = sql & "    FROM TBCMY023 "
    sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
    sql = sql & ") "
    sql = sql & "ORDER BY BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    j = rs.RecordCount
    
    'ﾃﾞｰﾀ無し
    If j <= 0 Then
        GoTo proc_exit
    End If
    
    '関連ﾌﾞﾛｯｸでない場合
    If rs.Fields("PROCCAT") = "D" Then
        GoTo proc_exit
    End If
        
    ReDim skanblock(j)
    'ﾌﾞﾛｯｸIDｾｯﾄ
    For i = 1 To j
        skanblock(i) = rs.Fields("BLOCKID")
        rs.MoveNext
    Next i
    rs.Close
    
    
    '④ ③で取得したﾌﾞﾛｯｸIDを条件に関連ﾌﾞﾛｯｸ情報を取得(XSDCAから)
    sql = "SELECT "
    sql = sql & "PLANTCATCA, "                          '向先
    sql = sql & "XTALCA, "                              '結晶番号
    sql = sql & "CRYNUMCA, "                            'ﾌﾞﾛｯｸID
    sql = sql & "HINBCA || "
    sql = sql & "TO_CHAR(NVL(REVNUMCA,0),'FM00') || "
    sql = sql & "FACTORYCA || "
    sql = sql & "OPECA AS HINBAN, "                     '品番(12桁)
    sql = sql & "SXLIDCA, "                             'SXLID
    sql = sql & "SXLIDCB, "                             'SXLID(XSDCB)　08/07/10 ooba
    sql = sql & "GNWKNTCA, "                            '現在工程(XSDCA)
    sql = sql & "GNWKNTCB, "                            '現在工程(XSDCB)
    sql = sql & "INPOSCA, "                             '結晶内開始位置
    sql = sql & "INPOSCS, "                             'ﾌﾞﾛｯｸ終了位置　08/07/10 ooba
    sql = sql & "GNLCA, "                               '現在長さ
    sql = sql & "GNMCA, "                               '現在枚数
    sql = sql & "WFHOLDFLGCA, "                         'WFﾎｰﾙﾄﾞ区分
    sql = sql & "KDAYCA "                               '更新日付
    sql = sql & "FROM XSDCA,XSDCB,XSDCS "
    sql = sql & "WHERE CRYNUMCA IN ( "
    
    '取得条件変更　08/10/28 ooba
    For i = 1 To UBound(skanblock)
        sql = sql & "'" & skanblock(i) & "' "
        If i <> UBound(skanblock) Then sql = sql & ","
    Next i
    
'''    '③ ②で取得したSXLIDを含むﾌﾞﾛｯｸIDを取得(XSDCAから)
'''    sql = sql & "    SELECT "
'''    sql = sql & "    CRYNUMCA "
'''    sql = sql & "    FROM XSDCA "
'''    sql = sql & "    WHERE SXLIDCA IN ( "
'''    '② ①で取得したSXLIDの中で関連ﾌﾞﾛｯｸのSXLIDを取得(XSDCBから)
'''    sql = sql & "        SELECT "
'''    sql = sql & "        SXLIDCB "
'''    sql = sql & "        FROM XSDCB "
'''    sql = sql & "        WHERE SXLIDCB IN ( "
'''    '① 選択したﾌﾞﾛｯｸIDを条件にSXLIDを取得(XSDCAから)
'''    sql = sql & "            SELECT "
'''    sql = sql & "            SXLIDCA "
'''    sql = sql & "            FROM XSDCA "
'''    sql = sql & "            WHERE CRYNUMCA = '" & sBlockID & "' "
'''    sql = sql & "            AND LIVKCA = '0' "
'''    sql = sql & "        ) "
'''    sql = sql & "        AND LIVKCB = '0' "
'''    sql = sql & "        AND KBLKFLGCB = '1' "
'''    sql = sql & "    ) "
    sql = sql & ") "
    sql = sql & "AND SXLIDCA LIKE '" & Mid(sBlockId, 1, 9) & "%' "      '08/10/28 ooba
    sql = sql & "AND (LIVKCA = '0' OR "
'    sql = sql & "     (LIVKCA = '1' AND LSTATBCA = 'H' AND LUFRBCA = 'H')) "        '全数廃棄ﾃﾞｰﾀ　08/07/10 ooba
    sql = sql & "     (LIVKCA = '1' AND LSTATBCA in ('H','M','E') AND LUFRBCA = 'H')) "     '月次廃却,欠量条件追加 09/02/22 ooba
    sql = sql & "AND TBKBNCS = 'B' "
    sql = sql & "AND XSDCA.SXLIDCA = XSDCB.SXLIDCB(+) "
    sql = sql & "AND XSDCA.CRYNUMCA = XSDCS.CRYNUMCS "
    sql = sql & "ORDER BY CRYNUMCA, INPOSCA "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    'ﾃﾞｰﾀ無し
    If rs.RecordCount <= 0 Then
        GoTo proc_exit
    End If
    
    iBlkCnt = 0
    iHinCnt = 0
    
    ReDim tKanrenDisp(0)
    
    For i = 1 To rs.RecordCount
        '1番目ﾃﾞｰﾀまたは前ﾃﾞｰﾀとﾌﾞﾛｯｸIDが異なる場合、行追加
        If iBlkCnt = 0 Or tKanrenDisp(iBlkCnt).BLOCKID <> rs.Fields("CRYNUMCA") Then
            iBlkCnt = iBlkCnt + 1       'ﾌﾞﾛｯｸ数＋1
            iHinCnt = 1                 '品番数=1
            ReDim Preserve tKanrenDisp(iBlkCnt)
            
            With tKanrenDisp(iBlkCnt)
                '--向先
                If IsNull(rs.Fields("PLANTCATCA")) = False Then
                    For j = 1 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs.Fields("PLANTCATCA") Then
                           .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                '--結晶番号
                If IsNull(rs.Fields("XTALCA")) = False Then .CRYNUM = rs.Fields("XTALCA")
                '--ﾌﾞﾛｯｸID
                If IsNull(rs.Fields("CRYNUMCA")) = False Then .BLOCKID = rs.Fields("CRYNUMCA")
                '--品番数
                .HINCNT = iHinCnt
                '--品番
                If IsNull(rs.Fields("HINBAN")) = False Then .hinban(iHinCnt) = rs.Fields("HINBAN")
                '--SXLID
                If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLID(iHinCnt) = rs.Fields("SXLIDCA")
                '--SXLID(XSDCB)　08/07/10 ooba
                If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLID_CB(iHinCnt) = rs.Fields("SXLIDCB")
                '--SXLID(更新用)
                .SXLID_NEW = ""
                '--仕掛工程
                If rs.Fields("GNWKNTCA") = PROCD_WFC_SOUGOUHANTEI Then
                    'CW750→CW740変換
                    .Koutei = PROCD_NUKISI_HENKOU
                Else
                    If IsNull(rs.Fields("GNWKNTCA")) = False Then .Koutei = rs.Fields("GNWKNTCA")
                End If
                '--結晶内開始位置
                If IsNull(rs.Fields("INPOSCA")) = False Then .INGOTPOS(iHinCnt) = rs.Fields("INPOSCA")
                '--ﾌﾞﾛｯｸ終了位置　08/07/10 ooba
                If IsNull(rs.Fields("INPOSCS")) = False Then .BLKEPOS = rs.Fields("INPOSCS")
                '--長さ
                If IsNull(rs.Fields("GNLCA")) = False Then .LENGTH(iHinCnt) = rs.Fields("GNLCA")
                '--枚数
                If IsNull(rs.Fields("GNMCA")) = False Then .MAISU(iHinCnt) = rs.Fields("GNMCA")
                '--WFﾎｰﾙﾄﾞ区分
                If IsNull(rs.Fields("WFHOLDFLGCA")) = False Then .HOLD = rs.Fields("WFHOLDFLGCA")
                '--日付
                If IsNull(rs.Fields("KDAYCA")) = False Then .KDATE = rs.Fields("KDAYCA")
            End With
        'ﾌﾞﾛｯｸIDが一致する場合、品番ﾃﾞｰﾀ追加
        Else
            iHinCnt = iHinCnt + 1       '品番数＋1
            '品番数ｴﾗｰ
            If iHinCnt > 5 Then
                GoTo proc_exit
            End If
            
            With tKanrenDisp(iBlkCnt)
                '--品番数
                .HINCNT = iHinCnt
                '--品番
                If IsNull(rs.Fields("HINBAN")) = False Then .hinban(iHinCnt) = rs.Fields("HINBAN")
                '--SXLID
                If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLID(iHinCnt) = rs.Fields("SXLIDCA")
                '--SXLID(XSDCB)　08/07/10 ooba
                If IsNull(rs.Fields("SXLIDCB")) = False Then .SXLID_CB(iHinCnt) = rs.Fields("SXLIDCB")
                '--結晶内開始位置
                If IsNull(rs.Fields("INPOSCA")) = False Then .INGOTPOS(iHinCnt) = rs.Fields("INPOSCA")
                '--長さ
                If IsNull(rs.Fields("GNLCA")) = False Then .LENGTH(iHinCnt) = rs.Fields("GNLCA")
                '--枚数
                If IsNull(rs.Fields("GNMCA")) = False Then .MAISU(iHinCnt) = rs.Fields("GNMCA")
            End With
        End If
        rs.MoveNext
    Next i
    
    rs.Close
    
    DBDRV_scmzc_fcmkc001k_Disp2 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'Add Start 2010/07/09 SMPK Nakamura
'概要      :関連ﾌﾞﾛｯｸ情報取得
'ﾊﾟﾗﾒｰﾀ　　:変数名         ,IO ,型               ,説明
'      　　:sBlockID       ,I  ,String         　,ﾌﾞﾛｯｸID
'      　　:tKanrenDisp()  ,I  ,typ_KanrenDisp   ,関連ﾌﾞﾛｯｸ一覧
'      　　:戻り値         ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2010/07/09 SMPK Nakamura
Public Function DBDRV_scmzc_fcmkc001k_Disp3(ByVal sBlockId As String, _
                                            ByRef tKanrenList() As typ_KanrenList) As FUNCTION_RETURN

    Dim i           As Integer
    Dim iBlkCnt     As Integer      'ﾌﾞﾛｯｸ数
    Dim rs          As OraDynaset
    Dim sql         As String
    
    'ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Disp3"
    
    DBDRV_scmzc_fcmkc001k_Disp3 = FUNCTION_RETURN_FAILURE
    
    '関連ﾌﾞﾛｯｸ紐切紐付ﾃｰﾌﾞﾙより関連ﾌﾞﾛｯｸ取得
    sql = "SELECT "
    sql = sql & "BLOCKID, "
    sql = sql & "PROCCAT, "
    sql = sql & "DECODE( GNWKNTCA, "
    sql = sql & "        '" & PROCD_NUKISI_HENKOU & "', 0,"
    sql = sql & "        '" & PROCD_WFC_SOUGOUHANTEI & "', 0,"
    sql = sql & "        1) as WAITFLG "
    sql = sql & "FROM TBCMY023, XSDCA, XSDCS "
    sql = sql & "WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TRANCNT = ( "
    sql = sql & "    SELECT "
    sql = sql & "    MAX(TRANCNT) "
    sql = sql & "    FROM TBCMY023 "
    sql = sql & "    WHERE CRYNUM LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "    AND BLOCKID = '" & sBlockId & "' "
    sql = sql & ") "
    sql = sql & "AND BLOCKID = CRYNUMCA "
    sql = sql & "AND (LIVKCA = '0' OR "
    sql = sql & "     (LIVKCA = '1' AND LSTATBCA in ('H','M','E') AND LUFRBCA = 'H')) " '月次廃却,欠量条件追加
    sql = sql & "AND SXLIDCA LIKE '" & Mid(sBlockId, 1, 9) & "%' "
    sql = sql & "AND TBKBNCS = 'B' "
    sql = sql & "AND XSDCA.CRYNUMCA = XSDCS.CRYNUMCS "
    sql = sql & "ORDER BY WAITFLG DESC, BLOCKID "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    'ﾃﾞｰﾀ無し
    If rs.RecordCount <= 0 Then
        GoTo proc_exit
    End If
    
    iBlkCnt = 0
    ReDim tKanrenList(0)
    'ﾌﾞﾛｯｸIDｾｯﾄ
    For i = 1 To rs.RecordCount
        '関連ﾌﾞﾛｯｸでない場合
        If rs.Fields("PROCCAT") = "D" Then
            GoTo proc_exit
        End If
        If iBlkCnt = 0 Or tKanrenList(iBlkCnt).BLOCKID <> rs.Fields("BLOCKID") Then
            iBlkCnt = iBlkCnt + 1       'ﾌﾞﾛｯｸ数＋1
            
            ReDim Preserve tKanrenList(iBlkCnt)

            tKanrenList(iBlkCnt).BLOCKID = rs.Fields("BLOCKID")
            tKanrenList(iBlkCnt).WAIT = rs.Fields("WAITFLG")
        End If
        rs.MoveNext
    Next i
    rs.Close
        
    DBDRV_scmzc_fcmkc001k_Disp3 = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'ｴﾗｰﾊﾝﾄﾞﾗ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function
'Add End 2010/07/09 SMPK Nakamura

'概要      :抜試変更指示用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sStaffID　　　,I  ,String         　,社員ID
'　　      :pBlkInf 　　　,I  ,typ_BlkInf3    　,ブロック情報
'      　　:pLackMap　　　,I  ,typ_LackMap    　,欠落ウェハー
'      　　:pSXLMng 　　　,I  ,typ_TBCME042   　,SXL管理
'      　　:pWafSmp 　　　,I  ,typ_XSDCW   　   ,新サンプル管理（SXL）
'      　　:pMesInd 　　　,I  ,typ_TBCMY003   　,測定評価方法指示
'      　　:pTrnScr 　　　,I  ,typ_TBCMW006   　,振替廃棄実績
'      　　:pSXLDcd 　　　,I  ,typ_TBCMY007   　,SXL確定指示
'      　　:pEpMesInd 　  ,I  ,typ_TBCMY020   　,EP測定評価指示
'      　　:sKanrenB 　   ,I  ,String         　,関連ﾌﾞﾛｯｸ　07/08/06 ooba
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/11 蔵本 作成
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -s-
Public Function DBDRV_scmzc_fcmkc001k_Exec(sStaffID As String, pBlkInf() As typ_BlkInf3, _
                                           pLackMap() As typ_LackMap, pSXLMng() As typ_TBCME042, _
                                           pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003, _
                                           pTrnScr() As typ_TBCMW006, pSXLDcd() As typ_TBCMY007, pEpMesInd() As typ_TBCMY020, sKanrenB() As String, sErrMsg As String) As FUNCTION_RETURN
''Public Function DBDRV_scmzc_fcmkc001k_Exec(sStaffID As String, pBlkInf() As typ_BlkInf3, _
''                                           pLackMap() As typ_LackMap, pSXLMng() As typ_TBCME042, _
''                                           pWafSmp() As typ_XSDCW, pMesInd() As typ_TBCMY003, _
''                                           pTrnScr() As typ_TBCMW006, pSXLDcd() As typ_TBCMY007, sErrMsg As String) As FUNCTION_RETURN
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh -e-
    Dim sql     As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim Blks    As String
    Dim sTmpSxl() As String     '仕掛工程再ﾁｪｯｸ用SXLID　06/03/14 ooba
    Dim recCnt  As Long
    Dim i       As Long
    Dim dynOra  As OraDynaset

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function DBDRV_scmzc_fcmkc001k_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE

    '仕掛工程再チェック機能追加　06/03/14 ooba START ========================================>
    sDbName = "XSDCA"
    sql = "SELECT SXLIDCA "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE CRYNUMCA LIKE '" & Left(pBlkInf(1).BLOCKID, 9) & "%'"
    sql = sql & "   AND (INPOSCA>=" & SIngotP
    sql = sql & "   AND  INPOSCA< " & EIngotP & ")"
    sql = sql & "   AND LIVKCA = '0' "
    sql = sql & "GROUP BY SXLIDCA"
    Set dynOra = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    ReDim sTmpSxl(0)
    If dynOra.RecordCount > 0 Then
        For i = 1 To dynOra.RecordCount
            If Not IsNull(dynOra.Fields("SXLIDCA")) Then
                ReDim Preserve sTmpSxl(i)
                sTmpSxl(i) = dynOra.Fields("SXLIDCA")
            End If
            dynOra.MoveNext
        Next i
    End If
    dynOra.Close
    If DBDRV_CheckCodeXSDCB(sTmpSxl, PROCD_NUKISI_HENKOU, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    '仕掛工程再チェック機能追加　06/03/14 ooba END ==========================================>

    '' SXL管理の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    recCnt = UBound(pSXLMng)
    If recCnt > 0 Then
        ''↓変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
        '' XSDCBに必要なデータが存在する可能性を考え、Delete→Insertはやめる
        sDbName = "XSDCB"
        If DBDRV_SXL_INS_CB(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            GoTo proc_exit
        End If
        '↓元ロジック（SXL管理（E042）→XSDCB機能移行）
'        sDbName = "E042"
'        sql = "delete from  TBCME042"
'        sql = sql & " where CRYNUM   ='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & "   and INGOTPOS>= " & pSXLMng(1).IngotPos
''       sql = sql & "   and INGOTPOS<  " & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & "   and INGOTPOS>= " & SIngotP
'        sql = sql & "   and INGOTPOS<  " & EIngotP
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        If DBDRV_SXL_INS(pSXLMng()) = FUNCTION_RETURN_FAILURE Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
'            GoTo proc_exit
'        End If
        ''↑変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
    End If

    '' WFサンプル管理の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    If recCnt > 0 Then
        sDbName = "XSDCW"
'        sDbName = "E044"
'        '範囲開始位置のWFサンプルを削除
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS=" & pSXLMng(1).IngotPos
'        sql = sql & " and INPOSCW=" & SIngotP
'        sql = sql & " and SMPKBNCW in ('T', 'D')"
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        '範囲に完全に含まれるWFサンプルを削除
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS>" & pSXLMng(1).IngotPos
''       sql = sql & " and INGOTPOS<" & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & " and INPOSCW>" & SIngotP
'        sql = sql & " and INPOSCW<" & EIngotP
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)
'        '範囲終了位置のWFサンプルを削除
'        sql = "delete from XSDCW"
'        sql = sql & " where XTALCW='" & pSXLMng(1).CRYNUM & "'"
''       sql = sql & " and INGOTPOS=" & pSXLMng(recCnt).IngotPos + pSXLMng(recCnt).LENGTH
'        sql = sql & " and INPOSCW=" & EIngotP
'        sql = sql & " and SMPKBNCW in ('B', 'U')"
'        WriteDBLog sql, sDbName
'        Call OraDB.ExecuteSQL(sql)

        '新サンプル管理にデータがあるか
        For i = 1 To UBound(pWafSmp)
            sql = "SELECT count(*) "
            sql = sql & "FROM  XSDCW "
            sql = sql & "WHERE SXLIDCW ='" & pWafSmp(i).SXLIDCW & "'"
        '   sql = sql & "  and SMPKBNCW='" & pWafSmp(i).SMPKBNCW & "'"
            sql = sql & "  and TBKBNCW ='" & pWafSmp(i).TBKBNCW & "'"

            Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
        '   If 0 < OraDB.ExecuteSQL(sql) Then
            If dynOra.Fields(0) <> 0 Then
                'データがあればUpdate
                If DBDRV_WfSmp_UPD(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            Else
                'なければInsert
                If DBDRV_WfSmp_INS(pWafSmp(), i) = FUNCTION_RETURN_FAILURE Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            End If
        Next i
    End If


    '' 抜試変更指示実績の挿入
    sDbName = "W003"
    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sCryNum = Left(.BLOCKID, 9) & "000"
            sql = "insert into TBCMW003 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE, BLOCKID, DELFLG, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"

            sql = sql & " select '"
            sql = sql & sCryNum & "', "
            sql = sql & .COF.TOPSMPLPOS & ", "
            sql = sql & "nvl(max(TRANCNT),0)+1, "
            sql = sql & .REALLEN & ", '"
            sql = sql & MGPRCD_NUKISI_HENKOU & "', '"
            sql = sql & PROCD_NUKISI_HENKOU & "', '"
            sql = sql & .BLOCKID & "', '"
            sql = sql & .DELFLG & "', '"
            sql = sql & sStaffID & "', "
            sql = sql & "sysdate, '"
            sql = sql & sStaffID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "sysdate"
            sql = sql & " from  TBCMW003"
            sql = sql & " where CRYNUM  ='" & sCryNum & "'"
            sql = sql & "   and INGOTPOS= " & .COF.TOPSMPLPOS

            '' WriteDBLog sql, sDbName

            Debug.Print sql

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                GoTo proc_exit
            End If
        End With
    Next i

    '' 測定評価方法指示の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        GoTo proc_exit
    End If

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
    '' エピ測定評価指示情報の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "Y020"
    If DBDRV_SokuSizi_EP_Ins(pEpMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        GoTo proc_exit
    End If
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    '' 振替廃棄実績の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "W006"
    recCnt = UBound(pTrnScr)
    For i = 1 To recCnt
        If DBDRV_Furikae_Ins(pTrnScr(i)) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            GoTo proc_exit
        End If
    Next i

    '' SXL確定指示の挿入
    sDbName = "Y007"
    recCnt = UBound(pSXLDcd)
    For i = 1 To recCnt
        With pSXLDcd(i)
            sql = "insert into TBCMY007 "
            ' 2007/09/03 SPK Tsutsumi Add Start
'            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE,PLANTCAT)"
            ' 2007/09/03 SPK Tsutsumi Add Start
            sql = sql & " values ('"
            sql = sql & .SXL_ID & "', '"        ' SXL-ID
            sql = sql & .SAMPLE_FROM & "', '"   ' サンプルID (From)
            sql = sql & .SAMPLE_TO & "', '"     ' サンプルID (To)
            sql = sql & .BLOCKID & "', '"       ' ブロックＩＤ
            sql = sql & .hinban & "', "         ' 確定品番
            sql = sql & "'S ', "                ' 区分コード
            sql = sql & "'TX853I', "            ' トランザクションID
            sql = sql & "sysdate, "             ' 登録日付
            sql = sql & "'0', "                 ' SUMMIT送信フラグ
            sql = sql & "'0', "                 ' 送信フラグ

            ' 2007/09/03 SPK Tsutsumi Add Start
            sql = sql & "sysdate,"              ' 送信日付
            sql = sql & "'" & sCmbMukesaki & "'"  ' 向先
            ' 2007/09/03 SPK Tsutsumi Add End

            '' WriteDBLog sql, sDbName

            Debug.Print sql

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

'2003/01/09 ooba チェックフラグ条件変更
    '' 欠落情報の更新
    sDbName = "Y012"
    Dim m As Integer
    Dim j As Long
    m = UBound(pBlkInf)
    recCnt = UBound(pLackMap)
    For i = 1 To m
        For j = 1 To recCnt
            If pBlkInf(i).BLOCKID = pLackMap(j).BLOCKID Then
                sql = "update TBCMY012 set CHKFLG='1' where LOTID='" & pLackMap(j).BLOCKID & "'"
                '' WriteDBLog sql, sDbName

                Debug.Print sql

                If OraDB.ExecuteSQL(sql) < 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    GoTo proc_exit
                End If
            End If
        Next j
    Next i

    'XSDCBのﾎｰﾙﾄﾞ区分(WF)更新　04/06/30 ooba
    Dim SxlCnt As Integer
    sDbName = "XSDCB"
    For i = 1 To UBound(pSXLMng)
        sql = "select count(*) from XSDCA "
        sql = sql & "where LIVKCA = '0' "
        sql = sql & "and SXLIDCA = '" & pSXLMng(i).SXLID & "' "
        Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
        SxlCnt = dynOra.Fields(0)
        'SXLﾃﾞｰﾀが存在する場合
        If SxlCnt > 0 Then
            sql = "select count(*) from XSDCA, XSDCB "
            sql = sql & "where LIVKCA = '0' "
            sql = sql & "and LIVKCB = '0' "
            sql = sql & "and (WFHOLDFLGCA != '1' "
            sql = sql & "or WFHOLDFLGCA is NULL) "
            sql = sql & "and WFHOLDFLGCB = '1' "
            sql = sql & "and SXLIDCA = SXLIDCB "
            sql = sql & "and SXLIDCA = '" & pSXLMng(i).SXLID & "' "
            Set dynOra = OraDB.DBCreateDynaset(sql, 0&)
            'XSDCBのﾎｰﾙﾄﾞ区分(WF)が「1」でXSDCAのﾎｰﾙﾄﾞ区分(WF)がすべて「1」以外の場合
            If dynOra.Fields(0) = SxlCnt Then
                'XSDCBのﾎｰﾙﾄﾞ区分(WF)を「0」に更新
                sql = "update XSDCB set WFHOLDFLGCB = '0' where SXLIDCB = '" & pSXLMng(i).SXLID & "' "
                '' WriteDBLog sql, sDbName
                Debug.Print sql

                If OraDB.ExecuteSQL(sql) < 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End If
    Next i

    '関連ブロック情報登録停止　08/01/23 ooba
''    '関連ﾌﾞﾛｯｸ情報登録　07/08/06 ooba START =====================================>
''    If UBound(sKanrenB) > 1 Then
''        sDbName = "Y023"
''        If DBDRV_KanrenBlk(left(pBlkInf(1).BLOCKID, 9) & "000", sKanrenB(), _
''                            SIngotP, EIngotP) = FUNCTION_RETURN_FAILURE Then
''
''            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
''            GoTo proc_exit
''        End If
''    End If
''    '関連ﾌﾞﾛｯｸ情報登録　07/08/06 ooba END =======================================>

'    '' 欠落情報の更新
'    sDBName = "Y012"
'    recCnt = UBound(pLackMap)
'    For i = 1 To recCnt
'        sql = "update TBCMY012 set CHKFLG='1'"
'        sql = sql & " where LOTID='" & pLackMap(i).BLOCKID & "'"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) < 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            GoTo proc_exit
'        End If
'    Next i
'送信済の測定評価指示も、新たなTRANCNTで作成される。
'そのため、既存レコードの再送は不要
'    '' 欠落を含むブロックにある測定評価指示の再送（欠落内のサンプルIDを除く）
'    sDBName = "Y003-2"
'    Blks = vbNullString
'    If UBound(pBlkInf) > 0 Then '必ず入るはずだが念のため
'        For i = 1 To UBound(pBlkInf)
'            Blks = Blks & "'" & pBlkInf(i).BLOCKID & "',"
'        Next i
'        Blks = Left$(Blks, Len(Blks) - 1)
'        sql = "update TBCMY003 set SENDFLAG='0'"
'        sql = sql & " where  (substr(SAMPLEID,1,12) in (" & Blks & "))"
'        sql = sql & " and    (substr(SAMPLEID,1,12)"
'        sql = sql & " in     (select distinct LOTID from TBCMY012"
'        sql = sql & " where  ((REJCAT='A') or (REJCAT='B')) and (ALLSCRAP<>'Y')))"
'        sql = sql & " and    (substr(SAMPLEID,1,15)"
'        sql = sql & " not in (select distinct LOTID || to_char(TOP_POS,'FM000')"
'        sql = sql & " as LOTPOS from TBCMY012"
'        sql = sql & " where ((REJCAT='A') or (REJCAT='B')) and (ALLSCRAP<>'Y')))"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) < 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            GoTo proc_exit
'        End If
'    End If

    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    '' WriteDBLog " ", "End"
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmkc001k_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

Public Function WFExistCheck(tblLackMap() As typ_LackMap, sBLK As String, iPos As Integer, sDirection As String, bAns As Integer) As FUNCTION_RETURN
    Dim iBseq   As Integer
    Dim sql     As String
    Dim rs      As OraDynaset
    Dim c0      As Integer


    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001k_SQL.bas -- Function WFExistCheck"

    WFExistCheck = FUNCTION_RETURN_FAILURE

    bAns = 1    '初期値：ＷＦが存在する

    sql = "select BLOCKSEQ"
    sql = sql & " from  TBCMY011"
    sql = sql & " where (LOTID='" & sBLK & "')"

    If (sDirection = "T") Or (sDirection = "D") Then
        sql = sql & " and (TOP_POS=ANY(select min(TOP_POS) from TBCMY011 where (LOTID='" & sBLK & "') and (TOP_POS>=" & iPos * 10 & ")))"
    Else
        sql = sql & " and (TOP_POS=ANY(select max(TOP_POS) from TBCMY011 where (LOTID='" & sBLK & "') and (TOP_POS<=" & iPos * 10 & ")))"
    End If

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        iBseq = rs("BLOCKSEQ")
        rs.Close
    Else
        bAns = 0    'そこにＷＦはない
        iBseq = -99
        rs.Close
    End If

    '' ブロックＰの欠落チェック
    For c0 = 1 To UBound(tblLackMap)
        With tblLackMap(c0)
'            If (.BLOCKID = sBLK And .REJCAT = "A" And .LACKCNTS <= iBseq And .LACKCNTE >= iBseq) _
'            Or (.BLOCKID = sBLK And .LACKCNTS < 0) Then
            'ﾌﾞﾛｯｸ状態での一部欠量対応 09/02/27 ooba
            If (.BLOCKID = sBLK And (.REJCAT = "A" Or .REJCAT = "E") And .LACKCNTS <= iBseq And .LACKCNTE >= iBseq) _
            Or (.BLOCKID = sBLK And .LACKCNTS < 0) Then
                bAns = -1   'そこは欠落している
                Exit For
            End If
        End With
    Next


    WFExistCheck = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    WFExistCheck = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'(2002/07 s_cmzcF_cmkc001g_SQL.basよりコピー)
'概要      :抜試指示用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:pHinSpec　　　,IO ,typ_HinSpec    　,製品仕様
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim sOT1    As String           '03/05/23 後藤
    Dim sOT2    As String
    Dim rtn     As FUNCTION_RETURN
    Dim sMAI1    As String           '04/06/28
    Dim sMAI2    As String           '04/06/28

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' 製品仕様の取得
    With pHinSpec
        sql = "select E021HWFRMIN,  E021HWFRMAX,  E021HWFRHWYS, E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS,"
        sql = sql & " E025HWFOS2HS, E025HWFOS3HS, E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS,"
        sql = sql & " E029HWFOF2HS, E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS "
        sql = sql & " from  VECME004"
        sql = sql & " where E018HINBAN  ='" & .hin.hinban & "'"
        sql = sql & "   and E018MNOREVNO= " & .hin.mnorevno
        sql = sql & "   and E018FACTORY ='" & .hin.factory & "'"
        sql = sql & "   and E018OPECOND ='" & .hin.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))
        .HWFRHWYS = rs("E021HWFRHWYS")
        .HWFMKHWS = rs("E024HWFMKHWS")
        .HWFONHWS = rs("E025HWFONHWS")
        .HWFOS1HS = rs("E025HWFOS1HS")
        .HWFOS2HS = rs("E025HWFOS2HS")
        .HWFOS3HS = rs("E025HWFOS3HS")
        .HWFDSOHS = rs("E026HWFDSOHS")
        .HWFSPVHS = rs("E028HWFSPVHS")
        .HWFDLHWS = rs("E028HWFDLHWS")
        .HWFOF1HS = rs("E029HWFOF1HS")
        .HWFOF2HS = rs("E029HWFOF2HS")
        .HWFOF3HS = rs("E029HWFOF3HS")
        .HWFOF4HS = rs("E029HWFOF4HS")
        .HWFBM1HS = rs("E029HWFBM1HS")
        .HWFBM2HS = rs("E029HWFBM2HS")
        .HWFBM3HS = rs("E029HWFBM3HS")
        'rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2)   '03/05/23
         'rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2)   '04/07/12 koyama update
        rtn = scmzc_getE036(pHinSpec.hin, sOT1, sOT2, sMAI1, sMAI2)   ''04/07/12 koyama update
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/23
        .HWFOTHER2 = sOT2
        .HWFOTHER1MAI = sMAI1  '04/06/28
        .HWFOTHER2MAI = sMAI2  '04/06/28
        rs.Close

        ''残存酸素仕様取得　03/12/11 ooba START ==============================>
        sql = "select HWFZOHWS from TBCME025 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") Else .HWFZOHWS = " "  '品WF残存酸素保証方法_処
        rs.Close
        ''残存酸素仕様取得　03/12/11 ooba END ================================>

        '' GD仕様取得　05/01/31 ooba START ================================================>
        sql = "select "
        sql = sql & "HWFDENHS, "
        sql = sql & "HWFLDLHS, "
        sql = sql & "HWFDVDHS "
        sql = sql & "from TBCME026 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS") Else .HWFDENHS = " "  '品WFDen保証方法_処
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS") Else .HWFLDLHS = " "  '品WFL/DL保証方法_処
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS") Else .HWFDVDHS = " "  '品WFDVD2保証方法_処

        rs.Close
        '' GD仕様取得　05/01/31 ooba END ==================================================>

        '' SPVNr濃度仕様取得　06/06/08 ooba START ===========================>
        sql = "select "
        sql = sql & "HWFNRHS "          '品WFSPVNR保証方法_処
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START ---注）HWFOF4HSのｴﾘｱが未使用のため、TBCME048.HWFSIRDHSで再利用
        sql = sql & ",HWFSIRDHS "       '軸状転位保証方法＿処
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END   ---注）HWFOF4HSのｴﾘｱが未使用のため、TBCME048.HWFSIRDHSで再利用
        sql = sql & "from TBCME048 "
        sql = sql & "where HINBAN = '" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO = " & .hin.mnorevno & " "
        sql = sql & "and FACTORY = '" & .hin.factory & "' "
        sql = sql & "and OPECOND = '" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HWFNRHS")) = False Then .HWFNRHS = rs("HWFNRHS") Else .HWFNRHS = " "
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD START ---注）HWFOF4HSのｴﾘｱが未使用のため、TBCME048.HWFSIRDHSで再利用
        If IsNull(rs("HWFSIRDHS")) = False Then .HWFOF4HS = rs("HWFSIRDHS") Else .HWFOF4HS = " "
        '◆--- 2010/01/20 SIRD対応 SPK habuki ADD END   ---注）HWFOF4HSのｴﾘｱが未使用のため、TBCME048.HWFSIRDHSで再利用

        rs.Close
        '' SPVNr濃度仕様取得　06/06/08 ooba START ===========================>

        '' WFカット単位取得　05/04/12 ffc)tanabe START =====================================>
        sql = "select "
        sql = sql & "TO_CHAR(WFCUTT) as WFCUTT "
        sql = sql & "from TBCME036 "
        sql = sql & "where HINBAN  ='" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO= " & .hin.mnorevno & " "
        sql = sql & "and FACTORY ='" & .hin.factory & "' "
        sql = sql & "and OPECOND ='" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("WFCUTT")) = False Then .WFCUTUNIT = rs("WFCUTT")   'WFカット単位

        rs.Close
        '' WFカット単位取得　05/04/12 ffc)tanabe END =======================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
        '' エピ仕様取得(OSF、BND)
        sql = "select "
        sql = sql & "HEPOF1HS, "        '品EPOSF1保証方法_処
        sql = sql & "HEPOF2HS, "        '品EPOSF2保証方法_処
        sql = sql & "HEPOF3HS, "        '品EPOSF3保証方法_処
        sql = sql & "HEPBM1HS, "        '品EPBMD1保証方法_処
        sql = sql & "HEPBM2HS, "        '品EPBMD2保証方法_処
        sql = sql & "HEPBM3HS "         '品EPBMD3保証方法_処
        sql = sql & "from TBCME050 "    '製品仕様エピデータ１
        sql = sql & "where HINBAN = '" & .hin.hinban & "' "
        sql = sql & "and MNOREVNO = " & .hin.mnorevno & " "
        sql = sql & "and FACTORY = '" & .hin.factory & "' "
        sql = sql & "and OPECOND = '" & .hin.opecond & "' "

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        If IsNull(rs("HEPOF1HS")) = False Then .HEPOF1HS = rs("HEPOF1HS") Else .HEPOF1HS = " "   '品EPOSF1保証方法_処
        If IsNull(rs("HEPOF2HS")) = False Then .HEPOF2HS = rs("HEPOF2HS") Else .HEPOF2HS = " "   '品EPOSF2保証方法_処
        If IsNull(rs("HEPOF3HS")) = False Then .HEPOF3HS = rs("HEPOF3HS") Else .HEPOF3HS = " "   '品EPOSF3保証方法_処
        If IsNull(rs("HEPBM1HS")) = False Then .HEPBM1HS = rs("HEPBM1HS") Else .HEPBM1HS = " "   '品EPBMD1保証方法_処
        If IsNull(rs("HEPBM2HS")) = False Then .HEPBM2HS = rs("HEPBM2HS") Else .HEPBM2HS = " "   '品EPBMD2保証方法_処
        If IsNull(rs("HEPBM3HS")) = False Then .HEPBM3HS = rs("HEPBM3HS") Else .HEPBM3HS = " "   '品EPBMD3保証方法_処
        rs.Close
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

    End With

    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME037 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME037_SQL.basより移動)
Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL全体
Dim sqlBase As String       'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         'レコード数
Dim i       As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME037"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .DELCLS = rs("DELCLS")           ' 削除区分
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCD = rs("PROCCD")           ' 工程コード
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
            .RPFACT = rs("RPFACT")           ' ねらい品番工場
            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
            .PRODCOND = rs("PRODCOND")       ' 製作条件
            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .DIAMETER = rs("DIAMETER")       ' 直径
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME039」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME039 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME039_SQL.basより移動)
Public Function DBDRV_GetTBCME039(records() As typ_TBCME039, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL全体
Dim sqlBase As String       'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         'レコード数
Dim i       As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACT, OPCOND, LENGTH, USECLASS, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME039"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME039 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 改訂番号
            .FACT = rs("FACT")               ' 工場
            .OPCOND = rs("OPCOND")           ' 操業条件
            .LENGTH = rs("LENGTH")           ' 長さ
            .USECLASS = rs("USECLASS")       ' 使用区分
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME039 = FUNCTION_RETURN_SUCCESS
End Function



'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME041」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME041 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME041_SQL.basより移動)
Public Function DBDRV_GetTBCME041(records() As typ_TBCME041, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL全体
Dim sqlBase As String       'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         'レコード数
Dim i       As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, HINBAN, REVNUM, FACTORY, OPECOND, LENGTH, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCME041"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME041 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .LENGTH = rs("LENGTH")           ' 長さ
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME041 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCME042」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME042 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME042_SQL.basより移動)
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL全体
Dim sqlBase As String       'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         'レコード数
Dim i       As Long

    ''SQLを組み立てる
    ''↓変更START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
'    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
'              " PASSFLAG "   '02/04/16 Yam
'    sqlBase = sqlBase & "From TBCME042"
'    sql = sqlBase
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   XTALCB CRYNUM"      '結晶番号
    sql = sql & "  ,INPOSCB INGOTPOS"   '結晶内開始位置
    sql = sql & "  ,RLENCB LENGTH"      '理論長さ
    sql = sql & "  ,SXLIDCB SXLID"      'SXLID
    sql = sql & "  ,'     ' KRPROCCD"   '管理工程(現行ブランク)
    sql = sql & "  ,GNWKNTCB NOWPROC"   '現在工程
    sql = sql & "  ,'     ' LPKRPROCCD" '最終通過管理工程(現行ブランク)
    sql = sql & "  ,NEWKNTCB LASTPASS"  '最終通過工程
    sql = sql & "  ,LIVKCB DELCLS"      '生死区分
    sql = sql & "  ,LSTCCB LSTATCLS"    '最終状態区分
    sql = sql & "  ,SHOLDCLSCB HOLDCLS" 'ホールド区分
    sql = sql & "  ,HINBCB HINBAN"      '品番
    sql = sql & "  ,REVNUMCB REVNUM"    '製品番号改訂番号
    sql = sql & "  ,FACTORYCB FACTORY"  '工場
    sql = sql & "  ,OPECB OPECOND"      '操業条件
    sql = sql & "  ,FURYCCB BDCAUS"     '不良理由
    sql = sql & "  ,MAICB COUNT"        '枚数
    sql = sql & "  ,TDAYCB REGDATE"     '登録日付
    sql = sql & "  ,KDAYCB UPDDATE"     '更新日付
    sql = sql & "  ,' ' SUMMITSENDFLAG" 'SUMMIT送信フラグ(未使用)
    sql = sql & "  ,SNDKCB SENDFLAG"    '送信フラグ
    sql = sql & "  ,SNDAYCB SENDDATE"   '送信日付
    sql = sql & "  ,' ' PASSFLAG"       'PASSFLAG(未使用)
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    ''↑変更START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            ''↓変更START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
'            .CRYNUM = rs("CRYNUM")           ' 結晶番号
'            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
'            .LENGTH = rs("LENGTH")           ' 長さ
'            .SXLID = rs("SXLID")             ' SXLID
'            .KRPROCCD = rs("KRPROCCD")       ' 管理工程
'            .NOWPROC = rs("NOWPROC")         ' 現在工程
'            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
'            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
'            .DELCLS = rs("DELCLS")           ' 削除区分
'            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
'            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
'            .hinban = rs("HINBAN")           ' 品番
'            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
'            .factory = rs("FACTORY")         ' 工場
'            .opecond = rs("OPECOND")         ' 操業条件
'            .BDCAUS = rs("BDCAUS")           ' 不良理由
'            .COUNT = rs("COUNT")             ' 枚数
'            .REGDATE = rs("REGDATE")         ' 登録日付
'            .UPDDATE = rs("UPDDATE")         ' 更新日付
'            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'            .SENDDATE = rs("SENDDATE")       ' 送信日付
'            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/04/16 Yam
'            If rs("PASSFLAG") = "1" Then
'                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/04/05 Yam
'            End If
            If IsNull(rs("CRYNUM")) = False Then .CRYNUM = rs("CRYNUM")     ' 結晶番号
            If IsNull(rs("INGOTPOS")) = False Then .INGOTPOS = rs("INGOTPOS")     ' 結晶内開始位置
            If IsNull(rs("LENGTH")) = False Then .LENGTH = rs("LENGTH")     ' 長さ
            If IsNull(rs("SXLID")) = False Then .SXLID = rs("SXLID")     ' SXLID
            If IsNull(rs("KRPROCCD")) = False Then .KRPROCCD = rs("KRPROCCD")     ' 管理工程
            If IsNull(rs("NOWPROC")) = False Then .NOWPROC = rs("NOWPROC")     ' 現在工程
            If IsNull(rs("LPKRPROCCD")) = False Then .LPKRPROCCD = rs("LPKRPROCCD")     ' 最終通過管理工程
            If IsNull(rs("LASTPASS")) = False Then .LASTPASS = rs("LASTPASS")     ' 最終通過工程
            If IsNull(rs("DELCLS")) = False Then .DELCLS = rs("DELCLS")     ' 削除区分
            If IsNull(rs("LSTATCLS")) = False Then .LSTATCLS = rs("LSTATCLS")     ' 最終状態区分
            If IsNull(rs("HOLDCLS")) = False Then .HOLDCLS = rs("HOLDCLS")     ' ホールド区分
            If IsNull(rs("HINBAN")) = False Then .hinban = rs("HINBAN")     ' 品番
            If IsNull(rs("REVNUM")) = False Then .REVNUM = rs("REVNUM")     ' 製品番号改訂番号
            If IsNull(rs("FACTORY")) = False Then .factory = rs("FACTORY")     ' 工場
            If IsNull(rs("OPECOND")) = False Then .opecond = rs("OPECOND")     ' 操業条件
            If IsNull(rs("BDCAUS")) = False Then .BDCAUS = rs("BDCAUS")     ' 不良理由
            If IsNull(rs("COUNT")) = False Then .Count = rs("COUNT")     ' 枚数
            If IsNull(rs("REGDATE")) = False Then .REGDATE = rs("REGDATE")     ' 登録日付
            If IsNull(rs("UPDDATE")) = False Then .UPDDATE = rs("UPDDATE")     ' 更新日付
            If IsNull(rs("SUMMITSENDFLAG")) = False Then .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")     '
            If IsNull(rs("SENDFLAG")) = False Then .SENDFLAG = rs("SENDFLAG")     ' 送信フラグ
            If IsNull(rs("SENDDATE")) = False Then .SENDDATE = rs("SENDDATE")     ' 送信日付
            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                If IsNull(rs("PASSFLAG")) = False Then .PASSFLAG = rs("PASSFLAG")     ' 通過フラグ
            End If
            ''↑変更START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
End Function


'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「XSDCW」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCW    ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME044_SQL.basより移動)
Public Function DBDRV_GetTBCME044(records() As typ_XSDCW, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql     As String       'SQL全体
Dim sqlBase As String       'SQL基本部(WHERE節の前まで)
Dim rs      As OraDynaset   'RecordSet
Dim recCnt  As Long         'レコード数
Dim i       As Long

    ''SQLを組み立てる
    'GD追加　05/01/31 ooba
    '2006/08/15 Add エピ先行評価追加対応 SMP)kondoh
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REPSMPLIDCW, XTALCW,INPOSCW ,HINBCW, REVNUMCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, NVL(WFSMPLIDRS1CW,'0') as RS1, NVL(WFSMPLIDRS2CW,'0') as RS2, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, " & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, NVL(WFINDOT1CW,'0') as DOT1, NVL(WFRESOT1CW,'0') as SOT1, " & _
              " WFSMPLIDOT2CW, NVL(WFINDOT2CW,'0') as DOT2, NVL(WFRESOT2CW,'0') as SOT2, NVL(WFSMPLIDAOICW,'0') as sAOI, NVL(WFINDAOICW,'0') as iAOI, NVL(WFRESAOICW,'0') as rAOI, NVL(SMPLNUMCW,'0') sNUM, " & _
              " NVL(SMPLPATCW,'0') as PAT, NVL(TSTAFFCW,'0') as STF, TDAYCW, NVL(KSTAFFCW,'0') as kSTF, KDAYCW, NVL(SNDKCW,'0') as SND, NVL(SNDDAYCW,'2003/09/18') as sDAY, " & _
              " WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW, " & _
              " EPSMPLIDB1CW, NVL(EPINDB1CW,'0') as EPINDB1CW, NVL(EPRESB1CW,'0') as EPRESB1CW," & _
              " EPSMPLIDB2CW, NVL(EPINDB2CW,'0') as EPINDB2CW, NVL(EPRESB2CW,'0') as EPRESB2CW," & _
              " EPSMPLIDB3CW, NVL(EPINDB3CW,'0') as EPINDB3CW, NVL(EPRESB3CW,'0') as EPRESB3CW," & _
              " EPSMPLIDL1CW, NVL(EPINDL1CW,'0') as EPINDL1CW, NVL(EPRESL1CW,'0') as EPRESL1CW," & _
              " EPSMPLIDL2CW, NVL(EPINDL2CW,'0') as EPINDL2CW, NVL(EPRESL2CW,'0') as EPRESL2CW," & _
              " EPSMPLIDL3CW, NVL(EPINDL3CW,'0') as EPINDL3CW, NVL(EPRESL3CW,'0') as EPRESL3CW "
    sqlBase = sqlBase & "From XSDCW"
    sql = sqlBase
'    sql = sql & "WHERE XTALCW =" & sqlOrder & " ORDER BY INPOSCW"
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME044 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
'        With records(i)
'            .CRYNUM = rs("CRYNUM")           ' 結晶番号
'            .INGOTPOS = rs("INGOTPOS")       ' 結晶内位置
'            .SMPKBN = rs("SMPKBN")           ' サンプル区分
'            .SMPLID = rs("SMPLID")           ' サンプルID
'            .hinban = rs("HINBAN")           ' 品番
'            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
'            .factory = rs("FACTORY")         ' 工場
'            .opecond = rs("OPECOND")         ' 操業条件
'            .KTKBN = rs("KTKBN")             ' 確定区分
'            .WFINDRS = rs("WFINDRS")         ' WF検査指示（Rs)
'            .WFINDOI = rs("WFINDOI")         ' WF検査指示（Oi)
'            .WFINDB1 = rs("WFINDB1")         ' WF検査指示（B1)
'            .WFINDB2 = rs("WFINDB2")         ' WF検査指示（B2）
'            .WFINDB3 = rs("WFINDB3")         ' WF検査指示（B3)
'            .WFINDL1 = rs("WFINDL1")         ' WF検査指示（L1)
'            .WFINDL2 = rs("WFINDL2")         ' WF検査指示（L2)
'            .WFINDL3 = rs("WFINDL3")         ' WF検査指示（L3)
'            .WFINDL4 = rs("WFINDL4")         ' WF検査指示（L4)
'            .WFINDDS = rs("WFINDDS")         ' WF検査指示（DS)
'            .WFINDDZ = rs("WFINDDZ")         ' WF検査指示（DZ)
'            .WFINDSP = rs("WFINDSP")         ' WF検査指示（SP)
'            .WFINDDO1 = rs("WFINDDO1")       ' WF検査指示（DO1)
'            .WFINDDO2 = rs("WFINDDO2")       ' WF検査指示（DO2)
'            .WFINDDO3 = rs("WFINDDO3")       ' WF検査指示（DO3)
'            '#####################################################03/05/23 後藤
'            .WFINDOT1 = rs("DOT1")       ' WF検査指示（OT1)
'            .WFINDOT2 = rs("DOT2")       ' WF検査指示（OT2)
'            '#####################################################03/05/23
'            .WFRESRS = rs("WFRESRS")         ' WF検査実績（Rs)
'            .WFRESOI = rs("WFRESOI")         ' WF検査実績（Oi)
'            .WFRESB1 = rs("WFRESB1")         ' WF検査実績（B1)
'            .WFRESB2 = rs("WFRESB2")         ' WF検査実績（B2）
'            .WFRESB3 = rs("WFRESB3")         ' WF検査実績（B3)
'            .WFRESL1 = rs("WFRESL1")         ' WF検査実績（L1)
'            .WFRESL2 = rs("WFRESL2")         ' WF検査実績（L2)
'            .WFRESL3 = rs("WFRESL3")         ' WF検査実績（L3)
'            .WFRESL4 = rs("WFRESL4")         ' WF検査実績（L4)
'            .WFRESDS = rs("WFRESDS")         ' WF検査実績（DS)
'            .WFRESDZ = rs("WFRESDZ")         ' WF検査実績（DZ)
'            .WFRESSP = rs("WFRESSP")         ' WF検査実績（SP)
'            .WFRESDO1 = rs("WFRESDO1")       ' WF検査実績（DO1)
'            .WFRESDO2 = rs("WFRESDO2")       ' WF検査実績（DO2)
'            .WFRESDO3 = rs("WFRESDO3")       ' WF検査実績（DO3)
'            '#####################################################03/05/23 後藤
'            .WFRESOT1 = rs("SOT1")       ' WF検査実績（OT1)
'            .WFRESOT2 = rs("SOT2")       ' WF検査実績（OT2)
'            '#####################################################03/05/23 後藤
'            .REGDATE = rs("REGDATE")         ' 登録日付
'            .UPDDATE = rs("UPDDATE")         ' 更新日付
'            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'            .SENDDATE = rs("SENDDATE")       ' 送信日付
'        End With

        With records(i)
'''''                    .SXLIDCW = rs!SXLIDCW
'''''                    .SMPKBNCW = rs!SMPKBNCW
'''''                    .TBKBNCW = rs!TBKBNCW
'''''                    .REPSMPLIDCW = rs!REPSMPLIDCW
'''''                    .XTALCW = rs!XTALCW
'''''                    .INPOSCW = rs!INPOSCW
'''''                    .HINBCW = rs!HINBCW
'''''                    .REVNUMCW = rs!REVNUMCW
'''''                    .FACTORYCW = rs!FACTORYCW
'''''                    .OPECW = rs!OPECW
'''''                    .KTKBNCW = rs!KTKBNCW
'''''                    .SMCRYNUMCW = rs!SMCRYNUMCW
'''''                    .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
'''''                    .WFSMPLIDRS1CW = rs!RS1
'''''                    .WFSMPLIDRS2CW = rs!rs2
'''''                    .WFINDRSCW = rs!WFINDRSCW
'''''                    .WFRESRS1CW = rs!WFRESRS1CW
'''''                    .WFSMPLIDOICW = rs!WFSMPLIDOICW
'''''                    .WFINDOICW = rs!WFINDOICW
'''''                    .WFRESOICW = rs!WFRESOICW
'''''                    .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
'''''                    .WFINDB1CW = rs!WFINDB1CW
'''''                    .WFRESB1CW = rs!WFRESB1CW
'''''                    .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
'''''                    .WFINDB2CW = rs!WFINDB2CW
'''''                    .WFRESB2CW = rs!WFRESB2CW
'''''                    .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
'''''                    .WFINDB3CW = rs!WFINDB3CW
'''''                    .WFRESB3CW = rs!WFRESB3CW
'''''                    .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
'''''                    .WFINDL1CW = rs!WFINDL1CW
'''''                    .WFRESL1CW = rs!WFRESL1CW
'''''                    .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
'''''                    .WFINDL2CW = rs!WFINDL2CW
'''''                    .WFRESL2CW = rs!WFRESL2CW
'''''                    .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
'''''                    .WFINDL3CW = rs!WFINDL3CW
'''''                    .WFRESL3CW = rs!WFRESL3CW
'''''                    .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
'''''                    .WFINDL4CW = rs!WFINDL4CW
'''''                    .WFRESL4CW = rs!WFRESL4CW
'''''                    .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
'''''                    .WFINDDSCW = rs!WFINDDSCW
'''''                    .WFRESDSCW = rs!WFRESDSCW
'''''                    .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
'''''                    .WFINDDZCW = rs!WFINDDZCW
'''''                    .WFRESDZCW = rs!WFRESDZCW
'''''                    .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
'''''                    .WFINDSPCW = rs!WFINDSPCW
'''''                    .WFRESSPCW = rs!WFRESSPCW
'''''                    .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
'''''                    .WFINDDO1CW = rs!WFINDDO1CW
'''''                    .WFRESDO1CW = rs!WFRESDO1CW
'''''                    .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
'''''                    .WFINDDO2CW = rs!WFINDDO2CW
'''''                    .WFRESDO2CW = rs!WFRESDO2CW
'''''                    .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
'''''                    .WFINDDO3CW = rs!WFINDDO3CW
'''''                    .WFRESDO3CW = rs!WFRESDO3CW
'''''                    .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
'''''                    .WFINDOT1CW = rs!DOT1
'''''                    .WFRESOT1CW = rs!sOT1
'''''                    .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
'''''                    .WFINDOT2CW = rs!DOT2
'''''                    .WFRESOT2CW = rs!sOT2
'''''''                    tHin.hinban = .hinban
'''''''                    tHin.factory = .factory
'''''''                    tHin.mnorevno = .REVNUM
'''''''                    tHin.opecond = .opecond
''''''                    rtn = scmzc_getE036(tHin, sOT1, sOT2)
''''''                    If rtn = FUNCTION_RETURN_FAILURE Then
''''''                        rs.Close
''''''                        DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
''''''                        GoTo proc_exit
''''''                    End If
''''''                    If sOT1 = "1" Then
''''''                        .WFINDOT1CW = rs!DOT1 '03/05/26
''''''                    Else
''''''                        .WFINDOT1CW = 0 '03/05/26
''''''                    End If
''''''                    If sOT2 = "1" Then
''''''                        .WFINDOT2CW = rs!DOT2 '03/05/26
''''''                    Else
''''''                        .WFINDOT2CW = 0 '03/05/26
''''''                    End If
'''''                    .WFSMPLIDAOICW = rs!sAOI
'''''                    .WFINDAOICW = rs!iAOI
'''''                    .WFRESAOICW = rs!rAOI
'''''                    .SMPLNUMCW = rs!sNUM
'''''                    .SMPLPATCW = rs!PAT
'''''                    .TSTAFFCW = rs!STF
'''''                    .TDAYCW = rs!TDAYCW
'''''                    .KSTAFFCW = rs!kSTF
'''''                    .KDAYCW = rs!KDAYCW
'''''                    .SNDKCW = rs!SND
'''''                    .SNDDAYCW = rs!sDAY

                    If IsNull(rs!SXLIDCW) = False Then .SXLIDCW = rs!SXLIDCW
                    If IsNull(rs!SMPKBNCW) = False Then .SMPKBNCW = rs!SMPKBNCW
                    If IsNull(rs!TBKBNCW) = False Then .TBKBNCW = rs!TBKBNCW
                    If IsNull(rs!REPSMPLIDCW) = False Then .REPSMPLIDCW = rs!REPSMPLIDCW
                    If IsNull(rs!XTALCW) = False Then .XTALCW = rs!XTALCW
                    If IsNull(rs!INPOSCW) = False Then .INPOSCW = rs!INPOSCW
                    If IsNull(rs!HINBCW) = False Then .HINBCW = rs!HINBCW
                    If IsNull(rs!REVNUMCW) = False Then .REVNUMCW = rs!REVNUMCW
                    If IsNull(rs!FACTORYCW) = False Then .FACTORYCW = rs!FACTORYCW
                    If IsNull(rs!OPECW) = False Then .OPECW = rs!OPECW
                    If IsNull(rs!KTKBNCW) = False Then .KTKBNCW = rs!KTKBNCW
                    If IsNull(rs!SMCRYNUMCW) = False Then .SMCRYNUMCW = rs!SMCRYNUMCW
                    If IsNull(rs!WFSMPLIDRSCW) = False Then .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                    If IsNull(rs!rs1) = False Then .WFSMPLIDRS1CW = rs!rs1
                    If IsNull(rs!rs2) = False Then .WFSMPLIDRS2CW = rs!rs2
                    If IsNull(rs!WFINDRSCW) = False Then .WFINDRSCW = rs!WFINDRSCW
                    If IsNull(rs!WFRESRS1CW) = False Then .WFRESRS1CW = rs!WFRESRS1CW
                    If IsNull(rs!WFSMPLIDOICW) = False Then .WFSMPLIDOICW = rs!WFSMPLIDOICW
                    If IsNull(rs!WFINDOICW) = False Then .WFINDOICW = rs!WFINDOICW
                    If IsNull(rs!WFRESOICW) = False Then .WFRESOICW = rs!WFRESOICW
                    If IsNull(rs!WFSMPLIDB1CW) = False Then .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                    If IsNull(rs!WFINDB1CW) = False Then .WFINDB1CW = rs!WFINDB1CW
                    If IsNull(rs!WFRESB1CW) = False Then .WFRESB1CW = rs!WFRESB1CW
                    If IsNull(rs!WFSMPLIDB2CW) = False Then .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                    If IsNull(rs!WFINDB2CW) = False Then .WFINDB2CW = rs!WFINDB2CW
                    If IsNull(rs!WFRESB2CW) = False Then .WFRESB2CW = rs!WFRESB2CW
                    If IsNull(rs!WFSMPLIDB3CW) = False Then .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                    If IsNull(rs!WFINDB3CW) = False Then .WFINDB3CW = rs!WFINDB3CW
                    If IsNull(rs!WFRESB3CW) = False Then .WFRESB3CW = rs!WFRESB3CW
                    If IsNull(rs!WFSMPLIDL1CW) = False Then .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                    If IsNull(rs!WFINDL1CW) = False Then .WFINDL1CW = rs!WFINDL1CW
                    If IsNull(rs!WFRESL1CW) = False Then .WFRESL1CW = rs!WFRESL1CW
                    If IsNull(rs!WFSMPLIDL2CW) = False Then .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                    If IsNull(rs!WFINDL2CW) = False Then .WFINDL2CW = rs!WFINDL2CW
                    If IsNull(rs!WFRESL2CW) = False Then .WFRESL2CW = rs!WFRESL2CW
                    If IsNull(rs!WFSMPLIDL3CW) = False Then .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                    If IsNull(rs!WFINDL3CW) = False Then .WFINDL3CW = rs!WFINDL3CW
                    If IsNull(rs!WFRESL3CW) = False Then .WFRESL3CW = rs!WFRESL3CW
                    If IsNull(rs!WFSMPLIDL4CW) = False Then .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                    If IsNull(rs!WFINDL4CW) = False Then .WFINDL4CW = rs!WFINDL4CW
                    If IsNull(rs!WFRESL4CW) = False Then .WFRESL4CW = rs!WFRESL4CW
                    If IsNull(rs!WFSMPLIDDSCW) = False Then .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                    If IsNull(rs!WFINDDSCW) = False Then .WFINDDSCW = rs!WFINDDSCW
                    If IsNull(rs!WFRESDSCW) = False Then .WFRESDSCW = rs!WFRESDSCW
                    If IsNull(rs!WFSMPLIDDZCW) = False Then .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                    If IsNull(rs!WFINDDZCW) = False Then .WFINDDZCW = rs!WFINDDZCW
                    If IsNull(rs!WFRESDZCW) = False Then .WFRESDZCW = rs!WFRESDZCW
                    If IsNull(rs!WFSMPLIDSPCW) = False Then .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                    If IsNull(rs!WFINDSPCW) = False Then .WFINDSPCW = rs!WFINDSPCW
                    If IsNull(rs!WFRESSPCW) = False Then .WFRESSPCW = rs!WFRESSPCW
                    If IsNull(rs!WFSMPLIDDO1CW) = False Then .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                    If IsNull(rs!WFINDDO1CW) = False Then .WFINDDO1CW = rs!WFINDDO1CW
                    If IsNull(rs!WFRESDO1CW) = False Then .WFRESDO1CW = rs!WFRESDO1CW
                    If IsNull(rs!WFSMPLIDDO2CW) = False Then .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                    If IsNull(rs!WFINDDO2CW) = False Then .WFINDDO2CW = rs!WFINDDO2CW
                    If IsNull(rs!WFRESDO2CW) = False Then .WFRESDO2CW = rs!WFRESDO2CW
                    If IsNull(rs!WFSMPLIDDO3CW) = False Then .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                    If IsNull(rs!WFINDDO3CW) = False Then .WFINDDO3CW = rs!WFINDDO3CW
                    If IsNull(rs!WFRESDO3CW) = False Then .WFRESDO3CW = rs!WFRESDO3CW
                    If IsNull(rs!WFSMPLIDOT1CW) = False Then .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                    If IsNull(rs!DOT1) = False Then .WFINDOT1CW = rs!DOT1
                    If IsNull(rs!sOT1) = False Then .WFRESOT1CW = rs!sOT1
                    If IsNull(rs!WFSMPLIDOT2CW) = False Then .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                    If IsNull(rs!DOT2) = False Then .WFINDOT2CW = rs!DOT2
                    If IsNull(rs!sOT2) = False Then .WFRESOT2CW = rs!sOT2

                    If IsNull(rs!sAOI) = False Then .WFSMPLIDAOICW = rs!sAOI
                    If IsNull(rs!iAOI) = False Then .WFINDAOICW = rs!iAOI
                    If IsNull(rs!rAOI) = False Then .WFRESAOICW = rs!rAOI
                    If IsNull(rs!sNum) = False Then .SMPLNUMCW = rs!sNum
                    If IsNull(rs!PAT) = False Then .SMPLPATCW = rs!PAT
                    If IsNull(rs!STF) = False Then .TSTAFFCW = rs!STF
                    If IsNull(rs!TDAYCW) = False Then .TDAYCW = rs!TDAYCW
                    If IsNull(rs!kSTF) = False Then .KSTAFFCW = rs!kSTF
                    If IsNull(rs!KDAYCW) = False Then .KDAYCW = rs!KDAYCW
                    If IsNull(rs!SND) = False Then .SNDKCW = rs!SND
                    If IsNull(rs!sDay) = False Then .SNDDAYCW = rs!sDay

                    '' GD追加　05/01/31 ooba START ===========================================>
                    If IsNull(rs!WFSMPLIDGDCW) = False Then .WFSMPLIDGDCW = rs!WFSMPLIDGDCW
                    If IsNull(rs!WFINDGDCW) = False Then .WFINDGDCW = rs!WFINDGDCW
                    If IsNull(rs!WFRESGDCW) = False Then .WFRESGDCW = rs!WFRESGDCW
                    If IsNull(rs!WFHSGDCW) = False Then .WFHSGDCW = rs!WFHSGDCW
                    '' GD追加　05/01/31 ooba END =============================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                    If IsNull(rs!EPSMPLIDB1CW) = False Then .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                    If IsNull(rs!EPINDB1CW) = False Then .EPINDB1CW = rs!EPINDB1CW
                    If IsNull(rs!EPRESB1CW) = False Then .EPRESB1CW = rs!EPRESB1CW
                    If IsNull(rs!EPSMPLIDB2CW) = False Then .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                    If IsNull(rs!EPINDB2CW) = False Then .EPINDB2CW = rs!EPINDB2CW
                    If IsNull(rs!EPRESB2CW) = False Then .EPRESB2CW = rs!EPRESB2CW
                    If IsNull(rs!EPSMPLIDB3CW) = False Then .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                    If IsNull(rs!EPINDB3CW) = False Then .EPINDB3CW = rs!EPINDB3CW
                    If IsNull(rs!EPRESB3CW) = False Then .EPRESB3CW = rs!EPRESB3CW
                    If IsNull(rs!EPSMPLIDL1CW) = False Then .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                    If IsNull(rs!EPINDL1CW) = False Then .EPINDL1CW = rs!EPINDL1CW
                    If IsNull(rs!EPRESL1CW) = False Then .EPRESL1CW = rs!EPRESL1CW
                    If IsNull(rs!EPSMPLIDL2CW) = False Then .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                    If IsNull(rs!EPINDL2CW) = False Then .EPINDL2CW = rs!EPINDL2CW
                    If IsNull(rs!EPRESL2CW) = False Then .EPRESL2CW = rs!EPRESL2CW
                    If IsNull(rs!EPSMPLIDL3CW) = False Then .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                    If IsNull(rs!EPINDL3CW) = False Then .EPINDL3CW = rs!EPINDL3CW
                    If IsNull(rs!EPRESL3CW) = False Then .EPRESL3CW = rs!EPRESL3CW
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                End With

        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME044 = FUNCTION_RETURN_SUCCESS
End Function

'概要      :XSDCWよりﾃﾞｰﾀ取得
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :sCryNum       ,I  ,String       ,結晶番号
'          :records()     ,O  ,typ_XSDCW    ,抽出レコード
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :08/02/04 ooba
Public Function DBDRV_GetXSDCW(sCryNum As String, records() As typ_XSDCW) As FUNCTION_RETURN

    Dim sql     As String       'SQL全体
    Dim rs      As OraDynaset   'RecordSet
    Dim recCnt  As Long         'ﾚｺｰﾄﾞ数
    Dim i       As Long

    sql = "SELECT "
    sql = sql & "SXLIDCB, "
    sql = sql & "SXLIDCW, "
    sql = sql & "NVL(SMPKBNCW,'T') as SMPKBNCW, "
    sql = sql & "NVL(TBKBNCW,'T') as TBKBN, "
    sql = sql & "NVL(REPSMPLIDCW,' ') as REPSMPLIDCW, "
    sql = sql & "NVL(XTALCB,' ') as XTALCB, "
    sql = sql & "NVL(INGOTPOS,0) as INGOTPOS, "
    sql = sql & "NVL(INPOSCW,0), "
    sql = sql & "NVL(HINBCB,' ') as HINBCB, "
    sql = sql & "NVL(REVNUMCB,0) as REVNUMCB, "
    sql = sql & "NVL(FACTORYCB,' ') as FACTORYCB, "
    sql = sql & "NVL(OPECB,' ') as OPECB, "
    sql = sql & "NVL(HINBCW,' ') as HINBCW, "
    sql = sql & "NVL(REVNUMCW,0) as REVNUMCW, "
    sql = sql & "NVL(FACTORYCW,' ') as FACTORYCW, "
    sql = sql & "NVL(OPECW,' ') as OPECW, "
    sql = sql & "NVL(KTKBNCW,' ') as KTKBNCW, "
    sql = sql & "NVL(SMCRYNUMCW,' ') as SMCRYNUMCW, "
    sql = sql & "NVL(WFSMPLIDRSCW,' ') as WFSMPLIDRSCW, "
    sql = sql & "NVL(WFSMPLIDRS1CW,' ') as WFSMPLIDRS1CW, "
    sql = sql & "NVL(WFSMPLIDRS2CW,' ') as WFSMPLIDRS2CW, "
    sql = sql & "NVL(WFINDRSCW,'0') as WFINDRSCW, "
    sql = sql & "NVL(WFRESRS1CW,'0') as WFRESRS1CW, "
    sql = sql & "NVL(WFRESRS2CW,'0') as WFRESRS2CW, "
    sql = sql & "NVL(WFSMPLIDOICW,' ') as WFSMPLIDOICW, "
    sql = sql & "NVL(WFINDOICW,'0') as WFINDOICW, "
    sql = sql & "NVL(WFRESOICW,'0') as WFRESOICW, "
    sql = sql & "NVL(WFSMPLIDB1CW,' ') as WFSMPLIDB1CW, "
    sql = sql & "NVL(WFINDB1CW,'0') as WFINDB1CW, "
    sql = sql & "NVL(WFRESB1CW,'0') as WFRESB1CW, "
    sql = sql & "NVL(WFSMPLIDB2CW,' ') as WFSMPLIDB2CW, "
    sql = sql & "NVL(WFINDB2CW,'0') as WFINDB2CW, "
    sql = sql & "NVL(WFRESB2CW,'0') as WFRESB2CW, "
    sql = sql & "NVL(WFSMPLIDB3CW,' ') as WFSMPLIDB3CW, "
    sql = sql & "NVL(WFINDB3CW,'0') as WFINDB3CW, "
    sql = sql & "NVL(WFRESB3CW,'0') as WFRESB3CW, "
    sql = sql & "NVL(WFSMPLIDL1CW,' ') as WFSMPLIDL1CW, "
    sql = sql & "NVL(WFINDL1CW,'0') as WFINDL1CW, "
    sql = sql & "NVL(WFRESL1CW,'0') as WFRESL1CW, "
    sql = sql & "NVL(WFSMPLIDL2CW,' ') as WFSMPLIDL2CW, "
    sql = sql & "NVL(WFINDL2CW,'0') as WFINDL2CW, "
    sql = sql & "NVL(WFRESL2CW,'0') as WFRESL2CW, "
    sql = sql & "NVL(WFSMPLIDL3CW,' ') as WFSMPLIDL3CW, "
    sql = sql & "NVL(WFINDL3CW,'0') as WFINDL3CW, "
    sql = sql & "NVL(WFRESL3CW,'0') as WFRESL3CW, "
    sql = sql & "NVL(WFSMPLIDL4CW,' ') as WFSMPLIDL4CW, "
    sql = sql & "NVL(WFINDL4CW,'0') as WFINDL4CW, "
    sql = sql & "NVL(WFRESL4CW,'0') as WFRESL4CW, "
    sql = sql & "NVL(WFSMPLIDDSCW,' ') as WFSMPLIDDSCW, "
    sql = sql & "NVL(WFINDDSCW,'0') as WFINDDSCW, "
    sql = sql & "NVL(WFRESDSCW,'0') as WFRESDSCW, "
    sql = sql & "NVL(WFSMPLIDDZCW,' ') as WFSMPLIDDZCW, "
    sql = sql & "NVL(WFINDDZCW,'0') as WFINDDZCW, "
    sql = sql & "NVL(WFRESDZCW,'0') as WFRESDZCW, "
    sql = sql & "NVL(WFSMPLIDSPCW,' ') as WFSMPLIDSPCW, "
    sql = sql & "NVL(WFINDSPCW,'0') as WFINDSPCW, "
    sql = sql & "NVL(WFRESSPCW,'0') as WFRESSPCW, "
    sql = sql & "NVL(WFSMPLIDDO1CW,' ') as WFSMPLIDDO1CW, "
    sql = sql & "NVL(WFINDDO1CW,'0') as WFINDDO1CW, "
    sql = sql & "NVL(WFRESDO1CW,'0') as WFRESDO1CW, "
    sql = sql & "NVL(WFSMPLIDDO2CW,' ') as WFSMPLIDDO2CW, "
    sql = sql & "NVL(WFINDDO2CW,'0') as WFINDDO2CW, "
    sql = sql & "NVL(WFRESDO2CW,'0') as WFRESDO2CW, "
    sql = sql & "NVL(WFSMPLIDDO3CW,' ') as WFSMPLIDDO3CW, "
    sql = sql & "NVL(WFINDDO3CW,'0') as WFINDDO3CW, "
    sql = sql & "NVL(WFRESDO3CW,'0') as WFRESDO3CW, "
    sql = sql & "NVL(WFSMPLIDOT1CW,' ') as WFSMPLIDOT1CW, "
    sql = sql & "NVL(WFINDOT1CW,'0') as WFINDOT1CW, "
    sql = sql & "NVL(WFRESOT1CW,'0') as WFRESOT1CW, "
    sql = sql & "NVL(WFSMPLIDOT2CW,' ') as WFSMPLIDOT2CW, "
    sql = sql & "NVL(WFINDOT2CW,'0') as WFINDOT2CW, "
    sql = sql & "NVL(WFRESOT2CW,'0') as WFRESOT2CW, "
    sql = sql & "NVL(WFSMPLIDAOICW,' ') as WFSMPLIDAOICW, "
    sql = sql & "NVL(WFINDAOICW,'0') as WFINDAOICW, "
    sql = sql & "NVL(WFRESAOICW,'0') as WFRESAOICW, "
    sql = sql & "NVL(SMPLNUMCW,0) as SMPLNUMCW, "
    sql = sql & "NVL(SMPLPATCW,' ') as SMPLPATCW, "
    sql = sql & "NVL(LIVKCW,'0') as LIVKCW, "
    sql = sql & "NVL(WFSMPLIDGDCW,' ') as WFSMPLIDGDCW, "
    sql = sql & "NVL(WFINDGDCW,'0') as WFINDGDCW, "
    sql = sql & "NVL(WFRESGDCW,'0') as WFRESGDCW, "
    sql = sql & "NVL(WFHSGDCW,'0') as WFHSGDCW, "
    sql = sql & "NVL(EPSMPLIDB1CW,' ') as EPSMPLIDB1CW, "
    sql = sql & "NVL(EPINDB1CW,'0') as EPINDB1CW, "
    sql = sql & "NVL(EPRESB1CW,'0') as EPRESB1CW, "
    sql = sql & "NVL(EPSMPLIDB2CW,' ') as EPSMPLIDB2CW, "
    sql = sql & "NVL(EPINDB2CW,'0') as EPINDB2CW, "
    sql = sql & "NVL(EPRESB2CW,'0') as EPRESB2CW, "
    sql = sql & "NVL(EPSMPLIDB3CW,' ') as EPSMPLIDB3CW, "
    sql = sql & "NVL(EPINDB3CW,'0') as EPINDB3CW, "
    sql = sql & "NVL(EPRESB3CW,'0') as EPRESB3CW, "
    sql = sql & "NVL(EPSMPLIDL1CW,' ') as EPSMPLIDL1CW, "
    sql = sql & "NVL(EPINDL1CW,'0') as EPINDL1CW, "
    sql = sql & "NVL(EPRESL1CW,'0') as EPRESL1CW, "
    sql = sql & "NVL(EPSMPLIDL2CW,' ') as EPSMPLIDL2CW, "
    sql = sql & "NVL(EPINDL2CW,'0') as EPINDL2CW, "
    sql = sql & "NVL(EPRESL2CW,'0') as EPRESL2CW, "
    sql = sql & "NVL(EPSMPLIDL3CW,' ') as EPSMPLIDL3CW, "
    sql = sql & "NVL(EPINDL3CW,'0') as EPINDL3CW, "
    sql = sql & "NVL(EPRESL3CW,'0') as EPRESL3CW "
    sql = sql & "FROM "
    sql = sql & "    (SELECT SXLIDCB, "
    sql = sql & "     XTALCB, "
    sql = sql & "     INPOSCB as INGOTPOS, "
    sql = sql & "     HINBCB, "
    sql = sql & "     REVNUMCB, "
    sql = sql & "     FACTORYCB, "
    sql = sql & "     OPECB "
    sql = sql & "     FROM XSDCB "
    sql = sql & "     WHERE XTALCB = '" & sCryNum & "' "
    sql = sql & "     AND LIVKCB = '0' "
    sql = sql & "    ), "
    sql = sql & "    (SELECT SXLIDCW, "
    sql = sql & "     SMPKBNCW, "
    sql = sql & "     TBKBNCW, "
    sql = sql & "     REPSMPLIDCW, "
    sql = sql & "     INPOSCW, "
    sql = sql & "     HINBCW, "
    sql = sql & "     REVNUMCW, "
    sql = sql & "     FACTORYCW, "
    sql = sql & "     OPECW, "
    sql = sql & "     KTKBNCW, "
    sql = sql & "     SMCRYNUMCW, "
    sql = sql & "     WFSMPLIDRSCW, "
    sql = sql & "     WFSMPLIDRS1CW, "
    sql = sql & "     WFSMPLIDRS2CW, "
    sql = sql & "     WFINDRSCW, "
    sql = sql & "     WFRESRS1CW, "
    sql = sql & "     WFRESRS2CW, "
    sql = sql & "     WFSMPLIDOICW, "
    sql = sql & "     WFINDOICW, "
    sql = sql & "     WFRESOICW, "
    sql = sql & "     WFSMPLIDB1CW, "
    sql = sql & "     WFINDB1CW, "
    sql = sql & "     WFRESB1CW, "
    sql = sql & "     WFSMPLIDB2CW, "
    sql = sql & "     WFINDB2CW, "
    sql = sql & "     WFRESB2CW, "
    sql = sql & "     WFSMPLIDB3CW, "
    sql = sql & "     WFINDB3CW, "
    sql = sql & "     WFRESB3CW, "
    sql = sql & "     WFSMPLIDL1CW, "
    sql = sql & "     WFINDL1CW, "
    sql = sql & "     WFRESL1CW, "
    sql = sql & "     WFSMPLIDL2CW, "
    sql = sql & "     WFINDL2CW, "
    sql = sql & "     WFRESL2CW, "
    sql = sql & "     WFSMPLIDL3CW, "
    sql = sql & "     WFINDL3CW, "
    sql = sql & "     WFRESL3CW, "
    sql = sql & "     WFSMPLIDL4CW, "
    sql = sql & "     WFINDL4CW, "
    sql = sql & "     WFRESL4CW, "
    sql = sql & "     WFSMPLIDDSCW, "
    sql = sql & "     WFINDDSCW, "
    sql = sql & "     WFRESDSCW, "
    sql = sql & "     WFSMPLIDDZCW, "
    sql = sql & "     WFINDDZCW, "
    sql = sql & "     WFRESDZCW, "
    sql = sql & "     WFSMPLIDSPCW, "
    sql = sql & "     WFINDSPCW, "
    sql = sql & "     WFRESSPCW, "
    sql = sql & "     WFSMPLIDDO1CW, "
    sql = sql & "     WFINDDO1CW, "
    sql = sql & "     WFRESDO1CW, "
    sql = sql & "     WFSMPLIDDO2CW, "
    sql = sql & "     WFINDDO2CW, "
    sql = sql & "     WFRESDO2CW, "
    sql = sql & "     WFSMPLIDDO3CW, "
    sql = sql & "     WFINDDO3CW, "
    sql = sql & "     WFRESDO3CW, "
    sql = sql & "     WFSMPLIDOT1CW, "
    sql = sql & "     WFINDOT1CW, "
    sql = sql & "     WFRESOT1CW, "
    sql = sql & "     WFSMPLIDOT2CW, "
    sql = sql & "     WFINDOT2CW, "
    sql = sql & "     WFRESOT2CW, "
    sql = sql & "     WFSMPLIDAOICW, "
    sql = sql & "     WFINDAOICW, "
    sql = sql & "     WFRESAOICW, "
    sql = sql & "     SMPLNUMCW, "
    sql = sql & "     SMPLPATCW, "
    sql = sql & "     LIVKCW, "
    sql = sql & "     WFSMPLIDGDCW, "
    sql = sql & "     WFINDGDCW, "
    sql = sql & "     WFRESGDCW, "
    sql = sql & "     WFHSGDCW, "
    sql = sql & "     EPSMPLIDB1CW, "
    sql = sql & "     EPINDB1CW, "
    sql = sql & "     EPRESB1CW, "
    sql = sql & "     EPSMPLIDB2CW, "
    sql = sql & "     EPINDB2CW, "
    sql = sql & "     EPRESB2CW, "
    sql = sql & "     EPSMPLIDB3CW, "
    sql = sql & "     EPINDB3CW, "
    sql = sql & "     EPRESB3CW, "
    sql = sql & "     EPSMPLIDL1CW, "
    sql = sql & "     EPINDL1CW, "
    sql = sql & "     EPRESL1CW, "
    sql = sql & "     EPSMPLIDL2CW, "
    sql = sql & "     EPINDL2CW, "
    sql = sql & "     EPRESL2CW, "
    sql = sql & "     EPSMPLIDL3CW, "
    sql = sql & "     EPINDL3CW, "
    sql = sql & "     EPRESL3CW "
    sql = sql & "     FROM XSDCW "
    sql = sql & "     WHERE XTALCW = '" & sCryNum & "' "
    sql = sql & "     AND TBKBNCW = 'T' "
    sql = sql & "    ) "
'    sql = sql & "WHERE INGOTPOS = INPOSCW(+) "
    sql = sql & "WHERE SXLIDCB = SXLIDCW(+) "           '08/07/10 ooba
    sql = sql & "AND NVL(LIVKCW,'0') = '0' "
    
    sql = sql & "UNION ALL "
    
    sql = sql & "SELECT "
    sql = sql & "SXLIDCB, "
    sql = sql & "SXLIDCW, "
    sql = sql & "NVL(SMPKBNCW,'B') as SMPKBNCW, "
    sql = sql & "NVL(TBKBNCW,'B') as TBKBN, "
    sql = sql & "NVL(REPSMPLIDCW,' ') as REPSMPLIDCW, "
    sql = sql & "NVL(XTALCB,' ') as XTALCB, "
    sql = sql & "NVL(INGOTPOS,0) as INGOTPOS, "
    sql = sql & "NVL(INPOSCW,0), "
    sql = sql & "NVL(HINBCB,' ') as HINBCB, "
    sql = sql & "NVL(REVNUMCB,0) as REVNUMCB, "
    sql = sql & "NVL(FACTORYCB,' ') as FACTORYCB, "
    sql = sql & "NVL(OPECB,' ') as OPECB, "
    sql = sql & "NVL(HINBCW,' ') as HINBCW, "
    sql = sql & "NVL(REVNUMCW,0) as REVNUMCW, "
    sql = sql & "NVL(FACTORYCW,' ') as FACTORYCW, "
    sql = sql & "NVL(OPECW,' ') as OPECW, "
    sql = sql & "NVL(KTKBNCW,' ') as KTKBNCW, "
    sql = sql & "NVL(SMCRYNUMCW,' ') as SMCRYNUMCW, "
    sql = sql & "NVL(WFSMPLIDRSCW,' ') as WFSMPLIDRSCW, "
    sql = sql & "NVL(WFSMPLIDRS1CW,' ') as WFSMPLIDRS1CW, "
    sql = sql & "NVL(WFSMPLIDRS2CW,' ') as WFSMPLIDRS2CW, "
    sql = sql & "NVL(WFINDRSCW,'0') as WFINDRSCW, "
    sql = sql & "NVL(WFRESRS1CW,'0') as WFRESRS1CW, "
    sql = sql & "NVL(WFRESRS2CW,'0') as WFRESRS2CW, "
    sql = sql & "NVL(WFSMPLIDOICW,' ') as WFSMPLIDOICW, "
    sql = sql & "NVL(WFINDOICW,'0') as WFINDOICW, "
    sql = sql & "NVL(WFRESOICW,'0') as WFRESOICW, "
    sql = sql & "NVL(WFSMPLIDB1CW,' ') as WFSMPLIDB1CW, "
    sql = sql & "NVL(WFINDB1CW,'0') as WFINDB1CW, "
    sql = sql & "NVL(WFRESB1CW,'0') as WFRESB1CW, "
    sql = sql & "NVL(WFSMPLIDB2CW,' ') as WFSMPLIDB2CW, "
    sql = sql & "NVL(WFINDB2CW,'0') as WFINDB2CW, "
    sql = sql & "NVL(WFRESB2CW,'0') as WFRESB2CW, "
    sql = sql & "NVL(WFSMPLIDB3CW,' ') as WFSMPLIDB3CW, "
    sql = sql & "NVL(WFINDB3CW,'0') as WFINDB3CW, "
    sql = sql & "NVL(WFRESB3CW,'0') as WFRESB3CW, "
    sql = sql & "NVL(WFSMPLIDL1CW,' ') as WFSMPLIDL1CW, "
    sql = sql & "NVL(WFINDL1CW,'0') as WFINDL1CW, "
    sql = sql & "NVL(WFRESL1CW,'0') as WFRESL1CW, "
    sql = sql & "NVL(WFSMPLIDL2CW,' ') as WFSMPLIDL2CW, "
    sql = sql & "NVL(WFINDL2CW,'0') as WFINDL2CW, "
    sql = sql & "NVL(WFRESL2CW,'0') as WFRESL2CW, "
    sql = sql & "NVL(WFSMPLIDL3CW,' ') as WFSMPLIDL3CW, "
    sql = sql & "NVL(WFINDL3CW,'0') as WFINDL3CW, "
    sql = sql & "NVL(WFRESL3CW,'0') as WFRESL3CW, "
    sql = sql & "NVL(WFSMPLIDL4CW,' ') as WFSMPLIDL4CW, "
    sql = sql & "NVL(WFINDL4CW,'0') as WFINDL4CW, "
    sql = sql & "NVL(WFRESL4CW,'0') as WFRESL4CW, "
    sql = sql & "NVL(WFSMPLIDDSCW,' ') as WFSMPLIDDSCW, "
    sql = sql & "NVL(WFINDDSCW,'0') as WFINDDSCW, "
    sql = sql & "NVL(WFRESDSCW,'0') as WFRESDSCW, "
    sql = sql & "NVL(WFSMPLIDDZCW,' ') as WFSMPLIDDZCW, "
    sql = sql & "NVL(WFINDDZCW,'0') as WFINDDZCW, "
    sql = sql & "NVL(WFRESDZCW,'0') as WFRESDZCW, "
    sql = sql & "NVL(WFSMPLIDSPCW,' ') as WFSMPLIDSPCW, "
    sql = sql & "NVL(WFINDSPCW,'0') as WFINDSPCW, "
    sql = sql & "NVL(WFRESSPCW,'0') as WFRESSPCW, "
    sql = sql & "NVL(WFSMPLIDDO1CW,' ') as WFSMPLIDDO1CW, "
    sql = sql & "NVL(WFINDDO1CW,'0') as WFINDDO1CW, "
    sql = sql & "NVL(WFRESDO1CW,'0') as WFRESDO1CW, "
    sql = sql & "NVL(WFSMPLIDDO2CW,' ') as WFSMPLIDDO2CW, "
    sql = sql & "NVL(WFINDDO2CW,'0') as WFINDDO2CW, "
    sql = sql & "NVL(WFRESDO2CW,'0') as WFRESDO2CW, "
    sql = sql & "NVL(WFSMPLIDDO3CW,' ') as WFSMPLIDDO3CW, "
    sql = sql & "NVL(WFINDDO3CW,'0') as WFINDDO3CW, "
    sql = sql & "NVL(WFRESDO3CW,'0') as WFRESDO3CW, "
    sql = sql & "NVL(WFSMPLIDOT1CW,' ') as WFSMPLIDOT1CW, "
    sql = sql & "NVL(WFINDOT1CW,'0') as WFINDOT1CW, "
    sql = sql & "NVL(WFRESOT1CW,'0') as WFRESOT1CW, "
    sql = sql & "NVL(WFSMPLIDOT2CW,' ') as WFSMPLIDOT2CW, "
    sql = sql & "NVL(WFINDOT2CW,'0') as WFINDOT2CW, "
    sql = sql & "NVL(WFRESOT2CW,'0') as WFRESOT2CW, "
    sql = sql & "NVL(WFSMPLIDAOICW,' ') as WFSMPLIDAOICW, "
    sql = sql & "NVL(WFINDAOICW,'0') as WFINDAOICW, "
    sql = sql & "NVL(WFRESAOICW,'0') as WFRESAOICW, "
    sql = sql & "NVL(SMPLNUMCW,0) as SMPLNUMCW, "
    sql = sql & "NVL(SMPLPATCW,' ') as SMPLPATCW, "
    sql = sql & "NVL(LIVKCW,'0') as LIVKCW, "
    sql = sql & "NVL(WFSMPLIDGDCW,' ') as WFSMPLIDGDCW, "
    sql = sql & "NVL(WFINDGDCW,'0') as WFINDGDCW, "
    sql = sql & "NVL(WFRESGDCW,'0') as WFRESGDCW, "
    sql = sql & "NVL(WFHSGDCW,'0') as WFHSGDCW, "
    sql = sql & "NVL(EPSMPLIDB1CW,' ') as EPSMPLIDB1CW, "
    sql = sql & "NVL(EPINDB1CW,'0') as EPINDB1CW, "
    sql = sql & "NVL(EPRESB1CW,'0') as EPRESB1CW, "
    sql = sql & "NVL(EPSMPLIDB2CW,' ') as EPSMPLIDB2CW, "
    sql = sql & "NVL(EPINDB2CW,'0') as EPINDB2CW, "
    sql = sql & "NVL(EPRESB2CW,'0') as EPRESB2CW, "
    sql = sql & "NVL(EPSMPLIDB3CW,' ') as EPSMPLIDB3CW, "
    sql = sql & "NVL(EPINDB3CW,'0') as EPINDB3CW, "
    sql = sql & "NVL(EPRESB3CW,'0') as EPRESB3CW, "
    sql = sql & "NVL(EPSMPLIDL1CW,' ') as EPSMPLIDL1CW, "
    sql = sql & "NVL(EPINDL1CW,'0') as EPINDL1CW, "
    sql = sql & "NVL(EPRESL1CW,'0') as EPRESL1CW, "
    sql = sql & "NVL(EPSMPLIDL2CW,' ') as EPSMPLIDL2CW, "
    sql = sql & "NVL(EPINDL2CW,'0') as EPINDL2CW, "
    sql = sql & "NVL(EPRESL2CW,'0') as EPRESL2CW, "
    sql = sql & "NVL(EPSMPLIDL3CW,' ') as EPSMPLIDL3CW, "
    sql = sql & "NVL(EPINDL3CW,'0') as EPINDL3CW, "
    sql = sql & "NVL(EPRESL3CW,'0') as EPRESL3CW "
    sql = sql & "FROM "
    sql = sql & "    (SELECT SXLIDCB, "
    sql = sql & "     XTALCB, "
    sql = sql & "     (INPOSCB+RLENCB) as INGOTPOS, "
    sql = sql & "     HINBCB, "
    sql = sql & "     REVNUMCB, "
    sql = sql & "     FACTORYCB, "
    sql = sql & "     OPECB "
    sql = sql & "     FROM XSDCB "
    sql = sql & "     WHERE XTALCB = '" & sCryNum & "' "
    sql = sql & "     AND LIVKCB = '0' "
    sql = sql & "    ), "
    sql = sql & "    (SELECT SXLIDCW, "
    sql = sql & "     SMPKBNCW, "
    sql = sql & "     TBKBNCW, "
    sql = sql & "     REPSMPLIDCW, "
    sql = sql & "     INPOSCW, "
    sql = sql & "     HINBCW, "
    sql = sql & "     REVNUMCW, "
    sql = sql & "     FACTORYCW, "
    sql = sql & "     OPECW, "
    sql = sql & "     KTKBNCW, "
    sql = sql & "     SMCRYNUMCW, "
    sql = sql & "     WFSMPLIDRSCW, "
    sql = sql & "     WFSMPLIDRS1CW, "
    sql = sql & "     WFSMPLIDRS2CW, "
    sql = sql & "     WFINDRSCW, "
    sql = sql & "     WFRESRS1CW, "
    sql = sql & "     WFRESRS2CW, "
    sql = sql & "     WFSMPLIDOICW, "
    sql = sql & "     WFINDOICW, "
    sql = sql & "     WFRESOICW, "
    sql = sql & "     WFSMPLIDB1CW, "
    sql = sql & "     WFINDB1CW, "
    sql = sql & "     WFRESB1CW, "
    sql = sql & "     WFSMPLIDB2CW, "
    sql = sql & "     WFINDB2CW, "
    sql = sql & "     WFRESB2CW, "
    sql = sql & "     WFSMPLIDB3CW, "
    sql = sql & "     WFINDB3CW, "
    sql = sql & "     WFRESB3CW, "
    sql = sql & "     WFSMPLIDL1CW, "
    sql = sql & "     WFINDL1CW, "
    sql = sql & "     WFRESL1CW, "
    sql = sql & "     WFSMPLIDL2CW, "
    sql = sql & "     WFINDL2CW, "
    sql = sql & "     WFRESL2CW, "
    sql = sql & "     WFSMPLIDL3CW, "
    sql = sql & "     WFINDL3CW, "
    sql = sql & "     WFRESL3CW, "
    sql = sql & "     WFSMPLIDL4CW, "
    sql = sql & "     WFINDL4CW, "
    sql = sql & "     WFRESL4CW, "
    sql = sql & "     WFSMPLIDDSCW, "
    sql = sql & "     WFINDDSCW, "
    sql = sql & "     WFRESDSCW, "
    sql = sql & "     WFSMPLIDDZCW, "
    sql = sql & "     WFINDDZCW, "
    sql = sql & "     WFRESDZCW, "
    sql = sql & "     WFSMPLIDSPCW, "
    sql = sql & "     WFINDSPCW, "
    sql = sql & "     WFRESSPCW, "
    sql = sql & "     WFSMPLIDDO1CW, "
    sql = sql & "     WFINDDO1CW, "
    sql = sql & "     WFRESDO1CW, "
    sql = sql & "     WFSMPLIDDO2CW, "
    sql = sql & "     WFINDDO2CW, "
    sql = sql & "     WFRESDO2CW, "
    sql = sql & "     WFSMPLIDDO3CW, "
    sql = sql & "     WFINDDO3CW, "
    sql = sql & "     WFRESDO3CW, "
    sql = sql & "     WFSMPLIDOT1CW, "
    sql = sql & "     WFINDOT1CW, "
    sql = sql & "     WFRESOT1CW, "
    sql = sql & "     WFSMPLIDOT2CW, "
    sql = sql & "     WFINDOT2CW, "
    sql = sql & "     WFRESOT2CW, "
    sql = sql & "     WFSMPLIDAOICW, "
    sql = sql & "     WFINDAOICW, "
    sql = sql & "     WFRESAOICW, "
    sql = sql & "     SMPLNUMCW, "
    sql = sql & "     SMPLPATCW, "
    sql = sql & "     LIVKCW, "
    sql = sql & "     WFSMPLIDGDCW, "
    sql = sql & "     WFINDGDCW, "
    sql = sql & "     WFRESGDCW, "
    sql = sql & "     WFHSGDCW, "
    sql = sql & "     EPSMPLIDB1CW, "
    sql = sql & "     EPINDB1CW, "
    sql = sql & "     EPRESB1CW, "
    sql = sql & "     EPSMPLIDB2CW, "
    sql = sql & "     EPINDB2CW, "
    sql = sql & "     EPRESB2CW, "
    sql = sql & "     EPSMPLIDB3CW, "
    sql = sql & "     EPINDB3CW, "
    sql = sql & "     EPRESB3CW, "
    sql = sql & "     EPSMPLIDL1CW, "
    sql = sql & "     EPINDL1CW, "
    sql = sql & "     EPRESL1CW, "
    sql = sql & "     EPSMPLIDL2CW, "
    sql = sql & "     EPINDL2CW, "
    sql = sql & "     EPRESL2CW, "
    sql = sql & "     EPSMPLIDL3CW, "
    sql = sql & "     EPINDL3CW, "
    sql = sql & "     EPRESL3CW "
    sql = sql & "     FROM XSDCW "
    sql = sql & "     WHERE XTALCW = '" & sCryNum & "' "
    sql = sql & "     AND TBKBNCW = 'B' "
    sql = sql & "    ) "
'    sql = sql & "WHERE INGOTPOS = INPOSCW(+) "
    sql = sql & "WHERE SXLIDCB = SXLIDCW(+) "           '08/07/10 ooba
    sql = sql & "AND NVL(LIVKCW,'0') = '0' "
    sql = sql & "ORDER BY INGOTPOS, TBKBN "
    
    'ﾃﾞｰﾀを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Or rs.RecordCount Mod 2 <> 0 Then
        ReDim records(0)
        DBDRV_GetXSDCW = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    '抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SXLIDCW = rs("SXLIDCB")                'SXLID
            .SMPKBNCW = rs("SMPKBNCW")              'サンプル区分
            .TBKBNCW = rs("TBKBN")                  'T/B区分
            .REPSMPLIDCW = rs("REPSMPLIDCW")        '代表サンプルID
            .XTALCW = rs("XTALCB")                  '結晶番号
            .INPOSCW = rs("INGOTPOS")               '結晶内位置
            'XSDCWの品番を優先
            If i Mod 2 = 0 And .SXLIDCW = records(i - 1).SXLIDCW And _
                Trim(rs("HINBCW")) <> "" And rs("HINBCW") <> records(i - 1).HINBCW Then
                
                records(i - 1).HINBCW = rs("HINBCW")        '品番
                records(i - 1).REVNUMCW = rs("REVNUMCW")    '製品番号改訂番号
                records(i - 1).FACTORYCW = rs("FACTORYCW")  '工場
                records(i - 1).OPECW = rs("OPECW")          '操業条件
                .HINBCW = rs("HINBCW")              '品番
                .REVNUMCW = rs("REVNUMCW")          '製品番号改訂番号
                .FACTORYCW = rs("FACTORYCW")        '工場
                .OPECW = rs("OPECW")                '操業条件
            Else
                .HINBCW = rs("HINBCB")              '品番
                .REVNUMCW = rs("REVNUMCB")          '製品番号改訂番号
                .FACTORYCW = rs("FACTORYCB")        '工場
                .OPECW = rs("OPECB")                '操業条件
            End If
            .KTKBNCW = rs("KTKBNCW")                '確定区分
            .SMCRYNUMCW = rs("SMCRYNUMCW")          'サンプルブロックID
            .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")      'サンプルID(Rs)
            .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")    '推定サンプルID1(Rs)
            .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")    '推定サンプルID2(Rs)
            .WFINDRSCW = rs("WFINDRSCW")            '状態FLG（Rs)
            .WFRESRS1CW = rs("WFRESRS1CW")          '実績FLG1（Rs)
            .WFRESRS2CW = rs("WFRESRS2CW")          '実績FLG2（Rs)
            .WFSMPLIDOICW = rs("WFSMPLIDOICW")      'サンプルID（Oi)
            .WFINDOICW = rs("WFINDOICW")            '状態FLG（Oi)
            .WFRESOICW = rs("WFRESOICW")            '実績FLG（Oi)
            .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")      'サンプルID（B1)
            .WFINDB1CW = rs("WFINDB1CW")            '状態FLG（B1)
            .WFRESB1CW = rs("WFRESB1CW")            '実績FLG（B1)
            .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")      'サンプルID（B2）
            .WFINDB2CW = rs("WFINDB2CW")            '状態FLG（B2）
            .WFRESB2CW = rs("WFRESB2CW")            '実績FLG（B2）
            .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")      'サンプルID（B3)
            .WFINDB3CW = rs("WFINDB3CW")            '状態FLG（B3)
            .WFRESB3CW = rs("WFRESB3CW")            '実績FLG（B3)
            .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")      'サンプルID（L1)
            .WFINDL1CW = rs("WFINDL1CW")            '状態FLG（L1)
            .WFRESL1CW = rs("WFRESL1CW")            '実績FLG（L1)
            .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")      'サンプルID（L2)
            .WFINDL2CW = rs("WFINDL2CW")            '状態FLG（L2)
            .WFRESL2CW = rs("WFRESL2CW")            '実績FLG（L2)
            .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")      'サンプルID（L3)
            .WFINDL3CW = rs("WFINDL3CW")            '状態FLG（L3)
            .WFRESL3CW = rs("WFRESL3CW")            '実績FLG（L3)
            .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")      'サンプルID（L4)
            .WFINDL4CW = rs("WFINDL4CW")            '状態FLG（L4)
            .WFRESL4CW = rs("WFRESL4CW")            '実績FLG（L4)
            .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")      'サンプルID（DS)
            .WFINDDSCW = rs("WFINDDSCW")            '状態FLG（DS)
            .WFRESDSCW = rs("WFRESDSCW")            '実績FLG（DS)
            .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")      'サンプルID（DZ)
            .WFINDDZCW = rs("WFINDDZCW")            '状態FLG（DZ)
            .WFRESDZCW = rs("WFRESDZCW")            '実績FLG（DZ)
            .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")      'サンプルID（SP)
            .WFINDSPCW = rs("WFINDSPCW")            '状態FLG（SP)
            .WFRESSPCW = rs("WFRESSPCW")            '実績FLG（SP)
            .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")    'サンプルID（DO1)
            .WFINDDO1CW = rs("WFINDDO1CW")          '状態FLG（DO1)
            .WFRESDO1CW = rs("WFRESDO1CW")          '実績FLG（DO1)
            .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")    'サンプルID（DO2)
            .WFINDDO2CW = rs("WFINDDO2CW")          '状態FLG（DO2)
            .WFRESDO2CW = rs("WFRESDO2CW")          '実績FLG（DO2)
            .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")    'サンプルID（DO3)
            .WFINDDO3CW = rs("WFINDDO3CW")          '状態FLG（DO3)
            .WFRESDO3CW = rs("WFRESDO3CW")          '実績FLG（DO3)
            .WFSMPLIDOT1CW = rs("WFSMPLIDOT1CW")    'サンプルID（OT1)
            .WFINDOT1CW = rs("WFINDOT1CW")          '状態FLG（OT1)
            .WFRESOT1CW = rs("WFRESOT1CW")          '実績FLG（OT1)
            .WFSMPLIDOT2CW = rs("WFSMPLIDOT2CW")    'サンプルID（OT2)
            .WFINDOT2CW = rs("WFINDOT2CW")          '状態FLG（OT2)
            .WFRESOT2CW = rs("WFRESOT2CW")          '実績FLG（OT2)
            .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")    'サンプルID（AOi)
            .WFINDAOICW = rs("WFINDAOICW")          '状態FLG（AOi)
            .WFRESAOICW = rs("WFRESAOICW")          '実績FLG（AOi)
            .SMPLNUMCW = rs("SMPLNUMCW")            'サンプル枚数
            .SMPLPATCW = rs("SMPLPATCW")            'サンプルパターン
            .LIVKCW = rs("LIVKCW")                  '生死区分
            .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")      'サンプルID（GD)
            .WFINDGDCW = rs("WFINDGDCW")            '状態FLG（GD)
            .WFRESGDCW = rs("WFRESGDCW")            '実績FLG（GD)
            .WFHSGDCW = rs("WFHSGDCW")              '保証FLG（GD)
            .EPSMPLIDB1CW = rs("EPSMPLIDB1CW")      'サンプルID（B1E)
            .EPINDB1CW = rs("EPINDB1CW")            '状態FLG（B1E)
            .EPRESB1CW = rs("EPRESB1CW")            '実績FLG（B1E)
            .EPSMPLIDB2CW = rs("EPSMPLIDB2CW")      'サンプルID（B2E）
            .EPINDB2CW = rs("EPINDB2CW")            '状態FLG（B2E）
            .EPRESB2CW = rs("EPRESB2CW")            '実績FLG（B2E）
            .EPSMPLIDB3CW = rs("EPSMPLIDB3CW")      'サンプルID（BE3)
            .EPINDB3CW = rs("EPINDB3CW")            '状態FLG（B3E)
            .EPRESB3CW = rs("EPRESB3CW")            '実績FLG（B3E)
            .EPSMPLIDL1CW = rs("EPSMPLIDL1CW")      'サンプルID（L1E)
            .EPINDL1CW = rs("EPINDL1CW")            '状態FLG（L1E)
            .EPRESL1CW = rs("EPRESL1CW")            '実績FLG（L1E)
            .EPSMPLIDL2CW = rs("EPSMPLIDL2CW")      'サンプルID（L2E)
            .EPINDL2CW = rs("EPINDL2CW")            '状態FLG（L2E)
            .EPRESL2CW = rs("EPRESL2CW")            '実績FLG（L2E)
            .EPSMPLIDL3CW = rs("EPSMPLIDL3CW")      'サンプルID（L3E)
            .EPINDL3CW = rs("EPINDL3CW")            '状態FLG（L3E)
            .EPRESL3CW = rs("EPRESL3CW")            '実績FLG（L3E)
            
            '仮ｻﾝﾌﾟﾙID登録
            If Trim(.REPSMPLIDCW) = "" Then
                .REPSMPLIDCW = Mid(.SXLIDCW, 1, 10) & Format(CStr(.INPOSCW), "000") & .SMPKBNCW
            End If
        End With
        rs.MoveNext
    Next i
    rs.Close
    
    DBDRV_GetXSDCW = FUNCTION_RETURN_SUCCESS
    
End Function

'概要      :抜試待ち一覧 表示用ＤＢドライバ
'パラメータ　　:変数名        ,IO ,型               ,説明
'      　　:pLackBlk　　　,O  ,typ_LackBlk    　,抜試待ちブロック一覧
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/10 蔵本 作成
'          :2003/05/27 筑　抜試待ち一覧取得用ドライバに変更
Public Function DBDRV_scmzc_fcmkc001j_Disp(pLackBlk() As typ_LackBlk) As FUNCTION_RETURN

'    Dim rs As OraDynaset
'    Dim sql As String
'    Dim recCnt As Integer
'    Dim i As Long
    Dim rs, rs2, rs3                As OraDynaset
    Dim sql, sql2, sql3             As String
    Dim recCnt, rec2Cnt, rec3Cnt    As Integer
    Dim i, j, k, cnt                As Long
    Dim SXLID, blkID                As String

''    ReDim wBLKID(0) As String
    Dim ChkCnt                      As Long
    Dim BLKIDFlg                    As Boolean
    Dim OldBlk                      As String
    Dim iCnt                        As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001j_SQL.bas -- Function DBDRV_scmzc_fcmkc001j_Disp"

    ''　抜試待ち一覧の取得方法変更　2003/08/27 Mori =================> START
    ''
    ''　抜試待ち一覧の取得方法変更　2003/05/27 tuku =================> START
    '  SXL管理（XSDCB）の現在工程がCW740のロットをすべて表示させる。
    sql = vbNullString
    ' SXL管理<XSDCB>の現在工程がCW740のロットを取得
    '                   但し、サンプルの確定区分が'9'の物は除く
    '                   その場合でもXSDCBのホールド区分<HOLDBCB>が優先される（'9ならば表示）
'    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day"
'    sql = sql & " from      "
'    sql = sql & " (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day "
'    sql = sql & "  from xsdca       "
'    sql = sql & "  where sxlidca in "
'    sql = sql & "    (select sxlidcb"
'    sql = sql & "     from xsdcb cb,"
'    sql = sql & "     vecme011 ve011"
'    sql = sql & "     where gnwkntcb = 'CW740'"
'    sql = sql & "     and cb.sxlidcb = ve011.e042sxlid"
'    sql = sql & "     and (   ve011.E044KTKBN != '9' "
'    sql = sql & "          or cb.holdbcb = '9')"
'    sql = sql & "    )  "
'    sql = sql & " and livkca = '0'      "
'    sql = sql & " group by crynumca) ca ,       "
'    sql = sql & " tbcme040 e040"
'    sql = sql & " where e040.blockid = ca.lot"
'    sql = sql & " order by ca.lot           "


'''''    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day"
'''''
'''''    sql = sql & ",c1.puptnc1"   '引上ﾊﾟﾀｰﾝ追加対応(2004/12/08) kubota
'''''
'''''    sql = sql & " from      "
'''''    sql = sql & " (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day "
'''''    sql = sql & "  from  xsdca      "
'''''    sql = sql & "  where sxlidca in "
'''''    sql = sql & "    (select sxlidcb"
'''''    sql = sql & "     from xsdcb cb, xsdcw cw "
'''''    sql = sql & "     where gnwkntcb  = 'CW740'"
'''''    sql = sql & "       and cb.sxlidcb=  cw.sxlidcw"
'''''    sql = sql & "       and (cw.ktkbncw!='9' "
'''''    sql = sql & "         or cb.holdbcb ='9')"
'''''    sql = sql & "       and cw.LIVKCW = '0'"
'''''    sql = sql & "       and cb.LIVKCB = '0'"
'''''    sql = sql & "    )"
'''''    sql = sql & " and livkca = '0'"
'''''    sql = sql & " group by crynumca) ca ,"
'''''    sql = sql & " tbcme040 e040"
'''''
'''''    sql = sql & ",xsdc1 c1,xsdc2 c2"    '引上ﾊﾟﾀｰﾝ追加対応(2004/12/08) kubota
'''''
'''''    sql = sql & " where e040.blockid = ca.lot"
'''''
'''''    '引上ﾊﾟﾀｰﾝ追加対応(2004/12/08) kubota
'''''    sql = sql & "   and c2.crynumc2  = ca.lot"
'''''    sql = sql & "   and c2.xtalc2    = c1.xtalc1"
'''''
'''''    sql = sql & " order by ca.lot"


''    ' 関連するSXLIDのﾌﾞﾛｯｸIDをすべて取得するSQL文に変更 2005/03/18 ffc)tanabe
''    sql = sql & " select ca.lot,e040.ingotpos,e040.REALLEN,ca.day,c1.puptnc1 from "
''    sql = sql & "   (select distinct(crynumca) lot,max(inposca) pos,max(kdayca) day from xsdca"
''    sql = sql & "     where crynumca in"
''    sql = sql & "       (select blockid from tbcmy001 where  sblockid in"
''    sql = sql & "         (select distinct sblockid from tbcmy001 where blockid in"
''    sql = sql & "           (select crynumca from xsdca where sxlidca in"
''    sql = sql & "             (select distinct sxlidcb from xsdcb cb, xsdcw cw"
''    sql = sql & "                where gnwkntcb  = 'CW740'"
''    sql = sql & "                and cb.sxlidcb= cw.sxlidcw"
'''    sql = sql & "                and (cw.ktkbncw!='9' or cb.holdbcb ='9')"  '←ｺﾒﾝﾄ化　06/01/12 ooba
''    sql = sql & "                and cw.LIVKCW = '0' and cb.LIVKCB = '0'"
''    sql = sql & "             )"
''    sql = sql & "             and livkca = '0' group by crynumca"
''    sql = sql & "           )"
''    sql = sql & "         )"
''    sql = sql & "       )"
''    sql = sql & "     and livkca = '0' group by crynumca"
''    sql = sql & "   ) ca ,"
''    sql = sql & " tbcme040 e040,xsdc1 c1,xsdc2 c2 where e040.blockid = ca.lot"
''    sql = sql & " and c2.crynumc2  = ca.lot"
''    sql = sql & " and c2.xtalc2    = c1.xtalc1"
''    sql = sql & " order by ca.lot"

''    '待ち一覧取得SQL変更　06/02/06 ooba START ===============================================>
''    sql = sql & "SELECT DISTINCT "
''    sql = sql & "CA_GP.CRYNUMCA, "
''    sql = sql & "E40.INGOTPOS, "
''    sql = sql & "E40.REALLEN, "
''    sql = sql & "C1.PUPTNC1, "
''    sql = sql & "CA_GP.MHOLDBCA, "
''    sql = sql & "CA_GP.MWFHOLDFLGCA, "
''    sql = sql & "C2.WFHUFLG, "
''    sql = sql & "CA_GP.MKDAYCA "
''    sql = sql & ", C2.PLANTCATC2 "  ' 向先 2007/09/03 SPK Tsutsumi Add
''    sql = sql & "FROM "
''    sql = sql & "   (SELECT CRYNUMCA, MAX(HOLDBCA) MHOLDBCA, "
''    sql = sql & "    MAX(WFHOLDFLGCA) MWFHOLDFLGCA, MAX(KDAYCA) MKDAYCA "
''    sql = sql & "    FROM XSDCA WHERE LIVKCA = '0' GROUP BY CRYNUMCA "
''    sql = sql & "   ) CA_GP, "
''    sql = sql & "   (SELECT CRYNUMCA, SXLIDCA FROM XSDCA WHERE LIVKCA = '0' "
''    sql = sql & "   ) CA_AL, "
''' 2007/09/03 SPK Tsutsumi Add Start
''    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG, PLANTCATC2 FROM XSDC2 WHERE LIVKC2 = '0' "
'''    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG FROM XSDC2 WHERE LIVKC2 = '0' "
''' 2007/09/03 SPK Tsutsumi Add End
''sql = sql & "   ) C2, "
''    sql = sql & "   (SELECT SXLIDCB FROM XSDCB WHERE LIVKCB = '0' AND GNWKNTCB = 'CW740' "
''    sql = sql & "   ) CB, "
''    sql = sql & "   (SELECT XTALC1, PUPTNC1 FROM XSDC1 "
''    sql = sql & "   ) C1, "
''    sql = sql & "   (SELECT BLOCKID, INGOTPOS, REALLEN FROM TBCME040 "
''    sql = sql & "   ) E40 "
''    sql = sql & "WHERE C1.XTALC1 = C2.XTALC2 "
''    sql = sql & "AND C2.CRYNUMC2 = CA_GP.CRYNUMCA "
''    sql = sql & "AND C2.CRYNUMC2 = E40.BLOCKID "
''    sql = sql & "AND CA_GP.CRYNUMCA = CA_AL.CRYNUMCA "
''    sql = sql & "AND CA_AL.SXLIDCA = CB.SXLIDCB "
''
''    ' 向先 2007/09/03 SPK Tsutsumi Add Start
''    If sCmbMukesaki <> "ALL" Then
''        sql = sql & "   AND C2.PLANTCATC2      = '" & sCmbMukesaki & "'"
''    End If
''    ' 2007/09/03 SPK Tsutsumi Add End
''
''    sql = sql & "ORDER BY CA_GP.CRYNUMCA"
''    '待ち一覧取得SQL変更　06/02/06 ooba END =================================================>

    '待ち一覧取得SQL変更　08/01/31 ooba START ===============================================>
    sql = vbNullString
    sql = sql & "SELECT "
    sql = sql & "CA_GP.CRYNUMCA, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "E40.INGOTPOS, "
'    sql = sql & "E40.REALLEN, "
    sql = sql & "C2.INPOSC2, "
    sql = sql & "C2.GNLC2, "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "C1.PUPTNC1, "
    sql = sql & "CA_GP.MHOLDBCA, "
    sql = sql & "CA_GP.MWFHOLDFLGCA, "
    sql = sql & "C2.WFHUFLG, "
    sql = sql & "CA_GP.MKDAYCA, "
    sql = sql & "C2.PLANTCATC2, "
    sql = sql & "MAX(CB.GNWKNTCB) MGNWKNTCB, "
    sql = sql & "MAX(CB.KBLKFLGCB) MKBLKFLGCB "
    ' 流動停止項目追加 add SETkimizuka Start  09/03/18
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUSY4 "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSEY4 "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNOY4 "
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || A9.NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNOY4 "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
    ' 流動停止項目追加 add SETkimizuka End    09/03/18
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & " , CA_GP.HIN_CNT as HIN_CNT "
    sql = sql & " , NVL(CB.CW740STSCB,' ') as CW740STS "
    sql = sql & " , MAX(CA_AL.HINBCA) as HINBAN "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "FROM "
    sql = sql & "   (SELECT CRYNUMCA, MAX(HOLDBCA) MHOLDBCA, "
    sql = sql & "    MAX(WFHOLDFLGCA) MWFHOLDFLGCA, MAX(KDAYCA) MKDAYCA "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , COUNT(CRYNUMCA) HIN_CNT "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCA WHERE LIVKCA = '0' "
    sql = sql & "    GROUP BY CRYNUMCA "
    sql = sql & "   ) CA_GP, "
    sql = sql & "   (SELECT CRYNUMCA, SXLIDCA "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , HINBCA "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCA "
    sql = sql & "    WHERE LIVKCA = '0' "
    sql = sql & "   ) CA_AL, "
    sql = sql & "   (SELECT CRYNUMC2, XTALC2, WFHUFLG, PLANTCATC2, GNWKNTC2, KBLKFLGC2 "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , INPOSC2, GNLC2 "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDC2 "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "    WHERE LIVKC2 = '0' AND GNWKNTC2 = 'CW750' "
    sql = sql & "    WHERE LIVKC2 = '0' AND (GNWKNTC2 = 'CW740' or GNWKNTC2 = 'CW750') "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "   ) C2, "
    sql = sql & "   (SELECT SXLIDCB, GNWKNTCB, DECODE(KBLKFLGCB,'1',KBLKFLGCB,'0') KBLKFLGCB "
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & "    , CW740STSCB "
    'Add End 2010/07/08 SMPK Nakamura
    sql = sql & "    FROM XSDCB "
    sql = sql & "    WHERE LIVKCB = '0' AND GNWKNTCB IN ('CST02', 'CW740') "
    sql = sql & "   ) CB, "
    sql = sql & "   (SELECT XTALC1, PUPTNC1 "
    sql = sql & "    FROM XSDC1 "
    sql = sql & "   ) C1, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "   (SELECT BLOCKID, INGOTPOS, REALLEN "
'    sql = sql & "    FROM TBCME040 "
'    sql = sql & "   ) E40 "
'    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
'    sql = sql & "    ,XODY3 Y3,XODY4 Y4,KODA9 A9 "
    sql = sql & "    XODY3 Y3,XODY4 Y4,KODA9 A9 "
    'Chg End 2010/07/08 SMPK Nakamura
    ' 流動停止項目追加 add SETkimizuka Start  09/03/18
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO ,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY4 = '0' AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' 流動停止項目追加 add SETkimizuka End  09/03/18
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
    sql = sql & "WHERE C1.XTALC1 = C2.XTALC2 "
    sql = sql & "AND C2.CRYNUMC2 = CA_GP.CRYNUMCA "
'    sql = sql & "AND C2.CRYNUMC2 = E40.BLOCKID "       2010/07/08 SMPK Nakamura
    sql = sql & "AND CA_GP.CRYNUMCA = CA_AL.CRYNUMCA "
    sql = sql & "AND CA_AL.SXLIDCA = CB.SXLIDCB "
    sql = sql & "AND ((CB.GNWKNTCB = 'CW740') OR "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "     (CB.GNWKNTCB = 'CST02' AND C2.GNWKNTC2 = 'CW750' AND C2.KBLKFLGC2 = '1')) "
    sql = sql & "     (CB.GNWKNTCB = 'CST02' AND (C2.GNWKNTC2 = 'CW740' or C2.GNWKNTC2 = 'CW750') AND C2.KBLKFLGC2 = '1')) "
    'Chg End 2010/07/08 SMPK Nakamura
    If sCmbMukesaki <> "ALL" Then
        sql = sql & "   AND C2.PLANTCATC2      = '" & sCmbMukesaki & "'"
    End If
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    'sql = sql & "   AND CA_GP.CRYNUMCA     = Y4.XTALNO(+) "            'add 09/03/18 SETkimizuka
    sql = sql & " AND CA_GP.CRYNUMCA = Y3.XTALNOY3(+) "
    sql = sql & " AND Y3.LIVKY3(+) = '0' "
    sql = sql & " AND Y4.LIVKY4(+) = '0' "
    sql = sql & " AND Y3.XTALNOY3 = Y4.XTALNOY4(+) "
    sql = sql & " AND Y3.RCNTY3 = Y4.RCNTY4(+) "
    sql = sql & " AND A9.SYSCA9(+) = 'X' AND A9.SHUCA9(+) = '30' AND Y4.CAUSEY4 = A9.CODEA9(+) "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
    sql = sql & " GROUP BY "
    sql = sql & "CA_GP.CRYNUMCA, "
    'Chg Start 2010/07/08 SMPK Nakamura
'    sql = sql & "E40.INGOTPOS, "
'    sql = sql & "E40.REALLEN, "
    sql = sql & "C2.INPOSC2, "
    sql = sql & "C2.GNLC2, "
    'Chg End 2010/07/08 SMPK Nakamura
    sql = sql & "C1.PUPTNC1, "
    sql = sql & "CA_GP.MHOLDBCA, "
    sql = sql & "CA_GP.MWFHOLDFLGCA, "
    sql = sql & "C2.WFHUFLG, "
    sql = sql & "CA_GP.MKDAYCA, "
    sql = sql & "C2.PLANTCATC2 "
    ' 流動停止項目追加 add SETkimizuka Start  09/03/18
    sql = sql & ",Y4.AGRSTATUSY4 "
    sql = sql & ",Y4.STOPY4 "
    sql = sql & ",Y4.CAUSEY4 "
    sql = sql & ",Y4.PRINTKINDY4 "
    sql = sql & ",Y4.PRINTNOY4 "
    sql = sql & ",Y4.WKKTY4 "   'add 09/06/29 SETkimizuka
    sql = sql & ",A9.NAMEJA9 "  'add 09/06/29 SETkimizuka
    'Add Start 2010/07/08 SMPK Nakamura
    sql = sql & ",CA_GP.HIN_CNT "
    sql = sql & ",CB.CW740STSCB "
    'Add End 2010/07/08 SMPK Nakamura
    ' 流動停止項目追加 add SETkimizuka End    09/03/18
    sql = sql & "ORDER BY CA_GP.CRYNUMCA"
    '待ち一覧取得SQL変更　08/01/31 ooba END =================================================>
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount

    ' 流動停止項目追加に伴い取得方法変更 add SETkimizuka Start  09/03/18
    iCnt = 0
    '検索結果を格納する（ブロックID,結晶内開始位置,ブロック長さ,入庫日時)
    For i = 1 To recCnt
        If OldBlk <> rs("CRYNUMCA") Then
            iCnt = iCnt + 1
            ReDim Preserve pLackBlk(iCnt)
            With pLackBlk(iCnt)
                'Chg Start 2010/07/08 SMPK Nakamura
'                .INGOTPOS = rs("INGOTPOS")  ' 結晶内開始位置
'                .REALLEN = rs("REALLEN")    ' 実長さ
                .INGOTPOS = rs("INPOSC2")  ' 結晶内開始位置
                .REALLEN = rs("GNLC2")    ' 実長さ
                'Chg End 2010/07/08 SMPK Nakamura
                .PUPTN = rs("PUPTNC1")
        
                .BLOCKID = rs("CRYNUMCA")       'ﾌﾞﾛｯｸID
                'ﾎｰﾙﾄﾞ区分
                If IsNull(rs("MHOLDBCA")) Then .HOLDFLG = " " Else .HOLDFLG = rs("MHOLDBCA")
                'WFﾎｰﾙﾄﾞ区分
                If IsNull(rs("MWFHOLDFLGCA")) Then .WFHOLDFLG = " " Else .WFHOLDFLG = rs("MWFHOLDFLGCA")
                'WF振替FLG
                If IsNull(rs("WFHUFLG")) Then .WFHUFLG = " " Else .WFHUFLG = rs("WFHUFLG")
                .REJDTTM = rs("MKDAYCA")        '更新日付
        
                If IsNull(rs("PLANTCATC2")) = False Then
                    For j = 0 To UBound(s_MukesakiBase)
                        If s_MukesakiBase(j).sMukeCode = rs("PLANTCATC2") Then
                           .MUKESAKI = s_MukesakiBase(j).sMukeName
                        End If
                    Next j
                End If
                .Koutei = rs("MGNWKNTCB")       '工程(XSDCB)　08/01/31 ooba
                .KANREN = rs("MKBLKFLGCB")      '関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
                
                ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
                '.STOP = rs("STOP")                   '停止区分
                '.AGRSTATUS = rs("AGRSTATUSY4")       '承認確認区分
                'If Trim(rs("CAUSEY4")) <> "" Then
                '    .CAUSE = rs("CAUSEY4") & vbTab       '停止理由
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW740" Or rs("WKKTY4") = "CW000") Then
                    .STOP = rs("STOP")                   '停止区分
                    .AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
                    If Trim(rs("CAUSE")) <> "" Then
                        .CAUSE = rs("CAUSE") & vbTab       '停止理由
                    End If
                End If
                ' 流動監視SQL修正 upd SETkimizuka End  09/06/26
                If Trim(rs("PRINTNOY4")) <> "" Then
                    .PRINTNO = rs("PRINTNOY4") & vbTab       '先行評価
                End If
                'Add Start 2010/07/08 SMPK Nakamura
                '品番数
                .HINCNT = rs("HIN_CNT")
                '品番
                .hinban = rs("HINBAN")
                'CW740ステータス
                .CW740STS = rs("CW740STS")
                'Add End 2010/07/08 SMPK Nakamura
            End With
        Else
            With pLackBlk(iCnt)
                ' 流動監視SQL修正 upd SETkimizuka Start  09/06/26
                'If Trim(rs("CAUSEY4")) <> "" And InStr(.CAUSE, rs("CAUSEY4")) = 0 Then
                '    .CAUSE = .CAUSE & rs("CAUSEY4") & vbTab        '停止区分
                'End If
                If rs("STOP") <> "2" And (rs("WKKTY4") = "CW740" Or rs("WKKTY4") = "CW000") Then
                    If Trim(.AGRSTATUS) = "" Or (rs("AGRSTATUS") < .AGRSTATUS And Trim(rs("AGRSTATUS")) <> "") Then
                         .AGRSTATUS = rs("AGRSTATUS")
                         .STOP = rs("STOP")
                    End If
                    If Trim(rs("CAUSE")) <> "" And InStr(.CAUSE, rs("CAUSE")) = 0 Then
                        .CAUSE = .CAUSE & rs("CAUSE") & vbTab        '停止区分
                    End If
                End If
                If Trim(rs("PRINTNOY4")) <> "" And InStr(.PRINTNO, rs("PRINTNOY4")) = 0 Then
                    .PRINTNO = .PRINTNO & rs("PRINTNOY4") & vbTab        '先行評価
                End If
            End With
            
        End If
        
        OldBlk = rs("CRYNUMCA")
        rs.MoveNext
    Next i
    ' 流動停止項目追加に伴い取得方法変更 add SETkimizuka End  09/03/18

    ' 流動停止項目追加に伴い取得方法変更 del SETkimizuka Start  09/03/18
'    For i = 1 To recCnt
'        iCnt = iCnt + 1
'        ReDim Preserve pLackBlk(i)
'        With pLackBlk(i)
''        .BLOCKID = rs("LOT")        ' ブロックID
'        .IngotPos = rs("INGOTPOS")  ' 結晶内開始位置
'        .REALLEN = rs("REALLEN")    ' 実長さ
''        .REJDTTM = rs("DAY")        ' 入庫日
'        .PUPTN = rs("PUPTNC1")
'
'        '06/02/06 ooba START ===============================================================>
'        .BLOCKID = rs("CRYNUMCA")       'ﾌﾞﾛｯｸID
'        'ﾎｰﾙﾄﾞ区分
'        If IsNull(rs("MHOLDBCA")) Then .HOLDFLG = " " Else .HOLDFLG = rs("MHOLDBCA")
'        'WFﾎｰﾙﾄﾞ区分
'        If IsNull(rs("MWFHOLDFLGCA")) Then .WFHOLDFLG = " " Else .WFHOLDFLG = rs("MWFHOLDFLGCA")
'        'WF振替FLG
'        If IsNull(rs("WFHUFLG")) Then .WFHUFLG = " " Else .WFHUFLG = rs("WFHUFLG")
'        .REJDTTM = rs("MKDAYCA")        '更新日付
'        '06/02/06 ooba END =================================================================>
'
'        ' 2007/09/03 SPK Tsutsumi Add Start
'        If IsNull(rs("PLANTCATC2")) = False Then
'            For j = 0 To UBound(s_MukesakiBase)
'                If s_MukesakiBase(j).sMukeCode = rs("PLANTCATC2") Then
'                   .MUKESAKI = s_MukesakiBase(j).sMukeName
'                End If
'            Next j
'        End If
'        ' 2007/09/03 SPK Tsutsumi Add End
'        .Koutei = rs("MGNWKNTCB")       '工程(XSDCB)　08/01/31 ooba
'        .KANREN = rs("MKBLKFLGCB")      '関連ﾌﾞﾛｯｸ有無　08/01/31 ooba
'
'        End With
'
'        OldBlk = rs("CRYNUMCA")  'add 09/03/18 SETkimizuka
'        rs.MoveNext
'    Next i
    ' 流動停止項目追加に伴い取得方法変更 del SETkimizuka End  09/03/18
    rs.Close


    ''　抜試待ち一覧の取得方法変更　2003/05/27 tuku =================> END
    ''　抜試待ち一覧の取得方法変更　2003/08/27 Mori =================> END

''    '' ﾎｰﾙﾄﾞ区分取得追加　05/01/31 ooba START ===================================>
''    If recCnt > 0 Then      '05/04/19 ooba
''        For i = 1 To UBound(pLackBlk)
''            sql2 = "SELECT "
''            sql2 = sql2 & "HOLDBCA, "
''            sql2 = sql2 & "WFHOLDFLGCA "
''            sql2 = sql2 & "FROM XSDCA "
''            sql2 = sql2 & "WHERE "
''            sql2 = sql2 & "CRYNUMCA = '" & pLackBlk(i).BLOCKID & "' "
''            sql2 = sql2 & "AND LIVKCA = '0' "
''
''            Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_NO_BLANKSTRIP)
''            rec2Cnt = rs2.RecordCount
''            If rec2Cnt > 0 Then
''                For j = 1 To rec2Cnt
''                    If IsNull(rs2("HOLDBCA")) = False Then
''                        pLackBlk(i).HOLDFLG = rs2("HOLDBCA")        'ﾎｰﾙﾄﾞ区分
''                    Else
''                        pLackBlk(i).HOLDFLG = ""
''                    End If
''                    If IsNull(rs2("WFHOLDFLGCA")) = False Then
''                        pLackBlk(i).WFHOLDFLG = rs2("WFHOLDFLGCA")  'WFﾎｰﾙﾄﾞ区分
''                    Else
''                        pLackBlk(i).WFHOLDFLG = ""
''                    End If
''
''                    If pLackBlk(i).HOLDFLG <> "0" Or pLackBlk(i).HOLDFLG <> " " Then
''                        Exit For
''                    End If
''                    rs2.MoveNext
''                Next
''            End If
''            rs2.Close
''        Next
''    End If
''    '' ﾎｰﾙﾄﾞ区分取得追加　05/01/31 ooba END =====================================>

    DBDRV_scmzc_fcmkc001j_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmkc001j_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function






'基本処理パラメータ作成 2002/09/10 ADD hitec)N.MATSUMOTO Start
'引数：frmFormID=処理画面の判定（1:WFセンタ総合判定　2:再抜試）
Public Function MakeParameter(ByVal StrCryNum As String) As FUNCTION_RETURN

    Dim lng                 As Long
    Dim dat                 As Variant
    Dim lRowCnt             As Long
    Dim rsMain              As OraDynaset
    Dim sql                 As String
    Dim intCnt              As Integer
    Dim errTbl              As String
    Dim sErrMsg             As String
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim strIngotpos         As String
    Dim varIngotpos         As Variant
    Dim i                   As Integer  'add 2003/05/17 hitec)matsumoto

    With f_cmbc036_2.sprExamine
    '品番を1列追加したことによる列の変更-------start iida 2003/09/03
''''        .GetText 3, 1, varIngotpos      'upd 2003/04/07 hitec)matsumoto 画面レイアウト変更に伴い修正
        .GetText 5, 1, varIngotpos
''''        lngBeginIngotpos = CInt(Trim(varIngotpos))
        lngBeginIngotpos = SIngotP  'upd 2003/05/16 hitec)matsumoto
''''        .GetText 3, .MaxRows, varIngotpos   'upd 2003/04/07 hitec)matsumoto 画面レイアウト変更に伴い修正
        .GetText 5, .MaxRows, varIngotpos
''''        lngEndIngotpos = CInt(Trim(varIngotpos))
        lngEndIngotpos = EIngotP    'upd 2003/05/16 hitec)matsumoto
    End With
    '品番を1列追加したことによる列の変更-------end iida 2003/09/03
    '構造体作成
    If cmbc036_2_CreateTable(StrCryNum, lngBeginIngotpos, lngEndIngotpos, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        MakeParameter = FUNCTION_RETURN_FAILURE
        f_cmbc036_2.lblMsg.Caption = sErrMsg
        Exit Function
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO End


'構造体作成処理　2002/09/10 ADD hitec)N.MATSUMOTO Start
Public Function cmbc036_2_CreateTable(ByVal StrCryNum As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs              As OraDynaset
    Dim errTbl          As String
    Dim StrBlockId()    As String
    Dim strDBName       As String
    Dim bNoData         As Boolean
    Dim intLoopCnt      As Integer
    Dim sql             As String
    Dim strCryNum9      As String

    bNoData = False

    giInpos = 9000  'add 2003/04/16 hitec)matsumoto 在庫減、振替情報の位置を初期化

    'ブロック管理からブロックＩＤを取得
    'upd start 2003/03/28 hitec)matsumoto ----------------------
''''    sql = "SELECT * from TBCME040 "
''''    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
''''    sql = sql & "   AND INGOTPOS>=" & lngBeginIngotpos & " AND (INGOTPOS + LENGTH) <=" & lngEndIngotpos
''''
''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''''    If rs.RecordCount = 0 Then
''''        rs.Close
''''        cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
''''        GoTo proc_exit
''''    End If
    '2003/04/27 hitec)okazaki 変更
'''''    sql = "SELECT DISTINCT(CRYNUMCA) "
'''''    sql = sql & " FROM XSDCA"
'''''    sql = sql & " WHERE CRYNUMCA = '" & strCryNum & "'"
    strCryNum9 = Left(StrCryNum, 9)
    sql = "SELECT CRYNUMCA "
    sql = sql & " FROM XSDCA"
    sql = sql & " WHERE SUBSTR(CRYNUMCA,1,9) = '" & strCryNum9 & "'"
    sql = sql & "   AND (INPOSCA>=" & lngBeginIngotpos
    sql = sql & "   AND  INPOSCA< " & lngEndIngotpos & ")"
    sql = sql & "   AND LIVKCA = '0' "
    sql = sql & "GROUP BY CRYNUMCA"
''''    strCryNum9 = Left(strCryNum, 9)
''''
''''    sql = "SELECT DISTINCT(CRYNUMCA) "
''''    sql = sql & " FROM XSDCA"
''''    sql = sql & " WHERE SUBSTR(CRYNUMCA,1,9) = '" & strCryNum9 & "'"
''''    sql = sql & " AND ( INPOSCA >= " & lngBeginIngotpos & ""
''''    sql = sql & " AND  INPOSCA < " & lngEndIngotpos & ")"
''''    sql = sql & " AND  LIVKCA = '0' "
    '変更end
    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    'ブロックIDを取得
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve StrBlockId(intLoopCnt) As String
        If IsNull(rs("CRYNUMCA")) = True Then
            StrBlockId(intLoopCnt) = ""
        Else
            StrBlockId(intLoopCnt) = rs("CRYNUMCA")            'ブロックID
        End If
        Debug.Print "cmbc036_2_CreateTable ループ開始 strBlockID(" & intLoopCnt & ") =" & StrBlockId(intLoopCnt)

        '基本情報構造体
        With Kihon
            .StaffID = Trim(f_cmbc036_2.txtStaffID.Text)
            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NOWPROC = PROCD_NUKISI_HENKOU
            .DIAMETER = 0      '--------------保留
            .ALLSCRAP = "N" '全数スクラップ
        End With

        '分割結晶（ブロック）から前工程実績取得
        strDBName = "XSDC2"
        If cmbc036_2_CreateXSDC2(StrBlockId(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc036_2_CreateTable = FUNCTION_RETURN_SUCCESS '処理は行わないが、正常で返す
                Debug.Print "cmbc036_2_CreateXSDC2(" & StrBlockId(intLoopCnt) & "," & bNoData & "):XSDC2前工程実績無し"
                Exit Function
            Else
                cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc036_2_CreateXSDC2(" & StrBlockId(intLoopCnt) & "," & bNoData & "):XSDC2前工程実績読込みエラー"
                Exit Function
            End If
        End If

        '分割結晶（品番）から前工程実績取得
        strDBName = "XSDCA"
        If cmbc036_2_CreateXSDCA(StrBlockId(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc036_2_CreateTable = FUNCTION_RETURN_SUCCESS '処理は行わないが、正常で返す
                Debug.Print "XSDCA：前工程実績無し"
                Exit Function
            Else
                cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "XSDCA：前工程実績読込みエラー"
                Exit Function
            End If
        End If

        '現在工程実績作成
        strErrMsg = GetMsgStr("EAPLY")
        If cmbc036_2_CreateNowProc(StrBlockId(intLoopCnt), lngBeginIngotpos, lngEndIngotpos, strErrMsg) = FUNCTION_RETURN_FAILURE Then
            cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
'            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "XSDC2,XSDCA：現在工程実績作成エラー"
            Exit Function
        End If
        strErrMsg = ""

        '基本処理
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc036_2_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "基本処理異常終了"
            Exit Function
        End If
        intLoopCnt = intLoopCnt + 1
        rs.MoveNext
    Loop
    rs.Close

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO End


'分割結晶（ブロック）前工程実績取得＆構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateXSDC2(ByVal StrBlockId As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False


    sql = "SELECT * from XSDC2 "
    sql = sql & " WHERE CRYNUMC2 ='" & StrBlockId & "'"
    sql = sql & "   AND LIVKC2= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    If rs.EOF = False Then
        With BlkOld
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")       '工程連番
            If IsNull(rs.Fields("XTALC2")) = False Then .XTALC2 = rs.Fields("XTALC2")
            If IsNull(rs.Fields("INPOSC2")) = False Then .INPOSC2 = rs.Fields("INPOSC2")
            If IsNull(rs.Fields("NEKKNTC2")) = False Then .NEKKNTC2 = rs.Fields("NEKKNTC2")
            If IsNull(rs.Fields("NEWKNTC2")) = False Then .NEWKNTC2 = rs.Fields("NEWKNTC2")
            If IsNull(rs.Fields("NEWKKBC2")) = False Then .NEWKKBC2 = rs.Fields("NEWKKBC2")
            If IsNull(rs.Fields("NEMACOC2")) = False Then .NEMACOC2 = rs.Fields("NEMACOC2")
            If IsNull(rs.Fields("GNKKNTC2")) = False Then .GNKKNTC2 = rs.Fields("GNKKNTC2")
            If IsNull(rs.Fields("GNWKNTC2")) = False Then .GNWKNTC2 = rs.Fields("GNWKNTC2")
            If IsNull(rs.Fields("GNWKKBC2")) = False Then .GNWKKBC2 = rs.Fields("GNWKKBC2")
            If IsNull(rs.Fields("GNMACOC2")) = False Then .GNMACOC2 = rs.Fields("GNMACOC2")
            If IsNull(rs.Fields("GNDAYC2")) = False Then .GNDAYC2 = rs.Fields("GNDAYC2")
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")          '現在長さ
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")          '現在重量
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")          '現在枚数
            If IsNull(rs.Fields("SUMITLC2")) = False Then .SUMITLC2 = rs.Fields("SUMITLC2")
            If IsNull(rs.Fields("SUMITWC2")) = False Then .SUMITWC2 = rs.Fields("SUMITWC2")
            If IsNull(rs.Fields("SUMITMC2")) = False Then .SUMITMC2 = rs.Fields("SUMITMC2")
            If IsNull(rs.Fields("CHGC2")) = False Then .CHGC2 = rs.Fields("CHGC2")
            If IsNull(rs.Fields("KAKOUBC2")) = False Then .KAKOUBC2 = rs.Fields("KAKOUBC2")
            If IsNull(rs.Fields("KEIDAYC2")) = False Then .KEIDAYC2 = rs.Fields("KEIDAYC2")
            If IsNull(rs.Fields("GNTKUBC2")) = False Then .GNTKUBC2 = rs.Fields("GNTKUBC2")
            If IsNull(rs.Fields("GNTNOC2")) = False Then .GNTNOC2 = rs.Fields("GNTNOC2")
            If IsNull(rs.Fields("XTWORKC2")) = False Then .XTWORKC2 = rs.Fields("XTWORKC2")
            If IsNull(rs.Fields("WFWORKC2")) = False Then .WFWORKC2 = rs.Fields("WFWORKC2")
            If IsNull(rs.Fields("LSTATBC2")) = False Then .LSTATBC2 = rs.Fields("LSTATBC2")
            If IsNull(rs.Fields("RSTATBC2")) = False Then .RSTATBC2 = rs.Fields("RSTATBC2")
            If IsNull(rs.Fields("LUFRCC2")) = False Then .LUFRCC2 = rs.Fields("LUFRCC2")
            If IsNull(rs.Fields("LUFRBC2")) = False Then .LUFRBC2 = rs.Fields("LUFRBC2")
            If IsNull(rs.Fields("LDFRCC2")) = False Then .LDFRCC2 = rs.Fields("LDFRCC2")
            If IsNull(rs.Fields("LDFRBC2")) = False Then .LDFRBC2 = rs.Fields("LDFRBC2")
            If IsNull(rs.Fields("HOLDCC2")) = False Then .HOLDCC2 = rs.Fields("HOLDCC2")
            If IsNull(rs.Fields("HOLDBC2")) = False Then .HOLDBC2 = rs.Fields("HOLDBC2")
            If IsNull(rs.Fields("EXKUBC2")) = False Then .EXKUBC2 = rs.Fields("EXKUBC2")
            If IsNull(rs.Fields("HENPKC2")) = False Then .HENPKC2 = rs.Fields("HENPKC2")
            If IsNull(rs.Fields("LIVKC2")) = False Then .LIVKC2 = rs.Fields("LIVKC2")
            If IsNull(rs.Fields("KANKC2")) = False Then .KANKC2 = rs.Fields("KANKC2")
            If IsNull(rs.Fields("NFC2")) = False Then .NFC2 = rs.Fields("NFC2")
            If IsNull(rs.Fields("SAKJC2")) = False Then .SAKJC2 = rs.Fields("SAKJC2")
            If IsNull(rs.Fields("TDAYC2")) = False Then .TDAYC2 = rs.Fields("TDAYC2")
            If IsNull(rs.Fields("KDAYC2")) = False Then .KDAYC2 = rs.Fields("KDAYC2")
            If IsNull(rs.Fields("SUMITBC2")) = False Then .SUMITBC2 = rs.Fields("SUMITBC2")
            If IsNull(rs.Fields("SNDKC2")) = False Then .SNDKC2 = rs.Fields("SNDKC2")
            If IsNull(rs.Fields("SNDDAYC2")) = False Then .SNDDAYC2 = rs.Fields("SNDDAYC2")
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2") ' 2007/09/04 SPK Tsutsumi Add
        End With
    End If

    rs.Close
    cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO

'現在工程構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateNowProc(ByVal StrBlockId As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, _
                                        ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs              As OraDynaset
    Dim sql             As String
    Dim intProcNo       As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim StrCryNum       As String
    Dim strLstatcls     As String
    Dim intBlkLength    As Integer  'ブロック管理データの長さ
    Dim intBlkIngotPos  As Integer  'ブロック管理データの位置
    Dim intSxlLength    As Integer  'シングル管理データの長さ
    Dim intSxlIngotPos  As Integer  'シングル管理データの位置
    Dim bFlg            As Boolean
    Dim sp              As Integer  '長さ判定用
    Dim ep              As Integer  '長さ判定用
    Dim sbp             As Integer  '長さ判定用
    Dim ebp             As Integer  '長さ判定用
    Dim intLength       As Integer  '長さ
    Dim intIngotPos     As Integer  '位置
    Dim rs2             As OraDynaset   'add 2003/04/15 hitec)matsumoto
    Dim iWFcnt          As Integer
    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0

    intBlkLength = 0
    intBlkIngotPos = 0
    intSxlLength = 0
    intSxlIngotPos = 0
    StrCryNum = ""

    'ブロック管理から長さを取得
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE BLOCKID='" & StrBlockId & "'"
''''    sql = sql & "   AND INGOTPOS=0"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then StrCryNum = rs("CRYNUM")               '結晶番号
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")            '長さ
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("INGOTPOS")      '位置
    End If

    rs.Close

    'ブロック管理で取得した長さをもとにシングル管理からデータを取得
    'upd start 2003/04/15 hitec)matsumoto 全数スクラップはTBCMY011で判断するように修正---------
''''    sql = "SELECT * from TBCME042 "
''''    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
''''    '↓ループ内で判定
''''    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
''''    sql = sql & "   AND LSTATCLS<>'H'"
    sql = "SELECT LOTID from TBCMY011 "
    sql = sql & " WHERE LOTID='" & StrBlockId & "'"     '2003/04/03 hitec)matsumoto 全数スクラップ="Y"はブロック単位なので、シングル範囲で取れない
''''    sql = sql & "   AND ((BLOCKSEQ >=" & lngWfBeginSeq & ") And (BLOCKSEQ <= " & lngWfEndSeq & "))"
    sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    'upd end   2003/04/15 hitec)matsumoto 全数スクラップはTBCMY011で判断するように修正---------

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then   '該当データ0件の場合、全数スクラップの処理
        '前工程実績を、現在工程実績にコピー
        BlkNow = BlkOld
        BlkNow.GNLC2 = "0"
        BlkNow.GNWC2 = "0"
        BlkNow.GNMC2 = "0"
        BlkNow.GNWKNTC2 = Kihon.NEWPROC
        BlkNow.NEWKNTC2 = Kihon.NOWPROC
        For intHinOldCnt = 0 To Kihon.CNTHINOLD - 1
            ReDim Preserve HinNow(intHinOldCnt) As typ_XSDCA_Update
            HinNow(intHinOldCnt) = HinOld(intHinOldCnt)
            HinNow(intHinOldCnt).GNLCA = "0"    '全数スクラップ=長さが0
            HinNow(intHinOldCnt).GNWCA = "0"    '重量 = 0
            HinNow(intHinOldCnt).GNMCA = "0"    '枚数 = 0
            HinNow(intHinOldCnt).GNWKNTCA = Kihon.NEWPROC
            HinNow(intHinOldCnt).NEWKNTCA = Kihon.NOWPROC
        Next
        Kihon.CNTHINNOW = 1
        Kihon.ALLSCRAP = "Y"

        '前工程の長さと現在工程の長さをくらべ、不良が存在するか判定
'        If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '不良なし
        If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '不良なし
            '基本情報構造体
            With Kihon
                .FURYOUMU = "N"
            End With
        Else                                            '不良あり
''            '基本情報構造体
''            With Kihon
''                .FURYOUMU = "Y"
''            End With
''            '不良構造体を作成
''            With Furyou
''                .XTALC4 = BlkNow.CRYNUMC2   'ブロックID
''                .INPOSC4 = BlkNow.INPOSC2   '結晶内開始位置
''                .KCKNTC4 = BlkNow.KCNTC2    '工程連番
''                .HINBC4 = "Z"               '品番
''    '            .REVNUMC4                   '製品番号改訂番号
''    '            .FACTORYC4                  '工場
''    '            .OPEC4                      '操業条件
''                .WKKTC4 = PROCD_NUKISI_HENKOU
''                .PUCUTLC4 = CLng(BlkOld.GNLC2) - CLng(BlkNow.GNLC2) '不良長さ(前工程-現在工程（良品）)
''                '不良重量
''                If GetDiameter(.XTALC4, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
''                    dblDiameter = 0
''    ''''                GoTo proc_wxit
''                End If
''                '取得した直径を元に重量を求める
''                .PUCUTWC4 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.PUCUTLC4))))
''                '不良枚数
''                If WfCount(.XTALC4, CLng(.PUCUTLC4), intNum) = FUNCTION_RETURN_FAILURE Then
''                    .PUCUTMC4 = 0
''    ''''                GoTo proc_wxit
''                Else
''                    .PUCUTMC4 = intNum
''                End If
''
''                .SUMITBC3 = "0"
''            End With
                rs.Close
                strErrMsg = GetMsgStr("EWFM5", "前工程=" & BlkOld.GNMC2 & "：現在工程=" & BlkNow.GNMC2) '03/06/06 後藤
'                lblMsg.Caption = "WF枚数不一致エラー"
                cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
        End If
        rs.Close
        cmbc036_2_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    '前工程の構造体を現在工程の構造体へコピー
    BlkNow = BlkOld
    '工程連番に＋１する
    With BlkNow
        If BlkNow.KCNTC2 = "" Then BlkNow.KCNTC2 = "0"
        .KCNTC2 = CInt(.KCNTC2) + 1         '工程連番
        .NEWKNTC2 = Kihon.NOWPROC           '前工程
        .GNWKNTC2 = Kihon.NEWPROC           '現在工程
        .SUMITLC2 = "0"                     'SUMMIT長さ
        .SUMITMC2 = "0"                     'SUMMIT枚数
        .SUMITWC2 = "0"                     'SUMMIT重量
        .SUMITBC2 = "0"
    End With

    ''↓変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   XTALCB"
    sql = sql & "  ,RLENCB"
    sql = sql & "  ,INPOSCB"
    sql = sql & "  ,HINBCB"
    sql = sql & "  ,REVNUMCB"
    sql = sql & "  ,FACTORYCB"
    sql = sql & "  ,OPECB"
    sql = sql & "  ,SXLIDCB"
    sql = sql & "  ,PLANTCATCB" ' 2007/09/05 SPK Tsutsumi Add
    sql = sql & " FROM"
    sql = sql & "   XSDCB"
    sql = sql & " WHERE XTALCB='" & StrCryNum & "'"
    '↓ループ内で判定
    sql = sql & "   AND ((INPOSCB >=" & lngBeginIngotpos & ")"
    sql = sql & "   And (INPOSCB + RLENCB <= " & lngEndIngotpos & "))"
    sql = sql & "   AND LSTCCB <> 'H'"
    '↓元ロジック(SXL管理（E042）→XSDCB機能移行)
'    sql = "SELECT * from TBCME042 "
'    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
'    '↓ループ内で判定
'    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
'''''    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS  <= " & lngEndIngotpos & "))"
'    sql = sql & "   AND LSTATCLS<>'H'"
    ''↑変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川


    Debug.Print sql
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)


    intLoopCnt = 0
''''    BlkNow.GNLC2 = 0    '現在工程（ブロック）の長さをクリアしておく
    BlkNow.GNMC2 = 0
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update
        '前工程の構造体を現在工程の構造体へコピー
''''        HinNow(intLoopCnt) = HinOld(intHinOldCnt)

        ''↓変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
        If IsNull(rs("XTALCB")) = False Then StrCryNum = rs("XTALCB")               '結晶番号
        If IsNull(rs("RLENCB")) = False Then intSxlLength = rs("RLENCB")            '長さ
        If IsNull(rs("INPOSCB")) = False Then intSxlIngotPos = rs("INPOSCB")        '位置
'        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '結晶番号
'        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")            '長さ
'        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")      '位置
        ''↑変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川

        '-- ブロックとシングルの位置関係を判定し、長さを算出 --------
        sp = intSxlIngotPos         'シングル開始位置
        ep = sp + intSxlLength      'シングル終端位置
        sbp = intBlkIngotPos        'ブロック開始位置
        ebp = sbp + intBlkLength    'ブロック終端位置

        '' ブロックがSXLの中に完全に含まれている場合 ---------
        If sp <= sbp And ep >= ebp Then

            intLength = intBlkLength                    'ブロック管理の長さを使用
            intIngotPos = intBlkIngotPos

        '' ブロックがSXLの開始位置より上にあり、かつ終端位置よりも長い場合 ---------
        ElseIf sp >= sbp And ep <= ebp Then

            intLength = intSxlLength                  'シングル管理の長さを使用
            intIngotPos = intSxlIngotPos

        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが上側。ただしブロックの終端とSXLの開始位置が一致しないこと) ------------
        ElseIf sp > sbp And sp < ebp And sp <> ebp Then

            intLength = ebp - sp                        'ブロックの終端位置 - シングルの開始位置
            intIngotPos = intSxlIngotPos

        '' ブロックが一部SXLにかかっている場合
        '' (ブロックが下側。ただしSXLの終端とブロックの開始位置が一致しないこと) ----------
        ElseIf sp < sbp And ep > sbp And ep <> sbp Then

            intLength = ep - sbp                        'シングルの終端位置 - ブロックの開始位置
            intIngotPos = intBlkIngotPos

        Else

''''            intLength = 0
''''            intIngotPos = intBlkIngotPos
            GoTo LoopNext

        End If
        '----------------------------------------------------

        '現在工程編集
        With HinNow(intLoopCnt)
            ''↓変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
            If IsNull(rs("XTALCB")) = False Then .XTALCA = rs("XTALCB")
            .CRYNUMCA = StrBlockId         'ブロックID
            If IsNull(rs("HINBCB")) = False Then .HINBCA = rs("HINBCB")             '品番
            If IsNull(rs("REVNUMCB")) = False Then .REVNUMCA = rs("REVNUMCB")       '製品番号改訂番号
            If IsNull(rs("FACTORYCB")) = False Then .FACTORYCA = rs("FACTORYCB")    '工場
            If IsNull(rs("OPECB")) = False Then .OPECA = rs("OPECB")        '操業条件

'            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")
'            .CRYNUMCA = strBlockID         'ブロックID
'            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '品番
'            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '製品番号改訂番号
'            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '工場
'            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '操業条件
            ''↑変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
            .INPOSCA = intIngotPos    '結晶内開始位置
            .GNLCA = intLength          '長さ
'            BlkNow.GNLC2 = CStr(CLng(BlkNow.GNLC2) + CLng(HinNow(intLoopCnt).GNLCA))  '長さ
            ''↓変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川
            If IsNull(rs("SXLIDCB")) = False Then .SXLIDCA = rs("SXLIDCB")          'シングルID
'            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          'シングルID
            ''↑変更START SXL管理（E042）→XSDCB機能移行 '06/1/5 SMP石川

            If IsNull(rs("PLANTCATCB")) = False Then .PLANTCATCA = rs("PLANTCATCB") '向先 2007/09/04 SPK Tsutsumi Add

            .SUMITBCA = 0
            .SUMITLCA = 0
            .SUMITMCA = 0
            .SUMITWCA = 0
            .NEWKNTCA = Kihon.NOWPROC   '前工程
            .GNWKNTCA = Kihon.NEWPROC   '現在工程
            .KCKNTCA = BlkNow.KCNTC2    '工程連番
            .NEMACOCA = BlkNow.NEMACOC2 '最終通過処理回数
            .GNMACOCA = BlkNow.GNMACOC2 '現在処理回数
''''        .XTALCA = strCryNum         '結晶番号
            '現在重量を求める
            If GetDiameter(StrBlockId, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '基本情報の直径セット
            Kihon.DIAMETER = dblDiameter

            '取得した直径を元に重量を求める
            .GNWCA = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLCA))))

            'add start hitec)matsumoto WFマップﾃｰﾌﾞﾙより枚数取得
            sql = "SELECT LOTID from TBCMY011 "
            sql = sql & " WHERE MSXLID='" & .SXLIDCA & "'"
            sql = sql & " AND LOTID='" & .CRYNUMCA & "'"
            sql = sql & " AND TO_NUMBER(WFSTA) <= 1"

            Debug.Print sql
            Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            iWFcnt = 0
            Do While Not rs2.EOF
                iWFcnt = iWFcnt + 1
                rs2.MoveNext
            Loop
            rs2.Close
            Debug.Print .SXLIDCA & " = " & iWFcnt & "枚"
            HinNow(intLoopCnt).GNMCA = iWFcnt   'add 2003/03/29 hitec)matsumoto 上でWFマップテーブルから枚数カウント取得しているので、それを良品枚数とする
            BlkNow.GNMC2 = BlkNow.GNMC2 + iWFcnt
            Debug.Print "BlkNow.GNMC2 = " & BlkNow.GNMC2 & "枚"
            .SUMITLCA = .GNLCA   '' 03/05/13 後藤
            .SUMITMCA = .GNMCA
            .SUMITWCA = .GNWCA
''''            '現在枚数を求める
''''            If WfCount(strBlockID, CLng(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
''''                .GNMCA = 0
''''''''                GoTo proc_wxit
''''            Else
''''                .GNMCA = intNum
''''            End If
        End With

        With BlkNow
            '現在重量を求める
            If GetDiameter(StrBlockId, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
    ''''                GoTo proc_wxit
            End If
            '基本情報の直径セット
            Kihon.DIAMETER = dblDiameter
            '取得した直径を元に重量を求める
'            .GNWC2 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLC2))))
            '現在枚数を求める
'            If WfCount(strBlockID, CLng(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
'                .GNMC2 = 0
''''                GoTo proc_wxit
'            Else
'                .GNMC2 = intNum
'            End If

        End With
        intLoopCnt = intLoopCnt + 1
        '良品件数セット
        With Kihon
            .CNTHINNOW = intLoopCnt
        End With

LoopNext:

        rs.MoveNext
    Loop



    rs.Close

    Debug.Print " WFマップﾃｰﾌﾞﾙより枚数取得 " & BlkNow.GNMC2 & "枚 : 前工程" & BlkOld.GNMC2 & "枚"

    '前工程の長さと現在工程の長さをくらべ、不良が存在するか判定
'    If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '不良なし
    If CInt(BlkNow.GNMC2) = CInt(BlkOld.GNMC2) Then '不良なし
        '基本情報構造体
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '不良あり
                strErrMsg = GetMsgStr("EWFM5", "前工程=" & BlkOld.GNMC2 & "：現在工程=" & BlkNow.GNMC2) '03/06/06 後藤
'                lblMsg.Caption = "WF枚数不一致エラー"
                cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
                GoTo proc_exit


    End If
    cmbc036_2_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO


'分割結晶（品番）前工程実績取得＆構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc036_2_CreateXSDCA(ByVal StrBlockId As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim iLoopCnt    As Integer
    Dim rs          As OraDynaset
    Dim sql         As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0

    'ブロックIDを得る
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & StrBlockId & "'"
    sql = sql & "   AND LIVKCA= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc036_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    rs.MoveFirst
    iLoopCnt = 0
    'add start 2003/05/14 hitec)matsumoto ----------------------------------------
    BlkOld.GNLC2 = 0
    BlkOld.GNWC2 = 0
    BlkOld.GNMC2 = 0
    'add end   2003/05/14 hitec)matsumoto ----------------------------------------
    Do While Not rs.EOF
        ReDim Preserve HinOld(iLoopCnt)
        ReDim Preserve HinNow(iLoopCnt)
        With HinOld(iLoopCnt)
            If IsNull(rs.Fields("CRYNUMCA")) = False Then .CRYNUMCA = rs.Fields("CRYNUMCA")
            If IsNull(rs.Fields("HINBCA")) = False Then .HINBCA = rs.Fields("HINBCA")
            If IsNull(rs.Fields("INPOSCA")) = False Then .INPOSCA = rs.Fields("INPOSCA")
            If IsNull(rs.Fields("REVNUMCA")) = False Then .REVNUMCA = rs.Fields("REVNUMCA")
            If IsNull(rs.Fields("FACTORYCA")) = False Then .FACTORYCA = rs.Fields("FACTORYCA")
            If IsNull(rs.Fields("OPECA")) = False Then .OPECA = rs.Fields("OPECA")
            If IsNull(rs.Fields("KCKNTCA")) = False Then .KCKNTCA = rs.Fields("KCKNTCA")
            If IsNull(rs.Fields("SXLIDCA")) = False Then .SXLIDCA = rs.Fields("SXLIDCA")
            If IsNull(rs.Fields("XTALCA")) = False Then .XTALCA = rs.Fields("XTALCA")
            If IsNull(rs.Fields("NEKKNTCA")) = False Then .NEKKNTCA = rs.Fields("NEKKNTCA")
            If IsNull(rs.Fields("NEWKNTCA")) = False Then .NEWKNTCA = rs.Fields("NEWKNTCA")
            If IsNull(rs.Fields("NEWKKBCA")) = False Then .NEWKKBCA = rs.Fields("NEWKKBCA")
            If IsNull(rs.Fields("NEMACOCA")) = False Then .NEMACOCA = rs.Fields("NEMACOCA")
            If IsNull(rs.Fields("GNKKNTCA")) = False Then .GNKKNTCA = rs.Fields("GNKKNTCA")
            If IsNull(rs.Fields("GNWKNTCA")) = False Then .GNWKNTCA = rs.Fields("GNWKNTCA")
            If IsNull(rs.Fields("GNWKKBCA")) = False Then .GNWKKBCA = rs.Fields("GNWKKBCA")
            If IsNull(rs.Fields("GNMACOCA")) = False Then .GNMACOCA = rs.Fields("GNMACOCA")
            If IsNull(rs.Fields("GNDAYCA")) = False Then .GNDAYCA = rs.Fields("GNDAYCA")
            If IsNull(rs.Fields("GNLCA")) = False Then .GNLCA = rs.Fields("GNLCA")
            If IsNull(rs.Fields("GNWCA")) = False Then .GNWCA = rs.Fields("GNWCA")
            If IsNull(rs.Fields("GNMCA")) = False Then .GNMCA = rs.Fields("GNMCA")
            'add start 2003/05/14 hitec)matsumoto ----------------------------------------
            BlkOld.GNLC2 = CLng(BlkOld.GNLC2) + CLng(.GNLCA)
            BlkOld.GNWC2 = CLng(BlkOld.GNWC2) + CLng(.GNWCA)
            BlkOld.GNMC2 = CLng(BlkOld.GNMC2) + CLng(.GNMCA)
            'add end   2003/05/14 hitec)matsumoto ----------------------------------------
            If IsNull(rs.Fields("SUMITLCA")) = False Then .SUMITLCA = rs.Fields("SUMITLCA")
            If IsNull(rs.Fields("SUMITWCA")) = False Then .SUMITWCA = rs.Fields("SUMITWCA")
            If IsNull(rs.Fields("SUMITMCA")) = False Then .SUMITMCA = rs.Fields("SUMITMCA")
            If IsNull(rs.Fields("CHGCA")) = False Then .CHGCA = rs.Fields("CHGCA")
            If IsNull(rs.Fields("KAKOUBCA")) = False Then .KAKOUBCA = rs.Fields("KAKOUBCA")
            If IsNull(rs.Fields("KEIDAYCA")) = False Then .KEIDAYCA = rs.Fields("KEIDAYCA")
            If IsNull(rs.Fields("GNTKUBCA")) = False Then .GNTKUBCA = rs.Fields("GNTKUBCA")
            If IsNull(rs.Fields("GNTNOCA")) = False Then .GNTNOCA = rs.Fields("GNTNOCA")
            If IsNull(rs.Fields("XTWORKCA")) = False Then .XTWORKCA = rs.Fields("XTWORKCA")
            If IsNull(rs.Fields("WFWORKCA")) = False Then .WFWORKCA = rs.Fields("WFWORKCA")
            If IsNull(rs.Fields("LSTATBCA")) = False Then .LSTATBCA = rs.Fields("LSTATBCA")
            If IsNull(rs.Fields("RSTATBCA")) = False Then .RSTATBCA = rs.Fields("RSTATBCA")
            If IsNull(rs.Fields("LUFRCCA")) = False Then .LUFRCCA = rs.Fields("LUFRCCA")
            If IsNull(rs.Fields("LUFRBCA")) = False Then .LUFRBCA = rs.Fields("LUFRBCA")
            If IsNull(rs.Fields("LDFRCCA")) = False Then .LDFRCCA = rs.Fields("LDFRCCA")
            If IsNull(rs.Fields("LDFRBCA")) = False Then .LDFRBCA = rs.Fields("LDFRBCA")
            If IsNull(rs.Fields("HOLDCCA")) = False Then .HOLDCCA = rs.Fields("HOLDCCA")
            If IsNull(rs.Fields("HOLDBCA")) = False Then .HOLDBCA = rs.Fields("HOLDBCA")
            If IsNull(rs.Fields("EXKUBCA")) = False Then .EXKUBCA = rs.Fields("EXKUBCA")
            If IsNull(rs.Fields("HENPKCA")) = False Then .HENPKCA = rs.Fields("HENPKCA")
            If IsNull(rs.Fields("LIVKCA")) = False Then .LIVKCA = rs.Fields("LIVKCA")
            If IsNull(rs.Fields("KANKCA")) = False Then .KANKCA = rs.Fields("KANKCA")
            If IsNull(rs.Fields("NFCA")) = False Then .NFCA = rs.Fields("NFCA")
            If IsNull(rs.Fields("SAKJCA")) = False Then .SAKJCA = rs.Fields("SAKJCA")
            If IsNull(rs.Fields("TDAYCA")) = False Then .TDAYCA = rs.Fields("TDAYCA")
            If IsNull(rs.Fields("KDAYCA")) = False Then .KDAYCA = rs.Fields("KDAYCA")
            If IsNull(rs.Fields("SUMITBCA")) = False Then .SUMITBCA = rs.Fields("SUMITBCA")
            If IsNull(rs.Fields("SNDKCA")) = False Then .SNDKCA = rs.Fields("SNDKCA")
            If IsNull(rs.Fields("SNDDAYCA")) = False Then .SNDDAYCA = rs.Fields("SNDDAYCA")
            If IsNull(rs.Fields("PLANTCATCA")) = False Then .PLANTCATCA = rs.Fields("PLANTCATCA") ' 2007/09/04 SPK Tsutsumi Add
        End With
        '良品件数セット
        With Kihon
            .CNTHINOLD = iLoopCnt + 1
        End With
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop

    rs.Close
    cmbc036_2_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc036_2_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO
'概要    :抜試指示 ブロックＩＤ(SXLが他のブロックに跨る場合）を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :tBL_SXLID    ,IO   ,type_DBDRV_scmzc_fcmlc001d_LOTSXL     ,ブロックＩＤ、ＳＸＬＩＤ構造体
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :SXLID→ブロックＩＤ(SXLが他のブロックに跨る場合）を取得する
'履歴    :2003/2/25 Hitec)okazaki
Public Function DBDRV_BLOCKIDGET() As FUNCTION_RETURN
    Dim sql             As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim inCnt           As Long
    Dim sDbName         As String
    Dim itUCount        As Integer
    Dim sBkBlkId        As String
    Dim sBkSxlId()      As String
    Dim iSxlCnt         As Integer
    Dim iLoopCnt        As Integer
    Dim iLoop           As Integer  '2003/05/28 HITEC)okazaki add
    Dim iLoop2          As Integer  '2003/05/28 HITEC)okazaki add
    Dim iLoop3          As Integer  '2003/05/28 HITEC)okazaki add
    Dim wkSXLID()       As type_DBDRV_LOTSXL
    Dim bCheckFlg       As Boolean
    Dim sBeforLotid     As String
    Dim iFLG            As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_BLOCKIDGET"
    DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
    sDbName = "(V001)"
    ReDim wkSXLID(1)
    wkSXLID(1) = tSXLID(1)
    'ループ開始
    For iLoop2 = 1 To 10        '永久ループ防止のため最大１０回で抜ける
        ReDim sBkSxlId(1)
        iSxlCnt = 1
        '======================================================================
        '前回のループで取得したブロック全てよりSXLを取得（重複も可となっている）
        '======================================================================
        For iLoopCnt = 1 To UBound(wkSXLID)
            iFLG = 0
            If iLoopCnt = 1 Then
                iFLG = 1
            ElseIf wkSXLID(iLoopCnt).LOTID <> wkSXLID(iLoopCnt - 1).LOTID Then
                iFLG = 1
            End If
            If iFLG <> 0 Then
                ' SXLIDの取得
                sql = "select"
                sql = sql & " SXLIDCA"
                sql = sql & " from XSDCA "
                sql = sql & " where CRYNUMCA ='" & wkSXLID(iLoopCnt).LOTID & "'"
                sql = sql & "   and LIVKCA = '0'"
                sql = sql & " ORDER BY SXLIDCA"

                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                '''抽出レコードが存在ならば該当
                If Not rs.EOF Then
                    rs.MoveFirst
                    Do While Not rs.EOF
                        ReDim Preserve sBkSxlId(iSxlCnt)
                        sBkSxlId(iSxlCnt) = rs.Fields("SXLIDCA")
                        iSxlCnt = iSxlCnt + 1
                        rs.MoveNext
                    Loop
                End If
                rs.Close
            End If
        Next iLoopCnt
        If iSxlCnt = 1 Then
            Exit Function
        End If
        itUCount = UBound(tSXLID)
        '=============================================
        '取得したSXLよりSXL・BLOCKの組み合わせ取得
        '=============================================
        For iLoopCnt = 1 To iSxlCnt - 1
            sql = "select"
            sql = sql & " CRYNUMCA,SXLIDCA"
            sql = sql & " from XSDCA "
            sql = sql & " where SXLIDCA ='" & sBkSxlId(iLoopCnt) & "'"
            sql = sql & "   and LIVKCA = '0'"
        '        sql = sql & " AND NOR Lotid ='" & tBL_SXLID(i).lotid & "'"
            sql = sql & " ORDER BY CRYNUMCA,SXLIDCA"

            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

            '''抽出レコードが存在ならば該当
            If Not rs.EOF Then
                '’抽出レコードをすべて取得（ループ）
                Do While Not rs.EOF
                    '’配列にその組み合わせを追加する
                    If itUCount = 1 Then  '1（まだロットを入れていない状態）
                        With tSXLID(itUCount)
                            .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                            .LOTID = rs.Fields("CRYNUMCA")
                        End With
                        itUCount = itUCount + 1

                    Else    '対象ロットが複数あった時
                        bCheckFlg = False
                        For iLoop3 = 1 To UBound(tSXLID)
                            If tSXLID(iLoop3).SXLID = rs.Fields("SXLIDCA") And _
                               tSXLID(iLoop3).LOTID = rs.Fields("CRYNUMCA") Then
                               bCheckFlg = True
                               Exit For
                            End If
                        Next iLoop3
                        If bCheckFlg = False Then
                            ReDim Preserve tSXLID(itUCount)  '配列の再定義
                            With tSXLID(itUCount)
                                .SXLID = rs.Fields("SXLIDCA")         'SXLIDCA
                                .LOTID = rs.Fields("CRYNUMCA")
                            End With
                            itUCount = itUCount + 1
                        End If
                    End If
                    rs.MoveNext
                Loop
                rs.Close
            End If
        Next iLoopCnt

        '==================================================
        '組み合わせが前回のループと同じなら終了
        '==================================================
        If UBound(tSXLID) = UBound(wkSXLID) Then
            Exit For
        End If
        ReDim wkSXLID(UBound(tSXLID))
        wkSXLID = tSXLID
    Next iLoop2


    ' 配列内のソートが必要
    iSxlCnt = UBound(tSXLID)
    itUCount = 1

    sql = "select"
    ''↓修正START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    ''  ブロック、SXLに対して複数品番が有り得るようになった為修正
    '　　結晶内位置を取得
'    sql = sql & " CRYNUMCA,SXLIDCA"
    sql = sql & " distinct CRYNUMCA,SXLIDCA,min(INPOSCA) INPOSCA"
    ''↑修正START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    sql = sql & " from  XSDCA "
    sql = sql & " where SXLIDCA in ("
    For iLoopCnt = 1 To iSxlCnt
        sql = sql & "'" & tSXLID(iLoopCnt).SXLID & "'"
        If iLoopCnt <> iSxlCnt Then sql = sql & ","
    Next
    sql = sql & " )"
    sql = sql & "   and LIVKCA ='0'"
    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    sql = sql & " GROUP BY CRYNUMCA,SXLIDCA"
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    sql = sql & " ORDER BY CRYNUMCA,SXLIDCA"

    ReDim tSXLID(0)  '配列の再定義

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
        '’抽出レコードをすべて取得（ループ）
        Do While Not rs.EOF
            ReDim Preserve tSXLID(itUCount)         '配列の再定義
            With tSXLID(itUCount)
                .SXLID = rs.Fields("SXLIDCA")       'SXLIDCA
                .LOTID = rs.Fields("CRYNUMCA")
                ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
                .INGOTPOS = rs.Fields("INPOSCA")
                ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
            End With
            itUCount = itUCount + 1
            rs.MoveNext
        Loop
        rs.Close
    End If

    DBDRV_BLOCKIDGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_BLOCKIDGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function



'概要    :抜試指示 MIN,MAX値を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :tBL_SXLID    ,IO   ,type_DBDRV_scmzc_fcmlc001d_LOTSXL     ,ブロックＩＤ、ＳＸＬＩＤ構造体
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :SXLID,BLOCKID→最大、最小（ブロックＰで判定）のデータを取得する
'履歴    :2003/2/25 Hitec)okazaki
Public Function DBDRV_MIN_MAX_SEQGET(ByRef iWfNum As Integer) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
    Dim inCnt       As Long
    Dim sDbName     As String
    Dim itUCount    As Integer
    Dim dblWFLen    As Double  '2003/04/25 hitec)okazaki
    Dim iRtn        As FUNCTION_RETURN
    Dim eps         As Double
    Dim sSmpKbn     As String
    Dim dblBlP      As Double
    Dim j, m, n     As Integer      '05/12/26 ooba

    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    Dim lsBackHinban    As String
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_MIN_MAX_SEQGET"

    eps = 0.000001        'εの設定
    iWfNum = 0
    itUCount = 1
    sDbName = "(Y011)"

    ReDim sWrpLOTID(0)      '05/12/26 ooba
    ReDim iWrpBLOCKSEQ(0)   '05/12/26 ooba

    'i = 0
    '対象SXLは一つなのでループなし（シングル単位の画面なので）
    '’ループ開始
    For i = 1 To UBound(tSXLID)
        ' SXLIDの取得
        sql = "select "
        sql = sql & "LOTID,"                ' ブロックID"
        sql = sql & "MSXLID,"               ' SXLID"   'upd hitec)matsumoto カラム名変更
        sql = sql & "blockseq,"             ' ブロック内連番"
        sql = sql & "WFSTA,"                ' WF状態"
        sql = sql & "MHINBAN,"              ' 品番" 'upd hitec)matsumoto カラム名変更
        sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
        sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
        sql = sql & "MSMPLEID,"             ' 抜試位置"    'upd hitec)matsumoto カラム名変更
        sql = sql & "SHAFLAG,"              ' サンプルフラグ"
        sql = sql & "INDTM,"
        sql = sql & "BASKETID,"
        sql = sql & "SLOTNO,"
        sql = sql & "CURRWPCS,"
        sql = sql & "EXISTFLG,"
        sql = sql & "TOP_POS,"
        sql = sql & "REJCAT,"
        sql = sql & "TXID,"
        sql = sql & "REGDATE,"
        sql = sql & "SUMMITSENDFLAG,"
        sql = sql & "SENDFLAG,"
        sql = sql & "SENDDATE,"
        sql = sql & "HREJCODE,"
        sql = sql & "UPDPROC,"
        sql = sql & "UPDDATE,"
        sql = sql & "MREVNUM,"
        sql = sql & "MFACTORY,"
        sql = sql & "MOPECOND,"
        sql = sql & "kankbn,"
        sql = sql & "NREJCODE"
        sql = sql & " from TBCMY011 "
        sql = sql & " where MSXLID='" & tSXLID(i).SXLID & "'"  'upd hitec)matsumoto カラム名変更
        sql = sql & "   AND Lotid ='" & tSXLID(i).LOTID & "'"
        sql = sql & " ORDER BY Lotid,blockseq ASC"

        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

        '05/12/26 ooba START ==================================>
        m = UBound(sWrpLOTID)
        n = rs.RecordCount
        j = 0
        ReDim Preserve sWrpLOTID(m + n)         'ﾌﾞﾛｯｸID
        ReDim Preserve iWrpBLOCKSEQ(m + n)      'ﾌﾞﾛｯｸ内連番
        '05/12/26 ooba END ====================================>

        '''抽出レコードが存在ならば該当
        iWfNum = 0
        Do While Not rs.EOF
            ''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
            '1SXL-1品番では無くなった為、計算方法を変える
'            If CInt(rs.Fields("WFSTA")) <= 1 Then
'                iWfNum = iWfNum + 1
'            End If
            ''↑削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
            '05/12/26 ooba START ==============================>
            j = j + 1
            'ﾌﾞﾛｯｸID
            If IsNull(rs("LOTID")) Then
                sWrpLOTID(m + j) = ""
            Else
                sWrpLOTID(m + j) = rs("LOTID")
            End If
            'ﾌﾞﾛｯｸ内連番
            If IsNull(rs("BLOCKSEQ")) Then
                iWrpBLOCKSEQ(m + j) = 0
            Else
                iWrpBLOCKSEQ(m + j) = rs("BLOCKSEQ")
            End If
            '05/12/26 ooba END ================================>
            rs.MoveNext
        Loop
        If rs.RecordCount = 0 Then
            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
            f_cmbc036_2.lblMsg.Caption = GetMsgStr("EWFM6", "Y011") '03/06/06 後藤
            rs.Close
            Exit Function
        End If

        rs.MoveFirst    '先頭ﾚｺｰﾄﾞに移動
        Do While Not rs.EOF
            ReDim Preserve tExamine(itUCount)   '配列の再定義
            With tExamine(itUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' ブロックID
                 Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 0                         ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' ブロック内連番
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF状態
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' 品番
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
                    'WF一枚の長さ取得                                   '2003/04/25 hitec)okazaki
                    iRtn = DBDRV_WFLENGET(tSXLID(i).LOTID, dblWFLen)
                    'ブロック先頭の表示位置はWF一枚の長さを引いたもの   '2003/04/25 hitec)okazaki
''''                .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.9 + eps)         ' 論理ブロック内位置  'add 2003/06/13 hitec)matsumoto [+ eps]追加
                    .RTOP_POS = Fix(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + 0.99999)           ' 論理ブロック内位置  'upd 2003/08/06 hitec)matsumoto
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                    .RINGOTPOS = 0
                Else
                    'ブロック先頭の表示位置はWF一枚の長さを引いたもの   '2003/04/25 hitec)okazaki
''''                .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.9 + eps)       ' 論理結晶内位置    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                    .RITOP_POS = Fix(CDbl(rs.Fields("RITOP_POS")) - dblWFLen + 0.99999)         ' 論理結晶内位置    'upd 2003/08/06 hitec)matsumoto
                    .RINGOTPOS = CDbl(rs.Fields("RITOP_POS"))     ' 2003/04/30 hitec)okazaki ソートの逆転を防ぐため追加
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' 抜試位置
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' サンプルフラグ
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto サンプルフラグが
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                    .INDTM = vbNullString
                Else
                    .INDTM = rs!INDTM               ' ウェハーセンター入庫日時
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' バスケットID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' スロットNO
                End If
                ''↓修正START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
                '1SXL-1品番では無くなった為、計算方法を変える
                iWfNum = 0
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = 0
                Else
                    If CInt(rs.Fields("WFSTA")) <= 1 Then
                        iWfNum = iWfNum + 1
                    End If
'                    .CURRWPCS = iWfNum              ' ウェハー枚数
                End If

                '↓元ロジック
'                If IsNull(rs!CURRWPCS) = True Then
'                    .CURRWPCS = 0
'                Else
'                    .CURRWPCS = iWfNum              ' ウェハー枚数
'                End If
                ''↑修正START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' 存在フラグ
                End If

                If IsNull(rs!TOP_POS) = True Then   ' ブロックのTOPからの位置
                    .TOP_POS = 0                    ' (画面表示のずれを防ぐため、サンプル区分により切り上げ
                Else                                '  切捨て処理追加 2003/05/05)
                    dblBlP = CDbl(rs!TOP_POS)
                    .TOP_POS = Int(dblBlP / 10)             '切捨て
                End If
                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' 欠落理由
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' トランザクションID
                End If
                If IsNull(rs!REGDATE) = True Then
                    .REGDATE = vbNullString
                Else
                    .REGDATE = rs!REGDATE           ' 登録日付
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT送信フラグ
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' 送信フラグ
                End If
                If IsNull(rs!SENDDATE) = True Then
                    .SENDDATE = vbNullString
                Else
                    .SENDDATE = rs!SENDDATE         ' 送信日付
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' 不良理由コード
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' 更新工程
                End If
                If IsNull(rs!UPDDATE) = True Then
                    .UPDDATE = vbNullString
                Else
                    .UPDDATE = rs!UPDDATE           ' 更新日付
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' 製品番号改訂番号
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' 工場
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' 操業条件
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' 完了区分
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' 抜試返答理由コード
                End If
            End With

            ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
            'SXLの変わり目だけでなく、品番の変わり目でもサンプル情報を取得する
            lsBackHinban = Trim(rs!mhinban)
            Do While (1)
                rs.MoveNext
                ''データの終わりの場合、一つ戻ってループ終了
                If rs.EOF Then
                    rs.MovePrevious
                    tExamine(itUCount).CURRWPCS = iWfNum              ' ウェハー枚数
                    Exit Do
                End If
                ''品番が変わったら、一つ戻ってループ終了
                If IsNull(rs!mhinban) = False Then
                    If lsBackHinban <> Trim(rs!mhinban) Then
                        rs.MovePrevious
                        tExamine(itUCount).CURRWPCS = iWfNum              ' ウェハー枚数
                        Exit Do
                    End If
                End If
                'ウェーハ枚数カウントアップ
                If CInt(rs.Fields("WFSTA")) <= 1 Then
                    iWfNum = iWfNum + 1
                End If
                lsBackHinban = Trim(rs!mhinban)
            Loop
            ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川


            '最終ﾚｺｰﾄﾞ
'            rs.MoveLast                             '最終ﾚｺｰﾄﾞに移動
            itUCount = itUCount + 1
            ReDim Preserve tExamine(itUCount)    '配列の再定義
            With tExamine(itUCount)
                If IsNull(rs!LOTID) = True Then
                    .LOTID = vbNullString           ' ブロックID
                Else
                    .LOTID = rs!LOTID
                End If
                If IsNull(rs!MSXLID) = True Then
                    .SXLID = vbNullString
                Else
                    .SXLID = rs!MSXLID               ' SXLID
                End If
                .MinMax = 1                          ' 0:MIN 1:MAX
                If IsNull(rs!BLOCKSEQ) = True Then
                    .BLOCKSEQ = vbNullString
                Else
                    .BLOCKSEQ = rs!BLOCKSEQ         ' ブロック内連番
                End If
                If IsNull(rs!WFSTA) = True Then
                    .WFSTA = vbNullString
                Else
                    .WFSTA = rs!WFSTA               ' WF状態
                End If
                If IsNull(rs!mhinban) = True Then
                    .hinban = vbNullString
                Else
                    .hinban = rs!mhinban             ' 品番
                End If
                If IsNull(rs!RTOP_POS) = True Then
                    .RTOP_POS = 0
                Else
'                   .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")))                    ' 論理ブロック内位置
                    .RTOP_POS = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)        ' 論理ブロック内位置   'add 2003/06/13 hitec)matsumoto [+ eps]追加
                End If
                If IsNull(rs!RITOP_POS) = True Then
                    .RITOP_POS = 0
                    .RINGOTPOS = 0
                Else
'                   .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")))                  ' 論理結晶内位置
                    .RITOP_POS = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)      ' 論理結晶内位置 'add 2003/06/13 hitec)matsumoto [+ eps]追加
                    .RINGOTPOS = CDbl(rs.Fields("RITOP_POS"))     ' 2003/04/30 hitec)okazaki ソートの逆転を防ぐため追加
                End If
                If IsNull(rs!MSMPLEID) = True Then
                    .SMPLEID = vbNullString
                Else
                    .SMPLEID = rs!MSMPLEID           ' 抜試位置
                End If
                If IsNull(rs!SHAFLAG) = True Then
                    .SHAFLAG = vbNullString
                Else
                    .SHAFLAG = rs!SHAFLAG           ' サンプルフラグ
                    If Trim(.SHAFLAG) = "1" Then
                        If Trim(.SMPLEID) = vbNullString Then   'add 2003/06/24 hitec)matsumoto サンプルフラグが
                            DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_FAILURE
                            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP4", "Y011")
                            rs.Close
                            Exit Function
                        End If
                    End If
                End If
                If IsNull(rs!INDTM) = True Then
                    .INDTM = vbNullString
                Else
                    .INDTM = rs!INDTM               ' ウェハーセンター入庫日時
                End If
                If IsNull(rs!BASKETID) = True Then
                    .BASKETID = vbNullString
                Else
                    .BASKETID = rs!BASKETID         ' バスケットID
                End If
                If IsNull(rs!SLOTNO) = True Then
                    .SLOTNO = vbNullString
                Else
                    .SLOTNO = rs!SLOTNO             ' スロットNO
                End If
                If IsNull(rs!CURRWPCS) = True Then
                    .CURRWPCS = vbNullString
                Else
                    .CURRWPCS = rs!CURRWPCS         ' ウェハー枚数
                End If
                If IsNull(rs!EXISTFLG) = True Then
                    .EXISTFLG = vbNullString
                Else
                    .EXISTFLG = rs!EXISTFLG         ' 存在フラグ
                End If
                If IsNull(rs!TOP_POS) = True Then   ' ブロックのTOPからの位置
                    .TOP_POS = 0                    ' (画面表示のずれを防ぐため、サンプル区分により切り上げ
                Else                                '  切捨て処理追加 2003/05/05)
                    dblBlP = CDbl(rs!TOP_POS)
                     .TOP_POS = Int((dblBlP / 10) + 0.9)     '切り上げ
                End If

                If IsNull(rs!REJCAT) = True Then
                    .REJCAT = vbNullString
                Else
                    .REJCAT = rs!REJCAT             ' 欠落理由
                End If
                If IsNull(rs!TXID) = True Then
                    .TXID = vbNullString
                Else
                    .TXID = rs!TXID                 ' トランザクションID
                End If
                If IsNull(rs!REGDATE) = True Then
                    .REGDATE = vbNullString
                Else
                    .REGDATE = rs!REGDATE           ' 登録日付
                End If
                If IsNull(rs!SUMMITSENDFLAG) = True Then
                    .SUMMITSENDFLAG = vbNullString
                Else
                    .SUMMITSENDFLAG = rs!SUMMITSENDFLAG ' SUMIT送信フラグ
                End If
                If IsNull(rs!SENDFLAG) = True Then
                    .SENDFLAG = vbNullString
                Else
                    .SENDFLAG = rs!SENDFLAG         ' 送信フラグ
                End If
                If IsNull(rs!SENDDATE) = True Then
                    .SENDDATE = vbNullString
                Else
                    .SENDDATE = rs!SENDDATE         ' 送信日付
                End If
                If IsNull(rs!HREJCODE) = True Then
                    .HREJCODE = vbNullString
                Else
                    .HREJCODE = rs!HREJCODE         ' 不良理由コード
                End If
                If IsNull(rs!UPDPROC) = True Then
                    .UPDPROC = vbNullString
                Else
                    .UPDPROC = rs!UPDPROC           ' 更新工程
                End If
                If IsNull(rs!UPDDATE) = True Then
                    .UPDDATE = vbNullString
                Else
                    .UPDDATE = rs!UPDDATE           ' 更新日付
                End If
                If IsNull(rs!MREVNUM) = True Then
                    .REVNUM = 0
                Else
                    .REVNUM = rs!MREVNUM             ' 製品番号改訂番号
                End If
                If IsNull(rs!Mfactory) = True Then
                    .factory = vbNullString
                Else
                    .factory = rs!Mfactory           ' 工場
                End If
                If IsNull(rs!Mopecond) = True Then
                    .opecond = vbNullString
                Else
                    .opecond = rs!Mopecond           ' 操業条件
                End If
                If IsNull(rs!KANKBN) = True Then
                    .KANKBN = vbNullString
                Else
                    .KANKBN = rs!KANKBN             ' 完了区分
                End If
                If IsNull(rs!NREJCODE) = True Then
                    .NREJCODE = vbNullString
                Else
                    .NREJCODE = rs!NREJCODE         ' 抜試返答理由コード
                End If
            End With ''''        End If
            itUCount = itUCount + 1
            rs.MoveNext
        Loop
    Next
    '’ループ終了

    DBDRV_MIN_MAX_SEQGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要    :抜試指示　検査項目を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :tBL_SXLID    ,IO   ,type_DBDRV_LOTSXL                     ,ブロックＩＤ、ＳＸＬＩＤ構造体
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :SXLID,BLOCKID→最大、最小（ブロックＰで判定）のデータを取得する
'履歴    :2003/2/25 Hitec)okazaki
Public Function DVDRV_KENSA_KOUMOKU(tKensa() As typ_XSDCW _
                                            ) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i           As Long
''''Dim inCnt       As Long
    Dim sDbName     As String
    Dim itUCount    As Integer
''''Dim tHIN        As tFullHinban
''''Dim sOT1        As String
''''Dim sOT2        As String
''''Dim rtn         As FUNCTION_RETURN

    Dim iIdx        As Integer
    Dim iCnt        As Integer
    Dim iChk        As Integer


    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KENSA_KOUMOKU"

    sDbName = "(V001)"
    'i = 0

''  --TEST-- Y011から取得したサンプルＩＤを使用すると共有部分でどちらも実績を取ってしまうので変更
''''itUCount = UBound(tExamine)
''''ReDim tKensa(itUCount)                      '領域再定義
    itUCount = UBound(tSXLID)
    ReDim tKensa(itUCount * 2)                  '領域再定義
    iIdx = 0

    ''↓追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川
    '後で比較の為に使用するので初期化する
    For iIdx = 0 To itUCount * 2
        tKensa(iIdx).SXLIDCW = ""
    Next iIdx
    iIdx = 0
    ''↑追加START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP石川

    '’ループ開始
    For i = 1 To itUCount
''''    If Trim(tExamine(i).SMPLEID) <> "" Then
        If Trim(tSXLID(i).SXLID) <> "" Then
            ' SXLIDの取得
'            sql = "select "
'            sql = sql & "CRYNUM,"              '結晶番号
'            sql = sql & "INGOTPOS,"            '結晶内位置
'            sql = sql & "SMPKBN,"              'サンプル区分
'            sql = sql & "SMPLID,"              'サンプルID
'            sql = sql & "HINBAN,"              '品番
'            sql = sql & "REVNUM,"              '製品番号改訂番号
'            sql = sql & "FACTORY,"             '工場
'            sql = sql & "OPECOND,"             '操業条件
'            sql = sql & "KTKBN,"               '確定区分
'            sql = sql & "WFINDRS,"             'WF検査指示（Rs)
'            sql = sql & "WFINDOI,"             'WF検査指示（Oi)
'            sql = sql & "WFINDB1,"             'WF検査指示（B1)
'            sql = sql & "WFINDB2,"             'WF検査指示（B2）
'            sql = sql & "WFINDB3,"             'WF検査指示（B3)
'            sql = sql & "WFINDL1,"             'WF検査指示（L1)
'            sql = sql & "WFINDL2,"             'WF検査指示（L2)
'            sql = sql & "WFINDL3,"             'WF検査指示（L3)
'            sql = sql & "WFINDL4,"             'WF検査指示（L4)
'            sql = sql & "WFINDDS,"             'WF検査指示（DS)
'            sql = sql & "WFINDDZ,"             'WF検査指示（DZ)
'            sql = sql & "WFINDSP,"             'WF検査指示（SP)
'            sql = sql & "WFINDDO1,"            'WF検査指示（DO1)
'            sql = sql & "WFINDDO2,"            'WF検査指示（DO2)
'            sql = sql & "WFINDDO3,"            'WF検査指示（DO3)
'            'add start 2003/05/23 hitec)後藤 -------------------------
'            sql = sql & "NVL(WFINDOT1,'0') as DOT1,"            ' WF検査指示（OT1)
'            sql = sql & "NVL(WFINDOT2,'0') as DOT2,"           ' WF検査指示（OT2)
'            'add end   2003/05/23 hitec)後藤 -------------------------
'            sql = sql & "WFRESRS,"             'WF検査実績（Rs)
'            sql = sql & "WFRESOI,"             'WF検査実績（Oi)
'            sql = sql & "WFRESB1,"             'WF検査実績（B1)
'            sql = sql & "WFRESB2,"             'WF検査実績（B2）
'            sql = sql & "WFRESB3,"             'WF検査実績（B3)
'            sql = sql & "WFRESL1,"             'WF検査実績（L1)
'            sql = sql & "WFRESL2,"             'WF検査実績（L2)
'            sql = sql & "WFRESL3,"             'WF検査実績（L3)
'            sql = sql & "WFRESL4,"             'WF検査実績（L4)
'            sql = sql & "WFRESDS,"             'WF検査実績（DS)
'            sql = sql & "WFRESDZ,"             'WF検査実績（DZ)
'            sql = sql & "WFRESSP,"             'WF検査実績（SP)
'            sql = sql & "WFRESDO1,"            'WF検査実績（DO1)
'            sql = sql & "WFRESDO2,"            'WF検査実績（DO2)
'            sql = sql & "WFRESDO3,"            'WF検査実績（DO3)
'            'add start 2003/05/23 hitec)後藤 -------------------------
'            sql = sql & "NVL(WFRESOT1,'0') as SOT1,"            ' WF検査実績（OT1)
'            sql = sql & "NVL(WFRESOT2,'0') as SOT2,"            ' WF検査実績（OT2)
'            'add end   2003/05/23 hitec)後藤 -------------------------
'            sql = sql & "REGDATE,"             '登録日付
'            sql = sql & "UPDDATE,"             '更新日付
'            sql = sql & "SENDFLAG,"            '送信フラグ
'            sql = sql & "SENDDATE"             '送信日付
'
'            sql = sql & " from XSDCW "
'            sql = sql & " where SMPLID ='" & tExamine(i).SMPLEID & "'"

            sql = "select "
            sql = sql & "SXLIDCW,"
            sql = sql & "SMPKBNCW,"
            sql = sql & "TBKBNCW,"
            sql = sql & "REPSMPLIDCW,"
            sql = sql & "XTALCW,"
            sql = sql & "INPOSCW,"
            sql = sql & "HINBCW,"
            sql = sql & "REVNUMCW,"
            sql = sql & "FACTORYCW,"
            sql = sql & "OPECW,"
            sql = sql & "KTKBNCW,"
            sql = sql & "SMCRYNUMCW,"
            sql = sql & "WFSMPLIDRSCW,"
            sql = sql & "NVL(WFSMPLIDRS1CW,'0') as RS1,"
            sql = sql & "NVL(WFSMPLIDRS2CW,'0') as RS2,"
            sql = sql & "WFINDRSCW,"
            sql = sql & "WFRESRS1CW,"
            sql = sql & "WFRESRS2CW,"
            sql = sql & "WFSMPLIDOICW,"
            sql = sql & "WFINDOICW,"
            sql = sql & "WFRESOICW,"
            sql = sql & "WFSMPLIDB1CW,"
            sql = sql & "WFINDB1CW,"
            sql = sql & "WFRESB1CW,"
            sql = sql & "WFSMPLIDB2CW,"
            sql = sql & "WFINDB2CW,"
            sql = sql & "WFRESB2CW,"
            sql = sql & "WFSMPLIDB3CW,"
            sql = sql & "WFINDB3CW,"
            sql = sql & "WFRESB3CW,"
            sql = sql & "WFSMPLIDL1CW,"
            sql = sql & "WFINDL1CW,"
            sql = sql & "WFRESL1CW,"
            sql = sql & "WFSMPLIDL2CW,"
            sql = sql & "WFINDL2CW,"
            sql = sql & "WFRESL2CW,"
            sql = sql & "WFSMPLIDL3CW,"
            sql = sql & "WFINDL3CW,"
            sql = sql & "WFRESL3CW,"
            sql = sql & "WFSMPLIDL4CW,"
            sql = sql & "WFINDL4CW,"
            sql = sql & "WFRESL4CW,"
            sql = sql & "WFSMPLIDDSCW,"
            sql = sql & "WFINDDSCW,"
            sql = sql & "WFRESDSCW,"
            sql = sql & "WFSMPLIDDZCW,"
            sql = sql & "WFINDDZCW,"
            sql = sql & "WFRESDZCW,"
            sql = sql & "WFSMPLIDSPCW,"
            sql = sql & "WFINDSPCW,"
            sql = sql & "WFRESSPCW,"
            sql = sql & "WFSMPLIDDO1CW,"
            sql = sql & "WFINDDO1CW,"
            sql = sql & "WFRESDO1CW,"
            sql = sql & "WFSMPLIDDO2CW,"
            sql = sql & "WFINDDO2CW,"
            sql = sql & "WFRESDO2CW,"
            sql = sql & "WFSMPLIDDO3CW,"
            sql = sql & "WFINDDO3CW,"
            sql = sql & "WFRESDO3CW,"
            sql = sql & "WFSMPLIDOT1CW,"
            sql = sql & "WFSMPLIDOT2CW,"
            'add start 2003/05/23 hitec)後藤 -------------------------
            sql = sql & "NVL(WFINDOT1CW,   '0')     as DOT1,"           ' WF検査指示（OT1)
            sql = sql & "NVL(WFINDOT2CW,   '0')     as DOT2,"           ' WF検査指示（OT2)
            'add end   2003/05/23 hitec)後藤 -------------------------
            'add start 2003/05/23 hitec)後藤 -------------------------
            sql = sql & "NVL(WFRESOT1CW,   '0')     as SOT1,"           ' WF検査実績（OT1)
            sql = sql & "NVL(WFRESOT2CW,   '0')     as SOT2,"           ' WF検査実績（OT2)
            'add end   2003/05/23 hitec)後藤 -------------------------
            sql = sql & "NVL(WFSMPLIDAOICW,'0')     as sAOI,"
            sql = sql & "NVL(WFINDAOICW,   '0')     as iAOI,"
            sql = sql & "NVL(WFRESAOICW,   '0')     as rAOI,"
            sql = sql & "NVL(SMPLNUMCW,    '0')     as sNUM,"
            sql = sql & "NVL(SMPLPATCW,    '0')     as PAT,"
            sql = sql & "NVL(TSTAFFCW,     '0')     as STF,"
            sql = sql & "TDAYCW,"
            sql = sql & "NVL(KSTAFFCW,     '0')     as kSTF,"
            sql = sql & "KDAYCW,"
            sql = sql & "NVL(SNDKCW,       '0')     as SND,"
            sql = sql & "NVL(SNDDAYCW,'2003/09/18') as sDAY,"

            '' GD追加　05/01/31 ooba START =====================================>
            sql = sql & "NVL(WFSMPLIDGDCW,'0')     as sGD,"
            sql = sql & "NVL(WFINDGDCW,   '0')     as iGD,"
            sql = sql & "NVL(WFRESGDCW,   '0')     as rGD,"
            sql = sql & "NVL(WFHSGDCW,   '0')      as hGD"
            '' GD追加　05/01/31 ooba END =======================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
            sql = sql & ",NVL(EPSMPLIDB1CW,   '0')  as EPSMPLIDB1CW,"
            sql = sql & "NVL(EPINDB1CW,   '0')      as EPINDB1CW,"
            sql = sql & "NVL(EPRESB1CW,   '0')      as EPRESB1CW,"
            sql = sql & "NVL(EPSMPLIDB2CW,   '0')   as EPSMPLIDB2CW,"
            sql = sql & "NVL(EPINDB2CW,   '0')      as EPINDB2CW,"
            sql = sql & "NVL(EPRESB2CW,   '0')      as EPRESB2CW,"
            sql = sql & "NVL(EPSMPLIDB3CW,   '0')   as EPSMPLIDB3CW,"
            sql = sql & "NVL(EPINDB3CW,   '0')      as EPINDB3CW,"
            sql = sql & "NVL(EPRESB3CW,   '0')      as EPRESB3CW,"
            sql = sql & "NVL(EPSMPLIDL1CW,   '0')   as EPSMPLIDL1CW,"
            sql = sql & "NVL(EPINDL1CW,   '0')      as EPINDL1CW,"
            sql = sql & "NVL(EPRESL1CW,   '0')      as EPRESL1CW,"
            sql = sql & "NVL(EPSMPLIDL2CW,   '0')   as EPSMPLIDL2CW,"
            sql = sql & "NVL(EPINDL2CW,   '0')      as EPINDL2CW,"
            sql = sql & "NVL(EPRESL2CW,   '0')      as EPRESL2CW,"
            sql = sql & "NVL(EPSMPLIDL3CW,   '0')   as EPSMPLIDL3CW,"
            sql = sql & "NVL(EPINDL3CW,   '0')      as EPINDL3CW,"
            sql = sql & "NVL(EPRESL3CW,   '0')      as EPRESL3CW"
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

            sql = sql & " from  XSDCW "
            sql = sql & " where SXLIDCW ='" & tSXLID(i).SXLID & "'"
            sql = sql & "   and LIVKCW  ='0'"                           ' 生死区分は必ず確認する事
            sql = sql & " order by INPOSCW"
'''''       sql = sql & " where REPSMPLIDCW ='" & tExamine(i).SMPLEID & "'"

            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            '''抽出レコードが存在ならば該当
            If Not rs.EOF Then
'                With tKensa(i)
'                   .CRYNUM = rs!CRYNUM
'                    .INGOTPOS = rs!INGOTPOS
'                    .SMPKBN = rs!SMPKBN
'                    .SMPLID = rs!SMPLID
'                    .hinban = rs!hinban
'                    .REVNUM = rs!REVNUM
'                    .factory = rs!factory
'                    .opecond = rs!opecond
'                    .KTKBN = rs!KTKBN
'                    .WFINDRS = rs!WFINDRS
'                    .WFINDOI = rs!WFINDOI
'                    .WFINDB1 = rs!WFINDB1
'                    .WFINDB2 = rs!WFINDB2
'                    .WFINDB3 = rs!WFINDB3
'                    .WFINDL1 = rs!WFINDL1
'                    .WFINDL2 = rs!WFINDL2
'                    .WFINDL3 = rs!WFINDL3
'                    .WFINDL4 = rs!WFINDL4
'                    .WFINDDS = rs!WFINDDS
'                    .WFINDDZ = rs!WFINDDZ
'                    .WFINDSP = rs!WFINDSP
'                    .WFINDDO1 = rs!WFINDDO1
'                    .WFINDDO2 = rs!WFINDDO2
'                    .WFINDDO3 = rs!WFINDDO3
'                    tHin.hinban = .hinban
'                    tHin.factory = .factory
'                    tHin.mnorevno = .REVNUM
'                    tHin.opecond = .opecond
'                    rtn = scmzc_getE036(tHin, sOT1, sOT2)
'                    If rtn = FUNCTION_RETURN_FAILURE Then
'                        rs.Close
'                        DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
'                        GoTo proc_exit
'                    End If
'                    If sOT1 = "1" Then
'                        .WFINDOT1 = rs!DOT1 '03/05/23
'                    Else
'                        .WFINDOT1 = 0 '03/05/23
'                    End If
'                    If sOT2 = "1" Then
'                        .WFINDOT2 = rs!DOT2 '03/05/23
'                    Else
'                        .WFINDOT2 = 0 '03/05/23
'                    End If
'                    .WFRESRS = rs!WFRESRS
'                    .WFRESOI = rs!WFRESOI
'                    .WFRESB1 = rs!WFRESB1
'                    .WFRESB2 = rs!WFRESB2
'                    .WFRESB3 = rs!WFRESB3
'                    .WFRESL1 = rs!WFRESL1
'                    .WFRESL2 = rs!WFRESL2
'                    .WFRESL3 = rs!WFRESL3
'                    .WFRESL4 = rs!WFRESL4
'                    .WFRESDS = rs!WFRESDS
'                    .WFRESDZ = rs!WFRESDZ
'                    .WFRESSP = rs!WFRESSP
'                    .WFRESDO1 = rs!WFRESDO1
'                    .WFRESDO2 = rs!WFRESDO2
'                    .WFRESDO3 = rs!WFRESDO3
'                    .WFRESOT1 = rs!sOT1 '03/05/23
'                    .WFRESOT2 = rs!sOT2 '03/05/23
'                    .REGDATE = rs!REGDATE
'                    .UPDDATE = rs!UPDDATE
'                    .SENDFLAG = rs!SENDFLAG
'                    .SENDDATE = rs!SENDDATE
'                End With

                iCnt = 0
                Do While Not rs.EOF
                    iIdx = iIdx + 1
                    iCnt = iCnt + 1
                    ' ３件目以降が存在する場合エラー
                    If iCnt > 2 Then
                        Exit Do
                    End If

                    If rs!TBKBNCW = "T" Then
                        For iChk = 1 To iIdx - 1
                            If tKensa(iChk).SXLIDCW = rs!SXLIDCW And tKensa(iChk).TBKBNCW = rs!TBKBNCW Then
                                Exit For
                            End If
                        Next
                    Else
                        iChk = iIdx
                    End If

                    If iChk = iIdx Then
                        With tKensa(iIdx)
                            .SXLIDCW = rs!SXLIDCW
                            .SMPKBNCW = rs!SMPKBNCW
                            .TBKBNCW = rs!TBKBNCW
                            .REPSMPLIDCW = rs!REPSMPLIDCW
                            .XTALCW = rs!XTALCW
                            .INPOSCW = rs!INPOSCW
                            .HINBCW = rs!HINBCW
                            .REVNUMCW = rs!REVNUMCW
                            .FACTORYCW = rs!FACTORYCW
                            .OPECW = rs!OPECW
                            .KTKBNCW = rs!KTKBNCW
                            .SMCRYNUMCW = rs!SMCRYNUMCW
                            .WFSMPLIDRSCW = rs!WFSMPLIDRSCW
                            .WFSMPLIDRS1CW = rs!rs1
                            .WFSMPLIDRS2CW = rs!rs2
                            .WFINDRSCW = rs!WFINDRSCW
                            .WFRESRS1CW = rs!WFRESRS1CW
                            .WFSMPLIDOICW = rs!WFSMPLIDOICW
                            .WFINDOICW = rs!WFINDOICW
                            .WFRESOICW = rs!WFRESOICW
                            .WFSMPLIDB1CW = rs!WFSMPLIDB1CW
                            .WFINDB1CW = rs!WFINDB1CW
                            .WFRESB1CW = rs!WFRESB1CW
                            .WFSMPLIDB2CW = rs!WFSMPLIDB2CW
                            .WFINDB2CW = rs!WFINDB2CW
                            .WFRESB2CW = rs!WFRESB2CW
                            .WFSMPLIDB3CW = rs!WFSMPLIDB3CW
                            .WFINDB3CW = rs!WFINDB3CW
                            .WFRESB3CW = rs!WFRESB3CW
                            .WFSMPLIDL1CW = rs!WFSMPLIDL1CW
                            .WFINDL1CW = rs!WFINDL1CW
                            .WFRESL1CW = rs!WFRESL1CW
                            .WFSMPLIDL2CW = rs!WFSMPLIDL2CW
                            .WFINDL2CW = rs!WFINDL2CW
                            .WFRESL2CW = rs!WFRESL2CW
                            .WFSMPLIDL3CW = rs!WFSMPLIDL3CW
                            .WFINDL3CW = rs!WFINDL3CW
                            .WFRESL3CW = rs!WFRESL3CW
                            .WFSMPLIDL4CW = rs!WFSMPLIDL4CW
                            .WFINDL4CW = rs!WFINDL4CW
                            .WFRESL4CW = rs!WFRESL4CW
                            .WFSMPLIDDSCW = rs!WFSMPLIDDSCW
                            .WFINDDSCW = rs!WFINDDSCW
                            .WFRESDSCW = rs!WFRESDSCW
                            .WFSMPLIDDZCW = rs!WFSMPLIDDZCW
                            .WFINDDZCW = rs!WFINDDZCW
                            .WFRESDZCW = rs!WFRESDZCW
                            .WFSMPLIDSPCW = rs!WFSMPLIDSPCW
                            .WFINDSPCW = rs!WFINDSPCW
                            .WFRESSPCW = rs!WFRESSPCW
                            .WFSMPLIDDO1CW = rs!WFSMPLIDDO1CW
                            .WFINDDO1CW = rs!WFINDDO1CW
                            .WFRESDO1CW = rs!WFRESDO1CW
                            .WFSMPLIDDO2CW = rs!WFSMPLIDDO2CW
                            .WFINDDO2CW = rs!WFINDDO2CW
                            .WFRESDO2CW = rs!WFRESDO2CW
                            .WFSMPLIDDO3CW = rs!WFSMPLIDDO3CW
                            .WFINDDO3CW = rs!WFINDDO3CW
                            .WFRESDO3CW = rs!WFRESDO3CW
                            .WFSMPLIDOT1CW = rs!WFSMPLIDOT1CW
                            .WFINDOT1CW = rs!DOT1
                            .WFRESOT1CW = rs!sOT1
                            .WFSMPLIDOT2CW = rs!WFSMPLIDOT2CW
                            .WFINDOT2CW = rs!DOT2
                            .WFRESOT2CW = rs!sOT2
                            .WFSMPLIDAOICW = rs!sAOI
                            .WFINDAOICW = rs!iAOI
                            .WFRESAOICW = rs!rAOI
                            .SMPLNUMCW = rs!sNum
                            .SMPLPATCW = rs!PAT
                            .TSTAFFCW = rs!STF
                            .TDAYCW = rs!TDAYCW
                            .KSTAFFCW = rs!kSTF
                            .KDAYCW = rs!KDAYCW
                            .SNDKCW = rs!SND
                            .SNDDAYCW = rs!sDay
                            .WFSMPLIDGDCW = rs!sGD      'ｻﾝﾌﾟﾙID(GD)    '05/01/31 ooba START ====>
                            .WFINDGDCW = rs!iGD         '状態FLG(GD)
                            .WFRESGDCW = rs!rGD         '実績FLG(GD)
                            .WFHSGDCW = rs!hGD          '保証FLG(GD)    '05/01/31 ooba END ======>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                            .EPSMPLIDB1CW = rs!EPSMPLIDB1CW
                            .EPINDB1CW = rs!EPINDB1CW
                            .EPRESB1CW = rs!EPRESB1CW
                            .EPSMPLIDB2CW = rs!EPSMPLIDB2CW
                            .EPINDB2CW = rs!EPINDB2CW
                            .EPRESB2CW = rs!EPRESB2CW
                            .EPSMPLIDB3CW = rs!EPSMPLIDB3CW
                            .EPINDB3CW = rs!EPINDB3CW
                            .EPRESB3CW = rs!EPRESB3CW
                            .EPSMPLIDL1CW = rs!EPSMPLIDL1CW
                            .EPINDL1CW = rs!EPINDL1CW
                            .EPRESL1CW = rs!EPRESL1CW
                            .EPSMPLIDL2CW = rs!EPSMPLIDL2CW
                            .EPINDL2CW = rs!EPINDL2CW
                            .EPRESL2CW = rs!EPRESL2CW
                            .EPSMPLIDL3CW = rs!EPSMPLIDL3CW
                            .EPINDL3CW = rs!EPINDL3CW
                            .EPRESL3CW = rs!EPRESL3CW
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                        End With
                    End If


                    If rs!TBKBNCW = "B" Then
                        For iChk = iIdx - 1 To 1 Step -1
                            If tKensa(iChk).SXLIDCW = rs!SXLIDCW And tKensa(iChk).TBKBNCW = rs!TBKBNCW Then
                                Exit For
                            End If
                        Next
                        If iChk > 0 Then
                            tKensa(iChk) = tKensa(0)
                        End If
                    End If


                    rs.MoveNext
                Loop
                rs.Close

                ' 取得件数が２件でない場合エラー
                If iCnt <> 2 Then
                    f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")
                    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                    Exit Function
                End If
           Else
                rs.Close
                f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")    '03/06/06 後藤
                DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_FAILURE
                Exit Function
            End If

        End If
    Next i
    '’ループ終了

    DVDRV_KENSA_KOUMOKU = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GoTo proc_exit
End Function

'概要    :抜試指示　検査項目を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                 ,説明
'        :tWafSmp      ,I    ,typ_XSDCW          ,ｻﾝﾌﾟﾙ管理構造体(結晶全体)
'        :tKensa       ,O    ,typ_XSDCW          ,ｻﾝﾌﾟﾙ管理構造体
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN    ,読み込み成否
'説明    :
'履歴    :08/02/04 ooba
Public Function DVDRV_KENSA_KOUMOKU_LOCAL(tWafSmp() As typ_XSDCW, tKensa() As typ_XSDCW) As FUNCTION_RETURN

    Dim i, j        As Integer
    Dim sql         As String
    Dim recCnt      As Integer
    Dim iChk        As Integer
    Dim bTflg       As Boolean
    Dim bBflg       As Boolean
    
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DVDRV_KENSA_KOUMOKU_LOCAL"
    
    ReDim tKensa(UBound(tSXLID) * 2)
    recCnt = 0
    
    '初期化
    For i = 0 To UBound(tKensa)
        tKensa(i).SXLIDCW = ""
    Next i
    
    For i = 1 To UBound(tSXLID)
        bTflg = False
        bBflg = False
        'TOP
        For iChk = 1 To UBound(tSXLID)
            If tSXLID(iChk).SXLID = tSXLID(i).SXLID Then Exit For
        Next iChk
        If iChk = i Then
            For j = 1 To UBound(tWafSmp)
                If tSXLID(i).SXLID = tWafSmp(j).SXLIDCW And tWafSmp(j).TBKBNCW = "T" Then
                    recCnt = recCnt + 1
                    tKensa(recCnt) = tWafSmp(j)
                    bTflg = True
                    Exit For
                End If
            Next j
        Else
            recCnt = recCnt + 1
            tKensa(recCnt) = tKensa(0)
            bTflg = True
        End If
        
        'BOT
        For iChk = UBound(tSXLID) To 1 Step -1
            If tSXLID(iChk).SXLID = tSXLID(i).SXLID Then Exit For
        Next iChk
        If iChk = i Then
            For j = 1 To UBound(tWafSmp)
                If tSXLID(i).SXLID = tWafSmp(j).SXLIDCW And tWafSmp(j).TBKBNCW = "B" Then
                    recCnt = recCnt + 1
                    tKensa(recCnt) = tWafSmp(j)
                    bBflg = True
                    Exit For
                End If
            Next j
        Else
            recCnt = recCnt + 1
            tKensa(recCnt) = tKensa(0)
            bBflg = True
        End If
        
        '存在ﾁｪｯｸ
        If bTflg = False Or bBflg = False Then
            f_cmbc036_2.lblMsg.Caption = GetMsgStr("ENSP2")
            DVDRV_KENSA_KOUMOKU_LOCAL = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    Next i
    
    DVDRV_KENSA_KOUMOKU_LOCAL = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    GoTo proc_exit
    
End Function

'概要    :抜試指示 欠落情報のブロック結晶内開始位置を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :SXLID,BLOCKID→最大、最小（ブロックＰで判定）のデータを取得する
'履歴    :2003/3/05 Hitec)okazaki
Public Function DVDRV_KETURAKU_Ingotget(ByVal sLotid As String, iIngotpos As Integer) As FUNCTION_RETURN
    Dim sql         As String
    Dim rs          As OraDynaset
''''Dim i           As Long
''''Dim inCnt       As Long
    Dim sDbName     As String
''''Dim itUCount    As Integer

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DVDRV_KETURAKU_Ingotget"

    sDbName = "(V001)"

    sql = "select INGOTPOS"            ' 結晶内開始位置
    sql = sql & " from  TBCME040"
    sql = sql & " where BLOCKID = '" & sLotid & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    '''抽出レコードが存在ならば該当
    If Not rs.EOF Then
        iIngotpos = Int(CDbl(rs!INGOTPOS))
    End If
    rs.Close

    DVDRV_KETURAKU_Ingotget = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function



'2003/02/28 hitec)okazaki ADD end
'********************************************************************************



'概要    :抜試指示 入力したブロックＰから、該当ＷＦを検索
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :iBlkP    　　,O    ,integer                               ,ブロックP
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :抜試指示 入力したブロックＰから、該当ＷＦを検索
'履歴    :2003/2/25 Hitec)matsumoto
Public Function DBDRV_GET_WFMAP(ByVal sBlkId As String, ByVal iBlkP As Integer, _
                                ByRef sBlkP As Variant, ByRef sKessyoP As Variant, _
                                ByRef sBlkSeq As Variant, ByRef sBlkSeq2 As Variant, ByRef sSmpId1 As Variant, _
                                ByRef sSmpId2 As Variant, ByRef iNextBlkP As Integer, _
                                ByRef vWfNum As Variant, iKbnFlg As Integer) _
                                    As FUNCTION_RETURN

    Dim sql         As String
    Dim rs          As OraDynaset
    Dim i, j        As Long
    Dim inCnt       As Long
    Dim sDbName     As String
    Dim iLoopCnt    As Integer
    Dim dChkBlkP    As Double
'   Dim dChkBlkP    As Double
    Dim iTopPos     As Integer
    Dim sAddSmpId1  As String
    Dim sAddSmpId2  As String
    Dim iBlkflg     As Integer
    Dim vBlkId      As Variant
    Dim sSXLID      As String
    Dim dblWFLen    As Double
    Dim eps         As Double

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_GET_WFMAP"

    sDbName = "(Y011)"
    i = 0
    eps = 0.000001

    sql = "select "
    sql = sql & "LOTID,"                ' ブロックID"
''''sql = sql & "SXLID,"                ' SXLID"
    sql = sql & "MSXLID,"               ' SXLID"
    sql = sql & "blockseq,"             ' ブロック内連番"
    sql = sql & "WFSTA,"                ' WF状態"
    sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
    sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
    sql = sql & "MSMPLEID,"             ' 抜試位置"
    sql = sql & "SHAFLAG,"              ' サンプルフラグ"
    sql = sql & "TOP_POS"               ' ブロック内位置
    sql = sql & " from TBCMY011 "
''''sql = sql & " where SXLID ='" & sSxlId & "'"
    sql = sql & " where LOTID ='" & sBlkId & "'"
    sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    sql = sql & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    iLoopCnt = 0
    vWfNum = 0
    Do While Not rs.EOF
        If IsNull(rs.Fields("RTOP_POS")) = True Then
            dChkBlkP = 0
        Else
            dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
        End If
        If (iBlkP < dChkBlkP) And (dChkBlkP <= iNextBlkP) Then
            vWfNum = CInt(vWfNum) + 1
        End If
        rs.MoveNext
    Loop
    rs.Close

    sql = "select "
    sql = sql & "LOTID,"                ' ブロックID"
    sql = sql & "MSXLID,"               ' SXLID"
    sql = sql & "blockseq,"             ' ブロック内連番"
    sql = sql & "WFSTA,"                ' WF状態"
    sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
    sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
    sql = sql & "MSMPLEID,"             ' 抜試位置"
    sql = sql & "SHAFLAG,"              ' サンプルフラグ"
    sql = sql & "TOP_POS"               ' ブロック内位置
    sql = sql & " from TBCMY011 "
''''sql = sql & " where SXLID ='" & sSxlId & "'"
    sql = sql & " where LOTID ='" & sBlkId & "'"
''''sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
    sql = sql & " ORDER BY BLOCKSEQ ASC"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    iLoopCnt = 0
    rs.MoveFirst
    Do While Not rs.EOF
        Select Case Right(sSmpId1, 1)
            Case "T"
                If iKbnFlg = 0 Then     '前ブロックの位置、と次ブロックのT
                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                    End If
                    If dChkBlkP > iBlkP Or dChkBlkP = iBlkP Then
                        If dChkBlkP > iBlkP Then
                            rs.MovePrevious
                        End If
                            'WFの欠落を判定（CW740ではほぼありえない)
                            If rs.Fields("WFSTA") = "4" Then
                                rs.Close
                                DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                                Exit Function
                            End If

                            If IsNull(rs.Fields("RTOP_POS")) = False Then
    ''''                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
                                sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            End If
                            If IsNull(rs.Fields("RITOP_POS")) = False Then
    ''''                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                                sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps) 'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            End If
                            If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                            End If
            ''''                    sBlkId = CStr(rs.Fields("LOTID"))
        '                    iTopPos = Int(CInt(rs.Fields("TOP_POS")) + 0.9) '切り上げ
        '                    sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "B"
                        rs.Close
                        'XXX-XXXT
                        '現在のブロックIDの次のブロックIDを取得
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If iBlkflg = 1 Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '次のBLID取得
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        iBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With


                        sql = "select "
                        sql = sql & "LOTID,"                ' ブロックID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' ブロック内連番"
                        sql = sql & "WFSTA,"                ' WF状態"
                        sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
                        sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
                        sql = sql & "MSMPLEID,"             ' 抜試位置"
                        sql = sql & "SHAFLAG,"              ' サンプルフラグ"
                        sql = sql & "TOP_POS"               ' ブロック内位置
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
''                      sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
                        sql = sql & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If

                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            If DBDRV_WFLENGET(sBlkId, dblWFLen) = FUNCTION_RETURN_SUCCESS Then
                                iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) - dblWFLen + eps)   'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            Else
                                iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            End If
                        End If
        '                        If IsNull(rs.Fields("RITOP_POS")) = False Then
        '                            sNextIngotP = rs.Fields("RITOP_POS")
        '                        End If

                         sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))

        '                rs.MoveFirst
        '                If IsNull(rs.Fields("RTOP_POS")) = False Then
        '                    sBlkP = rs.Fields("RTOP_POS")
        '                End If
        ''                If IsNull(rs.Fields("RITOP_POS")) = False Then
        ''                    sKessyoP = rs.Fields("RITOP_POS")
        ''                End If
        '                sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
        ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "T"
                        Exit Do
                    End If

                 Else   'そのブロックのT(以前のままのロジック）

                    rs.MoveFirst
                    'WFの欠落を判定（CW740ではほぼありえない)
                    If rs.Fields("WFSTA") = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If

                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        sBlkP = rs.Fields("RTOP_POS")
                    End If
                    If IsNull(rs.Fields("RITOP_POS")) = False Then
                        sKessyoP = rs.Fields("RITOP_POS")
                    End If
                    sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                    iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10 + eps) '切り捨て  'add 2003/06/13 hitec)matsumoto [+ eps]追加
                    sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "T"
                    Exit Do
                    End If
            Case "U"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
'===ADD okazaki 2003/04/18
'別ブロックをはさむサンプルへの対応
                If dChkBlkP > CInt(sBlkP) Or dChkBlkP = CInt(sBlkP) Then
                    If dChkBlkP > CInt(sBlkP) Then
                        rs.MovePrevious

                        If IsNull(rs.Fields("BLOCKSEQ")) = True Then    'add 2003/04/28 hitec)matsumoto  NULLの場合（該当WF無し）は、下に検索する
                            Do
                                rs.MoveNext
                                If IsNull(rs.Fields("RTOP_POS")) = False Then
                                    Exit Do
                                End If
                            Loop
                        End If
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
''''                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)   'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '切り上げ    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        rs.MoveNext
    '                    If sSmpId2 <> vbNullString Then 'Dのサンプルを作成
    '                        iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10)  '切り捨て
    '                        sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"
    '                    End If
    '                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
    '                    Exit Do
                    ElseIf dChkBlkP = CInt(sBlkP) Then
''''                    'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]追加
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
    ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '切り上げ    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        rs.MoveNext
                    End If

                    If Not rs.EOF Then
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'Dのサンプルを作成
                            '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                iTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)    '0.1mm引いて切捨て
                            Else
                                iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '切り捨て 'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    Else
                        '現在のブロックIDの次のブロックIDを取得
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For i = 1 To .MaxRows
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If iBlkflg = 1 Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '次のBLID取得
                                        Exit For
                                    ElseIf Right(sBlkId, 3) = vBlkId Then
                                        iBlkflg = 1
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sql = "select "
                        sql = sql & "LOTID,"                ' ブロックID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' ブロック内連番"
                        sql = sql & "WFSTA,"                ' WF状態"
                        sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
                        sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
                        sql = sql & "MSMPLEID,"             ' 抜試位置"
                        sql = sql & "SHAFLAG,"              ' サンプルフラグ"
                        sql = sql & "TOP_POS"               ' ブロック内位置
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
'                       sql = sql & "   AND TO_NUMBER(WFSTA) <= 1"
                        sql = sql & " ORDER BY BLOCKSEQ ASC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            iNextBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        End If
'                        If IsNull(rs.Fields("RITOP_POS")) = False Then
'                            sNextIngotP = rs.Fields("RITOP_POS")
'                        End If
                        If sSmpId2 <> vbNullString Then 'Dのサンプルを作成
                            '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                            If rs.Fields("TOP_POS") > 0 Then
                                iNextBlkP = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)  '0.1mm引いて切捨て
                            Else
                                iNextBlkP = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)  '切り捨て   'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            End If
                            sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iNextBlkP), "000") & "D"
                        End If
                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                        Exit Do
                    End If
                End If

            Case "D"
                If IsNull(rs.Fields("RTOP_POS")) = False Then
                    dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                End If
''''                    If iChkBlkP < iBlkP Then
''''                        sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
''''                        sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
''''                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
''''                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
''''                        sBlkId = CStr(rs.Fields("LOTID"))
''''                        iTopPos = Int(CInt(rs.Fields("TOP_POS")) / 10)  '切り捨て
''''                        sSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & sSmpId1
''''                        rs.MoveNext
''''                        sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
''''                        Exit Do
''''                    End If
                If dChkBlkP > iBlkP Then
'                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
'                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                    'WFの欠落を判定（CW740ではほぼありえない)
                    If rs.Fields("WFSTA") = "4" Then
                        rs.Close
                        DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                        Exit Function
                    End If
                    sBlkSeq2 = CStr(rs.Fields("BLOCKSEQ"))
                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + eps)  'add 2003/06/13 hitec)matsumoto [+ eps]追加
 ''''                        sBlkId = CStr(rs.Fields("LOTID"))
                    '0以外は0.1mm引いて切捨て(WF操業:Dは該当位置を含まずに下方向抜取り) 08/11/06 ooba
                    If rs.Fields("TOP_POS") > 0 Then
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS") - 1) / 10 + eps)    '0.1mm引いて切捨て
                    Else
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + eps)                        '切り捨て 'add 2003/06/13 hitec)matsumoto [+ eps]追加
                    End If
                    sAddSmpId2 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "D"        'DなのでsAddSmpId2に入れる
                    rs.MovePrevious

                    'Uのサンプル作成の修正(複数ﾌﾞﾛｯｸに対応) 2005/04/21 ffc)tanabe =============================> START
                    If Not rs.BOF Then
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'Uのサンプルを作成
                            iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps)              '切り上げ   'add 2003/06/13 hitec)matsumoto [+ eps]追加
                            sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"    '"U"なのでsAddSmpId1に入れる
                        End If
    '                    sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
    '                    sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        End If
                        If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                            sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        End If
                        Exit Do
                    Else
                         
                         '現在のブロックIDのRowsを取得     ## 2008.02.08
                         With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            For j = .MaxRows To 1 Step -1
                                .GetText 1, j, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If Right(sBlkId, 3) = vBlkId Then
                                        Exit For
                                    End If
                                End If
                            Next j
                        End With
                    
                        '現在のブロックIDの前のブロックIDを取得
                        With f_cmbc036_2.sprExamine
                            iBlkflg = 0
                            'For i = .MaxRows To 1 Step -1    '## 2008.02.08
                            For i = j To 1 Step -1
                                .GetText 1, i, vBlkId
                                If vBlkId <> "" And Len(vBlkId) <> 1 Then
                                    If Right(sBlkId, 3) <> vBlkId Then
                                        sBlkId = Left(sBlkId, 9) & CStr(vBlkId) '前のBLID取得
                                        Exit For
                                    End If
                                End If
                            Next i
                        End With
                        rs.Close

                        sql = "select "
                        sql = sql & "LOTID,"                ' ブロックID"
                        sql = sql & "MSXLID,"               ' SXLID"
                        sql = sql & "blockseq,"             ' ブロック内連番"
                        sql = sql & "WFSTA,"                ' WF状態"
                        sql = sql & "RTOP_POS,"             ' 論理ブロック内位置"
                        sql = sql & "RITOP_POS,"            ' 論理結晶内位置"
                        sql = sql & "MSMPLEID,"             ' 抜試位置"
                        sql = sql & "SHAFLAG,"              ' サンプルフラグ"
                        sql = sql & "TOP_POS"               ' ブロック内位置
                        sql = sql & " from TBCMY011 "
                        sql = sql & " where LOTID ='" & sBlkId & "'"
                        sql = sql & " ORDER BY BLOCKSEQ DESC"

                        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

                        iLoopCnt = 0
                        rs.MoveFirst
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If
                        If sSmpId2 <> vbNullString Then 'Uのサンプルを作成
                            iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps)              '切り上げ
                            sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "U"
                        End If
                        If IsNull(rs.Fields("RTOP_POS")) = False Then
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)
                        End If
                        If IsNull(rs.Fields("BLOCKSEQ")) = False Then
                            sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        End If
                        Exit Do
                    End If
                    'Uのサンプル作成の修正(複数ﾌﾞﾛｯｸに対応) 2005/04/21 ffc)tanabe =============================> END
                End If
            Case "B"
'                rs.MoveLast
                    If IsNull(rs.Fields("RTOP_POS")) = False Then
                        dChkBlkP = CDbl(rs.Fields("RTOP_POS"))
                    End If
                    If dChkBlkP > iBlkP Or dChkBlkP = iBlkP Then
                        If dChkBlkP > iBlkP Then
                            rs.MovePrevious
                        End If
                        'WFの欠落を判定（CW740ではほぼありえない)
                        If rs.Fields("WFSTA") = "4" Then
                            rs.Close
                            DBDRV_GET_WFMAP = FUNCTION_RETURN_FAILURE
                            Exit Function
                        End If

                        If IsNull(rs.Fields("RTOP_POS")) = False Then
''''                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")))
                            sBlkP = Int(CDbl(rs.Fields("RTOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        End If
                        If IsNull(rs.Fields("RITOP_POS")) = False Then
''''                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")))
                            sKessyoP = Int(CDbl(rs.Fields("RITOP_POS")) + 0.9 + eps)    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        End If
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
                        sBlkSeq = CStr(rs.Fields("BLOCKSEQ"))
        ''''                    sBlkId = CStr(rs.Fields("LOTID"))
                        iTopPos = Int(CDbl(rs.Fields("TOP_POS")) / 10 + 0.9 + eps) '切り上げ    'add 2003/06/13 hitec)matsumoto [+ eps]追加
                        sAddSmpId1 = Mid(sBlkId, 1, 12) & Format(CStr(iTopPos), "000") & "B"
                        Exit Do
                    End If

        End Select
        rs.MoveNext
    Loop
    sSmpId1 = sAddSmpId1
    sSmpId2 = sAddSmpId2
    rs.Close

    DBDRV_GET_WFMAP = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function




'概要    :WFマップテーブル更新
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                                    ,説明
'        :戻ﾘ値        ,O    ,FUNCTION_RETURN                       ,読み込み成否
'説明    :WFマップテーブル(TBCMY011)を更新する
'履歴    :2003/3/25 Hitec)matsumoto
Public Function DBDRV_UPD_WFMap() As FUNCTION_RETURN

    Dim sql             As String
    Dim rs              As OraDynaset
    Dim i               As Long
    Dim iLoopCnt        As Long
    Dim sDbName         As String
    Dim itUCount        As Integer
''''Dim nowtime         As Date
    Dim vGetMaxSeq      As Variant
    Dim sGetSxlId       As String
    Dim vGetSXLID1      As Variant
    Dim vGetSXLID2      As Variant
    Dim NowIngotPos     As Integer
    Dim iGetSmplLoop    As Integer
    Dim iFromBlkSeq     As Integer
    Dim iToBlkSeq       As Integer
    Dim iNextLoopCnt    As Integer
    Dim vGetSample      As Variant
    Dim iBlkflg         As Integer

    Dim sLotid          As String
    Dim iFromIngotPos   As Integer
    Dim iToIngotPos     As Integer
    Dim vGetHinban      As Variant
    Dim m               As Integer
    Dim k               As Integer      '2003/05/18 add
    Dim sOldSXLID       As String       '2003/05/18 add
    Dim sOldIngotP      As String       '2003/05/18 add
    Dim vGetBlockSEQ_S  As Variant      '2003/05/29
    Dim vGetBlockSEQ_E  As Variant      '2003/05/29

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc039_SQL.bas -- Function DBDRV_UPD_WFMap"

    sDbName = "(Y011)"
'品番を1列追加したことによる列の変更-------start iida 2003/09/03
    '2003/05/01
    With f_cmbc036_2.sprExamine
        m = .MaxRows
        'SXLの更新
        For iLoopCnt = 1 To m Step 2

            'サンプル行でSXLを判定する
            .row = iLoopCnt
            .col = 10
            If Len(Trim(.Text)) > 0 Then     'サンプル行の場合
                .GetText 5, iLoopCnt, gtSprWfMap(iLoopCnt).KESSYOUP 'add 2003/05/17 hitec)matsumoto 結晶位置を画面から取得
                sGetSxlId = Mid(gtSprWfMap(iLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(gtSprWfMap(iLoopCnt).KESSYOUP))
                If iLoopCnt = 1 Then    '先頭
                    sGetSxlId = Mid(gtSprWfMap(iLoopCnt).LOTID, 1, 10) & GetWafPos(CInt(SIngotP))
                End If

                '#######2003/05/18 okazaki
                If Get_OLDSXLID(CInt(gtSprWfMap(iLoopCnt).KESSYOUP), sOldSXLID, sOldIngotP) = FUNCTION_RETURN_SUCCESS Then

                    sGetSxlId = sOldSXLID
                    If iLoopCnt = 1 Then
                        gtSprWfMap(iLoopCnt).KESSYOUP = SIngotP   ' 2003/05/18 okazaki
                    Else
                        gtSprWfMap(iLoopCnt).KESSYOUP = sOldIngotP  '予備
                        For k = 0 To UBound(tmpSXLMng)
                            If sGetSxlId = tmpSXLMng(k).SXLID Then
                                gtSprWfMap(iLoopCnt).KESSYOUP = tmpSXLMng(k).INGOTPOS
                                Exit For
                            End If
                        Next k
                    End If
                End If
            End If


            .GetText 2, iLoopCnt, vGetHinban
            If vGetHinban <> "Z" Then
                '2003/05/29 hitec)okazaki ブロックSEQを画面から取得に変更
                .GetText 6, iLoopCnt, vGetBlockSEQ_S
                .GetText 6, iLoopCnt + 1, vGetBlockSEQ_E
                iFromBlkSeq = CInt(vGetBlockSEQ_S)                                  'ブロックSEQを取得
                iToBlkSeq = CInt(vGetBlockSEQ_E)                                    'ブロックSEQを取得
                '2003/05/29 end

                sql = "UPDATE TBCMY011 SET"
                sql = sql & " mhinban = '" & gtSprWfMap(iLoopCnt).hinban & "'"      ' 品番"
'''''                nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                If iLoopCnt = 1 Then
                    NowIngotPos = SIngotP
                Else
                    NowIngotPos = gtSprWfMap(iLoopCnt).KESSYOUP
                End If

                sql = sql & ",MSXLID = '" & sGetSxlId & "'"
                sql = sql & ",UPDPROC= 'CW740'"                                     ' 更新工程
                sql = sql & ",UPDDATE=  sysdate"                                    ' 更新日時

                '製品情報はtblWafIndから取得する 2003/05/28 okazaki start
                For i = 1 To UBound(tblWafInd)
                    If gtSprWfMap(iLoopCnt).hinban = tblWafInd(i).HINDN.hinban Then
                        Exit For
                    End If
                Next i
                sql = sql & ",MREVNUM =  " & tblWafInd(i).HINDN.mnorevno            ' 製品番号改訂番号
                sql = sql & ",MFACTORY= '" & tblWafInd(i).HINDN.factory & "'"       ' 工場
                sql = sql & ",MOPECOND= '" & tblWafInd(i).HINDN.opecond & "'"       ' 操業条件

                '2003/05/28 end
                sql = sql & " WHERE LOTID ='" & gtSprWfMap(iLoopCnt).LOTID & "'"    ' ブロックID"
                If (iFromBlkSeq <= iToBlkSeq) Then
                    sql = sql & " AND ((BLOCKSEQ >= " & iFromBlkSeq & ")"           ' ブロック内連番"
                    sql = sql & " AND  (BLOCKSEQ <= " & iToBlkSeq & "  ))"
                Else
                    sql = sql & " AND  (BLOCKSEQ >= " & iFromBlkSeq & ")"           ' ブロック内連番"
                End If

                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then
                    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If


                '範囲内の結果 (WFSTA=４)以外のサンプル情報をすべてクリア
                sql = "UPDATE TBCMY011 SET"
                sql = sql & " SHAFLAG = '0'"                                        ' サンプルフラグ"
                sql = sql & ",WFSTA   = '0'"                                        ' WF状態
                sql = sql & ",MSMPLEID= NULL"                                       ' 抜試位置"
                sql = sql & ",UPDDATE = sysdate"                                    ' 更新日時

                sql = sql & " WHERE LOTID ='" & gtSprWfMap(iLoopCnt).LOTID & "'"    ' ブロックID"
                If (iFromBlkSeq <= iToBlkSeq) Then
                    sql = sql & " AND ((BLOCKSEQ >= " & iFromBlkSeq & ")"           ' ブロック内連番"
                    sql = sql & " AND  (BLOCKSEQ <= " & iToBlkSeq & "  ))"
                Else
                    sql = sql & " AND  (BLOCKSEQ >= " & iFromBlkSeq & ")"           ' ブロック内連番"
                End If
                sql = sql & " AND WFSTA <> '4'"

                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then
                        '更新該当が0の場合も続行
                End If
            End If
        Next iLoopCnt


        'サンプルの更新（画面から「共有」を判定する）
        For iLoopCnt = 1 To UBound(gtSprWfMap())
            .GetText 10, iLoopCnt, vGetSample
            If (vGetSample <> vbNullString) Then

                sql = "UPDATE TBCMY011 SET"
                If vGetSample = gsWF_SMPL_JOINT Then
'                    .GetText 30, iLoopCnt, vGetSample
                    ''残存酸素検査項目追加による変更　03/12/09 ooba
'                    .GetText 31, iLoopCnt, vGetSample
                    'GD追加による変更　05/01/31 ooba
'                    .GetText 32, iLoopCnt, vGetSample
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                    .GetText 38, iLoopCnt, vGetSample
                    Call Cnv_GetSample(vGetSample)

                    sql = sql & " MSMPLEID= '" & vGetSample & "'"                   ' 抜試位置"
                    sql = sql & ",SHAFLAG = '1'"                                    ' サンプルフラグ"
                    sql = sql & ",WFSTA   = '1'"                                    ' WF状態サンプル  'del 2003/05/03 hitec)matsumoto
                Else
'                    .GetText 30, iLoopCnt, vGetSample                               ' 03/05/28
                    ''残存酸素検査項目追加による変更　03/12/09 ooba
'                    .GetText 31, iLoopCnt, vGetSample
                    'GD追加による変更　05/01/31 ooba
'                    .GetText 32, iLoopCnt, vGetSample
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                    .GetText 38, iLoopCnt, vGetSample
                    Call Cnv_GetSample(vGetSample)

                    sql = sql & " MSMPLEID= '" & vGetSample & "'"                   ' 抜試位置"  03/05/28
                    sql = sql & ",SHAFLAG = '1'"                                    ' サンプルフラグ"
                    sql = sql & ",WFSTA   = '0'"                                    ' WF状態サンプル  'del 2003/05/03 hitec)matsumoto
                End If

'''''                nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")

                sql = sql & ",UPDPROC = 'CW740'"                                    ' 更新工程
                sql = sql & ",UPDDATE = sysdate"                                    ' 更新日時
                sql = sql & " WHERE LOTID   = '" & gtSprWfMap(iLoopCnt).LOTID & "'" ' ブロックID"
                sql = sql & "   AND BLOCKSEQ=  " & gtSprWfMap(iLoopCnt).BLOCKSEQ    ' ブロック内連番"
                sql = sql & "   AND WFSTA   <>'4'"
                '' WriteDBLog sql
                Debug.Print sql
                If 0 >= OraDB.ExecuteSQL(sql) Then

                End If

            End If
        Next
    End With
    '品番を1列追加したことによる列の変更-------end iida 2003/09/03
    DBDRV_UPD_WFMap = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_UPD_WFMap = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function


Public Sub Cnv_GetSample(ByRef vGetSample As Variant)
    Dim i   As Integer
    Dim kbn As String

    For i = 1 To UBound(CngSmpID_UD)
        If CngSmpID_UD(i) = vGetSample Then
           kbn = Cnv_Smp_KB(Right(vGetSample, 1))
           vGetSample = Left(vGetSample, Len(vGetSample) - 1) + kbn
           Exit Sub
        End If
    Next
End Sub

Public Function Cnv_Smp_KB(SmpKb As String) As String
    If SmpKb = "U" Then
        Cnv_Smp_KB = "B"
        Exit Function
    End If

    If SmpKb = "D" Then
        Cnv_Smp_KB = "T"
        Exit Function
    End If
End Function


'概要    :該当ブロックのWF１枚の長さ（計算長）を取得
'ﾊﾟﾗﾒｰﾀ  :変数名       ,IO   ,型                    ,説明
'        :BLOCKID       ,I   ,STRING                ,ブロックＩＤ
'        :dblWFLen      ,O   ,DOUBLE        　　    ,WF1枚の計算長さ
'        :戻ﾘ値         ,O   ,FUNCTION_RETURN       ,読み込み成否
'説明    :該当ブロックのWF１枚の長さ（計算長）を取得
'履歴    :2003/4/25 Hitec)okazaki
Public Function DBDRV_WFLENGET(ByVal StrBlockId As String, ByRef dblWFLen As Double) As FUNCTION_RETURN

    Dim strSQL      As String
    Dim iRealLen    As Integer
    Dim iWFcnt      As Integer
    Dim rs          As OraDynaset
    Dim iKetuFrom   As Integer
    Dim iKetuTo     As Integer
    Dim iKetuLen    As Integer
    Dim sDbName     As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_WFLENGET"

    '実長さ、WF枚数取得
    sDbName = "(Y011)"

    strSQL = "select e40.blockid,e40.reallen,y11.cnt"
    strSQL = strSQL & " from tbcme040 e40,"
    strSQL = strSQL & " xsdca xa,"
    strSQL = strSQL & " (select lotid,count(lotid) cnt"
    strSQL = strSQL & "  from   tbcmy011"
    strSQL = strSQL & "  where  lotid ='" & StrBlockId & "'"
    strSQL = strSQL & "  group by lotid  ) y11"
    strSQL = strSQL & " where e40.blockid =  xa.CRYNUMCA"
    strSQL = strSQL & "   and y11.lotid   =  xa.CRYNUMCA"
    strSQL = strSQL & "   and y11.lotid   = '" & StrBlockId & "'"

    Set rs = OraDB.DBCreateDynaset(strSQL, ORADYN_NO_BLANKSTRIP)
    If Not rs.EOF Then
           iRealLen = CInt(rs!REALLEN)
           iWFcnt = CInt(rs!cnt)
    Else
        rs.Close
        DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    rs.Close

    '欠落長さ取得
    sDbName = "(Y012)"
    strSQL = "SELECT DISTINCT LENFROM,LENTO FROM TBCMY012"
    strSQL = strSQL & " Where "
    strSQL = strSQL & " LOTID   = '" & StrBlockId & "'"

    Set rs = OraDB.DBCreateDynaset(strSQL, ORADYN_NO_BLANKSTRIP)
    iKetuLen = 0
    Do While Not rs.EOF
        If (IsNull(rs.Fields("LENFROM")) = True) Or rs.Fields("LENFROM") = -1 Or _
            (IsNull(rs.Fields("LENTO")) = True) Or rs.Fields("LENTO") = -1 Then
        Else
            iKetuFrom = CInt(rs.Fields("LENFROM"))
            iKetuTo = CInt(rs.Fields("LENTO"))
            iKetuLen = iKetuLen + iKetuTo - iKetuFrom
        End If
        rs.MoveNext
    Loop
    rs.Close

    'WF長さ計算
    dblWFLen = (iRealLen - iKetuLen) / iWFcnt

    '小数点２桁目を四捨五入
''''    dblWFLen = Int((dblWFLen + 0.05) * 10) / 10 'del 2003/08/06 hitec)matusmoto

    DBDRV_WFLENGET = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print strSQL
    DBDRV_WFLENGET = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit


End Function

'概要      :SXLＩＤの取得&判定
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
'　　      :iIngotPos     ,I  ,Integer　,結晶位置
'          :sSXLID        ,O  ,STRING   ,SXLID
'　　      :戻り値　　　　,O  ,　       ,選択の有無
'説明      :画面より、A欠落を除くSXL先頭の品番に対し元のSXLIDを取得
'履歴      :2003/05/01   hitec)okazaki

Public Function Get_OLDSXLID(iIngotpos As Integer, sSXLID As String, sOldIngotP As String) As FUNCTION_RETURN

    Dim i           As Integer
    Dim iRowIngotP  As Integer
    Dim vGetIngotP  As Variant
    Dim vGetHinban  As Variant
    Dim vGetSXLID1  As Variant
    Dim vGetSXLID2  As Variant
    Dim vWFcnt      As Variant
    Dim iSumWFcnt   As Integer
    Dim j           As Integer
    Dim vGetSampl   As Variant
    Dim vGetSampl2  As Variant

    Dim idx2        As Integer

    Get_OLDSXLID = FUNCTION_RETURN_FAILURE

    '品番を1列追加したことに列の変更-------start iida 2003/09/03
    With f_cmbc036_2.sprExamine

        For i = 1 To .MaxRows - 1 Step 2
            If i > .MaxRows Then
                Exit Function
            End If
            .GetText 5, i, vGetIngotP
            If iIngotpos = CInt(vGetIngotP) Then
                iRowIngotP = i      '画面の該当位置を取得
                Exit For
            End If
        Next i

        idx2 = Get_YukouRow(iRowIngotP, "D")
        iRowIngotP = idx2

'        .GetText 36, iRowIngotP, vGetSXLID1                 '画面の該当位置の元SXLIDを取得
        ''残存酸素検査項目追加による変更　03/12/09 ooba
'        .GetText 37, iRowIngotP, vGetSXLID1
        'GD追加による変更　05/01/31 ooba
'        .GetText 38, iRowIngotP, vGetSXLID1
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
        .GetText 44, iRowIngotP, vGetSXLID1
        iSumWFcnt = 0
        For i = 2 To .MaxRows Step 2
            If iRowIngotP - i < 1 Then                      'その品番が元SXLの先頭の場合（A欠落除く）
                .GetText 5, 1, vGetIngotP
                sOldIngotP = CStr(vGetIngotP)
                sSXLID = CStr(vGetSXLID1)

                Get_OLDSXLID = FUNCTION_RETURN_SUCCESS
                Exit For
            End If
            .GetText 2, iRowIngotP - i, vGetHinban
            If vGetHinban <> "Z" Then                       'A欠落でない最初の前の品番
'                .GetText 36, iRowIngotP - i, vGetSXLID2
                ''残存酸素検査項目追加による変更　03/12/09 ooba
'                .GetText 37, iRowIngotP - i, vGetSXLID2
                'GD追加による変更　05/01/31 ooba
'                .GetText 38, iRowIngotP - i, vGetSXLID2
'--- 2006/08/15 Cng エピ先行評価追加対応 SMP)kondoh
                .GetText 44, iRowIngotP - i, vGetSXLID2
                If vGetSXLID2 <> vGetSXLID1 Then            'その品番が元SXLの先頭の場合（A欠落除く）
                    .GetText 5, iRowIngotP - i + 2, vGetIngotP
                    sOldIngotP = CStr(vGetIngotP)
                    sSXLID = CStr(vGetSXLID1)
                    Get_OLDSXLID = FUNCTION_RETURN_SUCCESS
                End If

                Exit For
            End If
        Next i

    End With
    '品番を1列追加したことに列の変更-------end iida 2003/09/03

End Function

'###############################################2003/05/19 okazaki

'概要      :A欠落を飛ばした有効な行を取得する(ブロックは考えない）
'ﾊﾟﾗﾒｰﾀ　　:変数名　　　　,IO ,型       ,説明
'　　      :iNowRow      ,I  ,Integer　,結晶位置
'          :sUD          ,I  ,STRING   ,方向（上か下か）"U":上　"D":下
'　　      :戻り値　　　　,O  ,INTEGER　,有効行
'説明      :画面より、A欠落を除く品番のある行番号(Spread)を取得
'履歴      :2003/05/19   hitec)okazaki

Public Function Get_YukouRow(iNowRow As Integer, ByRef sUD As String) As Integer

    Dim vGetHinban  As Variant
    Dim iCount      As Integer

    On Error Resume Next


    Get_YukouRow = iNowRow
    With f_cmbc036_2.sprExamine
        'パラメータチェック
        If iNowRow < 1 Or iNowRow > .MaxRows Then
            Exit Function
        End If

        If sUD <> "U" And sUD <> "D" Then
            Exit Function
        End If


        '上に検索
        If sUD = "U" Then
            For iCount = iNowRow To 1 Step -1
                If iCount Mod 2 = 1 Then
                    .GetText 2, iCount, vGetHinban
                    If vGetHinban <> "Z" Then
                        Get_YukouRow = iCount
                        Exit For
                    End If
                End If
            Next iCount

        '下に検索
        ElseIf sUD = "D" Then
            For iCount = iNowRow To .MaxRows - 1 Step 1
                If iCount Mod 2 = 1 Then
                    .GetText 2, iCount, vGetHinban
                    If vGetHinban <> "Z" Then
                        Get_YukouRow = iCount
                        Exit For
                    End If
                End If
            Next iCount
        End If
    End With

End Function
'###############################################2003/05/19 okazaki
'---------------------------------------------------------------
'
' 機能　　  : SXLチェックボックス詳細の表示
'
' 返り値　  : なし
'
' 引数　    : iIndex　１：表示 / ０：非表示
'
'
' 機能説明  : SXLチェックボックス表示の時、詳細を表示する
'
' 備考　　  : 03/05/31  後藤
'
'---------------------------------------------------------------
Public Sub Pic_Disp(iIndex As Integer)
    Dim iCnt    As Integer

    With f_cmbc036_2
        If iIndex = 0 Then
            For iCnt = 0 To 2
                .lbl_check(iCnt).Visible = False
            Next
            .pic_check(0).Visible = False
            .pic_check(1).Visible = False

        ElseIf iIndex = 1 Then
            For iCnt = 0 To 2
                .lbl_check(iCnt).Visible = True
            Next
            .pic_check(0).Visible = True
            .pic_check(1).Visible = True
        End If
    End With
End Sub


'概要      :WFサンプル管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型                 ,説明
'      　　:WFSMP 　　　,I  ,typ_XSDCW   　     ,新サンプル管理（SXL）
'      　　:戻り値      ,O  ,FUNCTION_RETURN　  ,書き込みの成否
'説明      :DBDRV_WfSmp_UpdInsに移行する予定
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_WfSmp_INS(WFSMP() As typ_XSDCW, i As Long) As FUNCTION_RETURN

    Dim sql As String
'    Dim i As Long '2003/09/22コメントにした
    Dim sDbName As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_INS"

    DBDRV_WfSmp_INS = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
'    For i = 1 To UBound(WFSMP)　'2003/09/22 コメントにした
        With WFSMP(i)

                sql = "insert into XSDCW ("
                sql = sql & "SXLIDCW, "         ' SXLID
                sql = sql & "SMPKBNCW, "        ' サンプル区分
                sql = sql & "TBKBNCW, "         ' T/B区分
                sql = sql & "REPSMPLIDCW, "     ' サンプルID
                sql = sql & "XTALCW, "          ' 結晶番号
                sql = sql & "INPOSCW, "         ' 結晶内位置
                sql = sql & "HINBCW, "          ' 品番
                sql = sql & "REVNUMCW, "        ' 製品番号改訂番号
                sql = sql & "FACTORYCW, "       ' 工場
                sql = sql & "OPECW, "           ' 操業条件
                sql = sql & "KTKBNCW, "         ' 確定区分
                sql = sql & "SMCRYNUMCW, "      ' サンプルブロックID
                sql = sql & "WFSMPLIDRSCW, "    ' サンプルID(Rs)
                sql = sql & "WFSMPLIDRS1CW, "   ' 推定サンプルID1（Rs）
                sql = sql & "WFSMPLIDRS2CW, "   ' 推定サンプルID2（Rs）
                sql = sql & "WFINDRSCW, "       ' 状態FLG（Rs)
                sql = sql & "WFRESRS1CW, "      ' 実績FLG1（Rs)
                sql = sql & "WFRESRS2CW, "      ' 実績FLG2（Rs)
                sql = sql & "WFSMPLIDOICW, "    ' サンプルID（Oi）
                sql = sql & "WFINDOICW, "       ' 状態FLG（Oi)
                sql = sql & "WFRESOICW, "       ' 実績FLG（Oi)
                sql = sql & "WFSMPLIDB1CW, "    ' サンプルID（B1）
                sql = sql & "WFINDB1CW, "       ' 状態FLG（B1)
                sql = sql & "WFRESB1CW, "       ' 実績FLG（B1)
                sql = sql & "WFSMPLIDB2CW, "    ' サンプルID（B2）
                sql = sql & "WFINDB2CW, "       ' 状態FLG（B2)
                sql = sql & "WFRESB2CW, "       ' 実績FLG（B2)
                sql = sql & "WFSMPLIDB3CW, "    ' サンプルID（B3）
                sql = sql & "WFINDB3CW, "       ' 状態FLG（B3)
                sql = sql & "WFRESB3CW, "       ' 実績FLG（B3)
                sql = sql & "WFSMPLIDL1CW, "    ' サンプルID（L1）
                sql = sql & "WFINDL1CW, "       ' 状態FLG（L1)
                sql = sql & "WFRESL1CW, "       ' 実績FLG（L1)
                sql = sql & "WFSMPLIDL2CW, "    ' サンプルID（L2）
                sql = sql & "WFINDL2CW, "       ' 状態FLG（L2)
                sql = sql & "WFRESL2CW, "       ' 実績FLG（L2)
                sql = sql & "WFSMPLIDL3CW, "    ' サンプルID（L3）
                sql = sql & "WFINDL3CW, "       ' 状態FLG（L3)
                sql = sql & "WFRESL3CW, "       ' 実績FLG（L3)
                sql = sql & "WFSMPLIDL4CW, "    ' サンプルID（L4）
                sql = sql & "WFINDL4CW, "       ' 状態FLG（L4)
                sql = sql & "WFRESL4CW, "       ' 実績FLG（L4)
                sql = sql & "WFSMPLIDDSCW, "    ' サンプルID（DS）
                sql = sql & "WFINDDSCW, "       ' 状態FLG（DS)
                sql = sql & "WFRESDSCW, "       ' 実績FLG（DS)
                sql = sql & "WFSMPLIDDZCW, "    ' サンプルID（DZ）
                sql = sql & "WFINDDZCW, "       ' 状態FLG（DZ)
                sql = sql & "WFRESDZCW, "       ' 実績FLG（DZ)
                sql = sql & "WFSMPLIDSPCW, "    ' サンプルID（SP）
                sql = sql & "WFINDSPCW, "       ' 状態FLG（SP)
                sql = sql & "WFRESSPCW, "       ' 実績FLG（SP)
                sql = sql & "WFSMPLIDDO1CW,"    ' サンプルID（DO1）
                sql = sql & "WFINDDO1CW, "      ' 状態FLG（DO1)
                sql = sql & "WFRESDO1CW, "      ' 実績FLG（DO1)
                sql = sql & "WFSMPLIDDO2CW, "   ' サンプルID（DO2）
                sql = sql & "WFINDDO2CW, "      ' 状態FLG（DO2)
                sql = sql & "WFRESDO2CW, "      ' 実績FLG（DO2)
                sql = sql & "WFSMPLIDDO3CW, "   ' サンプルID（DO3）
                sql = sql & "WFINDDO3CW, "      ' 状態FLG（DO3)
                sql = sql & "WFRESDO3CW, "      ' 実績FLG（DO3)
                sql = sql & "WFSMPLIDOT1CW, "   ' サンプルID（OT1）
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW, "      ' 状態FLG（OT1)
                sql = sql & "WFRESOT1CW, "      ' 実績FLG（OT1)
                sql = sql & "WFSMPLIDOT2CW, "   ' サンプルID（OT2）
                sql = sql & "WFINDOT2CW, "      ' 状態FLG（OT2)
                sql = sql & "WFRESOT2CW, "      ' 実績FLG（OT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFSMPLIDAOICW, "   ' サンプルID（AOi）
                sql = sql & "WFINDAOICW, "      ' 状態FLG（AOi）
                sql = sql & "WFRESAOICW, "      ' 実績FLG（AOi）
                '' GD追加　05/01/31 ooba START =====================================>
                sql = sql & "WFSMPLIDGDCW, "    ' サンプルID (GD)
                sql = sql & "WFINDGDCW, "       ' 状態FLG (GD)
                sql = sql & "WFRESGDCW, "       ' 実績FLG (GD)
                sql = sql & "WFHSGDCW, "        ' 保証FLG (GD)
                '' GD追加　05/01/31 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & "EPSMPLIDB1CW, "    ' サンプルID (BMD1E)
                sql = sql & "EPINDB1CW, "       ' 状態FLG (BMD1E)
                sql = sql & "EPRESB1CW, "       ' 実績FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW, "    ' サンプルID (BMD2E)
                sql = sql & "EPINDB2CW, "       ' 状態FLG (BMD2E)
                sql = sql & "EPRESB2CW, "       ' 実績FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW, "    ' サンプルID (BMD3E)
                sql = sql & "EPINDB3CW, "       ' 状態FLG (BMD3E)
                sql = sql & "EPRESB3CW, "       ' 実績FLG (BMD3E)
                sql = sql & "EPSMPLIDL1CW, "    ' サンプルID (OSF1E)
                sql = sql & "EPINDL1CW, "       ' 状態FLG (OSF1E)
                sql = sql & "EPRESL1CW, "       ' 実績FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW, "    ' サンプルID (OSF2E)
                sql = sql & "EPINDL2CW, "       ' 状態FLG (OSF2E)
                sql = sql & "EPRESL2CW, "       ' 実績FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW, "    ' サンプルID (OSF3E)
                sql = sql & "EPINDL3CW, "       ' 状態FLG (OSF3E)
                sql = sql & "EPRESL3CW, "       ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                sql = sql & "SMPLNUMCW, "       ' サンプル枚数
                sql = sql & "SMPLPATCW, "       ' サンプルパターン
                sql = sql & "NUKISIFLGCW, "     ' 抜試指示通過フラグ 09/05/26 ooba
                sql = sql & "TSTAFFCW,"         ' 登録社員ID
                sql = sql & "TDAYCW, "          ' 登録日付
                sql = sql & "KSTAFFCW, "        ' 更新社員ID
                sql = sql & "KDAYCW, "          ' 更新日付
                sql = sql & "SNDKCW, "          ' 送信フラグ
                sql = sql & "SNDDAYCW, "        ' 送信日付
                sql = sql & "LIVKCW)"           ' 生死区分

                sql = sql & " values ('"
                sql = sql & .SXLIDCW & "', '"           ' SXLID
                sql = sql & .SMPKBNCW & "', '"          ' サンプル区分
                sql = sql & .TBKBNCW & "', '"           ' T/B区分
                sql = sql & .REPSMPLIDCW & "', '"       ' サンプルID
                sql = sql & .XTALCW & "', "             ' 結晶番号
                sql = sql & .INPOSCW & ", '"            ' 結晶内位置
                sql = sql & .HINBCW & "', "             ' 品番
                sql = sql & .REVNUMCW & ", '"           ' 製品番号改訂番号
                sql = sql & .FACTORYCW & "', '"         ' 工場
                sql = sql & .OPECW & "', '"             ' 操業条件
                sql = sql & .KTKBNCW & "', '"           ' 確定区分
                sql = sql & .SMCRYNUMCW & "', '"        ' サンプルブロックID
                sql = sql & .WFSMPLIDRSCW & "', "       ' サンプルID（Rs）
'               sql = sql & .WFSMPLIDRS1CW & "',"       ' 推定サンプルID1（Rs）
'               sql = sql & .WFSMPLIDRS2CW & "', "      ' 推定サンプルID2（Rs）
                sql = sql & "Null, "                    ' 推定サンプルID1（Rs）
                sql = sql & "Null, '"                   ' 推定サンプルID2（Rs）
''              sql = sql & .WFINDRSCW & "', "          ' 状態FLG（Rs)
''              sql = sql & "Null, "                    ' 実績FLG1（Rs)
                sql = sql & .WFINDRSCW & "', '"         ' 状態FLG（Rs)
                sql = sql & .WFRESRS1CW & "', "         ' 実績FLG1（Rs)
                sql = sql & "Null, '"                   ' 実績FLG2（Rs)
                sql = sql & .WFSMPLIDOICW & "', '"      ' サンプルID（Oi）
                sql = sql & .WFINDOICW & "', '"         ' 状態FLG（Oi)
                sql = sql & .WFRESOICW & "', '"         ' 実績FLG（Oi)
                sql = sql & .WFSMPLIDB1CW & "', '"      ' サンプルID（B1）
                sql = sql & .WFINDB1CW & "', '"         ' 状態FLG（B1)
                sql = sql & .WFRESB1CW & "', '"         ' 実績FLG（B1)
                sql = sql & .WFSMPLIDB2CW & "', '"      ' サンプルID（B2）
                sql = sql & .WFINDB2CW & "', '"         ' 状態FLG（B2)
                sql = sql & .WFRESB2CW & "', '"         ' 実績FLG（B2)
                sql = sql & .WFSMPLIDB3CW & "', '"      ' サンプルID（B3）
                sql = sql & .WFINDB3CW & "', '"         ' 状態FLG（B3)
                sql = sql & .WFRESB3CW & "', '"         ' 実績FLG（B3)
                sql = sql & .WFSMPLIDL1CW & "', '"      ' サンプルID（L1）
                sql = sql & .WFINDL1CW & "', '"         ' 状態FLG（L1)
                sql = sql & .WFRESL1CW & "', '"         ' 実績FLG（L1)
                sql = sql & .WFSMPLIDL2CW & "', '"      ' サンプルID（L2）
                sql = sql & .WFINDL2CW & "', '"         ' 状態FLG（L2)
                sql = sql & .WFRESL2CW & "', '"         ' 実績FLG（L2)
                sql = sql & .WFSMPLIDL3CW & "', '"      ' サンプルID（L3）
                sql = sql & .WFINDL3CW & "', '"         ' 状態FLG（L3)
                sql = sql & .WFRESL3CW & "', '"         ' 実績FLG（L3)
                sql = sql & .WFSMPLIDL4CW & "', '"      ' サンプルID（L4）
                sql = sql & .WFINDL4CW & "', '"         ' 状態FLG（L4)
                sql = sql & .WFRESL4CW & "', '"         ' 実績FLG（L4)
                sql = sql & .WFSMPLIDDSCW & "', '"      ' サンプルID（DS）
                sql = sql & .WFINDDSCW & "', '"         ' 状態FLG（DS)
                sql = sql & .WFRESDSCW & "', '"         ' 実績FLG（DS)
                sql = sql & .WFSMPLIDDZCW & "', '"      ' サンプルID（DZ）
                sql = sql & .WFINDDZCW & "', '"         ' 状態FLG（DZ)
                sql = sql & .WFRESDZCW & "', '"         ' 実績FLG（DZ)
                sql = sql & .WFSMPLIDSPCW & "', '"      ' サンプルID（SP）
                sql = sql & .WFINDSPCW & "', '"         ' 状態FLG（SP)
                sql = sql & .WFRESSPCW & "', '"         ' 実績FLG（SP)
                sql = sql & .WFSMPLIDDO1CW & "', '"     ' サンプルID（DO1）
                sql = sql & .WFINDDO1CW & "', '"        ' 状態FLG（DO1)
                sql = sql & .WFRESDO1CW & "', '"        ' 実績FLG（DO1)
                sql = sql & .WFSMPLIDDO2CW & "', '"     ' サンプルID（DO2）
                sql = sql & .WFINDDO2CW & "', '"        ' 状態FLG（DO2)
                sql = sql & .WFRESDO2CW & "', '"        ' 実績FLG（DO2)
                sql = sql & .WFSMPLIDDO3CW & "', '"     ' サンプルID（DO3）
                sql = sql & .WFINDDO3CW & "', '"        ' 状態FLG（DO3)
                sql = sql & .WFRESDO3CW & "', '"        ' 実績FLG（DO3)
                sql = sql & .WFSMPLIDOT1CW & "', '"     ' サンプルID（OT1）
                sql = sql & .WFINDOT1CW & "', '"        ' 状態FLG（OT1)
                sql = sql & .WFRESOT1CW & "', '"        ' 実績FLG（OT1)
                sql = sql & .WFSMPLIDOT2CW & "', '"     ' サンプルID（OT2）
                sql = sql & .WFINDOT2CW & "', '"        ' 状態FLG（OT2)
                sql = sql & .WFRESOT2CW & "', '"        ' 実績FLG（OT2)
'                sql = sql & "NULL, "                    ' サンプルID（AOi）
'                sql = sql & "NULL, "                    ' 状態FLG（AOi）
'                sql = sql & "NULL, "                    ' 実績FLG（AOi）
                ''ｺﾒﾝﾄ解除－残存酸素データ登録追加　03/12/11 ooba
                sql = sql & .WFSMPLIDAOICW & "', '"     ' サンプルID（AOi）
                sql = sql & .WFINDAOICW & "', '"        ' 状態FLG（AOi）
                sql = sql & .WFRESAOICW & "', '"        ' 実績FLG（AOi）
                '' GD追加　05/01/31 ooba START =====================================>
                sql = sql & .WFSMPLIDGDCW & "', '"      ' サンプルID (GD)
                sql = sql & .WFINDGDCW & "', '"         ' 状態FLG (GD)
                sql = sql & .WFRESGDCW & "', '"         ' 実績FLG (GD)
                sql = sql & .WFHSGDCW & "', "           ' 保証FLG (GD)
                '' GD追加　05/01/31 ooba END =======================================>
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & "'" & .EPSMPLIDB1CW & "', '"      ' サンプルID (BMD1E)
                sql = sql & .EPINDB1CW & "', '"         ' 状態FLG (BMD1E)
                sql = sql & .EPRESB1CW & "', '"         ' 実績FLG (BMD1E)
                sql = sql & .EPSMPLIDB2CW & "', '"      ' サンプルID (BMD2E)
                sql = sql & .EPINDB2CW & "', '"         ' 状態FLG (BMD2E)
                sql = sql & .EPRESB2CW & "', '"         ' 実績FLG (BMD2E)
                sql = sql & .EPSMPLIDB3CW & "', '"      ' サンプルID (BMD3E)
                sql = sql & .EPINDB3CW & "', '"         ' 状態FLG (BMD3E)
                sql = sql & .EPRESB3CW & "', '"         ' 実績FLG (BMD3E)
                sql = sql & .EPSMPLIDL1CW & "', '"      ' サンプルID (OSF1E)
                sql = sql & .EPINDL1CW & "', '"         ' 状態FLG (OSF1E)
                sql = sql & .EPRESL1CW & "', '"         ' 実績FLG (OSF1E)
                sql = sql & .EPSMPLIDL2CW & "', '"      ' サンプルID (OSF2E)
                sql = sql & .EPINDL2CW & "', '"         ' 状態FLG (OSF2E)
                sql = sql & .EPRESL2CW & "', '"         ' 実績FLG (OSF2E)
                sql = sql & .EPSMPLIDL3CW & "', '"      ' サンプルID (OSF3E)
                sql = sql & .EPINDL3CW & "', '"         ' 状態FLG (OSF3E)
                sql = sql & .EPRESL3CW & "',  "         ' 実績FLG (OSF3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
'''''           sql = sql & "NULL,"                     ' サンプル枚数
'''''           sql = sql & .SMPLPATCW & "', '"         ' サンプルパターン
                sql = sql & "NULL, "                    ' サンプル枚数
                sql = sql & "NULL, "                    ' サンプルパターン
                sql = sql & "'1', '"                    ' 抜試指示通過フラグ 09/05/26 ooba
                sql = sql & .TSTAFFCW & "', "           ' 登録社員ID
                sql = sql & "sysdate, '"                ' 登録日付
                sql = sql & .KSTAFFCW & "', "           ' 更新社員ID
                sql = sql & "sysdate, "                 ' 更新日付
                sql = sql & "'0', "                     ' 送信フラグ
                sql = sql & "sysdate, "                 ' 送信日付
                sql = sql & "'0')"                      ' 生死区分

                '' WriteDBLog sql, sDbName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_INS = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :WFサンプル管理の更新
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型                 ,説明
'      　　:WFSMP 　　　,I  ,typ_XSDCW   　     ,新サンプル管理（SXL）
'      　　:戻り値      ,O  ,FUNCTION_RETURN　  ,書き込みの成否
'説明      :新サンプル管理のデータを更新する
'履歴      :2003/09/22  作成 飯田
Public Function DBDRV_WfSmp_UPD(WFSMP() As typ_XSDCW, i As Long) As FUNCTION_RETURN

    Dim sql As String
'    Dim i As Long
    Dim sDbName As String


    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_WfSmp_UPD"

    DBDRV_WfSmp_UPD = FUNCTION_RETURN_SUCCESS

    sDbName = "XSDCW"
'    For i = 1 To UBound(WFSMP)
       With WFSMP(i)
                sql = "UPDATE XSDCW "
                sql = sql & "SET "
'               sql = sql & "SXLIDCW      ='" & .SXLIDCW & "',"         ' SXLID"
                sql = sql & "SMPKBNCW     ='" & .SMPKBNCW & "',"        ' サンプル区分"
'               sql = sql & "TBKBNCW      ='" & .TBKBNCW & "',"         ' T/B区分"
                sql = sql & "REPSMPLIDCW  ='" & .REPSMPLIDCW & "',"     ' サンプルID"
                sql = sql & "XTALCW       ='" & .XTALCW & "',"          ' 結晶番号"
                sql = sql & "INPOSCW      ='" & .INPOSCW & "',"         ' 結晶内位置"
                sql = sql & "HINBCW       ='" & .HINBCW & "',"          ' 品番
                sql = sql & "REVNUMCW     ='" & .REVNUMCW & "',"        ' 製品番号改訂番号
                sql = sql & "FACTORYCW    ='" & .FACTORYCW & "',"       ' 工場
                sql = sql & "OPECW        ='" & .OPECW & "',"           ' 操業条件
                sql = sql & "KTKBNCW      ='" & .KTKBNCW & "',"         ' 確定区分
                sql = sql & "SMCRYNUMCW   ='" & .SMCRYNUMCW & "',"      ' サンプルブロックID
                sql = sql & "WFSMPLIDRSCW ='" & .WFSMPLIDRSCW & "',"    ' サンプルID(Rs)
                sql = sql & "WFSMPLIDRS1CW= NULL,"                      ' 推定サンプルID1（Rs）
                sql = sql & "WFSMPLIDRS2CW= NULL,"                      ' 推定サンプルID2（Rs）
                sql = sql & "WFINDRSCW    ='" & .WFINDRSCW & "',"       ' 状態FLG（Rs)
'''''           sql = sql & "WFRESRS1CW   = NULL,"                      ' 実績FLG1（Rs)
'''''           sql = sql & "WFRESRS2CW   = NULL,"                      ' 実績FLG2（Rs)
                sql = sql & "WFRESRS1CW   ='" & .WFRESRS1CW & "',"      ' 実績FLG1（Rs)
                sql = sql & "WFSMPLIDOICW ='" & .WFSMPLIDOICW & "',"    ' サンプルID（Oi）
                sql = sql & "WFINDOICW    ='" & .WFINDOICW & "',"       ' 状態FLG（Oi)
                sql = sql & "WFRESOICW    ='" & .WFRESOICW & "',"       ' 実績FLG（Oi)
                sql = sql & "WFSMPLIDB1CW ='" & .WFSMPLIDB1CW & "',"    ' サンプルID（B1）
                sql = sql & "WFINDB1CW    ='" & .WFINDB1CW & "',"       ' 状態FLG（B1)
                sql = sql & "WFRESB1CW    ='" & .WFRESB1CW & "',"       ' 実績FLG（B1)
                sql = sql & "WFSMPLIDB2CW ='" & .WFSMPLIDB2CW & "',"    ' サンプルID（B2）
                sql = sql & "WFINDB2CW    ='" & .WFINDB2CW & "',"       ' 状態FLG（B2)
                sql = sql & "WFRESB2CW    ='" & .WFRESB2CW & "',"       ' 実績FLG（B2)
                sql = sql & "WFSMPLIDB3CW ='" & .WFSMPLIDB3CW & "',"    ' サンプルID（B3）
                sql = sql & "WFINDB3CW    ='" & .WFINDB3CW & "',"       ' 状態FLG（B3)
                sql = sql & "WFRESB3CW    ='" & .WFRESB3CW & "',"       ' 実績FLG（B3)
                sql = sql & "WFSMPLIDL1CW ='" & .WFSMPLIDL1CW & "',"    ' サンプルID（L1）
                sql = sql & "WFINDL1CW    ='" & .WFINDL1CW & "',"       ' 状態FLG（L1)
                sql = sql & "WFRESL1CW    ='" & .WFRESL1CW & "',"       ' 実績FLG（L1)
                sql = sql & "WFSMPLIDL2CW ='" & .WFSMPLIDL2CW & "',"    ' サンプルID（L2）
                sql = sql & "WFINDL2CW    ='" & .WFINDL2CW & "',"       ' 状態FLG（L2)
                sql = sql & "WFRESL2CW    ='" & .WFRESL2CW & "',"       ' 実績FLG（L2)
                sql = sql & "WFSMPLIDL3CW ='" & .WFSMPLIDL3CW & "',"    ' サンプルID（L3）
                sql = sql & "WFINDL3CW    ='" & .WFINDL3CW & "',"       ' 状態FLG（L3)
                sql = sql & "WFRESL3CW    ='" & .WFRESL3CW & "',"       ' 実績FLG（L3)
                sql = sql & "WFSMPLIDL4CW ='" & .WFSMPLIDL4CW & "',"    ' サンプルID（L4）
                sql = sql & "WFINDL4CW    ='" & .WFINDL4CW & "',"       ' 状態FLG（L4)
                sql = sql & "WFRESL4CW    ='" & .WFRESL4CW & "',"       ' 実績FLG（L4)
                sql = sql & "WFSMPLIDDSCW ='" & .WFSMPLIDDSCW & "',"    ' サンプルID（DS）
                sql = sql & "WFINDDSCW    ='" & .WFINDDSCW & "',"       ' 状態FLG（DS)
                sql = sql & "WFRESDSCW    ='" & .WFRESDSCW & "',"       ' 実績FLG（DS)
                sql = sql & "WFSMPLIDDZCW ='" & .WFSMPLIDDZCW & "',"    ' サンプルID（DZ）
                sql = sql & "WFINDDZCW    ='" & .WFINDDZCW & "',"       ' 状態FLG（DZ)
                sql = sql & "WFRESDZCW    ='" & .WFRESDZCW & "',"       ' 実績FLG（DZ)
                sql = sql & "WFSMPLIDSPCW ='" & .WFSMPLIDSPCW & "',"    ' サンプルID（SP）
                sql = sql & "WFINDSPCW    ='" & .WFINDSPCW & "',"       ' 状態FLG（SP)
                sql = sql & "WFRESSPCW    ='" & .WFRESSPCW & "',"       ' 実績FLG（SP)
                sql = sql & "WFSMPLIDDO1CW='" & .WFSMPLIDDO1CW & "',"   ' サンプルID（DO1）
                sql = sql & "WFINDDO1CW   ='" & .WFINDDO1CW & "',"      ' 状態FLG（DO1)
                sql = sql & "WFRESDO1CW   ='" & .WFRESDO1CW & "',"      ' 実績FLG（DO1)
                sql = sql & "WFSMPLIDDO2CW='" & .WFSMPLIDDO2CW & "',"   ' サンプルID（DO2）
                sql = sql & "WFINDDO2CW   ='" & .WFINDDO2CW & "',"      ' 状態FLG（DO2)
                sql = sql & "WFRESDO2CW   ='" & .WFRESDO2CW & "',"      ' 実績FLG（DO2)
                sql = sql & "WFSMPLIDDO3CW='" & .WFSMPLIDDO3CW & "',"   ' サンプルID（DO3）
                sql = sql & "WFINDDO3CW   ='" & .WFINDDO3CW & "',"      ' 状態FLG（DO3)
                sql = sql & "WFRESDO3CW   ='" & .WFRESDO3CW & "',"      ' 実績FLG（DO3)
                sql = sql & "WFSMPLIDOT1CW='" & .WFSMPLIDOT1CW & "',"   ' サンプルID（OT1）
                sql = sql & "WFSMPLIDOT2CW='" & .WFSMPLIDOT2CW & "',"   ' サンプルID（OT2）
               'add start 2003/05/21 hitec)matsumoto -------------------------
                sql = sql & "WFINDOT1CW   ='" & .WFINDOT1CW & "',"     ' 状態FLG（OT1)
                sql = sql & "WFRESOT1CW   ='" & .WFRESOT1CW & "',"     ' 実績FLG（OT1)
                sql = sql & "WFINDOT2CW   ='" & .WFINDOT2CW & "',"     ' 状態FLG（OT2)
                sql = sql & "WFRESOT2CW   ='" & .WFRESOT2CW & "',"     ' 実績FLG（OT2)
               'add end   2003/05/21 hitec)matsumoto -------------------------
'                sql = sql & "WFSMPLIDAOICW= NULL,"                      ' サンプルID（AOi）
'                sql = sql & "WFINDAOICW   = NULL,"                      ' 状態FLG（AOi）
'                sql = sql & "WFRESAOICW   = NULL,"                      ' 実績FLG（AOi）

                ''残存酸素データ登録追加　03/12/11 ooba START ===============================>
                sql = sql & "WFSMPLIDAOICW='" & .WFSMPLIDAOICW & "',"   ' サンプルID（AOi）
                sql = sql & "WFINDAOICW   ='" & .WFINDAOICW & "',"      ' 状態FLG（AOi）
                sql = sql & "WFRESAOICW   ='" & .WFRESAOICW & "',"      ' 実績FLG（AOi）
                ''残存酸素データ登録追加　03/12/11 ooba END =================================>

                '' GD追加　05/01/31 ooba START =============================================>
                sql = sql & "WFSMPLIDGDCW ='" & .WFSMPLIDGDCW & "',"    ' サンプルID (GD)
                sql = sql & "WFINDGDCW    ='" & .WFINDGDCW & "', "      ' 状態FLG (GD)
                sql = sql & "WFRESGDCW    ='" & .WFRESGDCW & "', "      ' 実績FLG (GD)
                sql = sql & "WFHSGDCW     ='" & .WFHSGDCW & "', "       ' 保証FLG (GD)
                '' GD追加　05/01/31 ooba END ===============================================>

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                sql = sql & "EPSMPLIDL1CW = '" & .EPSMPLIDL1CW & "', "  ' サンプルID (OSF1E)
                sql = sql & "EPINDL1CW = '" & .EPINDL1CW & "', "        ' 状態FLG (OSF1E)
                sql = sql & "EPRESL1CW = '" & .EPRESL1CW & "', "        ' 実績FLG (OSF1E)
                sql = sql & "EPSMPLIDL2CW = '" & .EPSMPLIDL2CW & "', "  ' サンプルID (OSF2E)
                sql = sql & "EPINDL2CW = '" & .EPINDL2CW & "', "        ' 状態FLG (OSF2E)
                sql = sql & "EPRESL2CW = '" & .EPRESL2CW & "', "        ' 実績FLG (OSF2E)
                sql = sql & "EPSMPLIDL3CW = '" & .EPSMPLIDL3CW & "', "  ' サンプルID (OSF3E)
                sql = sql & "EPINDL3CW = '" & .EPINDL3CW & "', "        ' 状態FLG (OSF3E)
                sql = sql & "EPRESL3CW = '" & .EPRESL3CW & "', "        ' 実績FLG (OSF3E)
                sql = sql & "EPSMPLIDB1CW = '" & .EPSMPLIDB1CW & "', "  ' サンプルID (BMD1E)
                sql = sql & "EPINDB1CW = '" & .EPINDB1CW & "', "        ' 状態FLG (BMD1E)
                sql = sql & "EPRESB1CW = '" & .EPRESB1CW & "', "        ' 実績FLG (BMD1E)
                sql = sql & "EPSMPLIDB2CW = '" & .EPSMPLIDB2CW & "', "  ' サンプルID (BMD2E)
                sql = sql & "EPINDB2CW = '" & .EPINDB2CW & "', "        ' 状態FLG (BMD2E)
                sql = sql & "EPRESB2CW = '" & .EPRESB2CW & "', "        ' 実績FLG (BMD2E)
                sql = sql & "EPSMPLIDB3CW = '" & .EPSMPLIDB3CW & "', "  ' サンプルID (BMD3E)
                sql = sql & "EPINDB3CW = '" & .EPINDB3CW & "', "        ' 状態FLG (BMD3E)
                sql = sql & "EPRESB3CW = '" & .EPRESB3CW & "', "        ' 実績FLG (BMD3E)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                sql = sql & "SMPLNUMCW    = NULL,"                      ' サンプル枚数
                sql = sql & "SMPLPATCW    = NULL,"                      ' サンプルパターン
                sql = sql & "NUKISIFLGCW  = '1',"                       ' 抜試指示通過フラグ 09/05/26 ooba
'               sql = sql & "TSTAFFCW     ='" & .TSTAFFCW & "' "        ' 登録社員ID
'''''           sql = sql & "TDAYCW       = sysdate,"                   ' 登録日付
                sql = sql & "KSTAFFCW     ='" & .KSTAFFCW & "',"        ' 更新社員ID
                sql = sql & "KDAYCW       = sysdate, "                  ' 更新日付"
                sql = sql & "SNDKCW       ='0',"                        ' 送信フラグ"
                sql = sql & "SNDDAYCW     = sysdate "                   ' 送信日付"

                sql = sql & "WHERE "
                sql = sql & "SXLIDCW ='" & .SXLIDCW & "'and "           ' SXLID"
'               sql = sql & "SMPKBNCW='" & .SMPKBNCW & "'"              ' サンプル区分"
                sql = sql & "TBKBNCW ='" & .TBKBNCW & "'"               ' TB区分"

                '' WriteDBLog sql, sDbName
        End With
        If OraDB.ExecuteSQL(sql) <= 0 Then
            DBDRV_WfSmp_UPD = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_WfSmp_UPD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'ホールドロット検索処理     '04/06/29 ooba 作成
Public Function HoldLot_Get740(xtal As String, HOLDBCA As String, WFHOLDFLGCA As String) As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim blkcnt As Integer
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function HoldLot_Get740"

    HoldLot_Get740 = FUNCTION_RETURN_SUCCESS

    sql = "select CRYNUMCA, HOLDBCA, WFHOLDFLGCA "
    sql = sql & "from XSDCA "
    sql = sql & "where LIVKCA = '0' "
    sql = sql & "and CRYNUMCA in ( "
    sql = sql & "     select BLOCKID "
    sql = sql & "     from TBCMY001 "
    sql = sql & "     where SBLOCKID in ( "
    sql = sql & "             select SBLOCKID "
    sql = sql & "             from TBCMY001 "
    sql = sql & "             where BLOCKID = '" & xtal & "' "
    sql = sql & "             ) "
    sql = sql & ") "
    Set rs = OraDB.CreateDynaset(sql, ORADYN_DEFAULT)

    If rs.EOF = False Then
        For blkcnt = 1 To rs.RecordCount
            If rs("HOLDBCA") = "1" Then
                HOLDBCA = rs("HOLDBCA")
                Exit For
            Else
                HOLDBCA = " "
            End If
            If rs("WFHOLDFLGCA") = "1" Then
                WFHOLDFLGCA = rs("WFHOLDFLGCA")
                Exit For
            Else
                WFHOLDFLGCA = " "
            End If
            rs.MoveNext
        Next blkcnt
    Else
        HoldLot_Get740 = FUNCTION_RETURN_FAILURE
    End If
    rs.Close
    Set rs = Nothing

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    HoldLot_Get740 = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function


'＠(f)
'
'機能       :品番管理 - 表示用TBCME041データ取得
'
'返り値     :0 - 正常終了
'           :1 - 異常終了
'
'引き数     :records()  - 抽出レコード
'           :sCryNum    - 結晶番号
'           :sBlockId   - ブロックID
'
'機能説明   :表示用品番データを取得する
'
'履歴       :2005/12/26　SMP)石川　作成
'
'備考       :SXL管理（E042）→XSDCB機能移行
'           WF情報変更では、TBCME041が更新されないので、XSDCAとXSDCBを使用してデータを作成する
Private Function DBDRV_GetTBCME041_Clone(records() As typ_TBCME041, _
                                        sCryNum As String, _
                                        sBlockId() As String) As FUNCTION_RETURN
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Long
    Dim lsSXL()     As String
    Dim llSXLTop    As Long         'SXLの結晶内開始位置
    Dim llLastCBLen As Long         '最終SXLの長さ
    Dim tmpXSDCA    As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_GetTBCME041_Clone"

    tmpXSDCA = "   AND a.CRYNUMCA IN ("
    For i = 1 To UBound(sBlockId)
        tmpXSDCA = tmpXSDCA & "'" & sBlockId(i) & "',"
    Next i
    tmpXSDCA = Mid(tmpXSDCA, 1, Len(tmpXSDCA) - 1)
    tmpXSDCA = tmpXSDCA & ") "

    ''SQLを組み立てる
    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.XTALCA"
    sql = sql & "  ,a.HINBCA"
    sql = sql & "  ,a.INPOSCA"
    sql = sql & "  ,a.REVNUMCA"
    sql = sql & "  ,a.FACTORYCA"
    sql = sql & "  ,a.OPECA"
    sql = sql & "  ,b.RLENCB"
    sql = sql & "  ,a.SXLIDCA"
    sql = sql & "  ,NVL(b.INPOSCB,0) INPOSCB"
    sql = sql & " FROM"
    sql = sql & "   XSDCA A"
    sql = sql & "  ,XSDCB B"
    sql = sql & "  ,XSDC2 C"
    sql = sql & " WHERE a.SXLIDCA = b.SXLIDCB"
    sql = sql & "   AND a.CRYNUMCA  = c.CRYNUMC2"
'    sql = sql & "   AND c.WFHUFLG  = '1'"
    sql = sql & "   AND a.LIVKCA  = '0'"
    sql = sql & "   AND a.XTALCA = '" & Trim(sCryNum) & "'"
    sql = sql & tmpXSDCA
    sql = sql & " ORDER BY"
    sql = sql & "   a.SXLIDCA"
    sql = sql & "  ,a.INPOSCA"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    ReDim records(0) As typ_TBCME041
    ReDim lsSXL(0) As String
    i = 0
    llSXLTop = 0
    llLastCBLen = 0

    ''抽出結果を格納する
    Do Until rs.EOF 'データがなくなるまで取得
        i = i + 1
        ReDim Preserve records(i) As typ_TBCME041
        ReDim Preserve lsSXL(i) As String

        With records(i)
            .CRYNUM = rs("XTALCA")          ' 結晶番号
            .INGOTPOS = rs("INPOSCA")       ' 結晶内開始位置
            .hinban = rs("HINBCA")          ' 品番
            .REVNUM = rs("REVNUMCA")        ' 製品番号改訂番号
            .factory = rs("FACTORYCA")      ' 工場
            .opecond = rs("OPECA")          ' 操業条件
'            .LENGTH = rs("RLENCB")          ' XSDCBのSXL長さ
            lsSXL(i) = rs("SXLIDCA")        ' SXLID

            '長さを再計算する
            records(i - 1).LENGTH = .INGOTPOS - records(i - 1).INGOTPOS
            '最終SXLの結晶内TOP位置を保持
            If lsSXL(i) <> lsSXL(i - 1) Then
                llSXLTop = rs("INPOSCB")
            End If

            'ブロックの変わり目で長さを計算する
            If records(i).CRYNUM <> records(i - 1).CRYNUM And i <> 1 Then
                records(UBound(records)).LENGTH = (llSXLTop + llLastCBLen) - records(UBound(records)).INGOTPOS
            End If

            '最終SXLIDのXSCB.RLENBを保持する
            llLastCBLen = rs("RLENCB")

        End With
        rs.MoveNext
    Loop

    'ブロックの最後の品番の長さは (SXLのINGOTPOS + XSDCB.RLENB)-INGOTPOS
    records(UBound(records)).LENGTH = (llSXLTop + llLastCBLen) - records(UBound(records)).INGOTPOS

    rs.Close

    'データ無しの場合エラー
    If i = 0 Then
        ReDim records(0) As typ_TBCME041
        DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_GetTBCME041_Clone = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'＠(s)
'
'機能       :品番構造体データ合成
'
'返り値     :なし
'
'引き数     :tMotoHinMng    :全体の品番データ
'            tUpdateHinMng  :合成する品番データ
'
'機能説明   :tMotoHinMngのデータのうち、tUpdateHinMngの部分をtUpdateHinMngに置換える。
'
'履歴       :2005/12/26　SMP)石川　作成
'
'備考       :SXL管理（E042）→XSDCB機能移行
'           tUpdateHinMngは1ブロック分の品番データとする。
'
Private Sub s_cmbc036_2_F_SynHinban(tMotoHinMng() As typ_TBCME041, tUpdateHinMng() As typ_TBCME041)
    Dim tHinMng()   As typ_TBCME041         '品番管理テーブルワーク
    Dim j           As Long
    Dim k           As Long
    Dim i           As Long
    Dim sHinban1    As String               '一覧上の品番その1
    Dim sHinban2    As String               '一覧上の品番その2
    Dim sRev        As String               '品番その2から取得した製品番号改訂番号
    Dim sFac        As String               '品番その2から取得した工場
    Dim sOPE        As String               '品番その2から取得した操業条件
    Dim sCrystalNo  As String               '結晶番号
    Dim lCrystalPos As Integer              '結晶内位置
    Dim lBlockTP    As Long                 '結晶内位置（ブロックTOP）
    Dim lBlockBP    As Long                 '結晶内位置（ブロックBottom）
    Dim llRow       As Long
    Dim lLen        As Long                 '長さ
    Dim llUpdateFlg As Long                 '置換え済みフラグ（関係ブロック分を処理した:1 してない:0）

    ''品番管理テーブルのデータ構造体更新
    ReDim tHinMng(0) As typ_TBCME041
    i = 1
    j = 1
    ''結晶番号取得
    sCrystalNo = tUpdateHinMng(1).CRYNUM
    ''結晶内位置（ブロックTOP）取得
    lBlockTP = tUpdateHinMng(1).INGOTPOS
    ''結晶内位置（ブロックBottom）取得
    lBlockBP = tUpdateHinMng(UBound(tUpdateHinMng)).INGOTPOS + tUpdateHinMng(UBound(tUpdateHinMng)).LENGTH
    ''元品番の配列0の結晶内開始位置を初期化
    tMotoHinMng(0).INGOTPOS = 0
    '置換え済みフラグ初期化
    llUpdateFlg = 0

    For i = 1 To UBound(tMotoHinMng)
'        '元品番の結晶内位置が関係ブロックの最TOP位置より小さい、または、
'        '元品番の結晶内位置が関係ブロックの最Bottom位置以上かつ、
'        '  一つ前の元品番の結晶内位置が関係ブロックの最Bottom位置より大きい場合
'        If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
'         (tMotoHinMng(i).INGOTPOS > lBlockBP And tMotoHinMng(i - 1).INGOTPOS >= lBlockBP) Then

        ''元品番の結晶内位置が関係ブロックの最TOP位置より小さい、または
        ''元品番の結晶内位置が関係ブロックの最BOTTOM位置より大きいかつ、関係ブロックを反映済みの場合
        If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
          (tMotoHinMng(i).INGOTPOS >= lBlockBP And llUpdateFlg = 1) Then
            ''結晶内位置がブロックの範囲外
            ReDim Preserve tHinMng(j) As typ_TBCME041
            tHinMng(j) = tMotoHinMng(i)
            j = j + 1

        Else
            ''結晶内位置がブロックの範囲内
            '置換えフラグセット
            llUpdateFlg = 1
            ' 一つ上の品番の長さが変わる場合、調整する
            ' (関係ブロックの結晶内位置 - 一つ上の品番の結晶内位置)
            If tMotoHinMng(i).INGOTPOS <> lBlockTP Then
                tHinMng(j - 1).LENGTH = lBlockTP - tHinMng(j - 1).INGOTPOS
            End If

            For llRow = 1 To UBound(tUpdateHinMng)
                '品番取得
                sHinban1 = tUpdateHinMng(llRow).hinban
                '製品番号改訂番号
                sRev = tUpdateHinMng(llRow).REVNUM
                '工場
                sFac = tUpdateHinMng(llRow).factory
                '操業条件
                sOPE = tUpdateHinMng(llRow).opecond
                '結晶内位置取得
                lCrystalPos = CLng(tUpdateHinMng(llRow).INGOTPOS)
                '長さ
                lLen = CLng(tUpdateHinMng(llRow).LENGTH)

                'ワーク領域に設定
                ReDim Preserve tHinMng(j) As typ_TBCME041

                tHinMng(j).CRYNUM = sCrystalNo
                tHinMng(j).INGOTPOS = lCrystalPos
                tHinMng(j).hinban = sHinban1
                tHinMng(j).REVNUM = CInt(sRev)
                tHinMng(j).factory = sFac
                tHinMng(j).opecond = sOPE
                tHinMng(j).LENGTH = CInt(lLen)

                j = j + 1

            Next llRow
            ''ブロック範囲を抜けるまで進める
            Do While (1)
'                If tMotoHinMng(i).INGOTPOS < lBlockTP Or _
'                (tMotoHinMng(i).INGOTPOS > lBlockBP And tMotoHinMng(i - 1).INGOTPOS >= lBlockBP) Then

                ''結晶内開始位置が関係ブロックのBottom位置より大きい品番を探す
                If tMotoHinMng(i).INGOTPOS >= lBlockBP Then
                    'もともとの品番の長さが関係ブロックのBottom位置より長い場合、品番を一つ付け足す
                    'このとき、長さと結晶内開始位置を調整する
                    If tMotoHinMng(i).INGOTPOS <> lBlockBP Then
                        ReDim Preserve tHinMng(j) As typ_TBCME041
                        tHinMng(j) = tMotoHinMng(i - 1)
                        tHinMng(j).LENGTH = tMotoHinMng(i - 1).LENGTH - (lBlockBP - tMotoHinMng(i - 1).INGOTPOS)
                        tHinMng(j).INGOTPOS = lBlockBP
                        j = j + 1
                    End If

                    'ループのカウントアップでカウンタが進むので、一つ戻す
                    i = i - 1
                    Exit Do
                End If
                i = i + 1
                'データが無くなったら終了
                If i > UBound(tMotoHinMng) Then
                    Exit Do
                End If
            Loop
        End If
    Next i

    ''最終品番後に追加した場合を考慮
    If tMotoHinMng(UBound(tMotoHinMng)).INGOTPOS < lBlockTP Then
        For llRow = 1 To UBound(tUpdateHinMng)
            '品番取得
            sHinban1 = tUpdateHinMng(llRow).hinban
            '製品番号改訂番号
            sRev = tUpdateHinMng(llRow).REVNUM
            '工場
            sFac = tUpdateHinMng(llRow).factory
            '操業条件
            sOPE = tUpdateHinMng(llRow).opecond
            '結晶内位置取得
            lCrystalPos = CLng(tUpdateHinMng(llRow).INGOTPOS)
            '長さ
            lLen = CLng(tUpdateHinMng(llRow).LENGTH)

            'ワーク領域に設定
            ReDim Preserve tHinMng(j) As typ_TBCME041

            tHinMng(j).CRYNUM = sCrystalNo
            tHinMng(j).INGOTPOS = lCrystalPos
            tHinMng(j).hinban = sHinban1
            tHinMng(j).REVNUM = CInt(sRev)
            tHinMng(j).factory = sFac
            tHinMng(j).opecond = sOPE
            tHinMng(j).LENGTH = CInt(lLen)

            j = j + 1

        Next llRow
    End If

    ''tMotoHinMngに設定
    ReDim tMotoHinMng(UBound(tHinMng)) As typ_TBCME041
    For i = 1 To UBound(tHinMng)
        tMotoHinMng(i) = tHinMng(i)
    Next i

End Sub

'＠(f)
'
'機能       :新サンプル管理(SXL)データ補完
'
'返り値     :0 - 正常終了
'           :1 - 異常終了
'
'引き数     :tXSDCW()   - 取得XSDCWデータ
'           :sCryNum    - 結晶番号
'           :sBlockId   - 関係ブロックID
'
'機能説明   :XSDCWに補完するデータをXSDCAから取得し、引数で渡されたXSDCWのデータの修正を行う
'
'履歴       :2005/12/26　SMP)石川　作成
'
'備考       :SXL管理（E042）→XSDCB機能移行
'           WF情報変更では、XSDCWが更新されないので、XSDCAとXSDC2を使用してデータを作成する
Public Function DBDRV_GetXSDCWUpdate(tXSDCW() As typ_XSDCW, _
                                      sCryNum As String, _
                                      sBlockId() As String) As FUNCTION_RETURN
    Dim sql         As String       'SQL全体
    Dim rs          As OraDynaset   'RecordSet
    Dim i           As Long
    Dim j           As Long
    Dim lsWhere     As String
    Dim lsHinban    As String       '補完するかチェック用の品番
    Dim tmpXSDCA()  As typ_XSDCA    '関係ブロックID内のXSDCAデータ
    Dim tmpXSDCA2() As typ_XSDCA    '関係ブロックID内のXSDCAデータ
    Dim tmpXSDCW()  As typ_XSDCW    'XSDCWワーク領域
    Dim liUpdateFLG As Integer      '更新フラグ
    Dim liUpFLG2    As Integer      '更新フラグ
    Dim liEndSxlFLG As Integer      '最終SXLフラグ

    Dim lsHinWork   As String
    Dim lsSXLWork   As String
    Dim lsSendFlgWork   As String
    Dim llCnt       As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_GetXSDCWUpdate"

    '↓削除 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応
'    ''初期化
'    ReDim tSXLID(0)
    '↑削除 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応

    ''SQLを組み立てる

    'ブロックの条件を個別で作成
    lsWhere = "   AND a.CRYNUMCA IN ("
    For i = 1 To UBound(sBlockId)
        lsWhere = lsWhere & "'" & sBlockId(i) & "',"
    Next i
    lsWhere = Mid(lsWhere, 1, Len(lsWhere) - 1)
    lsWhere = lsWhere & ") "

    sql = ""
    sql = sql & " SELECT"
    sql = sql & "   a.XTALCA"       '結晶番号
    sql = sql & "  ,a.HINBCA"       '品番
    sql = sql & "  ,a.INPOSCA"      '結晶内開始位置
    sql = sql & "  ,a.REVNUMCA"     '製品番号改訂番号
    sql = sql & "  ,a.FACTORYCA"    '工場
    sql = sql & "  ,a.OPECA"        '操業条件
    sql = sql & "  ,a.GNLCA"        '現在長さ
    sql = sql & "  ,a.SXLIDCA"      'SXLID
    sql = sql & "  ,a.CRYNUMCA"     'ブロックID
    sql = sql & "  ,NVL(b.WFHUFLG,' ') WFHUFLG"      'WF振替FLG
    sql = sql & " FROM"
    sql = sql & "   XSDCA A"
    sql = sql & "  ,XSDC2 B"
    sql = sql & " WHERE a.LIVKCA  = '0'"
    sql = sql & "   AND a.CRYNUMCA = b.CRYNUMC2"
    sql = sql & "   AND a.XTALCA = '" & Trim(sCryNum) & "'"
    sql = sql & lsWhere
    sql = sql & " ORDER BY"
    sql = sql & "   a.XTALCA"
    sql = sql & "  ,a.INPOSCA"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    ReDim tmpXSDCA(0) As typ_XSDCA
    ReDim tmpXSDCW(0) As typ_XSDCW
    '↓削除 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応
'    ReDim tSXLID(0)
    '↑削除 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応
    i = 0
    liUpdateFLG = 0
    tmpXSDCA(0).SNDKCA = 0
    ''抽出結果を格納する
    Do Until rs.EOF 'データがなくなるまで取得
'        i = i + 1
'        ReDim Preserve tmpXSDCA(i) As typ_XSDCA

        lsHinWork = rs("HINBCA")        ' 品番
        lsSendFlgWork = rs("WFHUFLG")   ' WF振替FLG
        lsSXLWork = rs("SXLIDCA")       ' SXLID

        '1つ前のレコードの品番、SXLIDが同じで、ともに振替フラグが立っている場合、そのレコードは取得しない
'        If tmpXSDCA(i).HINBCA = lsHinWork _
'          And tmpXSDCA(i).SNDKCA = "1" _
'          And lsSendFlgWork = "1" _
'          And tmpXSDCA(i).SXLIDCA = lsSXLWork Then
'
'        Else

            i = i + 1
            ReDim Preserve tmpXSDCA(i) As typ_XSDCA
            With tmpXSDCA(i)

                .XTALCA = rs("XTALCA")          ' 結晶番号
                .INPOSCA = rs("INPOSCA")        ' 結晶内開始位置
                .HINBCA = rs("HINBCA")          ' 品番
                .REVNUMCA = rs("REVNUMCA")      ' 製品番号改訂番号
                .FACTORYCA = rs("FACTORYCA")    ' 工場
                .OPECA = rs("OPECA")            ' 操業条件
                .GNLCA = rs("GNLCA")            ' 長さ
                .SXLIDCA = rs("SXLIDCA")        ' SXLID
                .CRYNUMCA = rs("CRYNUMCA")      ' SXLID
                'WF振替FLGはCAに項目が無いので、変わりに送信フラグに入れる
                .SNDKCA = rs("WFHUFLG")         ' WF振替FLG
                If .SNDKCA = "1" Then
                    liUpdateFLG = 1
                End If

            End With
'        End If

        rs.MoveNext
    Loop

    rs.Close

    'データ無しの場合エラー
    If i = 0 Then
'        ReDim records(0) As typ_XSDCA
        DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''関係ブロックのWF振替FLGがすべて立っていない場合、WF情報変更で更新されていないので、CWの補完はしない
    If liUpdateFLG = 0 Then
        ReDim tXSDCW(0)
        DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If

    '↓追加 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応
    ''初期化
    ReDim tSXLID(0)
    '↑追加 2006/03/20 障害対応 SMP石川 WF情報変更されていない場合、欠落が表示されない障害に対応

    '' 引数で渡されたXSDCWのデータに対して、WF情報変更で変更した分を補完してやる
    i = 0
    j = 1
    liUpdateFLG = 0
    'XSDCWはTopとBottomでセットなので2つづ進める
    For i = 1 To UBound(tXSDCW) Step 2
        liUpFLG2 = 0
        liUpdateFLG = 0
        '1SXL内でXSDCWの品番と違う品番が存在した場合 = WF情報変更された場合、補完する。
        lsHinban = tXSDCW(i).HINBCW
        For j = 1 To UBound(tmpXSDCA)
            If tmpXSDCA(j).SXLIDCA = tXSDCW(i).SXLIDCW Then
                If tmpXSDCA(j).HINBCA <> lsHinban Then
                    liUpFLG2 = 1

                    For llCnt = 1 To UBound(tmpXSDCA)
                        ReDim Preserve tSXLID(llCnt)
                        tSXLID(llCnt).LOTID = tmpXSDCA(llCnt).CRYNUMCA
                        tSXLID(llCnt).SXLID = tmpXSDCA(llCnt).SXLIDCA
                        tSXLID(llCnt).INGOTPOS = tmpXSDCA(llCnt).INPOSCA
                    Next llCnt

                    Exit For
                End If
            End If
        Next j

        If liUpFLG2 = 1 Then
            ReDim tmpXSDCA2(0)
            For j = 1 To UBound(tmpXSDCA)
                '' 一つ前の品番が違う、又は　品番が同じでも、どちらかが振替えられていない場合
                    'SNDKCA=0→SNDKCA="0"に変更　06/06/15 ooba
                    'And (tmpXSDCA(j).SNDKCA = 0 Or tmpXSDCA(j - 1).SNDKCA = 0)
                If (tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA) _
                Or (tmpXSDCA(j).HINBCA = tmpXSDCA(j - 1).HINBCA _
                    And (tmpXSDCA(j).SNDKCA = "0" Or tmpXSDCA(j - 1).SNDKCA = "0") _
                   ) Then
                    ReDim Preserve tmpXSDCA2(UBound(tmpXSDCA2) + 1)
                    tmpXSDCA2(UBound(tmpXSDCA2)) = tmpXSDCA(j)

                End If
            Next j
            ReDim tmpXSDCA(UBound(tmpXSDCA2))
            For j = 1 To UBound(tmpXSDCA2)
                tmpXSDCA(j) = tmpXSDCA2(j)
            Next j


            liUpdateFLG = 0
            For j = 1 To UBound(tmpXSDCA)
                ''XSDCAのトップの位置が、XSDCWのトップ以上、Bottomより小さい場合、
                ''かつ品番が変わった場合に補完する
                If tXSDCW(i).INPOSCW <= tmpXSDCA(j).INPOSCA _
                 And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j).INPOSCA _
                 And tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA Then

'                ''XSDCAのトップの位置が、XSDCWのトップ以上、Bottomより小さい場合、
'                ''かつ品番が変わった場合、又は品番が同じで、両方振替えられているに補完する
'                If tXSDCW(i).INPOSCW <= tmpXSDCA(j).INPOSCA _
'                 And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j).INPOSCA _
'                 And ( _
'                           (tmpXSDCA(j).HINBCA = tmpXSDCA(j - 1).HINBCA _
'                              And ((CLng(tmpXSDCA(j).SNDKCA) + CLng(tmpXSDCA(j - 1).SNDKCA)) <> 2)) _
'                       Or _
'                           (tmpXSDCA(j).HINBCA <> tmpXSDCA(j - 1).HINBCA) _
'                     ) Then

                    ReDim Preserve tmpXSDCW(UBound(tmpXSDCW) + 2) As typ_XSDCW

                    'SXL内での品番の数をカウント
                    liUpdateFLG = liUpdateFLG + 1

                    '' TOP位置設定 ------------------------------------------------------------------------------------
                    '元XSDCWのTOPの位置は、XSDCWのデータを使用する
                    If liUpdateFLG = 1 Then
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW = tXSDCW(i).SXLIDCW              'SXLID
                        tmpXSDCA(j).SXLIDCA = tXSDCW(i).SXLIDCW 'SXLID保存
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = tXSDCW(i).SMPKBNCW            'サンプル区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TBKBNCW = tXSDCW(i).TBKBNCW              'T/B区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REPSMPLIDCW = tXSDCW(i).REPSMPLIDCW      '代表サンプルID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).XTALCW = tXSDCW(i).XTALCW                '結晶番号
                        tmpXSDCW(UBound(tmpXSDCW) - 1).INPOSCW = tXSDCW(i).INPOSCW              '結晶内位置
                        tmpXSDCW(UBound(tmpXSDCW) - 1).HINBCW = tmpXSDCA(j).HINBCA              '品番(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REVNUMCW = tmpXSDCA(j).REVNUMCA          '製品番号改訂番号(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).FACTORYCW = tmpXSDCA(j).FACTORYCA        '工場(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).OPECW = tmpXSDCA(j).OPECA                '操業条件(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KTKBNCW = tXSDCW(i).KTKBNCW              '確定区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMCRYNUMCW = tXSDCW(i).SMCRYNUMCW        'サンプルブロックID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRSCW = tXSDCW(i).WFSMPLIDRSCW    'サンプルID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS1CW = tXSDCW(i).WFSMPLIDRS1CW  '推定サンプルID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS2CW = tXSDCW(i).WFSMPLIDRS2CW  '推定サンプルID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDRSCW = tXSDCW(i).WFINDRSCW          '状態FLG（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS1CW = tXSDCW(i).WFRESRS1CW        '実績FLG1（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS2CW = tXSDCW(i).WFRESRS2CW        '実績FLG2（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOICW = tXSDCW(i).WFSMPLIDOICW    'サンプルID（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOICW = tXSDCW(i).WFINDOICW          '状態FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOICW = tXSDCW(i).WFRESOICW          '実績FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB1CW = tXSDCW(i).WFSMPLIDB1CW    'サンプルID（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB1CW = tXSDCW(i).WFINDB1CW          '状態FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB1CW = tXSDCW(i).WFRESB1CW          '実績FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB2CW = tXSDCW(i).WFSMPLIDB2CW    'サンプルID（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB2CW = tXSDCW(i).WFINDB2CW          '状態FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB2CW = tXSDCW(i).WFRESB2CW          '実績FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB3CW = tXSDCW(i).WFSMPLIDB3CW    'サンプルID（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB3CW = tXSDCW(i).WFINDB3CW          '状態FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB3CW = tXSDCW(i).WFRESB3CW          '実績FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL1CW = tXSDCW(i).WFSMPLIDL1CW    'サンプルID（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL1CW = tXSDCW(i).WFINDL1CW          '状態FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL1CW = tXSDCW(i).WFRESL1CW          '実績FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL2CW = tXSDCW(i).WFSMPLIDL2CW    'サンプルID（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL2CW = tXSDCW(i).WFINDL2CW          '状態FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL2CW = tXSDCW(i).WFRESL2CW          '実績FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL3CW = tXSDCW(i).WFSMPLIDL3CW    'サンプルID（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL3CW = tXSDCW(i).WFINDL3CW          '状態FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL3CW = tXSDCW(i).WFRESL3CW          '実績FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL4CW = tXSDCW(i).WFSMPLIDL4CW    'サンプルID（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL4CW = tXSDCW(i).WFINDL4CW          '状態FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL4CW = tXSDCW(i).WFRESL4CW          '実績FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDSCW = tXSDCW(i).WFSMPLIDDSCW    'サンプルID（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDSCW = tXSDCW(i).WFINDDSCW          '状態FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDSCW = tXSDCW(i).WFRESDSCW          '実績FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDZCW = tXSDCW(i).WFSMPLIDDZCW    'サンプルID（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDZCW = tXSDCW(i).WFINDDZCW          '状態FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDZCW = tXSDCW(i).WFRESDZCW          '実績FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDSPCW = tXSDCW(i).WFSMPLIDSPCW    'サンプルID（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDSPCW = tXSDCW(i).WFINDSPCW          '状態FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESSPCW = tXSDCW(i).WFRESSPCW          '実績FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO1CW = tXSDCW(i).WFSMPLIDDO1CW  'サンプルID（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO1CW = tXSDCW(i).WFINDDO1CW        '状態FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO1CW = tXSDCW(i).WFRESDO1CW        '実績FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO2CW = tXSDCW(i).WFSMPLIDDO2CW  'サンプルID（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO2CW = tXSDCW(i).WFINDDO2CW        '状態FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO2CW = tXSDCW(i).WFRESDO2CW        '実績FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO3CW = tXSDCW(i).WFSMPLIDDO3CW  'サンプルID（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO3CW = tXSDCW(i).WFINDDO3CW        '状態FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO3CW = tXSDCW(i).WFRESDO3CW        '実績FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT1CW = tXSDCW(i).WFSMPLIDOT1CW  'サンプルID（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT1CW = tXSDCW(i).WFINDOT1CW        '状態FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT1CW = tXSDCW(i).WFRESOT1CW        '実績FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT2CW = tXSDCW(i).WFSMPLIDOT2CW  'サンプルID（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT2CW = tXSDCW(i).WFINDOT2CW        '状態FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT2CW = tXSDCW(i).WFRESOT2CW        '実績FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDAOICW = tXSDCW(i).WFSMPLIDAOICW  'サンプルID（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDAOICW = tXSDCW(i).WFINDAOICW        '状態FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESAOICW = tXSDCW(i).WFRESAOICW        '実績FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLNUMCW = tXSDCW(i).SMPLNUMCW          'サンプル枚数
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLPATCW = tXSDCW(i).SMPLPATCW          'サンプルパターン
                        tmpXSDCW(UBound(tmpXSDCW) - 1).LIVKCW = tXSDCW(i).LIVKCW                '生死区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TSTAFFCW = tXSDCW(i).TSTAFFCW            '登録社員ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TDAYCW = tXSDCW(i).TDAYCW                '登録日付
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KSTAFFCW = tXSDCW(i).KSTAFFCW            '更新社員ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KDAYCW = tXSDCW(i).KDAYCW                '更新日付
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDKCW = tXSDCW(i).SNDKCW                '送信フラグ
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDDAYCW = tXSDCW(i).SNDDAYCW            '送信日付
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDGDCW = tXSDCW(i).WFSMPLIDGDCW    'サンプルID（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDGDCW = tXSDCW(i).WFINDGDCW          '状態FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESGDCW = tXSDCW(i).WFRESGDCW          '実績FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFHSGDCW = tXSDCW(i).WFHSGDCW            '保証FLG（GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB1CW = tXSDCW(i).EPSMPLIDB1CW    'サンプルID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB1CW = tXSDCW(i).EPINDB1CW          '状態FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB1CW = tXSDCW(i).EPRESB1CW          '実績FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB2CW = tXSDCW(i).EPSMPLIDB2CW    'サンプルID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB2CW = tXSDCW(i).EPINDB2CW          '状態FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB2CW = tXSDCW(i).EPRESB2CW          '実績FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB3CW = tXSDCW(i).EPSMPLIDB3CW    'サンプルID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB3CW = tXSDCW(i).EPINDB3CW          '状態FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB3CW = tXSDCW(i).EPRESB3CW          '実績FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL1CW = tXSDCW(i).EPSMPLIDL1CW    'サンプルID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL1CW = tXSDCW(i).EPINDL1CW          '状態FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL1CW = tXSDCW(i).EPRESL1CW          '実績FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL2CW = tXSDCW(i).EPSMPLIDL2CW    'サンプルID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL2CW = tXSDCW(i).EPINDL2CW          '状態FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL2CW = tXSDCW(i).EPRESL2CW          '実績FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL3CW = tXSDCW(i).EPSMPLIDL3CW    'サンプルID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL3CW = tXSDCW(i).EPINDL3CW          '状態FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL3CW = tXSDCW(i).EPRESL3CW          '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                    Else
                        '' WF情報変更で分割されて増えた分のレコードを補完する
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = tXSDCW(i).SMPKBNCW            'サンプル区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPKBNCW = "D"                           'サンプル区分

                        tmpXSDCW(UBound(tmpXSDCW) - 1).TBKBNCW = tXSDCW(i).TBKBNCW              'T/B区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).XTALCW = tXSDCW(i).XTALCW                '結晶番号
                        tmpXSDCW(UBound(tmpXSDCW) - 1).HINBCW = tmpXSDCA(j).HINBCA              '品番(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).REVNUMCW = tmpXSDCA(j).REVNUMCA          '製品番号改訂番号(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).FACTORYCW = tmpXSDCA(j).FACTORYCA        '工場(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).OPECW = tmpXSDCA(j).OPECA                '操業条件(XSDCA)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KTKBNCW = tXSDCW(i).KTKBNCW              '確定区分

                        tmpXSDCW(UBound(tmpXSDCW) - 1).REPSMPLIDCW = "                "         '代表サンプルID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).INPOSCW = tmpXSDCA(j).INPOSCA            '結晶内位置
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMCRYNUMCW = fncBsmpID(tmpXSDCW, UBound(tmpXSDCW) - 1)     'サンプルブロックID

                        tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW = Left(tmpXSDCA(j).CRYNUMCA, 10) & GetWafPos(tmpXSDCA(j).INPOSCA) 'SXLID
                        tmpXSDCA(j).SXLIDCA = tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW 'SXLID保存
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRSCW = ""                       'サンプルID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS1CW = "0"                     '推定サンプルID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDRS2CW = "0"                     '推定サンプルID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDRSCW = "0"                         '状態FLG（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS1CW = "0"                        '実績FLG1（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESRS2CW = "0"                        '実績FLG2（Rs)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOICW = ""                       'サンプルID（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOICW = "0"                         '状態FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOICW = "0"                         '実績FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB1CW = ""                       'サンプルID（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB1CW = "0"                         '状態FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB1CW = "0"                         '実績FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB2CW = ""                       'サンプルID（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB2CW = "0"                         '状態FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB2CW = "0"                         '実績FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDB3CW = ""                       'サンプルID（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDB3CW = "0"                         '状態FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESB3CW = "0"                         '実績FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL1CW = ""                       'サンプルID（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL1CW = "0"                         '状態FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL1CW = "0"                         '実績FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL2CW = ""                       'サンプルID（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL2CW = "0"                         '状態FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL2CW = "0"                         '実績FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL3CW = ""                       'サンプルID（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL3CW = "0"                         '状態FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL3CW = "0"                         '実績FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDL4CW = ""                       'サンプルID（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDL4CW = "0"                         '状態FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESL4CW = "0"                         '実績FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDSCW = ""                       'サンプルID（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDSCW = "0"                         '状態FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDSCW = "0"                         '実績FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDZCW = ""                       'サンプルID（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDZCW = "0"                          '状態FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDZCW = "0"                          '実績FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDSPCW = ""                        'サンプルID（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDSPCW = "0"                          '状態FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESSPCW = "0"                          '実績FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO1CW = ""                       'サンプルID（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO1CW = "0"                         '状態FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO1CW = "0"                         '実績FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO2CW = ""                       'サンプルID（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO2CW = "0"                         '状態FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO2CW = "0"                         '実績FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDDO3CW = ""                       'サンプルID（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDDO3CW = "0"                         '状態FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESDO3CW = "0"                         '実績FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT1CW = ""                       'サンプルID（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT1CW = "0"                         '状態FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT1CW = "0"                         '実績FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDOT2CW = ""                       'サンプルID（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDOT2CW = "0"                         '状態FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESOT2CW = "0"                         '実績FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDAOICW = ""                       'サンプルID（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDAOICW = "0"                         '状態FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESAOICW = "0"                         '実績FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLNUMCW = "0"                           'サンプル枚数
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SMPLPATCW = ""                           'サンプルパターン
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFSMPLIDGDCW = ""                        'サンプルID（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFINDGDCW = "0"                          '状態FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFRESGDCW = "0"                          '実績FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).WFHSGDCW = "0"                           '保証FLG（GD)

                        tmpXSDCW(UBound(tmpXSDCW) - 1).LIVKCW = tXSDCW(i).LIVKCW                '生死区分
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TSTAFFCW = tXSDCW(i).TSTAFFCW            '登録社員ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).TDAYCW = tXSDCW(i).TDAYCW                '登録日付
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KSTAFFCW = tXSDCW(i).KSTAFFCW            '更新社員ID
                        tmpXSDCW(UBound(tmpXSDCW) - 1).KDAYCW = tXSDCW(i).KDAYCW                '更新日付
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDKCW = tXSDCW(i).SNDKCW                '送信フラグ
                        tmpXSDCW(UBound(tmpXSDCW) - 1).SNDDAYCW = tXSDCW(i).SNDDAYCW            '送信日付

'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB1CW = ""                        'サンプルID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB1CW = "0"                          '状態FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB1CW = "0"                          '実績FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB2CW = ""                        'サンプルID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB2CW = "0"                          '状態FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB2CW = "0"                          '実績FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDB3CW = ""                        'サンプルID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDB3CW = "0"                          '状態FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESB3CW = "0"                          '実績FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL1CW = ""                        'サンプルID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL1CW = "0"                          '状態FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL1CW = "0"                          '実績FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL2CW = ""                        'サンプルID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL2CW = "0"                          '状態FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL2CW = "0"                          '実績FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPSMPLIDL3CW = ""                        'サンプルID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPINDL3CW = "0"                          '状態FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW) - 1).EPRESL3CW = "0"                          '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                    End If



                    '' Bottom位置設定 --------------------------------------------------------------------------------------
                    tmpXSDCW(UBound(tmpXSDCW)).SXLIDCW = tmpXSDCW(UBound(tmpXSDCW) - 1).SXLIDCW 'SXLID
'                    tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = tXSDCW(i + 1).SMPKBNCW            'サンプル区分
                    tmpXSDCW(UBound(tmpXSDCW)).TBKBNCW = tXSDCW(i + 1).TBKBNCW              'T/B区分
                    tmpXSDCW(UBound(tmpXSDCW)).XTALCW = tXSDCW(i).XTALCW                    '結晶番号
                    tmpXSDCW(UBound(tmpXSDCW)).HINBCW = tmpXSDCA(j).HINBCA                  '品番(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).REVNUMCW = tmpXSDCA(j).REVNUMCA              '製品番号改訂番号(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).FACTORYCW = tmpXSDCA(j).FACTORYCA            '工場(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).OPECW = tmpXSDCA(j).OPECA                    '操業条件(XSDCA)
                    tmpXSDCW(UBound(tmpXSDCW)).KTKBNCW = tXSDCW(i + 1).KTKBNCW              '確定区分

                    ''フラグ初期化
                    liEndSxlFLG = 0
                    If j + 1 <= UBound(tmpXSDCA) Then
                        If tXSDCW(i).INPOSCW <= tmpXSDCA(j + 1).INPOSCA _
                            And tXSDCW(i + 1).INPOSCW > tmpXSDCA(j + 1).INPOSCA Then
                            '' SXLの中（追加データ中）の場合フラグを立てる
                            liEndSxlFLG = 1
                            '' XSCAのデータが次のレコードもある場合
                            tmpXSDCW(UBound(tmpXSDCW)).REPSMPLIDCW = "                "         '代表サンプルID
                            tmpXSDCW(UBound(tmpXSDCW)).INPOSCW = tmpXSDCA(j + 1).INPOSCA        '結晶内位置
                            tmpXSDCW(UBound(tmpXSDCW)).SMCRYNUMCW = fncBsmpID(tmpXSDCW, UBound(tmpXSDCW))        'サンプルブロックID
                            tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = "U"                           'サンプル区分

                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRSCW = ""                        'サンプルID(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS1CW = "0"                      '推定サンプルID1(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS2CW = "0"                      '推定サンプルID2(Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDRSCW = "0"                          '状態FLG（Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESRS1CW = "0"                         '実績FLG1（Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESRS2CW = "0"                         '実績FLG2（Rs)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOICW = ""                        'サンプルID（Oi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOICW = "0"                          '状態FLG（Oi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOICW = "0"                          '実績FLG（Oi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB1CW = ""                        'サンプルID（B1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB1CW = "0"                          '状態FLG（B1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB1CW = "0"                          '実績FLG（B1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB2CW = ""                        'サンプルID（B2）
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB2CW = "0"                          '状態FLG（B2）
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB2CW = "0"                          '実績FLG（B2）
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB3CW = ""                        'サンプルID（B3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDB3CW = "0"                          '状態FLG（B3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESB3CW = "0"                          '実績FLG（B3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL1CW = ""                        'サンプルID（L1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL1CW = "0"                          '状態FLG（L1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL1CW = "0"                          '実績FLG（L1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL2CW = ""                        'サンプルID（L2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL2CW = "0"                          '状態FLG（L2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL2CW = "0"                          '実績FLG（L2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL3CW = ""                        'サンプルID（L3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL3CW = "0"                          '状態FLG（L3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL3CW = "0"                          '実績FLG（L3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL4CW = ""                        'サンプルID（L4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDL4CW = "0"                          '状態FLG（L4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESL4CW = "0"                          '実績FLG（L4)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDSCW = ""                        'サンプルID（DS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDSCW = "0"                          '状態FLG（DS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDSCW = "0"                          '実績FLG（DS)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDZCW = ""                        'サンプルID（DZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDZCW = "0"          '状態FLG（DZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDZCW = "0"          '実績FLG（DZ)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDSPCW = ""        'サンプルID（SP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDSPCW = "0"          '状態FLG（SP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESSPCW = "0"          '実績FLG（SP)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO1CW = ""       'サンプルID（DO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO1CW = "0"         '状態FLG（DO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO1CW = "0"         '実績FLG（DO1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO2CW = ""       'サンプルID（DO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO2CW = "0"         '状態FLG（DO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO2CW = "0"         '実績FLG（DO2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO3CW = ""       'サンプルID（DO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDDO3CW = "0"         '状態FLG（DO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESDO3CW = "0"         '実績FLG（DO3)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT1CW = ""       'サンプルID（OT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOT1CW = "0"         '状態FLG（OT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOT1CW = "0"         '実績FLG（OT1)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT2CW = ""       'サンプルID（OT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDOT2CW = "0"         '状態FLG（OT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESOT2CW = "0"         '実績FLG（OT2)
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDAOICW = ""       'サンプルID（AOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDAOICW = "0"         '状態FLG（AOi)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESAOICW = "0"         '実績FLG（AOi)
                            tmpXSDCW(UBound(tmpXSDCW)).SMPLNUMCW = "0"           'サンプル枚数
                            tmpXSDCW(UBound(tmpXSDCW)).SMPLPATCW = ""           'サンプルパターン
                            tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDGDCW = ""        'サンプルID（GD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFINDGDCW = "0"          '状態FLG（GD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFRESGDCW = "0"          '実績FLG（GD)
                            tmpXSDCW(UBound(tmpXSDCW)).WFHSGDCW = "0"           '保証FLG（GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB1CW = ""                    'サンプルID(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB1CW = "0"                      '状態FLG(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB1CW = "0"                      '実績FLG(BMD1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB2CW = ""                    'サンプルID(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB2CW = "0"                      '状態FLG(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB2CW = "0"                      '実績FLG(BMD2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB3CW = ""                    'サンプルID(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDB3CW = "0"                      '状態FLG(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESB3CW = "0"                      '実績FLG(BMD3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL1CW = ""                    'サンプルID(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL1CW = "0"                      '状態FLG(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL1CW = "0"                      '実績FLG(OSF1)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL2CW = ""                    'サンプルID(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL2CW = "0"                      '状態FLG(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL2CW = "0"                      '実績FLG(OSF2)
                            tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL3CW = ""                    'サンプルID(OSF3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPINDL3CW = "0"                      '状態FLG(OSF3)
                            tmpXSDCW(UBound(tmpXSDCW)).EPRESL3CW = "0"                      '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-
                        End If
                    End If
                    ''
                    If liEndSxlFLG = 0 Then
                        '' XSCAのデータがこれで最後の場合
                        tmpXSDCW(UBound(tmpXSDCW)).REPSMPLIDCW = tXSDCW(i + 1).REPSMPLIDCW         '代表サンプルID
                        tmpXSDCW(UBound(tmpXSDCW)).INPOSCW = tXSDCW(i + 1).INPOSCW              '結晶内位置
                        tmpXSDCW(UBound(tmpXSDCW)).SMCRYNUMCW = tXSDCW(i + 1).SMCRYNUMCW        'サンプルブロックID
                        tmpXSDCW(UBound(tmpXSDCW)).SMPKBNCW = tXSDCW(i + 1).SMPKBNCW            'サンプル区分

                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRSCW = tXSDCW(i + 1).WFSMPLIDRSCW   'サンプルID(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS1CW = tXSDCW(i + 1).WFSMPLIDRS1CW '推定サンプルID1(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDRS2CW = tXSDCW(i + 1).WFSMPLIDRS2CW '推定サンプルID2(Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDRSCW = tXSDCW(i + 1).WFINDRSCW         '状態FLG（Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESRS1CW = tXSDCW(i + 1).WFRESRS1CW       '実績FLG1（Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESRS2CW = tXSDCW(i + 1).WFRESRS2CW       '実績FLG2（Rs)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOICW = tXSDCW(i + 1).WFSMPLIDOICW   'サンプルID（Oi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOICW = tXSDCW(i + 1).WFINDOICW         '状態FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOICW = tXSDCW(i + 1).WFRESOICW         '実績FLG（Oi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB1CW = tXSDCW(i + 1).WFSMPLIDB1CW   'サンプルID（B1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB1CW = tXSDCW(i + 1).WFINDB1CW         '状態FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB1CW = tXSDCW(i + 1).WFRESB1CW         '実績FLG（B1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB2CW = tXSDCW(i + 1).WFSMPLIDB2CW   'サンプルID（B2）
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB2CW = tXSDCW(i + 1).WFINDB2CW         '状態FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB2CW = tXSDCW(i + 1).WFRESB2CW         '実績FLG（B2）
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDB3CW = tXSDCW(i + 1).WFSMPLIDB3CW   'サンプルID（B3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDB3CW = tXSDCW(i + 1).WFINDB3CW         '状態FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESB3CW = tXSDCW(i + 1).WFRESB3CW         '実績FLG（B3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL1CW = tXSDCW(i + 1).WFSMPLIDL1CW   'サンプルID（L1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL1CW = tXSDCW(i + 1).WFINDL1CW         '状態FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL1CW = tXSDCW(i + 1).WFRESL1CW         '実績FLG（L1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL2CW = tXSDCW(i + 1).WFSMPLIDL2CW   'サンプルID（L2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL2CW = tXSDCW(i + 1).WFINDL2CW         '状態FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL2CW = tXSDCW(i + 1).WFRESL2CW         '実績FLG（L2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL3CW = tXSDCW(i + 1).WFSMPLIDL3CW   'サンプルID（L3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL3CW = tXSDCW(i + 1).WFINDL3CW         '状態FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL3CW = tXSDCW(i + 1).WFRESL3CW         '実績FLG（L3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDL4CW = tXSDCW(i + 1).WFSMPLIDL4CW   'サンプルID（L4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDL4CW = tXSDCW(i + 1).WFINDL4CW         '状態FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESL4CW = tXSDCW(i + 1).WFRESL4CW         '実績FLG（L4)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDSCW = tXSDCW(i + 1).WFSMPLIDDSCW   'サンプルID（DS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDSCW = tXSDCW(i + 1).WFINDDSCW         '状態FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDSCW = tXSDCW(i + 1).WFRESDSCW         '実績FLG（DS)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDZCW = tXSDCW(i + 1).WFSMPLIDDZCW   'サンプルID（DZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDZCW = tXSDCW(i + 1).WFINDDZCW         '状態FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDZCW = tXSDCW(i + 1).WFRESDZCW         '実績FLG（DZ)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDSPCW = tXSDCW(i + 1).WFSMPLIDSPCW   'サンプルID（SP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDSPCW = tXSDCW(i + 1).WFINDSPCW         '状態FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESSPCW = tXSDCW(i + 1).WFRESSPCW         '実績FLG（SP)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO1CW = tXSDCW(i + 1).WFSMPLIDDO1CW 'サンプルID（DO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO1CW = tXSDCW(i + 1).WFINDDO1CW       '状態FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO1CW = tXSDCW(i + 1).WFRESDO1CW       '実績FLG（DO1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO2CW = tXSDCW(i + 1).WFSMPLIDDO2CW 'サンプルID（DO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO2CW = tXSDCW(i + 1).WFINDDO2CW       '状態FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO2CW = tXSDCW(i + 1).WFRESDO2CW       '実績FLG（DO2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDDO3CW = tXSDCW(i + 1).WFSMPLIDDO3CW 'サンプルID（DO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDDO3CW = tXSDCW(i + 1).WFINDDO3CW       '状態FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESDO3CW = tXSDCW(i + 1).WFRESDO3CW       '実績FLG（DO3)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT1CW = tXSDCW(i + 1).WFSMPLIDOT1CW 'サンプルID（OT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOT1CW = tXSDCW(i + 1).WFINDOT1CW       '状態FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOT1CW = tXSDCW(i + 1).WFRESOT1CW       '実績FLG（OT1)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDOT2CW = tXSDCW(i + 1).WFSMPLIDOT2CW 'サンプルID（OT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDOT2CW = tXSDCW(i + 1).WFINDOT2CW       '状態FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESOT2CW = tXSDCW(i + 1).WFRESOT2CW       '実績FLG（OT2)
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDAOICW = tXSDCW(i + 1).WFSMPLIDAOICW 'サンプルID（AOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDAOICW = tXSDCW(i + 1).WFINDAOICW       '状態FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESAOICW = tXSDCW(i + 1).WFRESAOICW       '実績FLG（AOi)
                        tmpXSDCW(UBound(tmpXSDCW)).SMPLNUMCW = tXSDCW(i + 1).SMPLNUMCW         'サンプル枚数
                        tmpXSDCW(UBound(tmpXSDCW)).SMPLPATCW = tXSDCW(i + 1).SMPLPATCW         'サンプルパターン
                        tmpXSDCW(UBound(tmpXSDCW)).WFSMPLIDGDCW = tXSDCW(i + 1).WFSMPLIDGDCW   'サンプルID（GD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFINDGDCW = tXSDCW(i + 1).WFINDGDCW         '状態FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFRESGDCW = tXSDCW(i + 1).WFRESGDCW         '実績FLG（GD)
                        tmpXSDCW(UBound(tmpXSDCW)).WFHSGDCW = tXSDCW(i + 1).WFHSGDCW           '保証FLG（GD)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -s-
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB1CW = tXSDCW(i + 1).EPSMPLIDB1CW    'サンプルID(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB1CW = tXSDCW(i + 1).EPINDB1CW          '状態FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB1CW = tXSDCW(i + 1).EPRESB1CW          '実績FLG(BMD1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB2CW = tXSDCW(i + 1).EPSMPLIDB2CW    'サンプルID(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB2CW = tXSDCW(i + 1).EPINDB2CW          '状態FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB2CW = tXSDCW(i + 1).EPRESB2CW          '実績FLG(BMD2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDB3CW = tXSDCW(i + 1).EPSMPLIDB3CW    'サンプルID(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDB3CW = tXSDCW(i + 1).EPINDB3CW          '状態FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESB3CW = tXSDCW(i + 1).EPRESB3CW          '実績FLG(BMD3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL1CW = tXSDCW(i + 1).EPSMPLIDL1CW    'サンプルID(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL1CW = tXSDCW(i + 1).EPINDL1CW          '状態FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL1CW = tXSDCW(i + 1).EPRESL1CW          '実績FLG(OSF1)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL2CW = tXSDCW(i + 1).EPSMPLIDL2CW    'サンプルID(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL2CW = tXSDCW(i + 1).EPINDL2CW          '状態FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL2CW = tXSDCW(i + 1).EPRESL2CW          '実績FLG(OSF2)
                        tmpXSDCW(UBound(tmpXSDCW)).EPSMPLIDL3CW = tXSDCW(i + 1).EPSMPLIDL3CW    'サンプルID(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPINDL3CW = tXSDCW(i + 1).EPINDL3CW          '状態FLG(OSF3)
                        tmpXSDCW(UBound(tmpXSDCW)).EPRESL3CW = tXSDCW(i + 1).EPRESL3CW          '実績FLG(OSF3)
'--- 2006/08/15 Add エピ先行評価追加対応 SMP)kondoh -e-

                    End If
                    tmpXSDCW(UBound(tmpXSDCW)).LIVKCW = tXSDCW(i).LIVKCW                '生死区分
                    tmpXSDCW(UBound(tmpXSDCW)).TSTAFFCW = tXSDCW(i).TSTAFFCW            '登録社員ID
                    tmpXSDCW(UBound(tmpXSDCW)).TDAYCW = tXSDCW(i).TDAYCW                '登録日付
                    tmpXSDCW(UBound(tmpXSDCW)).KSTAFFCW = tXSDCW(i).KSTAFFCW            '更新社員ID
                    tmpXSDCW(UBound(tmpXSDCW)).KDAYCW = tXSDCW(i).KDAYCW                '更新日付
                    tmpXSDCW(UBound(tmpXSDCW)).SNDKCW = tXSDCW(i).SNDKCW                '送信フラグ
                    tmpXSDCW(UBound(tmpXSDCW)).SNDDAYCW = tXSDCW(i).SNDDAYCW            '送信日付

                End If
            Next j
'            For j = 1 To UBound(tmpXSDCA)
'                ReDim Preserve tSXLID(j)
'                tSXLID(j).LOTID = tmpXSDCA(j).CRYNUMCA
'                tSXLID(j).SXLID = tmpXSDCA(j).SXLIDCA
'                tSXLID(j).IngotPos = tmpXSDCA(j).INPOSCA
'            Next j
        Else
            '補完されていない場合、そのままのデータを使用
            If liUpdateFLG = 0 Then
                ReDim Preserve tmpXSDCW(UBound(tmpXSDCW) + 2) As typ_XSDCW
                'TOP部分の設定
                tmpXSDCW(UBound(tmpXSDCW) - 1) = tXSDCW(i)
                'Bottom部分の設定
                tmpXSDCW(UBound(tmpXSDCW)) = tXSDCW(i + 1)
            End If
        End If
    Next i

    ''ワーク領域のデータを反映させる
    ReDim tXSDCW(UBound(tmpXSDCW)) As typ_XSDCW
    For i = 0 To UBound(tmpXSDCW)
        tXSDCW(i) = tmpXSDCW(i)
    Next i

    DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_GetXSDCWUpdate = FUNCTION_RETURN_FAILURE
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'概要      :SXL管理の挿入
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:SXL   　　　,I  ,typ_TBCME042   　,SXL管理
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :DBDRV_SXL_UpdInsに移行する予定
'履歴      :2001/07/12  作成 蔵本
'           2006/01/20 SXL管理（E042）→XSDCB機能移行 SMP石川
'           s_cmzcDBdriverCOM_SQL.DBDRV_SXL_INSより移植

Private Function DBDRV_SXL_INS_CB(SXL() As typ_TBCME042) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset   'RecordSet
    Dim liRecCnt        As Long
    Dim lsMotoHinban    As String      '元品番
    Dim iLoopBkHinGet   As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcDBdriverCOM_SQL.bas -- Function DBDRV_SXL_INS_CB"

    DBDRV_SXL_INS_CB = FUNCTION_RETURN_SUCCESS

    For i = 1 To UBound(SXL)
        If SXL(i).LENGTH > 0 Then
            'E042の時は、結晶番号と結晶内開始位置で見ていたが、
            'XSDCBに変更に伴いSXLIDで検索するように変える
            sql = "select count(XTALCB) cnt from XSDCB where SXLIDCB='" & SXL(i).SXLID & "'"
            ''データを抽出する
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
            liRecCnt = CLng(rs("CNT"))
            rs.Close

            '元品番取得
            lsMotoHinban = ""
            For iLoopBkHinGet = 0 To Kihon.CNTHINOLD - 1
                If (CInt(HinOld(iLoopBkHinGet).INPOSCA) <= CInt(SXL(0).INGOTPOS)) And (CInt(SXL(0).INGOTPOS) <= CInt(HinOld(iLoopBkHinGet).INPOSCA) + CInt(HinOld(iLoopBkHinGet).GNLCA)) Then
                     lsMotoHinban = HinOld(iLoopBkHinGet).HINBCA
                     Exit For
                End If
            Next
            If lsMotoHinban = "" Then 'もし該当HINOLDが無かったら自分の品番を元品番とする
                lsMotoHinban = SXL(0).hinban
            End If

            'データがない場合はInsert、あった場合はUpdateにする
            With SXL(i)

                If liRecCnt = 0 Then
                    sql = ""
                    sql = sql & " INSERT INTO XSDCB"
                    sql = sql & " ("
                    sql = sql & "   XTALCB"                         ' 結晶番号
                    sql = sql & "  ,INPOSCB"                        ' 結晶内開始位置
                    sql = sql & "  ,RLENCB"                         ' 長さ
                    sql = sql & "  ,SXLIDCB"                        ' SXLID
                    sql = sql & "  ,GNWKNTCB"                       ' 現在工程
                    sql = sql & "  ,NEWKNTCB"                       ' 最終通過工程
                    sql = sql & "  ,LIVKCB"                         ' 削除区分
                    sql = sql & "  ,LSTCCB"                         ' 最終状態区分
                    sql = sql & "  ,SHOLDCLSCB"                     ' ホールド区分
                    sql = sql & "  ,HINBCB"                         ' 品番
                    sql = sql & "  ,REVNUMCB"                       ' 製品番号改訂番号
                    sql = sql & "  ,FACTORYCB"                      ' 工場
                    sql = sql & "  ,OPECB"                          ' 操業条件
                    sql = sql & "  ,FURYCCB"                        ' 不良理由
                    sql = sql & "  ,MAICB"                          ' 枚数
                    sql = sql & "  ,TDAYCB"                         ' 登録日付
                    sql = sql & "  ,KDAYCB"                         ' 更新日付
                    sql = sql & "  ,SNDKCB"                         ' 送信フラグ
                    sql = sql & "  ,WSRMAICB"                       ' WS洗後枚数
                    sql = sql & "  ,WSNMAICB"                       ' WS洗浄欠落枚数
                    sql = sql & "  ,WFCMAICB"                       ' 受入枚数
                    sql = sql & "  ,SXLRMAICB"                      ' SXL指示(良品)
                    sql = sql & "  ,WFCNMAICB"                      ' WFC内欠落枚数
                    sql = sql & "  ,SXLEMAICB"                      ' SXL確定枚数
                    sql = sql & "  ,SRMAICB"                        ' サンプル抜指示枚数
                    sql = sql & "  ,SNMAICB"                        ' サンプル抜指示不良枚数
                    sql = sql & "  ,STMAICB"                        ' サンプル枚数
                    sql = sql & "  ,FURIMAICB"                      ' 振替枚数
                    sql = sql & "  ,XTWORKCB"                       ' 製造工場
                    sql = sql & "  ,WFWORKCB"                       ' ウェーハ製造
                    sql = sql & "  ,LUFRCCB"                        ' 格上コード
                    sql = sql & "  ,LUFRBCB"                        ' 格上区分
                    sql = sql & "  ,LDERCCB"                        ' 格下コード
                    sql = sql & "  ,HOLDCCB"                        ' ホールドコード
                    sql = sql & "  ,HOLDBCB"                        ' ホールド区分
                    sql = sql & "  ,EXKUBCB"                        ' 例外区分
                    sql = sql & "  ,HENPKCB"                        ' 返品区分
                    sql = sql & "  ,KANKCB"                         ' 完了区分
                    sql = sql & "  ,NFCB"                           ' 入庫区分
                    sql = sql & "  ,SAKJCB"                         ' 削除区分
                    sql = sql & "  ,SUMITCB"                        ' SUMIT送信フラグ
                    sql = sql & " )"
                    sql = sql & " VALUES"
                    sql = sql & " ("
                    sql = sql & "   '" & .CRYNUM & "'"              ' 結晶番号
                    sql = sql & "  ," & .INGOTPOS & ""              ' 結晶内開始位置
                    sql = sql & "  ," & .LENGTH & ""                ' 長さ
                    sql = sql & "  ,'" & .SXLID & "'"               ' SXLID
                    sql = sql & "  ,'" & .NOWPROC & "'"             ' 現在工程
                    sql = sql & "  ,'" & .LPKRPROCCD & "'"          ' 最終通過工程
                    sql = sql & "  ,'" & .DELCLS & "'"              ' 削除区分
                    sql = sql & "  ,'" & .LSTATCLS & "'"            ' 最終状態区分
                    sql = sql & "  ,'" & .HOLDCLS & "'"             ' ホールド区分
                    sql = sql & "  ,'" & .hinban & "'"              ' 品番
                    sql = sql & "  ," & .REVNUM & ""                ' 製品番号改訂番号
                    sql = sql & "  ,'" & .factory & "'"             ' 工場
                    sql = sql & "  ,'" & .opecond & "'"             ' 操業条件
                    sql = sql & "  ,'" & .BDCAUS & "'"              ' 不良理由
                    sql = sql & "  ," & .Count & ""                 ' 枚数
                    sql = sql & "  ,sysdate"                        ' 登録日付
                    sql = sql & "  ,sysdate"                        ' 更新日付
                    sql = sql & "  ,'0'"                            ' 送信フラグ
                    sql = sql & "  ,'0'"                            ' WS洗後枚数
                    sql = sql & "  ,'0'"                            ' WS洗浄欠落枚数
                    sql = sql & "  ,'0'"                            ' 受入枚数
                    sql = sql & "  ,'0'"                            ' SXL指示(良品)
                    sql = sql & "  ,'0'"                            ' WFC内欠落枚数
                    sql = sql & "  ,'0'"                            ' SXL確定枚数
                    sql = sql & "  ,'0'"                            ' サンプル抜指示枚数
                    sql = sql & "  ,'0'"                            ' サンプル抜指示不良枚数
                    sql = sql & "  ,'0'"                            ' サンプル枚数
                    sql = sql & "  ,'0'"                            ' 振替枚数
                    sql = sql & "  ,'42'"                           ' 製造工場
                    sql = sql & "  ,'  '"                           ' ウェーハ製造
                    sql = sql & "  ,'   '"                          ' 格上コード
                    sql = sql & "  ,' '"                            ' 格上区分
                    sql = sql & "  ,'   '"                          ' 格下コード
                    sql = sql & "  ,'   '"                          ' ホールドコード
                    sql = sql & "  ,'0'"                            ' ホールド区分
                    sql = sql & "  ,' '"                            ' 例外区分
                    sql = sql & "  ,' '"                            ' 返品区分
                    sql = sql & "  ,'0'"                            ' 完了区分
                    sql = sql & "  ,'0'"                            ' 入庫区分
                    sql = sql & "  ,'0'"                            ' 削除区分
                    sql = sql & "  ,'0'"                            ' SUMIT送信フラグ
                    sql = sql & " )"
                Else
                    sql = ""
                    sql = sql & " UPDATE XSDCB"
                    sql = sql & " SET XTALCB   = '" & .CRYNUM & "'"
                    sql = sql & "  ,INPOSCB    = " & .INGOTPOS & ""
                    sql = sql & "  ,RLENCB     = " & .LENGTH & ""
                    sql = sql & "  ,SXLIDCB    = '" & .SXLID & "'"
                    sql = sql & "  ,GNWKNTCB   = '" & .NOWPROC & "'"
                    sql = sql & "  ,NEWKNTCB   = '" & .LPKRPROCCD & "'"
                    sql = sql & "  ,LIVKCB     = '" & .DELCLS & "'"
                    sql = sql & "  ,LSTCCB     = '" & .LSTATCLS & "'"
                    sql = sql & "  ,SHOLDCLSCB = '" & .HOLDCLS & "'"
                    sql = sql & "  ,HINBCB     = '" & .hinban & "'"
                    sql = sql & "  ,REVNUMCB   = " & .REVNUM & ""
                    sql = sql & "  ,FACTORYCB  = '" & .factory & "'"
                    sql = sql & "  ,OPECB      = '" & .opecond & "'"
                    sql = sql & "  ,FURYCCB    = '" & .BDCAUS & "'"
                    sql = sql & "  ,MAICB      = " & .Count & ""
                    sql = sql & "  ,TDAYCB     = sysdate"
                    sql = sql & "  ,KDAYCB     = sysdate"
                    sql = sql & "  ,SNDKCB     = '0'"
                    sql = sql & "  ,SNDAYCB    = sysdate"
                    sql = sql & " where SXLIDCB='" & SXL(i).SXLID & "'"
                End If
            End With
            '' WriteDBLog sql
            If OraDB.ExecuteSQL(sql) <= 0 Then
                DBDRV_SXL_INS_CB = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End If
    Next i

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    DBDRV_SXL_INS_CB = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit

End Function

'ﾌﾞﾛｯｸｻﾝﾌﾟﾙ管理(XSDCS)からﾌﾞﾛｯｸ終了位置を取得　08/07/10 ooba
Public Function GetCSpos(sBlkId As String, iPos As Integer) As Integer
    Dim sql         As String
    Dim rs          As OraDynaset
    
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function GetCSpos"
    
    GetCSpos = iPos
    
    sql = "select INPOSCS from XSDCS "
    sql = sql & "where CRYNUMCS = '" & sBlkId & "' "
    sql = sql & "and TBKBNCS = 'B' "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount <> 1 Then
        GoTo proc_exit
    End If

    If IsNull(rs("INPOSCS")) = False Then GetCSpos = rs("INPOSCS")
    rs.Close
    
proc_exit:
    gErr.Pop
    Exit Function

proc_err:
    Resume proc_exit
    
End Function

'概要      :WFﾏｯﾌﾟ(TBCMY011)登録(関連ﾌﾞﾛｯｸ変更時)
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型                ,説明
'      　　:sSxlid       ,I  ,String            ,SXLID
'      　　:sqlWhere     ,I  ,String            ,条件式
'      　　:戻り値       ,O  ,FUNCTION_RETURN　 ,書き込みの成否
'説明      :
'履歴      :08/01/28 ooba
Public Function DBDRV_KanrenBlkMap(sSXLID As String, sqlWhere As String) As FUNCTION_RETURN
    
    Dim sql     As String
    
    'ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlkMap"

    DBDRV_KanrenBlkMap = FUNCTION_RETURN_FAILURE
    
    '条件式なし
    If Trim(sqlWhere) = "" Then GoTo proc_exit
    
    'SXLID更新
    sql = "UPDATE TBCMY011 "
    sql = sql & "SET MSXLID = '" & sSXLID & "' "
    sql = sql & sqlWhere
    
    If OraDB.ExecuteSQL(sql) < 0 Then
        '更新失敗
        GoTo proc_exit
    End If
    
    DBDRV_KanrenBlkMap = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' ｴﾗｰﾊﾝﾄﾞﾗ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'概要      :関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)登録
'ﾊﾟﾗﾒｰﾀ　　:変数名           ,IO ,型                ,説明
'      　　:tKanrenDisp()    ,I  ,typ_KanrenDisp    ,関連ﾌﾞﾛｯｸ一覧
'      　　:戻り値           ,O  ,FUNCTION_RETURN　 ,書き込みの成否
'説明      :
'履歴      :08/01/28 ooba
Public Function DBDRV_KanrenBlk(tKanrenDisp() As typ_KanrenDisp) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ﾚｺｰﾄﾞ数
    Dim iTrnCnt         As Integer          '処理回数
    Dim bSaveFlg        As Boolean          '登録有無
    Dim KanrenData()    As typ_TBCMY023     '関連ﾌﾞﾛｯｸ紐付紐切ﾃﾞｰﾀ
    
    'ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '処理回数取得
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & tKanrenDisp(1).CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '処理回数(最大) + 1
    End If
    rs.Close

    lRecCnt = 0             '登録ﾚｺｰﾄﾞ数
    
    '関連ﾌﾞﾛｯｸ紐切ﾃﾞｰﾀｾｯﾄ
    For i = 1 To UBound(tKanrenDisp)
        lRecCnt = lRecCnt + 1
        ReDim Preserve KanrenData(lRecCnt)
        With KanrenData(lRecCnt)
            .CRYNUM = tKanrenDisp(i).CRYNUM     '結晶番号
            .TRANCNT = iTrnCnt                  '処理回数
            .BLOCKID = tKanrenDisp(i).BLOCKID   'ﾌﾞﾛｯｸID
            .PROCCAT = "D"                      '処理区分(D:紐切)
            .TXID = "TX879I"                    'ﾄﾗﾝｻﾞｸｼｮﾝID
        End With
    Next i
    
    '関連ﾌﾞﾛｯｸ紐付ﾃﾞｰﾀｾｯﾄ
    For i = 1 To UBound(tKanrenDisp)
        bSaveFlg = False
        '関連ﾌﾞﾛｯｸの先頭ﾌﾞﾛｯｸ(処理前)
        If i = 1 Then
            If tKanrenDisp(i).KANREN = 0 Then
                iTrnCnt = iTrnCnt + 1       '処理回数+1
                bSaveFlg = True
            End If
        Else
            '前のﾌﾞﾛｯｸと関連ﾌﾞﾛｯｸ
            If tKanrenDisp(i - 1).KANREN = 0 Then bSaveFlg = True
            '関連ﾌﾞﾛｯｸの先頭ﾌﾞﾛｯｸ(処理後)
            If tKanrenDisp(i - 1).KANREN = 1 And tKanrenDisp(i).KANREN = 0 Then
                iTrnCnt = iTrnCnt + 1       '処理回数+1
                bSaveFlg = True
            End If
        End If
        
        If bSaveFlg Then
            lRecCnt = lRecCnt + 1
            ReDim Preserve KanrenData(lRecCnt)
            With KanrenData(lRecCnt)
                .CRYNUM = tKanrenDisp(i).CRYNUM     '結晶番号
                .TRANCNT = iTrnCnt                  '処理回数
                .BLOCKID = tKanrenDisp(i).BLOCKID   'ﾌﾞﾛｯｸID
                .PROCCAT = "C"                      '処理区分(C:付替え)
                .TXID = "TX879I"                    'ﾄﾗﾝｻﾞｸｼｮﾝID
            End With
        End If
    Next i
    
    '関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)に登録
    For i = 1 To UBound(KanrenData)
        With KanrenData(i)
            sql = "INSERT INTO TBCMY023"
            sql = sql & " (CRYNUM,"
            sql = sql & " TRANCNT,"
            sql = sql & " BLOCKID,"
            sql = sql & " PROCCAT,"
            sql = sql & " TXID,"
            sql = sql & " REGDATE,"
            sql = sql & " SUMITFLAG,"
            sql = sql & " SUMITSND,"
            sql = sql & " SSENDNO,"
            sql = sql & " SENDFLAG,"
            sql = sql & " SENDDATE,"
            sql = sql & " PLANTCAT)"
            sql = sql & " VALUES"
            sql = sql & " ('" & .CRYNUM & "',"      '結晶番号
            sql = sql & .TRANCNT & ","              '処理回数
            sql = sql & " '" & .BLOCKID & "',"      'ﾌﾞﾛｯｸID
            sql = sql & " '" & .PROCCAT & "',"      '処理区分
            sql = sql & " '" & .TXID & "',"         'ﾄﾗﾝｻﾞｸｼｮﾝID
            sql = sql & " SYSDATE,"                 '登録日付
            sql = sql & " '0',"                     'SUMIT送信ﾌﾗｸﾞ
            sql = sql & " NULL,"                    'SUMIT送信日付
            sql = sql & " NULL,"                    '送信順連番
            sql = sql & " '0',"                     '送信ﾌﾗｸﾞ
            sql = sql & " NULL,"                    '送信日付
            sql = sql & " '" & sCmbMukesaki & "')"  '向先
        End With

        '登録失敗
        If OraDB.ExecuteSQL(sql) <= 0 Then
            GoTo proc_exit
        End If
    Next i
    
    DBDRV_KanrenBlk = FUNCTION_RETURN_SUCCESS
    
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' ｴﾗｰﾊﾝﾄﾞﾗ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
    
End Function

'概要      :関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)登録
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型               ,説明
'      　　:sCrynum     ,I  ,String         　,結晶番号
'      　　:sKblockid() ,I  ,String         　,関連ﾌﾞﾛｯｸ
'      　　:iSpos       ,I  ,Integer        　,結晶内開始位置
'      　　:iEpos       ,I  ,Integer        　,結晶内終了位置
'      　　:戻り値      ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :07/08/06 ooba
Public Function DBDRV_KanrenBlk_BK(sCryNum As String, sKblockid() As String, _
                                iSpos As Integer, iEpos As Integer) As FUNCTION_RETURN

    Dim sql             As String
    Dim i, j            As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ﾚｺｰﾄﾞ数
    Dim sLotid          As String           'ﾌﾞﾛｯｸID(WFﾏｯﾌﾟ)
    Dim sSXLID          As String           'SXLID(WFﾏｯﾌﾟ)
    Dim KanrenData()    As typ_TBCMY023     '関連ﾌﾞﾛｯｸ紐付紐切ﾃﾞｰﾀ
    Dim bCutFlg         As Boolean          '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ
    Dim iTrnCnt         As Integer          '処理回数


    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function DBDRV_KanrenBlk_BK"

    DBDRV_KanrenBlk_BK = FUNCTION_RETURN_FAILURE

    '処理回数取得
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '処理回数(最大) + 1
    End If
    rs.Close


    lRecCnt = 0             '登録ﾚｺｰﾄﾞ数
    bCutFlg = False         '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ(False:紐切り無)

    '関連ﾌﾞﾛｯｸ紐切ﾃﾞｰﾀｾｯﾄ
    For i = 1 To UBound(sKblockid)
        lRecCnt = lRecCnt + 1
        ReDim Preserve KanrenData(lRecCnt)
        With KanrenData(lRecCnt)
            .CRYNUM = sCryNum               '結晶番号
            .TRANCNT = iTrnCnt              '処理回数
            .BLOCKID = sKblockid(i)         'ﾌﾞﾛｯｸID
            .PROCCAT = "D"                  '処理区分(D:紐切)
            .TXID = "TX879I"                'ﾄﾗﾝｻﾞｸｼｮﾝID
        End With
    Next i


    'WFﾏｯﾌﾟよりﾌﾞﾛｯｸID,SXLIDを取得
    sql = "SELECT LOTID, MSXLID FROM TBCMY011"
    sql = sql & " WHERE LOTID LIKE '" & Left(sCryNum, 9) & "%'"
    sql = sql & " AND (WFSTA = '0' OR WFSTA = '1')"
    sql = sql & " AND RITOP_POS > " & iSpos
    sql = sql & " AND RITOP_POS <= " & iEpos
    sql = sql & " AND MSXLID IS NOT NULL"
    sql = sql & " GROUP BY LOTID, MSXLID"
    sql = sql & " ORDER BY LOTID, MAX(BLOCKSEQ)"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    'ﾃﾞｰﾀなし
    If rs.RecordCount = 0 Then
        rs.Close
        GoTo proc_exit
    End If

    '関連ﾌﾞﾛｯｸ紐付ﾃﾞｰﾀｾｯﾄ
    For i = 1 To rs.RecordCount
        If i > 1 Then
            '別ﾌﾞﾛｯｸで同一SXL(関連ﾌﾞﾛｯｸ○)
            If sLotid <> rs("LOTID") And sSXLID = rs("MSXLID") Then
                '関連ﾌﾞﾛｯｸ(上)
                If KanrenData(lRecCnt).BLOCKID <> sLotid Then
                    iTrnCnt = iTrnCnt + 1       '処理回数
                    lRecCnt = lRecCnt + 1
                    ReDim Preserve KanrenData(lRecCnt)
                    With KanrenData(lRecCnt)
                        .CRYNUM = sCryNum               '結晶番号
                        .TRANCNT = iTrnCnt              '処理回数
                        .BLOCKID = sLotid               'ﾌﾞﾛｯｸID
                        .PROCCAT = "C"                  '処理区分(C:付替え)
                        .TXID = "TX879I"                'ﾄﾗﾝｻﾞｸｼｮﾝID
                    End With
                End If
                '関連ﾌﾞﾛｯｸ(下)
                lRecCnt = lRecCnt + 1
                ReDim Preserve KanrenData(lRecCnt)
                With KanrenData(lRecCnt)
                    .CRYNUM = sCryNum                   '結晶番号
                    .TRANCNT = iTrnCnt                  '処理回数
                    .BLOCKID = rs("LOTID")              'ﾌﾞﾛｯｸID
                    .PROCCAT = "C"                      '処理区分(C:付替え)
                    .TXID = "TX879I"                    'ﾄﾗﾝｻﾞｸｼｮﾝID
                End With

            '別ﾌﾞﾛｯｸで別SXL(関連ﾌﾞﾛｯｸ×)
            ElseIf sLotid <> rs("LOTID") And sSXLID <> rs("MSXLID") Then
                bCutFlg = True          '関連ﾌﾞﾛｯｸ紐切りﾌﾗｸﾞ(True:紐切り有)
            End If
        End If
        sLotid = rs("LOTID")        'ﾌﾞﾛｯｸID
        sSXLID = rs("MSXLID")       'SXLID
        rs.MoveNext
    Next i
    rs.Close


    '関連ﾌﾞﾛｯｸ紐切りが発生した場合、関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)に登録
    If bCutFlg Then
        For i = 1 To UBound(KanrenData)
            With KanrenData(i)
                sql = "INSERT INTO TBCMY023"
                sql = sql & " (CRYNUM,"
                sql = sql & " TRANCNT,"
                sql = sql & " BLOCKID,"
                sql = sql & " PROCCAT,"
                sql = sql & " TXID,"
                sql = sql & " REGDATE,"
                sql = sql & " SUMITFLAG,"               '07/12/21 ooba
                sql = sql & " SUMITSND,"                '07/12/21 ooba
                sql = sql & " SSENDNO,"                 '07/12/21 ooba
                sql = sql & " SENDFLAG,"

                ' 2007/09/03 SPK Tsutsumi Add Start
                sql = sql & " SENDDATE,"
                sql = sql & " PLANTCAT)"
'                sql = sql & " SENDDATE)"
                ' 2007/09/03 SPK Tsutsumi Add End

                sql = sql & " VALUES"
                sql = sql & " ('" & .CRYNUM & "',"      '結晶番号
                sql = sql & .TRANCNT & ","              '処理回数
                sql = sql & " '" & .BLOCKID & "',"      'ﾌﾞﾛｯｸID
                sql = sql & " '" & .PROCCAT & "',"      '処理区分
                sql = sql & " '" & .TXID & "',"         'ﾄﾗﾝｻﾞｸｼｮﾝID
                sql = sql & " SYSDATE,"                 '登録日付
                sql = sql & " '0',"                     'SUMIT送信ﾌﾗｸﾞ  07/12/21 ooba
                sql = sql & " NULL,"                    'SUMIT送信日付  07/12/21 ooba
                sql = sql & " NULL,"                    '送信順連番  07/12/21 ooba
                sql = sql & " '0',"                     '送信ﾌﾗｸﾞ

                ' 2007/09/03 SPK Tsutsumi Add Start
                sql = sql & " NULL,"                    '送信日付
                sql = sql & " '" & sCmbMukesaki & "')"  '向先
'                sql = sql & " NULL)"                    '送信日付
                ' 2007/09/03 SPK Tsutsumi Add End
            End With

            '登録失敗
            If OraDB.ExecuteSQL(sql) <= 0 Then
                GoTo proc_exit
            End If
        Next i
    End If

    DBDRV_KanrenBlk_BK = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

'◆--- 2010/01/20 SIRD対応 SPK habuki ADD START
'---------------------------------------------------------------------------
'概要      :結晶番号上７桁よりTBCMJ022を検索し、SIRD検査情報を返す
'---------------------------------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :pCRYNUM     ,I  　　,String                 ,結晶番号
'          :pflgSird    ,O  　　,Boolean                ,SIRD検査情報の有無(True:有、False：無)
'          :pSMPLID     ,O  　　,String                 ,SIRD検査情報の代表ｻﾝﾌﾟﾙID
'          :戻り値      ,O      ,Boolean                ,[True:OK／False:NG]
'---------------------------------------------------------------------------
Public Function fncGetSirdSample(ByVal pCRYNUM As String, ByRef pflgSird As Boolean, ByRef pSMPLID As String) As Boolean

    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet

    '--ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetSirdSample"
    
    '--初期化
    fncGetSirdSample = False: pflgSird = False: pSMPLID = ""
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL文生成
    sql = "select SMPLNO  "
    sql = sql & "from TBCMJ022 " & vbCrLf
    sql = sql & "where" & vbCrLf
    sql = sql & "     substr(CRYNUM,1,7) = '" & left(pCRYNUM, 7) & "'" & vbCrLf     '結晶番号(上7桁)
    sql = sql & " and TRANCNT   = 0" & vbCrLf                                       '処理回数

    '--ﾃﾞｰﾀを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_READONLY Or ORADYN_NOCACHE)
    If rs Is Nothing Then
        GoTo proc_exit
    End If
    
    '--抽出結果参照
    If Not (rs.EOF) Then
        '<< ﾃﾞｰﾀ有り >>
        rs.MoveFirst
        pflgSird = True                 '[SIRD検査情報有り]
        pSMPLID = rs("SMPLNO")          '[代表ｻﾝﾌﾟﾙID]
    End If
    
    fncGetSirdSample = True

proc_exit:
    '<< 終了 >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< ｴﾗｰﾊﾝﾄﾞﾗ >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'◆--- 2010/01/20 SIRD対応 SPK habuki ADD END

'Add Start 2010/07/08 SMPK Nakamura
'---------------------------------------------------------------------------
'概要      :ブロックIDよりHINBCAを検索し、品番情報を返す
'---------------------------------------------------------------------------
'ﾊﾟﾗﾒｰﾀ    :変数名      ,IO     ,型                     ,説明
'          :pCRYNUM     ,I  　　,String                 ,ブロックID
'          :psHinban    ,O  　　,String                 ,品番
'          :戻り値      ,O      ,Boolean                ,[True:OK／False:NG]
'---------------------------------------------------------------------------
Public Function fncGetMultiHinban(ByVal pCRYNUM As String, ByRef psHinban As String) As Boolean

    Dim sql         As String           'SQL全体
    Dim rs          As OraDynaset       'RecordSet
    Dim i           As Integer          'ループカウント
    Dim sBlockId()    As String

    '--ｴﾗｰﾊﾝﾄﾞﾗの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc036_SQL.bas -- Function fncGetMultiHinban"
    
    sBlockId = Split(pCRYNUM, Chr(9))
    
    '--初期化
    fncGetMultiHinban = False
    Set rs = Nothing      'Oracle RecordSet Free

    '--SQL文生成
    sql = "select distinct HINBCA from ("
    sql = sql & "select HINBCA "
    sql = sql & "from XSDCA " & vbCrLf
    sql = sql & "where" & vbCrLf
    sql = sql & "     CRYNUMCA in ( "
    If UBound(sBlockId) > 1 Then
        For i = 0 To UBound(sBlockId) - 1
            If InStr(sBlockId(i), "Wait") > 0 Then
                sql = sql & "'" & Trim(Mid(sBlockId(i), 1, InStr(sBlockId(i), "Wait") - 1)) & "' " & vbCrLf
            Else
                sql = sql & "'" & sBlockId(i) & "' " & vbCrLf
            End If
            If i <> UBound(sBlockId) - 1 Then sql = sql & ","
        Next i
    Else
        sql = sql & "'" & pCRYNUM & "' " & vbCrLf
    End If
    sql = sql & ") "
    sql = sql & "     and LIVKCA = '0' " & vbCrLf                '生死フラグ
    sql = sql & "order by INPOSCA)" & vbCrLf

    '--ﾃﾞｰﾀを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Or rs.RecordCount = 0 Then
        GoTo proc_exit
    End If
    
    '--抽出結果参照
    psHinban = ""
    For i = 1 To rs.RecordCount
        psHinban = psHinban & rs("HINBCA") & vbTab  '品番
        rs.MoveNext
    Next i
    
    fncGetMultiHinban = True

proc_exit:
    '<< 終了 >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing
    
    gErr.Pop
    Exit Function

proc_err:
    '<< ｴﾗｰﾊﾝﾄﾞﾗ >>
    'Oracle RecordSet Free
    If Not (rs Is Nothing) Then
        rs.Close
    End If
    Set rs = Nothing

    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    
    gErr.HandleError
    Resume proc_exit

End Function
'Add End 2010/07/08 SMPK Nakamura
