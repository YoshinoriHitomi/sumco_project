Attribute VB_Name = "s_cmbc017_SQL"
Option Explicit

' 結晶研削加工実績入力

' 製品仕様
Public Type typ_HinSpec2
    HIN As tFullHinban          ' 品番
    HSXTYPE As String * 1       ' タイプ
    HSXCDIR As String * 1       ' 方位
    HSXD1CEN As Double          ' 直径
    
    HSXDPDIR As String * 2      ' ノッチ位置
    HSXDPMIN As Double          ' ノッチ角度の判定(Notch位置規格 下限)　2009/09 SUMCO Akizuki
    HSXDPMAX As Double          ' ノッチ角度の判定(Notch位置規格 上限)　2009/09 SUMCO Akizuki
    
    HSXDPAMN As Integer         ' ノッチ角度(下限)
    HSXDPAMX As Integer         ' ノッチ角度(上限)
    HSXDPACN As Integer         ' ノッチ角度(中心) 2005/08
    
    HSXDWMIN As Double          ' ノッチ幅(下限)
    HSXDWMAX As Double          ' ノッチ幅(上限)
    
    HSXDDMIN As Double          ' ノッチ深さ(下限)
    HSXDDMAX As Double          ' ノッチ深さ(上限)
    HSXDDCEN As Double          ' ノッチ深さ(中心)  2005/08
End Type

' 待ち一覧
Public Type typ_DispData
    CRYNUM As String
    HIN As String
    DIAK As String
    GNDAY As Date
    GNL As Double
    GNW As Double
    PRIORITY As String
    PUPTN As String
    NOUKI As Date
    MUKE As String
    HLDCAUSE As String
    HOLDKT As String
    BIKOU As String
    HLDCMNT As String
    HLDTRCLS As String
    PUHINB As String   '2005/10
    XTALCA As String   '2005/10
    RPCRYNUMCA As String   '2005/10
    NEWKNTCA As String   '2005/11
    KIKBN    As String  '期判別区分 2006/11/14 SETsw kubota
    PLANTCATCA As String    '向先 2007/08/21 SPK Tsutsumi Add
    DPDIR   As String   'ノッチ位置方位 2008/01/09
    AGRSTATUS As String             ' 承認確認区分 add SETkimizuka
    STOP    As String               ' 停止 add SETkimizuka
    CAUSE   As String               ' 停止理由 add SETkimizuka
    PRINTNO As String               ' 先行評価 add SETkimizuka
End Type

' 2007/08/17 SPK Tsutsumi Add Start
Public Type typ_Mukesaki
    sMukeCode As String     '' 向先コード
    sMukeName As String     '' 向先名
End Type

Public s_Mukesaki() As typ_Mukesaki
' 2007/08/17 SPK Tsutsumi Add End

'' ストッカ対応 2006/11/08 SETsw J.W -->
Public gsGrTim As String     ' 研削時間
Public gsNchLength As String 'ノッチ長さ
'' ストッカ対応 2006/11/08 SETsw J.W -->


'概要      :結晶研削加工実績入力用 結晶番号入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'　　      :sCryNum　　　,I  ,String         　,結晶番号
'　　      :pCryInf　　　,O  ,typ_TBCME037   　,結晶情報
'　　      :pHinDsn　　　,O  ,typ_TBCME039   　,品番設計
'　　      :pCutIns　　　,O  ,typ_TBCME045   　,切断指示
'　　      :pProcBR　　　,O  ,typ_TBCMI001   　,加工払出実績
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'　　      :戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001c_Disp(ByVal sCryNum As String, ByVal sBlockId As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pCutIns() As typ_TBCME045, _
                                           pProcBR As typ_TBCMI001, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpProcBR() As typ_TBCMI001
    Dim rs  As OraDynaset
    Dim rs2 As OraDynaset   ' add 2006/11/09 SETsw J.W
    Dim sql As String
    Dim sDbName As String
    Dim recCnt As Long
    Dim i As Long
    Dim sans As String
    Dim sNowproc As String
    '2004.09.08 Y.K 紐付け変更
    Dim sSijiNo As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_Disp"
    sErrMsg = ""

    '2004.09.08 Y.K 紐付け変更  <=== START
    sDbName = "XSDC1"
    sSijiNo = ""
    sSijiNo = F_Get_SijiNoGet(sCryNum)
    '2004.09.08 Y.K 紐付け変更  == > END

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' 工程チェック
    If DBDRV_get_xGR(sCryNum, sans) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET")
        Exit Function
    End If
    If sans = "A" Then
        'AGRの場合
        If pCryInf.PROCCD <> PROCD_KENNSAKU_KAKOU Then
            sErrMsg = GetMsgStr("EPRC0")
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Else
        'MGRの場合
        If GetTBCME040_NOWPROC(sBlockId, sNowproc) = FUNCTION_RETURN_FAILURE Then
            sDbName = "E040"
            sErrMsg = GetMsgStr("ENG11", sDbName)
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

'        If (pCryInf.PROCCD <> PROCD_KESSYOU_SOUGOUHANTEI) And (pCryInf.PROCCD <> PROCD_SETUDAN) Then
        If (sNowproc <> PROCD_KENNSAKU_KAKOU) Then    '仕掛かり工程ﾁｪｯｸ変更　2002/11/28
            sErrMsg = GetMsgStr("EPRC2")
            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If

    '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
    sDbName = "E039"
    '2004.09.08 Y.K 紐付け変更
'    sql = " where substr(CRYNUM,1,7)='" & Left(sCrynum, 7) & "' order by INGOTPOS"
    sql = " where substr(CRYNUM,1,9)='" & sSijiNo & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinDsn) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 切断指示の取得 ----> 2005/11 XSDCAに変更
    'sDbName = "E045"
    'sql = "select "
    'sql = sql & "CRYNUM, INGOTPOS, TRANCNT "
    'sql = sql & " from TBCME045 T1"
    'sql = sql & " where CRYNUM='" & sCryNum & "'"
    'sql = sql & " and TRANCNT=any(select max(TRANCNT) from TBCME045 T2 where CRYNUM='" & sCryNum & "'"
    'sql = sql & " and T1.INGOTPOS=T2.INGOTPOS ) "
    'Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    'recCnt = rs.RecordCount
    'If recCnt = 0 Then
    '    rs.Close
    '    sErrMsg = GetMsgStr("EGET2", sDbName)
    '    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    '    GoTo proc_exit
    'End If

    'ReDim pCutIns(recCnt)
    'For i = 1 To recCnt
    '    With pCutIns(i)
    '        .CRYNUM = rs("CRYNUM")          ' 結晶番号
    '        .INGOTPOS = rs("INGOTPOS")      ' 結晶内開始位置
    '        .TRANCNT = rs("TRANCNT")        ' 処理回数
    '    End With
    '    rs.MoveNext
    'Next i
    'rs.Close

    'For i = 1 To recCnt
    '    With pCutIns(i)
    '        sql = "select "
    '        sql = sql & "LENGTH, HINBAN, REVNUM, FACTORY, OPECOND"
    '        sql = sql & " from TBCME045"
    '        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    '        sql = sql & " and INGOTPOS=" & .INGOTPOS
    '        sql = sql & " and TRANCNT=" & .TRANCNT
    '        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '        If rs.RecordCount = 0 Then
    '            rs.Close
    '            sErrMsg = GetMsgStr("EGET2", sDbName)
    '            DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    '            GoTo proc_exit
    '        End If
    '        .LENGTH = rs("LENGTH")      ' 長さ
    '        .hinban = rs("HINBAN")      ' 品番
    '        .REVNUM = rs("REVNUM")      ' 製品番号改訂番号
    '        .factory = rs("FACTORY")    ' 工場
    '        .opecond = rs("OPECOND")    ' 操業条件
    '    End With
    '    rs.Close
    'Next i
    '’↓更新 SPT用実績作成方法変更 2006/04/17 SMP松田

    sDbName = "XSDCZ"
    sql = "select DISTINCT "
    sql = sql & "HINBCZ, REVNUMCZ, FACTORYCZ, OPECZ "
    sql = sql & " from XSDCZ "
    sql = sql & " where RPCRYNUMCZ ='" & sBlockId & "'"
    sql = sql & " and GNWKNTCZ = 'CC400'"

'    sDbName = "XSDCA"
'    sql = "select DISTINCT "
'    sql = sql & "HINBCA, REVNUMCA, FACTORYCA, OPECA "
'    sql = sql & " from XSDCA "
'    sql = sql & " where RPCRYNUMCA ='" & sCrynum & "'"
'    sql = sql & " and GNWKNTCA = 'CC400'"

    '’↑更新 SPT用実績作成方法変更 2006/04/17 SMP松田
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ReDim pCutIns(recCnt)
    For i = 1 To recCnt
        With pCutIns(i)
        '’↓更新 SPT用実績作成方法変更 2006/04/17 SMP松田
            .hinban = rs("HINBCZ")      ' 品番
            .REVNUM = rs("REVNUMCZ")    ' 製品番号改訂番号
            .factory = rs("FACTORYCZ")  ' 工場
            .opecond = rs("OPECZ")      ' 操業条件

'            .hinban = rs("HINBCA")      ' 品番
'            .REVNUM = rs("REVNUMCA")      ' 製品番号改訂番号
'            .factory = rs("FACTORYCA")    ' 工場
'            .opecond = rs("OPECA")    ' 操業条件
        '’↑更新 SPT用実績作成方法変更 2006/04/17 SMP松田
        End With
        rs.MoveNext
    Next i
    rs.Close
    '' 加工払出実績の取得(s_cmzcTBCMI001_SQL.bas が必要)
    sDbName = "I001"
    sql = " where CRYNUM='" & sCryNum & "'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT)"
    sql = sql & " from TBCMI001 where CRYNUM='" & sCryNum & "')"
    If DBDRV_GetTBCMI001(tmpProcBR(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpProcBR) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pProcBR = tmpProcBR(1)

    '' ストッカ対応 2006/11/09 SETsw J.W -->
    '' 研削時間･ノッチの取得
    gsNchLength = ""
    gsGrTim = ""

    sql = "      select UPDDATE"
    sql = sql & "     , -1 NOTCH"  '<== 該当カラムが無いため空欄 (2006/11/09 J.W)
    sql = sql & "     , TO_CHAR(CYGRTIM,'FM000000') GRTIM"
    sql = sql & "  from TBCMF002"
    sql = sql & " where INGOTNO='" & sCryNum & "'"
    sql = sql & "   and TRANCNT=(select MAX(TRANCNT) from TBCMF002 where INGOTNO='" & sCryNum & "')" & vbCrLf
    sql = sql & " union "
    sql = sql & "select NVL(UPDDATE,REGDATE) UPDDATE"   '新テーブル更新日付がNULLの場合、登録日付
    sql = sql & "     , TRWLEN NOTCH"
    sql = sql & "     , TO_CHAR(CYGRTIM,'FM000000') GRTIM"
    sql = sql & "  from TBCMF010"
    sql = sql & " where CRYNUM='" & sBlockId & "'"
    sql = sql & "   and PROCNUM=(select MAX(PROCNUM) from TBCMF010 where CRYNUM='" & sBlockId & "')"
    sql = sql & " order by UPDDATE desc"

    Set rs2 = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)

    If rs2.RecordCount > 0 Then
        gsNchLength = NulltoStr(rs2("NOTCH").Value)
        '' ノッチ長さがマイナス(TBCMF002から取得した場合は空欄にする) 2006/11/10 SETsw J.W
        If (val(gsNchLength) < 0) Then
            gsNchLength = ""
        End If
        gsGrTim = Format(NulltoStr(rs2("GRTIM").Value), "@@:@@:@@")
        If Left(gsGrTim, 1) = "0" Then
            gsGrTim = Mid(gsGrTim, 2)
        End If
    End If
    rs2.Close
    '' ストッカ対応 2006/11/09 SETsw J.W <--

    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmic001c_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2004.09.08　Y.K　紐付け変更
'指定結晶番号の指示No取得処理
'ＸＳＤＣ１の指示を取得する
'但し、取得できない場合は、結晶番号７桁＋’０’＋結晶番号９桁を返す
Private Function F_Get_SijiNoGet(sCryNum As String) As String
  Dim sSql As String
  Dim rs As OraDynaset    'RecordSet

    sSql = ""
    sSql = sSql & "SELECT"
    sSql = sSql & "  hisijiC1 "
    sSql = sSql & "FROM"
    sSql = sSql & "  XSDC1 C1 "
    sSql = sSql & "WHERE"
    sSql = sSql & "  substr(C1.XTALC1,1,9) = '" & Mid(sCryNum, 1, 9) & "'"

    Set rs = OraDB.DBCreateDynaset(sSql, ORADYN_DEFAULT)

    If (rs.RecordCount = 1) Then
        If (IsNull(rs.Fields("hisijiC1")) = False) Then
            F_Get_SijiNoGet = rs.Fields("hisijiC1")
        Else
            F_Get_SijiNoGet = Mid(sCryNum, 1, 7) & "0" & Mid(sCryNum, 9, 1)
        End If
    Else
        F_Get_SijiNoGet = Mid(sCryNum, 1, 7) & "0" & Mid(sCryNum, 9, 1)
    End If

End Function


'概要      :結晶研削加工実績入力用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'　　      :pHinSpec　　　,IO ,typ_HinSpec2   　,製品仕様
'　　      :戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001c_GetSpec(pHinSpec As typ_HinSpec2) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_GetSpec"

    '' 製品仕様の取得
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDPDIR, HSXDPMIN, HSXDPMAX, "
    sql = sql & "HSXDWMIN, HSXDWMAX, HSXDDMIN, HSXDDMAX, HSXDDCEN, HSXDPACN "  '2009/09
    sql = sql & " from TBCME018"
    sql = sql & " where HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and OPECOND='" & pHinSpec.HIN.opecond & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
'NULL対応 ----- START ----- 2003/12/10
    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")                          ' タイプ
        .HSXCDIR = rs("HSXCDIR")                          ' 方位
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))          ' 直径
        .HSXDPDIR = rs("HSXDPDIR")                        ' ノッチ位置
        .HSXDWMIN = fncNullCheck(rs("HSXDWMIN"))          ' ノッチ幅(下限)
        .HSXDWMAX = fncNullCheck(rs("HSXDWMAX"))          ' ノッチ幅(上限)
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' ノッチ深さ(下限)
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' ノッチ深さ(上限)
        .HSXDDCEN = fncNullCheck(rs("HSXDDCEN"))          ' ノッチ深さ(中心) 2005/08
        .HSXDPACN = fncNullCheck(rs("HSXDPACN"))          ' ノッチ角度(中心) 2005/08
        
        '値に「-1」もあるため、Nullの場合は[999]を返す     ' ノッチ位置(下限) 2009/09 Akizuki
        If IsNull(rs("HSXDPMIN")) Then
            .HSXDPMIN = 999
        Else
            .HSXDPMIN = rs("HSXDPMIN")
        End If
        
        '値に「-1」もあるため、Nullの場合は[999]を返す   　' ノッチ位置(上限) 2009/09 Akizuki
        If IsNull(rs("HSXDPMAX")) Then
            .HSXDPMAX = 999
        Else
            .HSXDPMAX = rs("HSXDPMAX")
        End If
    End With
    
    rs.Close
'NULL対応 -----  END  ----- 2003/12/10

    DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001c_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :結晶研削加工実績入力用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:sCryNum　　　,I  ,String         　,結晶番号
'      　　:pPlshPR　　　,I  ,typ_TBCMI002   　,研削加工実績
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DBDRV_scmzc_fcmic001c_Exec(sCryNum As String, pPlshPR As typ_TBCMI002, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sans As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001c_Exec"
    sErrMsg = ""

    '工程コード設定ロジック統一   2002/11/27 tuku START
    'AGR MGR のチェック
    If DBDRV_get_xGR(sCryNum, sans) = FUNCTION_RETURN_FAILURE Then
        sDbName = "I001"
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        Exit Function
    End If
    If sans = "A" Then
        'AGRの場合結晶情報（TBCME037)のみ更新
        '' 結晶情報の更新
        sDbName = "E037"
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & MGPRCD_SETUDAN & "', "
        sql = sql & "PROCCD='" & nextCd & "', "
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
        sql = sql & "LASTPASS='" & nowCd & "', "
        sql = sql & "DIAMETER=" & pPlshPR.DMTOP1 & ", "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & sCryNum & "'"
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    ElseIf sans = "M" Then
        'MGRの場合結晶情報（TBCME037)とブロック管理（TBCME040)を更新
        '' 結晶情報の更新
        sDbName = "E037"
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "
        sql = sql & "PROCCD='" & nextCd & "', "
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
        sql = sql & "LASTPASS='" & nowCd & "', "
        sql = sql & "DIAMETER=" & pPlshPR.DMTOP1 & ", "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & sCryNum & "'"
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        ''ブロック管理の更新
        '' ブロック管理テーブルの更新
        sDbName = "E040"
        sql = "update TBCME040 set "
        sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "      ' 現在管理工程
        sql = sql & "NOWPROC='" & nextCd & "', "                        ' 現在工程
        sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "                ' 最終通過管理工程
        sql = sql & "LASTPASS='" & nowCd & "', "                        ' 最終通過工程
        sql = sql & "UPDDATE=sysdate "                                 ' 更新日付
        sql = sql & "where CRYNUM='" & sCryNum & "' "
        sql = sql & "and  BLOCKID='" & pPlshPR.CRYNUM & "' "
        sql = sql & "and INGOTPOS=" & pPlshPR.INGOTPOS
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '                           2002/11/27 tuku END



    '' 研削加工実績の挿入
    sDbName = "I002"
    With pPlshPR
        sql = "insert into TBCMI002 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, "
        sql = sql & "DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH, "
        sql = sql & "BDLNTOP, BDCDTOP, BDLNTAIL, BDCDTAIL, "
        sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, INGOTPOS, LENGTH, "
        sql = sql & "GOUKI, NCHWTAIL ,BLOCKID, "             '2006/02/01 tuku
        sql = sql & "NCHLENGTH, CYGRTIM, NCHANGLE )"          ' ストッカ対応 2006/11/09 SETsw J.W
        sql = sql & " select '"
        sql = sql & sCryNum & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', "
        sql = sql & .DMTOP1 & ", "
        sql = sql & .DMTOP2 & ", "
        sql = sql & .DMTAIL1 & ", "
        sql = sql & .DMTAIL2 & ", '"
        sql = sql & .NCHPOS & "', "
        sql = sql & .NCHDPTH & ", "
        sql = sql & .NCHWIDTH & ", "
        sql = sql & .BDLNTOP & ", '"
        sql = sql & .BDCDTOP & "', "
        sql = sql & .BDLNTAIL & ", '"
        sql = sql & .BDCDTAIL & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate,"
        sql = sql & .INGOTPOS & ", "
        sql = sql & .LENGTH & " ,  "
        sql = sql & .GOUKI & " , "                      '2003/06/12 osawa 号機追加
        sql = sql & .NCHWTAIL & ", '"                  '2004/05/25
        sql = sql & .BLOCKID & "',"                    '2006/02/01 tuku ﾌﾞﾛｯｸID追加
        sql = sql & .NCHLENGTH & ", "                  ' ストッカ対応 2006/11/09 SETsw J.W
        sql = sql & .CYGRTIM & ", "                    ' ストッカ対応 2006/11/09 SETsw J.W
        sql = sql & .NCHANGLE                          ' 2009/09 SUMCO Akizuki Notch角度追加"
        sql = sql & " from TBCMI002"
        sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .INGOTPOS
    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_scmzc_fcmic001c_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'概要      :INGOTPOS,LENGTHの取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :BLOCKID        ,I   ,String            ,結晶番号orブロックID
'          :iIngotpos      ,O   ,Integer           ,結晶内開始位置
'          :iLength        ,O   ,Integer           ,長さ
'      　　:戻り値          , O  , FUNCTION_RETURN　, 読み込みの成否
'説明      :
'履歴      :2002/04/17 佐野 信哉 作成
Public Function scmzc_getIngotposLength(BLOCKID As String, iIngotpos As Integer, iLength As Integer) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim AGRFlag As Boolean
    Dim Ans As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function scmzc_getIngotposLength"
    scmzc_getIngotposLength = FUNCTION_RETURN_FAILURE

    '引き上げ結晶の場合
    '加工払い出し実績からAGRかMGRかを求める
    If DBDRV_get_xGR(BLOCKID, Ans) = FUNCTION_RETURN_FAILURE Then
        GoTo proc_exit
    End If
    AGRFlag = (Trim(Ans) = "A")

    If AGRFlag Then
        'AGRの場合
        'INGOTPOS=0で加工実績から実績を求める
        sql = "select UPLENGTH from TBCMI001 "
        sql = sql & "where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "' and "
        sql = sql & "TRANCNT = any(select max(TRANCNT) from TBCMI001 where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "')"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            scmzc_getIngotposLength = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        iIngotpos = 0
        iLength = rs("UPLENGTH")
        rs.Close
    Else
        'MGRの場合
        'ブロック管理からINGOTPOSを求める
        'そのブロックの初回切断時のINGOTPOSを求める
        sql = "select INGOTPOS,LENGTH from TBCME040 "
        sql = sql & "where CRYNUM = '" & Left(BLOCKID, 9) & "000" & "' and "
        sql = sql & "BLOCKID = '" & BLOCKID & "' "
        sql = sql & "order by INGOTPOS"

        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            GoTo proc_exit
        End If
        iIngotpos = rs("INGOTPOS")
        iLength = rs("LENGTH")
        rs.Close

    End If

    scmzc_getIngotposLength = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getIngotposLength = FUNCTION_RETURN_FAILURE
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
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

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
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

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

'概要      :テーブル「TBCMI001」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMI001 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMI001_SQL.basより移動)
Public Function DBDRV_GetTBCMI001(records() As typ_TBCMI001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, TRANCNT, KRPROCCD, PROCCODE, UPLENGTH, FREELENG, UPWEIGHT, SEED, PRCMCN, TSTAFFID, REGDATE," & _
              " KSTAFFID, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMI001"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMI001 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .TRANCNT = rs("TRANCNT")         ' 処理回数
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
            .FREELENG = rs("FREELENG")       ' フリー長
            .UPWEIGHT = rs("UPWEIGHT")       ' 引上げ重量
            .SEED = rs("SEED")               ' シード
            .PRCMCN = rs("PRCMCN")           ' 研削機
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMI001 = FUNCTION_RETURN_SUCCESS
End Function

'
'概要    : TBCME040よりフィールド値LENGTHの取得
'ﾊﾟﾗﾒｰﾀ  :変数名        ,IO  ,型                                     ,説明
'
'        :戻ﾘ値         ,O   ,FUNCTION_RETURN                        ,読み込み成否
'説明    :
'履歴    :2002.8 追加 H.Kakizawa    2005/11 C2に変更
Public Function GetTBCME040_NOWPROC(ByVal pBlockid As String, pNowproc As String) As FUNCTION_RETURN

    Dim sql As String
    Dim rs As OraDynaset

    '初期値
    GetTBCME040_NOWPROC = FUNCTION_RETURN_FAILURE

    '引上実績の加工区分を取得
    sql = ""
    'sql = sql & "select NOWPROC from TBCME040 "
    'sql = sql & "where BLOCKID = '" & pBlockid & "' "
    sql = sql & "select GNWKNTC2 from XSDC2 "
    sql = sql & "where CRYNUMC2 = '" & pBlockid & "' "

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP) 'データを抽出する

    If rs.RecordCount = 0 Then 'レコードがない場合は否
        rs.Close
        Exit Function
    Else
        'pNowproc = rs.Fields("NOWPROC")
        pNowproc = rs.Fields("GNWKNTC2")
        rs.Close
    End If

    GetTBCME040_NOWPROC = FUNCTION_RETURN_SUCCESS

End Function

'概要      :加工払出一覧用 画面表示時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pCutMap　　　,O  ,typ_CutMap     　,切断指示一覧
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001d_Disp(pDispData() As typ_DispData, pCrynum As String, pHinb As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Long
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_scmzc_fcmic001d_Disp"
    sql = ""
    'sql = sql & "SELECT PUHINBC1, PUPTNC1, DIA1C1, XTALCA, CRYNUMCA, HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA as HINBAN12, GNLCA, GNWCA, GNDAYCA "
    sql = sql & "SELECT PUHINBC1, PUPTNC1, DIA1C1, XTALCA, CRYNUMCA, HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA as HINBAN12,  "
''↓削除 START SPT用実績作成方法変更 2006/05/18 SMP-OKAMOTO
    sql = sql & " GNLCA, GNWCA, GNDAYCA, CRYNUMCA as RPCRYNUMCA, NEWKNTCA "
'    sql = sql & " GNLCA, GNWCA, GNDAYCA, RPCRYNUMCA, NEWKNTCA "
''↑削除 END   SPT用実績作成方法変更 2006/05/18 SMP-OKAMOTO
    sql = sql & "  ,HOLDBC2, HOLDCC2, HOLDKTC2 "
    sql = sql & "  ,PLANTCATCA"     ' 2007/08/21 SPK Tsutsumi Add
    sql = sql & " , HSXDPDIR"       ' 2008/01/09 ノッチ位置方位
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
    ' 流動停止項目追加 add SETkimizuka Start  09/03/25
    'sql = sql & " , NVL(TO_CHAR(Y4.AGRSTATUS),' ') as AGRSTATUS "
    'sql = sql & " , NVL(TO_CHAR(Y4.STOP),'0') as STOP "
    'sql = sql & " , NVL(Y4.CAUSE,' ') as CAUSE "
    'sql = sql & " , NVL(Y4.PRINTKIND || Y4.PRINTNO,' ') as PRINTNO "
    ' 流動停止項目追加 add SETkimizuka End    09/03/25
    sql = sql & " , NVL(TO_CHAR(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",'" & WF_KBN & "'," & WF_SORT & ",AGRSTATUSY4)),' ') as AGRSTATUS "
    sql = sql & " , NVL(TO_CHAR(Y4.STOPY4),' ') as STOP "
    sql = sql & " , DECODE(TRIM(Y4.CAUSEY4),NULL,' ',TRIM(Y4.CAUSEY4) || ':' || NAMEJA9) as CAUSE "
    sql = sql & " , NVL(Y4.PRINTKINDY4 || Y4.PRINTNOY4,' ') as PRINTNO "
    sql = sql & " , NVL(Y4.WKKTY4,'0') as WKKTY4 "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
    sql = sql & " from XSDC1, XSDCA, XSDC2 "
    sql = sql & " 　　,TBCME018 "   ' 2008/01/09 ノッチ位置方位
    ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
    sql = sql & "    ,XODY3,XODY4 Y4,KODA9  "
    ' 流動停止項目追加 add SETkimizuka Start  09/03/25
    'sql = sql & "    ,(SELECT XTALNOY3 as XTALNO,MIN(DECODE(AGRSTATUSY4,'" & SYONIN_KBN & "'," & SYONIN_SORT & ",'" & KAKUNIN_KBN & "'," & KAKUNIN_SORT & ",'" & SIJI_KBN & "'," & SIJI_SORT & ",'" & VB_KBN & "'," & VB_SORT & ",AGRSTATUSY4)) as AGRSTATUS  "
    'sql = sql & "      ,MAX(STOPY4) as STOP,DECODE(TRIM(CAUSEY4),'',TRIM(CAUSEY4),TRIM(CAUSEY4) || ':' || TRIM(NAMEJA9)) as CAUSE ,Y5.PRINTNO,Y5.PRINTKIND "
    'sql = sql & "      FROM XODY3  "
    'sql = sql & "           LEFT OUTER JOIN XODY4 on ( XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND STOPY4 <> '2'  AND WKKTY4 in " & CreateWkktSQL(WATCH_PROCCD) & ") "
    'sql = sql & "           LEFT OUTER JOIN KODA9 on ( SYSCA9 = 'X' AND SHUCA9 = '30' AND CAUSEY4 = CODEA9 ) "
    'sql = sql & "           LEFT OUTER JOIN (SELECT XTALNOY4 as XTALNO,SXLIDY4 as SXLID,PRINTNOY5 as PRINTNO,PRINTKINDY5 as PRINTKIND "
    'sql = sql & "                FROM XODY3,XODY4,XODY5 "
    'sql = sql & "              WHERE XTALNOY3 = XTALNOY4 AND RCNTY3 = RCNTY4 AND LIVKY3 = '0' "
    'sql = sql & "                AND PRINTKINDY4 = PRINTKINDY5 AND PRINTNOY4 = PRINTNOY5  "
    'sql = sql & "                AND HKBNY5 ='0' GROUP BY XTALNOY4,SXLIDY4,PRINTNOY5,PRINTKINDY5) Y5 ON (XTALNOY3 = XTALNO AND SXLIDY3 = SXLID ) "
    'sql = sql & "      WHERE  "
    'sql = sql & "       LIVKY3    = '0' "
    'sql = sql & "       GROUP BY XTALNOY3,AGRSTATUSY4,CAUSEY4,Y5.PRINTNO,Y5.PRINTKIND,NAMEJA9) Y4 "
    ' 流動停止項目追加 add SETkimizuka End  09/03/25
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
    sql = sql & " where XTALC1 = XTALC2 "
    sql = sql & " AND CRYNUMCA = CRYNUMC2 "
    sql = sql & " AND CRYNUMCA LIKE '" & pCrynum & "%'"
    sql = sql & " AND GNWKNTCA = 'CC400'  "
    sql = sql & " AND LIVKCA = '0'  "
    sql = sql & " AND HINBCA||LTRIM(to_char(REVNUMCA,'00'))||FACTORYCA||OPECA  LIKE '" & pHinb & "%'"
    sql = sql & " AND HINBCA = HINBAN "         ' 2008/01/09 ノッチ位置方位
    sql = sql & " AND REVNUMCA = MNOREVNO "     ' 2008/01/09 ノッチ位置方位
    sql = sql & " AND FACTORYCA = FACTORY "     ' 2008/01/09 ノッチ位置方位
    sql = sql & " AND OPECA = OPECOND "         ' 2008/01/09 ノッチ位置方位
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
    'sql = sql & " AND CRYNUMCA    = Y4.XTALNO(+) "            'add 09/03/25 SETkimizuka
    sql = sql & " AND CRYNUMCA = XTALNOY3(+) "
    sql = sql & " AND LIVKY3(+) = '0' "
    sql = sql & " AND LIVKY4(+) = '0' "
    sql = sql & " AND XTALNOY3 = XTALNOY4(+) "
    sql = sql & " AND RCNTY3 = RCNTY4(+) "
    sql = sql & " AND SYSCA9(+) = 'X' AND SHUCA9(+) = '30' AND CAUSEY4 = CODEA9(+) "
    ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
    sql = sql & " order by CRYNUMCA,INPOSCA "
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
    End If

    ReDim pDispData(recCnt)
    For i = 1 To recCnt
        With pDispData(i)
            .CRYNUM = rs("CRYNUMCA")  '
            .HIN = rs("HINBAN12")
            .PUPTN = rs("PUPTNC1")
            .GNL = rs("GNLCA")
            .GNW = rs("GNWCA")
            .GNDAY = rs("GNDAYCA")
            .DIAK = rs("DIA1C1")
            .HLDCAUSE = rs("HOLDCC2")
            .HLDTRCLS = rs("HOLDBC2")
            If IsNull(rs("HOLDKTC2")) = False Then .HOLDKT = rs("HOLDKTC2")
            .PUHINB = rs("PUHINBC1")  '2005/10
            .XTALCA = rs("XTALCA")
            If IsNull(rs("RPCRYNUMCA")) = False Then .RPCRYNUMCA = rs("RPCRYNUMCA") '2005/10
            .NEWKNTCA = rs("NEWKNTCA")
            If IsNull(rs("PLANTCATCA")) = False Then .MUKE = rs("PLANTCATCA")  ' 2007/08/21 SPK Tsutsumi Add
            .DPDIR = rs("HSXDPDIR")     '2008/01/09  ノッチ位置方位
            ' 流動停止項目追加 add SETkimizuka Start  09/03/25
            ' 流動監視SQL修正 upd SETkimizuka Start  09/06/29
            '.STOP = rs("STOP")                   '停止区分
            '.AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
            'If Trim(rs("CAUSE")) <> "" Then
            '    .CAUSE = rs("CAUSE") & vbTab       '停止理由
            'End If
            If rs("STOP") <> "2" And rs("WKKTY4") = "CC400" Then
                .STOP = rs("STOP")                   '停止区分
                .AGRSTATUS = rs("AGRSTATUS")       '承認確認区分
                If Trim(rs("CAUSE")) <> "" Then
                    .CAUSE = rs("CAUSE") & vbTab       '停止理由
                End If
            End If
            ' 流動監視SQL修正 upd SETkimizuka End  09/06/29
            If Trim(rs("PRINTNO")) <> "" Then
                .PRINTNO = rs("PRINTNO") & vbTab       '先行評価
            End If
            ' 流動停止項目追加 add SETkimizuka End    09/03/25
        End With
        rs.MoveNext
    Next i
    rs.Close
    DBDRV_scmzc_fcmic001d_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001d_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

Public Function DBDRV_SELECT_HOLD(pTblDispData As typ_DispData) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim sCryNum As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_HOLD"

    With pTblDispData

        sCryNum = Left(.CRYNUM, 9) & "000"
        ''SQLを組み立てる
        sql = "SELECT HLDCMNT FROM TBCMJ012 "
        sql = sql & " WHERE CRYNUM = '" & sCryNum & "'"
        'sql = sql & " AND   XTALC2 = CRYNUM   "
        'sql = sql & " AND   INPOSC2 = INGOTPOS   "
        'sql = sql & " AND TRANCNT = (SELECT MAX(TRANCNT) FROM TBCMJ012 WHERE CRYNUM = '" & pCrynum & "')"
        sql = sql & " ORDER BY TRANCNT"

        'データを抽出する
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

        If rs Is Nothing Then
            DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        If rs.RecordCount > 0 Then
           rs.MoveLast
            If IsNull(rs("HLDCMNT")) = False Then .HLDCMNT = rs("HLDCMNT")
        End If
    End With
    rs.Close

    DBDRV_SELECT_HOLD = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_SELECT_HOLD = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function
Public Function DBDRV_SELECT_BLOCK(pBlockData() As typ_XSDC2, sXtal As String, sMgr As String) As FUNCTION_RETURN

    Dim sql As String       'SQL全体
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long
    Dim sCryNum As String

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_BLOCK"

        ''SQLを組み立てる
        '’↓更新 SPT用実績作成方法変更 2006/04/17 SMP松田
        '' 2007/08/22 SPK Tsutsumi Add Start
        sql = "SELECT CRYNUMCZ, GNLCZ, GNWCZ, INPOSCZ, PLANTCATCZ FROM XSDCZ "
'        sql = "SELECT CRYNUMCZ, GNLCZ, GNWCZ, INPOSCZ FROM XSDCZ "
        '' 2007/08/22 SPK Tsutsumi Add Start
        sql = sql & " WHERE RPCRYNUMCZ = '" & sXtal & "'"
        sql = sql & " AND LIVKCZ = '0'"
        sql = sql & " AND GNWKNTCZ = 'CC400'"
''↓更新 START SPT用実績作成方法変更 2006/05/25 SMP-OKAMOTO
        ''結晶内位置でソートする
        sql = sql & " ORDER BY INPOSCZ "
'        sql = sql & " ORDER BY CRYNUMCZ "
''↑更新 END   SPT用実績作成方法変更 2006/05/25 SMP-OKAMOTO

'        sql = "SELECT CRYNUMC2, GNLC2, GNWC2, INPOSC2 FROM XSDC2 "
'        If sMgr = "M" Then  'MGRの場合
'            sql = sql & " WHERE CRYNUMC2 = '" & sXtal & "'"
'        Else                'AGRの場合
'            sql = sql & " WHERE RPCRYNUMC2 = '" & sXtal & "'"
'        End If
'        sql = sql & " AND LIVKC2 = '0'"
'        sql = sql & " AND GNWKNTC2 = 'CC400'"
'        sql = sql & " ORDER BY CRYNUMC2 "
        '’↑更新 SPT用実績作成方法変更 2006/04/17 SMP松田

        'データを抽出する
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

        If rs Is Nothing Then
            DBDRV_SELECT_BLOCK = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
        End If
        ''↓追加 START SPT用実績作成方法変更 2006/06/28 SMP-OKAMOTO
        sCryNum = ""
        i = 0
        ReDim pBlockData(i)
        Do Until rs.EOF
            ''ブロックIDで纏める
            If sCryNum <> CStr(rs("CRYNUMCZ")) Then
                i = i + 1
                ReDim Preserve pBlockData(i)
                sCryNum = CStr(rs("CRYNUMCZ"))
                pBlockData(i).CRYNUMC2 = rs("CRYNUMCZ")
                pBlockData(i).GNLC2 = rs("GNLCZ")
                pBlockData(i).GNWC2 = rs("GNWCZ")
                pBlockData(i).INPOSC2 = rs("INPOSCZ")

                If IsNull(rs("PLANTCATCZ")) = False Then pBlockData(i).PLANTCATC2 = rs("PLANTCATCZ") ' 2007/09/10 SPK Tsutsumi Add Start
            Else
                pBlockData(i).GNLC2 = CLng(pBlockData(i).GNLC2) + CLng(rs("GNLCZ"))
                pBlockData(i).GNWC2 = CLng(pBlockData(i).GNWC2) + CLng(rs("GNWCZ"))

                If IsNull(rs("PLANTCATCZ")) = False Then pBlockData(i).PLANTCATC2 = rs("PLANTCATCZ") ' 2007/09/10 SPK Tsutsumi Add Start
            End If
            rs.MoveNext
        Loop
        ''↑追加 END   SPT用実績作成方法変更 2006/06/28 SMP-OKAMOTO
        ''↓削除 START SPT用実績作成方法変更 2006/06/28 SMP-OKAMOTO
'        ReDim pBlockData(recCnt)
'        For i = 1 To recCnt
'            With pBlockData(i)
'                '’↓更新 SPT用実績作成方法変更 2006/04/17 SMP松田
'                .CRYNUMC2 = rs("CRYNUMCZ")
'                .GNLC2 = rs("GNLCZ")
'                .GNWC2 = rs("GNWCZ")
'                .INPOSC2 = rs("INPOSCZ")
'
''                .CRYNUMC2 = rs("CRYNUMC2")  '
''                .GNLC2 = rs("GNLC2")
''                .GNWC2 = rs("GNWC2")
''                .INPOSC2 = rs("INPOSC2")
'                '’↑更新 SPT用実績作成方法変更 2006/04/17 SMP松田
'            End With
'            rs.MoveNext
'        Next i
        ''↑削除 END   SPT用実績作成方法変更 2006/06/28 SMP-OKAMOTO
        rs.Close
    DBDRV_SELECT_BLOCK = FUNCTION_RETURN_SUCCESS


proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    DBDRV_SELECT_BLOCK = FUNCTION_RETURN_FAILURE
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :結晶研削加工実績 ﾌﾞﾛｯｸ管理ＤＢ更新用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:sCryNum　　　,I  ,String         　,結晶番号(BLOCKID)
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DBDRV_TBCME040_UPDATE(sCryNum As String, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sans As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc017_SQL.bas -- Function DBDRV_TBCME040_UPDATE"
    sErrMsg = ""

    '' ブロック管理テーブルの更新
    sDbName = "E040"
    sql = "update TBCME040 set "
    sql = sql & "KRPROCCD='" & MGPRCD_SETUDAN & "', "      ' 現在管理工程
    sql = sql & "NOWPROC='" & nextCd & "', "                        ' 現在工程
    sql = sql & "LPKRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "                ' 最終通過管理工程
    sql = sql & "LASTPASS='" & nowCd & "', "                        ' 最終通過工程
    sql = sql & "UPDDATE=sysdate "                                 ' 更新日付
    '’↓更新 SPT用実績作成方法変更 2006/04/21 SMP松田
    sql = sql & "where CRYNUM ='" & sCryNum & "' "          ' 結晶番号
    sql = sql & "and   nowproc = 'CC400'"                   ' 現在工程
'    sql = sql & "where BLOCKID ='" & sCryNum & "' "
    '’↑更新 SPT用実績作成方法変更 2006/04/21 SMP松田
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If


    DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
    DBDRV_TBCME040_UPDATE = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

''↓追加 START SPT用実績作成方法変更 2006/08/01 SMP-OKAMOTO
Public Function DBDRV_SELECT_HINBAN(BLOCKID, pHinban, ByRef Hinban12() As String) As FUNCTION_RETURN
Dim sql As String
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc017_SQL.bas -- Function DBDRV_SELECT_HINBAN"

    sql = "select distinct A.hinban12 " & _
          "From (  " & _
          "select hinbcz||ltrim(to_char(revnumcz,'00'))||factorycz||opecz as hinban12,inposcz " & _
          "From XSDCZ  " & _
          "Where (RPCRYNUMCZ =  '" & BLOCKID & "')" & _
          " and (LIVKCZ <> '1' )" & _
          " AND hinbcz||ltrim(to_char(revnumcz,'00'))||factorycz||opecz LIKE '" & pHinban & "%'" & _
          " AND (trim(hinbcz) <> 'Z' AND trim(hinbcz) <> 'G')" & _
          " order by INPOSCZ ) A "

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs Is Nothing Then
        ReDim Hinban12(0)
        DBDRV_SELECT_HINBAN = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim Hinban12(recCnt)
    For i = 1 To recCnt
        If IsNull(rs("hinban12")) = False Then Hinban12(i) = rs("hinban12")
        rs.MoveNext
    Next
    rs.Close

    DBDRV_SELECT_HINBAN = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    gErr.HandleError
    Resume proc_exit
End Function
''↑追加 END   SPT用実績作成方法変更 2006/08/01 SMP-OKAMOTO

'2007/08/17 SPK Tsutsumi Add Start
Public Function GetMukeCode() As FUNCTION_RETURN
    Dim sql As String
    Dim rs As OraDynaset
    Dim recCnt As Long      'レコード数
    Dim i  As Long

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "f_cmbc016_0.frm -- Function Getstaffauthority"

    GetMukeCode = FUNCTION_RETURN_FAILURE

    sql = "Select CODEA9,NAMEJA9 "
    sql = sql & "from KODA9 "
    sql = sql & "where SYSCA9 = 'X' "
    sql = sql & "and SHUCA9 = '20' "
    sql = sql & "and (CODEA9 = '14' "
    sql = sql & "or CODEA9 = '15' "
    sql = sql & "or CODEA9 = '16' "
    sql = sql & "or CODEA9 = 'ZX' "         '08/07/01 ooba
    sql = sql & "or CODEA9 = 'ZZ') "

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs.RecordCount = 0 Then
        GoTo proc_exit
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim s_Mukesaki(recCnt)

    If recCnt = 0 Then
        Exit Function
    End If

    For i = 1 To recCnt
        With s_Mukesaki(i)
            If IsNull(rs.Fields("CODEA9")) = False Then .sMukeCode = rs.Fields("CODEA9")    ' 向先コード
            If IsNull(rs.Fields("NAMEJA9")) = False Then .sMukeName = rs.Fields("NAMEJA9")  ' 向先名
        End With
        rs.MoveNext
    Next
    rs.Close

    GetMukeCode = FUNCTION_RETURN_SUCCESS
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

'概要      :全品番の加工仕様データの取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名          , IO , 型               , 説明
'          :HIN()          ,I   ,tFullHinban       ,品番リスト
'          :Spec()         ,O   ,Judg_Kakou        ,加工仕様
'      　　:戻り値          , O  , FUNCTION_RETURN　, 読み込みの成否
'説明      :
'
'履歴      :2002/04/17 佐野 信哉 作成
'           2009/09    SUMCO Akizuki scmzc_getKakouSpecを元に作成
'                                    Notch規格判定を追加

Public Function scmzc_getKakouSpec_cmbc017(HIN() As tFullHinban, Spec() As Judg_Kakou_cmbc017) As FUNCTION_RETURN
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouSpec_cmbc017"
    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_FAILURE
    
    '求めた全品番の加工仕様を求める
    sql = "select HSXD1CEN, HSXD1MIN, HSXD1MAX, HSXD2CEN, HSXD2MIN, HSXD2MAX, HSXDPACN, HSXDPMIN, HSXDPMAX, HSXDPDIR, HSXDDCEN, HSXDDMIN, HSXDDMAX, HSXDWCEN, HSXDWMIN, HSXDWMAX from TBCME018 "
    sql = sql & "Where " & SQLMake_HINBAN(HIN())

    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim Spec(recCnt)
    For c0 = 1 To recCnt
        Spec(c0).TOP(0) = fncNullCheck(rs("HSXD1CEN"))
        Spec(c0).TOP(1) = fncNullCheck(rs("HSXD1MIN"))
        Spec(c0).TOP(2) = fncNullCheck(rs("HSXD1MAX"))
        Spec(c0).TAIL(0) = fncNullCheck(rs("HSXD2CEN"))
        Spec(c0).TAIL(1) = fncNullCheck(rs("HSXD2MIN"))
        Spec(c0).TAIL(2) = fncNullCheck(rs("HSXD2MAX"))
        Spec(c0).DPTH(0) = fncNullCheck(rs("HSXDDCEN"))
        
        Spec(c0).POS = rs("HSXDPDIR")
        Spec(c0).DPTH(1) = fncNullCheck(rs("HSXDDMIN"))
        Spec(c0).DPTH(2) = fncNullCheck(rs("HSXDDMAX"))
        Spec(c0).WIDH(0) = fncNullCheck(rs("HSXDWCEN"))
        Spec(c0).WIDH(1) = fncNullCheck(rs("HSXDWMIN"))
        Spec(c0).WIDH(2) = fncNullCheck(rs("HSXDWMAX"))
        
        Spec(c0).ANGLE(0) = fncNullCheck(rs("HSXDPACN"))   '2009/09 SUMCO Akizuki

'       仕様規格データに｢-1｣もあるため、Nullチェックでの｢-1｣置換えを廃止    2009/09 SUMCO Akizuki
        If IsNull(rs("HSXDPMIN")) Then
            Spec(c0).ANGLE(1) = 999
        Else
            Spec(c0).ANGLE(1) = rs("HSXDPMIN")
        End If
        
'       仕様規格データに｢-1｣もあるため、Nullチェックでの｢-1｣置換えを廃止    2009/09 SUMCO Akizuki
        If IsNull(rs("HSXDPMAX")) Then
            Spec(c0).ANGLE(2) = 999
        Else
            Spec(c0).ANGLE(2) = rs("HSXDPMAX")
        End If

        rs.MoveNext
    Next
    
    rs.Close

    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouSpec_cmbc017 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :加工実績の取得ドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名          ,IO     , 型               , 説明
'          :BLOCKID         ,I      ,String            ,結晶番号orブロックID
'          :Jiltuseki       ,O      ,Judg_Kakou        ,加工実績
'      　　:戻り値          ,O      , FUNCTION_RETURN　, 読み込みの成否
'説明      :
'
'履歴      :2002/04/17 佐野 信哉 作成
'           2009/09 SUMCO Akizuki
'               共通関数s_mzccjude_SQL(scmzc_getKakouJiltuseki)を参考
'               背景：総合判定も同じ関数を使用して、影響があったために作成

Public Function scmzc_getKakouJiltuseki_cmbc017 _
(BLOCKID As String, Jiltuseki As Judg_Kakou_cmbc017) As FUNCTION_RETURN
    
    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Integer
    Dim c0 As Integer
    Dim AGRFlag As Boolean
    Dim Ans As String
    Dim tINGOTPOS As Integer
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcjudg_SQL.bas -- Function scmzc_getKakouJiltuseki_cmbc017"
    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_FAILURE
    
    '対象ブロックの加工実績の初期化
    For c0 = 1 To 2
        Jiltuseki.TAIL(c0) = -1
        Jiltuseki.TOP(c0) = -1
        Jiltuseki.DPTH(c0) = -1
        Jiltuseki.WIDH(c0) = -1
    Next
    
    Jiltuseki.POS = ""
        '引き上げ結晶の場合
        sql = "select DMTOP1, DMTOP2, DMTAIL1, DMTAIL2, NCHPOS, NCHDPTH, NCHWIDTH, NCHANGLE from TBCMI002 "
        sql = sql & "where CRYNUM='" & Left(BLOCKID, 9) & "000" & "'"
        
        'ﾌﾞﾛｯｸ管理(TBCME040)参照停止　05/10/03 ooba
        sql = sql & " and (select INPOSC2 from XSDC2 where CRYNUMC2 = '" & BLOCKID & "') between INGOTPOS and INGOTPOS+LENGTH-1 "
        sql = sql & "order by INGOTPOS desc, TRANCNT desc"
        sql = "select * from (" & sql & ") where rownum=1"
        Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        recCnt = rs.RecordCount
        If recCnt = 0 Then
            rs.Close
            scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_SUCCESS
            GoTo proc_exit
        End If
        Jiltuseki.TAIL(1) = rs("DMTAIL1")
        Jiltuseki.TAIL(2) = rs("DMTAIL2")
        Jiltuseki.TOP(1) = rs("DMTOP1")
        Jiltuseki.TOP(2) = rs("DMTOP2")
        Jiltuseki.DPTH(1) = rs("NCHDPTH")
        Jiltuseki.DPTH(2) = -1
        Jiltuseki.WIDH(1) = rs("NCHWIDTH")
        Jiltuseki.WIDH(2) = -1
        Jiltuseki.POS = rs("NCHPOS")
        '2009/09 SUMCO Akizuki
        Jiltuseki.ANGLE(1) = rs("NCHANGLE")
        Jiltuseki.ANGLE(2) = 999
        rs.Close

    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_SUCCESS
proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    scmzc_getKakouJiltuseki_cmbc017 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
