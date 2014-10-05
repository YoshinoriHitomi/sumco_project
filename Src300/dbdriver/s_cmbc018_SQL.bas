Attribute VB_Name = "s_cmbc018_SQL"
Option Explicit

' 切断待ち一覧

' 切断待ち一覧
Public Type typ_CutMap
    PRIORITY    As String * 1       ' 優先順位
    BLOCKID     As String * 12      ' ブロックID
    REGDATE     As Date             ' 登録日付
    KENSAKU     As Integer          ' 研削加工済
    CRYNUM      As String           ' 結晶番号
    TRANCNT     As Integer          ' 処理回数
    PRCMCN      As String           ' 研削機
End Type
' 切断仕様 (by SUMCO)
Public Type typ_CutSpec1
    HIN As tFullHinban          ' 品番
    HSXTYPE As String * 1       ' タイプ
    HSXCDIR As String * 1       ' 方位
    HSXD1CEN As Double          ' 直径
    HSXCDOP As String * 1       ' 結晶ドープ  '4/2 Yam
    HSXDPDIR As String * 2      ' ノッチ位置  '3/20 Yam
    HSXDDMIN As Double          ' ノッチ深さ（ＭＩＮ）
    HSXDDMAX As Double          ' ノッチ深さ（ＭＡＸ）
    HSXSDSLP As String * 1      ' シード傾き
    HSXCTCEN As Double          ' シード傾き用（傾縦中心　N(3,2)）4/2 Yam
    HSXCYCEN As Double          ' シード傾き用（傾横中心）N(3,2)) 4/2 Yam
End Type

' 製品仕様
Public Type typ_HinSpec1
    HIN As tFullHinban          ' 品番
    HSXTYPE As String * 1       ' タイプ
    HSXCDIR As String * 1       ' 方位
    HSXD1CEN As Double          ' 直径
    HSXDOP As String * 1        ' 結晶ドープ
    HSXDPDIR As String * 2      ' ノッチ位置
    HSXDDMIN As Double          ' ノッチ深さ（ＭＩＮ）
    HSXDDMAX As Double          ' ノッチ深さ（ＭＡＸ）
    HSXSDSLP As Integer         ' シード傾き
End Type

' 切断指示
Public Type typ_CutInd
    INGOTPOS As Integer         ' カット位置
    TRANCNT As Integer          ' 処理回数
    LENGTH As Integer           ' 長さ
    PROCCODE As String * 5      ' 工程コード
    BDCAUS As String * 3        ' 区分
    HINUP As tFullHinban        ' 上品番
    HINDN As tFullHinban        ' 下品番
    BLOCKID As String * 12      ' ブロックID
    SMP As typ_SXLSample        ' 検査項目
    PALTNUM As String * 4       ' パレット番号
    ERRUPFLG As Boolean         ' 上品番エラーフラグ
    ERRDNFLG As Boolean         ' 下品番エラーフラグ
    RECOMMEND(1 To 13) As String * 1    'お勧め検査(Rs～EPD)
End Type


'概要      :切断指示一覧用 画面表示時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pCutMap　　　,O  ,typ_CutMap     　,切断指示一覧
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001d_Disp(pCutMap() As typ_CutMap) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim recCnt As Long
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001d_Disp"

    '================= 暫定対応 2001/10/17 T.Nomura ==================
    '無効な切断指示を「切断済」とする
    sql = "update TBCME045 CUT set" & _
          " STATCLS='1', " & _
          " UPDDATE=sysdate, " & _
          " SENDFLAG='9' " & _
          "Where (Cut.STATCLS = 0)" & _
          "  and ((select NOWPROC from TBCME040 where BLOCKID=CUT.BLOCKID)<>'CC450')"
    OraDB.ExecuteSQL sql
    '=================================================================

    '' 切断指示の取得
    sql = ""
    sql = sql & "select B.PRIORITY, B.BLOCKID, B.MaxDate, decode(A.CRYNUM,null,0,1) as KENSAKU,"
    sql = sql & "       B.CRYNUM, B.TRANCNT, B.PRCMCN"
    sql = sql & "  from (select distinct CRYNUM from TBCMI002) A, "
    sql = sql & "       (select E045.PRIORITY, E045.BLOCKID, nvl(max(E045.REGDATE),to_date('1900','YYYY')) as MaxDate,"
    sql = sql & "               E045.CRYNUM, I001.TRANCNT, I001.PRCMCN "
    sql = sql & "          from TBCME045 E045, "
    sql = sql & "               (select CRYNUM,max(TRANCNT) as TRANCNT,PRCMCN from TBCMI001 group by CRYNUM,PRCMCN) I001"
    sql = sql & "         where (E045.STATCLS='0')"
    sql = sql & "           and (substr(E045.BLOCKID,1,9)||'000' = I001.CRYNUM(+))"
    sql = sql & "      group by PRIORITY, BLOCKID, E045.CRYNUM, I001.TRANCNT, I001.PRCMCN order by BLOCKID) B"
    sql = sql & " where (substr(B.BLOCKID,1,9)||'000' = A.CRYNUM(+))"
    sql = sql & " order by B.PRIORITY, B.BLOCKID"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
    End If

    ReDim pCutMap(recCnt)
    For i = 1 To recCnt
        With pCutMap(i)
            .PRIORITY = rs("PRIORITY")  ' 優先順位
            .BLOCKID = rs("BLOCKID")    ' ブロックID
            .REGDATE = rs("MaxDate")    ' 登録日付
            If rs("PRCMCN") = "M" Then  'MGRなら未研削で切断可
                .KENSAKU = 1    ' 研削加工済
            Else                        'AGRなら研削加工済かどうかで判断する
                .KENSAKU = rs("KENSAKU")    ' 研削加工済
            End If
            .CRYNUM = rs("CRYNUM")      ' 結晶番号
            .TRANCNT = rs("TRANCNT")    ' 処理回数
            .PRCMCN = IIf(IsNull(rs("PRCMCN")), "", (rs("PRCMCN")))     ' 研削機
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
' 切断

'概要      :切断用 画面表示時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sBlockID　　　,I  ,String         　,ブロックID
'      　　:pCryInf 　　　,O  ,typ_TBCME037   　,結晶情報
'      　　:pBlkMng 　　　,O  ,typ_TBCME040   　,ブロック管理
'      　　:pHinMng 　　　,O  ,typ_TBCME041   　,品番管理（初期時は品番設計）
'      　　:pCrySmp 　　　,O  ,typ_XSDCS   　   ,新サンプル管理（ブロック）
'      　　:pProcBR 　　　,O  ,typ_TBCMI001   　,加工払出実績
'      　　:pCutInd 　　　,O  ,typ_CutInd     　,切断指示
'      　　:pNotCut 　　　,O  ,typ_CutInd     　,切断指示（無切断部）
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001e_Disp(ByVal sBlockID As String, pCryInf As typ_TBCME037, _
                                           pBlkMng() As typ_TBCME040, pHinMng() As typ_TBCME041, _
                                           pCrySmp() As typ_XSDCS, pProcBR As typ_TBCMI001, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpHinDsn() As typ_TBCME039
    Dim tmpProcBR() As typ_TBCMI001
    Dim rs As OraDynaset
    Dim sql As String
    Dim sCryNum As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim i As Long
    '----2002/05/10 追加-------
    Dim newLength As Integer
    Dim cutLength As Integer
    Dim desLength As Integer
    '--------------------------

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001e_Disp"
    sErrMsg = ""

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sCryNum = Left(sBlockID, 9) & "000"
    sDBName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' ブロック管理の取得(s_cmzcTBCME040_SQL.bas が必要)
    sDBName = "E040"
    sql = " where CRYNUM='" & sCryNum & "' and INGOTPOS>=0 order by INGOTPOS"
    If DBDRV_GetTBCME040(pBlkMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 品番管理の取得(s_cmzcTBCME041_SQL.bas が必要)
    sDBName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
        sDBName = "E039"
        sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
        If DBDRV_GetTBCME039(tmpHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("EGET2", sDBName)
            DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        recCnt = UBound(tmpHinDsn)
        If recCnt = 0 Then
            sErrMsg = GetMsgStr("EGET2", sDBName)
            DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        ReDim pHinMng(recCnt)
        For i = 1 To recCnt
            With pHinMng(i)
                .CRYNUM = sCryNum
                .INGOTPOS = tmpHinDsn(i).INGOTPOS
                .hinban = tmpHinDsn(i).hinban
                .REVNUM = tmpHinDsn(i).REVNUM
                .factory = tmpHinDsn(i).FACT
                .opecond = tmpHinDsn(i).OPCOND
                .LENGTH = tmpHinDsn(i).LENGTH
                .REGDATE = tmpHinDsn(i).REGDATE
                .UPDDATE = tmpHinDsn(i).UPDDATE
                .SENDFLAG = tmpHinDsn(i).SENDFLAG
                .SENDDATE = tmpHinDsn(i).SENDDATE
            End With
        Next i
    End If

    '' 結晶サンプル管理の取得
    sDBName = "E043"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME043(pCrySmp(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 加工払出実績の取得(s_cmzcTBCMI001_SQL.bas が必要)
    sDBName = "I001"
    sql = " where CRYNUM='" & sCryNum & "'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMI001"
    sql = sql & " where CRYNUM='" & sCryNum & "')"
    If DBDRV_GetTBCMI001(tmpProcBR(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpProcBR) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pProcBR = tmpProcBR(1)

    '' 切断指示の取得
    sDBName = "E045"
    sql = "select INGOTPOS, TRANCNT from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "' and INGOTPOS>=0 and STATCLS='0'"
    sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "') order by INGOTPOS"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDBName)
        DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' 結晶内開始位置
            .TRANCNT = rs("TRANCNT")        ' 処理回数
        End With
        rs.MoveNext
    Next i
    rs.Close

    For i = 1 To recCnt
        With pCutInd(i)
            sql = "select "
            sql = sql & "LENGTH, PROCCODE, BDCAUS, "
            sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BLOCKID, "
            sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, "
            sql = sql & "CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, "
            sql = sql & "CRYINDGD, CRYINDT, CRYINDEP, PALTNUM"
            sql = sql & " from TBCME045"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & pCutInd(i).INGOTPOS
            sql = sql & " and TRANCNT=" & pCutInd(i).TRANCNT
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDBName)
                DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            .LENGTH = rs("LENGTH")
            .PROCCODE = rs("PROCCODE")
            .BDCAUS = rs("BDCAUS")
            .HINDN.hinban = rs("HINBAN")
            .HINDN.mnorevno = rs("REVNUM")
            .HINDN.factory = rs("FACTORY")
            .HINDN.opecond = rs("OPECOND")
            .BLOCKID = rs("BLOCKID")
            .SMP.CRYINDRS = rs("CRYINDRS")
            .SMP.CRYINDOI = rs("CRYINDOI")
            .SMP.CRYINDB1 = rs("CRYINDB1")
            .SMP.CRYINDB2 = rs("CRYINDB2")
            .SMP.CRYINDB3 = rs("CRYINDB3")
            .SMP.CRYINDL1 = rs("CRYINDL1")
            .SMP.CRYINDL2 = rs("CRYINDL2")
            .SMP.CRYINDL3 = rs("CRYINDL3")
            .SMP.CRYINDL4 = rs("CRYINDL4")
            .SMP.CRYINDCS = rs("CRYINDCS")
            .SMP.CRYINDGD = rs("CRYINDGD")
            .SMP.CRYINDT = rs("CRYINDT")
            .SMP.CRYINDEP = rs("CRYINDEP")
            .PALTNUM = rs("PALTNUM")
            rs.Close
        End With
    Next i

    '' 切断指示（無切断部）の取得
    sql = "select INGOTPOS, TRANCNT, LENGTH, BDCAUS from TBCME045"
    sql = sql & " where BLOCKID='" & sBlockID & "' and INGOTPOS<=-99 and STATCLS='0'"
    sql = sql & " and TRANCNT=" & pCutInd(1).TRANCNT & " order by INGOTPOS"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    ReDim pNotCut(recCnt)
    For i = 1 To recCnt
        With pNotCut(i)
            .INGOTPOS = rs("INGOTPOS")      ' 結晶内開始位置
            .TRANCNT = rs("TRANCNT")        ' 処理回数
            .LENGTH = rs("LENGTH")          ' 長さ
            .BDCAUS = rs("BDCAUS")          ' 区分
        End With
        rs.MoveNext
    Next i
    rs.Close
    
    ''無切断部の正しい長さを取得
    sql = "select C.CRYNUM, min(BODYLENG) as BODYLENG, min(INGOTPOS) as FIRSTCUT, max(INGOTPOS) as LASTCUT "
    sql = sql & "from TBCME045 C, TBCME037 XL "
    sql = sql & "where C.INGOTPOS>=0 and C.STATCLS='0'"
    sql = sql & "  and C.CRYNUM=XL.CRYNUM "
    sql = sql & "  and C.CRYNUM='" & sCryNum & "' "
    sql = sql & "group by C.CRYNUM"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    For i = 1 To UBound(pNotCut)
        If pNotCut(i).INGOTPOS = -99 Then
            pNotCut(i).LENGTH = rs("FIRSTCUT")
        Else
            pNotCut(i).LENGTH = rs("BODYLENG") - rs("LASTCUT")
        End If
    Next
    rs.Close
    Set rs = Nothing
    
    '----2002/05/10--------------
    '' 品番管理の更新（設計時と最下切断位置が違っていた場合）
    '' 最下位置の検査結果が入力できるように品番管理の最下位置をブロック管理と合わせる
    If Right$(sBlockID, 3) = "000" Then
        desLength = tmpHinDsn(UBound(tmpHinDsn)).INGOTPOS + tmpHinDsn(UBound(tmpHinDsn)).LENGTH
        cutLength = pCutInd(UBound(pCutInd)).INGOTPOS + pCutInd(UBound(pCutInd)).LENGTH

        newLength = 0
        If desLength < cutLength Then
            newLength = cutLength - tmpHinDsn(UBound(tmpHinDsn)).INGOTPOS
        End If

        If newLength > 0 Then
            '品番設計の最下品番を、最下切断位置まで伸ばす
            pHinMng(UBound(tmpHinDsn)).LENGTH = newLength
        End If
    End If
    '----------------------------

    
    DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("EGET2", sDBName)
    DBDRV_scmzc_fcmic001e_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :切断用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sCryNum 　　　,I  ,String         　,結晶番号
'      　　:pBlkMng 　　　,I  ,typ_TBCME040   　,ブロック管理
'      　　:pBlkOld 　　　,I  ,typ_TBCME040   　,変更前ブロック管理
'      　　:pHinMng 　　　,I  ,typ_TBCME041   　,品番管理
'      　　:pHinOld 　　　,I  ,typ_TBCME041   　,変更前品番管理
'      　　:pCrySmp 　　　,IO ,typ_XSDCS   　   ,新サンプル管理（ブロック）
'      　　:pCryOld 　　　,I  ,typ_XSDCS   　   ,変更前新サンプル管理（ブロック）
'      　　:pCryCat 　　　,I  ,typ_TBCMG007   　,クリスタルカタログ受入実績
'      　　:pCutRslt　　　,I  ,typ_TBCMI003   　,切断実績
'      　　:pCutInd 　　　,I  ,typ_CutInd     　,切断指示
'      　　:pNotCut 　　　,I  ,typ_CutInd     　,切断指示（無切断部）
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DBDRV_scmzc_fcmic001e_Exec(sCryNum As String, _
                                           pBlkMng() As typ_TBCME040, pBlkOld() As typ_TBCME040, _
                                           pHinMng() As typ_TBCME041, pHinOld() As typ_TBCME041, _
                                           pCrySmp() As typ_XSDCS, pCryOld() As typ_XSDCS, _
                                           pCryCat() As typ_TBCMG007, pCutRslt As typ_TBCMI003, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDBName As String
    Dim recCnt As Long
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_scmzc_fcmic001e_Exec"
    sErrMsg = ""
    
    '' WriteDBLog " ", "Start"

    '' 結晶情報の更新
    sDBName = "E037"
    sql = "update TBCME037 set "
    sql = sql & "KRPROCCD='" & MGPRCD_KESSYOU_SOUGOUHANTEI & "', "
    sql = sql & "PROCCD='" & nextCd & "', "
    sql = sql & "LPKRPROCCD='" & MGPRCD_SETUDAN & "', "
    sql = sql & "LASTPASS='" & nowCd & "', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where CRYNUM='" & sCryNum & "'"
    '' WriteDBLog sql, sDBName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ブロック管理の挿入／更新
    sDBName = "E040"
    If DBDRV_BlockMng_UpdIns(pBlkOld(), pBlkMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 品番管理の挿入／更新
    sDBName = "E041"
    If DBDRV_Hinban_UpdIns(pHinOld(), pHinMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' サンプル№の取得
    recCnt = UBound(pCrySmp)
    For i = 1 To recCnt
        If pCrySmp(i).REPSMPLIDCS = 0 Then
            pCrySmp(i).REPSMPLIDCS = GetNewID_SampleNo()
        End If
    Next i

    '' 結晶サンプル管理の挿入／更新
    sDBName = "E043"
    If DBDRV_CrySmp_UpdIns(pCryOld(), pCrySmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 結晶総合判定測定値の更新
    sDBName = "J014"
    recCnt = UBound(pCrySmp)
    For i = 1 To recCnt
        With pCrySmp(i)
            If .KTKBNCS = "1" Then
                sql = "update TBCMJ014 set SMPKBN='" & .SMPKBNCS & "'"
                sql = sql & " where CRYNUM='" & .XTALCS & "' and POSITION=" & .INPOSCS
                '' WriteDBLog sql, sDBName
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
                    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

    '' 切断指示の更新
    sDBName = "E045"
    recCnt = UBound(pCutInd)
    For i = 1 To recCnt
        With pCutInd(i)
            sql = "update TBCME045 set "
            sql = sql & "STATCLS='1', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0', "
            sql = sql & "SENDDATE=sysdate"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .INGOTPOS
            sql = sql & " and TRANCNT=" & .TRANCNT
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' 切断指示の更新（無切断部）
    recCnt = UBound(pNotCut)
    For i = 1 To recCnt
        With pNotCut(i)
            sql = "update TBCME045 set "
            sql = sql & "STATCLS='1', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0', "
            sql = sql & "SENDDATE=sysdate"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .INGOTPOS
            sql = sql & " and TRANCNT=" & .TRANCNT
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' クリスタルカタログ受入実績の挿入
    sDBName = "G007"
    recCnt = UBound(pCryCat)
    For i = 1 To recCnt
        With pCryCat(i)
            sql = "insert into TBCMG007 "
            sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, BDCODE, PALTNUM, "
            sql = sql & "TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & .CRYNUM & "', "
            sql = sql & "nvl(max(TRANCNT),0)+1, '"
            sql = sql & .KRPROCCD & "', '"
            sql = sql & .PROCCODE & "', '"
            sql = sql & .BDCODE & "', '"
            sql = sql & .PALTNUM & "', '"
            sql = sql & .TSTAFFID & "', "
            sql = sql & "sysdate, '"
            sql = sql & .KSTAFFID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "sysdate "
            sql = sql & " from TBCMG007"
            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        End With
        '' WriteDBLog sql, sDBName
        If OraDB.ExecuteSQL(sql) <= 0 Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
            DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' 切断実績の挿入
    sDBName = "I003"
    With pCutRslt
        sql = "insert into TBCMI003 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, TSTAFFID, "
        sql = sql & "REGDATE, KSTAFFID, UPDDATE, SENDFLAG, SENDDATE, GOUKI)" ' YAM
        sql = sql & " select '"
        sql = sql & .CRYNUM & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "sysdate, "
        sql = sql & "'0', "
        sql = sql & "sysdate, '"    ' Yam
        sql = sql & .GOUKI & "'"    ' Yam
        sql = sql & " from TBCMI003"
        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
    End With
    '' WriteDBLog sql, sDBName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
        DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_SUCCESS

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
    sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
    DBDRV_scmzc_fcmic001e_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :ブロックID用連番の取得
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型       ,説明
'　　      :CryNum       ,I  ,String   ,結晶番号
'　　      :戻り値       ,O  ,String 　,ブロックID連番(max)
'説明      :ブロックIDの最大連番を取得する
'履歴      :2001/09/26　蔵本 作成
Public Function DBDRV_GetBlockNum(CRYNUM As String) As Integer
    
    Dim rs As OraDynaset
    Dim sql As String
    Dim sNum As String
    

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetBlockNum"

    DBDRV_GetBlockNum = 0

    sql = "select "
    sql = sql & "nvl(max(substr(BLOCKID,12,1)),'0') as NUM "
    sql = sql & "from TBCME040 "
    sql = sql & "where BLOCKID like '" & Left$(CRYNUM, 10) & "$_'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    
    If rs Is Nothing Then
        rs.Close
        GoTo proc_exit
    End If
    
    If rs.RecordCount = 0 Then
        DBDRV_GetBlockNum = 0
    Else
        sNum = rs("NUM")
        If StrComp(sNum, "9", vbTextCompare) = 1 Then
            DBDRV_GetBlockNum = Asc(sNum) - 55
        Else
            DBDRV_GetBlockNum = Val(sNum)
        End If
    End If
        
    rs.Close

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

'概要      :テーブル「TBCME040」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME040 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME040_SQL.basより移動)
Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
              " PASSFLAG "   '02/07/05 hama
    
    sqlBase = sqlBase & "From TBCME040"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .LENGTH = rs("LENGTH")           ' 長さ
            .REALLEN = rs("REALLEN")         ' 実長さ
            .BLOCKID = rs("BLOCKID")         ' ブロックID
            .KRPROCCD = rs("KRPROCCD")       ' 現在管理工程
            .NOWPROC = rs("NOWPROC")         ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .DELCLS = rs("DELCLS")           ' 削除区分
            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
            .RSTATCLS = rs("RSTATCLS")       ' 流動状態区分
            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
            .BDCAUS = rs("BDCAUS")           ' 不良理由
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/07/05 hama
             If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/07/05 hama
            End If

        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
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
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

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

'概要      :テーブル「XSDCS」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_XSDCS    ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME043_SQL.basより移動)
Public Function DBDRV_GetTBCME043(records() As typ_XSDCS, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLNO, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, CRYINDRS, CRYINDOI, CRYINDB1," & _
'              " CRYINDB2, CRYINDB3, CRYINDL1, CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, CRYINDEP, CRYRESRS," & _
'              " CRYRESOI, CRYRESB1, CRYRESB2, CRYRESB3, CRYRESL1, CRYRESL2, CRYRESL3, CRYRESL4, CRYRESCS, CRYRESGD, CRYREST," & _
'              " CRYRESEP, SMPLNUM, SMPLPAT, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME043"
    sqlBase = "Select CRYNUMCS, SMPKBNCS, TBKBNCS, REPSMPLIDCS, XTALCS, INPOSCS, HINBCS, REVNUMCS, FACTORYCS, OPECS, KTKBNCS, " & _
              " BLKKTFLAGCS, CRYSMPLIDRSCS, CRYSMPLIDRS1CS, CRYSMPLIDRS2CS, CRYINDRSCS, CRYRESRS1CS, CRYRESRS2CS, CRYSMPLIDOICS, " & _
              " CRYINDOICS, CRYRESOICS, CRYSMPLIDB1CS, CRYINDB1CS, CRYRESB1CS, CRYSMPLIDB2, CRYINDB2CS, CRYRESB2CS, CRYSMPLIDB3CS, " & _
              " CRYINDB3CS, CRYRESB3CS, CRYSMPLIDL1CS, CRYINDL1CS, CRYRESL1CS, CRYSMPLIDL2CS,  CRYINDL2CS, CRYRESL2CS, CRYSMPLIDL3CS, " & _
              " CRYINDL3CS, CRYRESL3CS, CRYSMPLIDL4CS, CRYINDL4CS, CRYRESL4CS, CRYSMPLIDCSCS, CRYINDCSCS, CRYRESCSCS, CRYSMPLIDGDCS, " & _
              " CRYINDGDCS, CRYRESGDCS, CRYSMPLIDTCS, CRYINDTCS, CRYRESTCS, CRYSMPLIDEPCS, CRYINDEPCS,CRYRESEPCS, SMPLNUMCS, " & _
              " SMPLPATCS, TSTAFFCS, TDAYCS, KSTAFFCS, KDAYCS, SNDKCS, SNDDAYCS "
    sqlBase = sqlBase & "From XSDCS"

    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCME043 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内位置
            .SMPKBN = rs("SMPKBN")           ' サンプル区分
            .SMPLNO = rs("SMPLNO")           ' サンプルNo
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .KTKBN = rs("KTKBN")             ' 確定区分
            .CRYINDRS = rs("CRYINDRS")       ' 結晶検査指示（Rs)
            .CRYINDOI = rs("CRYINDOI")       ' 結晶検査指示（Oi)
            .CRYINDB1 = rs("CRYINDB1")       ' 結晶検査指示（B1)
            .CRYINDB2 = rs("CRYINDB2")       ' 結晶検査指示（B2）
            .CRYINDB3 = rs("CRYINDB3")       ' 結晶検査指示（B3)
            .CRYINDL1 = rs("CRYINDL1")       ' 結晶検査指示（L1)
            .CRYINDL2 = rs("CRYINDL2")       ' 結晶検査指示（L2)
            .CRYINDL3 = rs("CRYINDL3")       ' 結晶検査指示（L3)
            .CRYINDL4 = rs("CRYINDL4")       ' 結晶検査指示（L4)
            .CRYINDCS = rs("CRYINDCS")       ' 結晶検査指示（Cs)
            .CRYINDGD = rs("CRYINDGD")       ' 結晶検査指示（GD)
            .CRYINDT = rs("CRYINDT")         ' 結晶検査指示（T)
            .CRYINDEP = rs("CRYINDEP")       ' 結晶検査指示（EPD)
            .CRYRESRS = rs("CRYRESRS")       ' 結晶検査実績（Rs)
            .CRYRESOI = rs("CRYRESOI")       ' 結晶検査実績（Oi)
            .CRYRESB1 = rs("CRYRESB1")       ' 結晶検査実績（B1)
            .CRYRESB2 = rs("CRYRESB2")       ' 結晶検査実績（B2）
            .CRYRESB3 = rs("CRYRESB3")       ' 結晶検査実績（B3)
            .CRYRESL1 = rs("CRYRESL1")       ' 結晶検査実績（L1)
            .CRYRESL2 = rs("CRYRESL2")       ' 結晶検査実績（L2)
            .CRYRESL3 = rs("CRYRESL3")       ' 結晶検査実績（L3)
            .CRYRESL4 = rs("CRYRESL4")       ' 結晶検査実績（L4)
            .CRYRESCS = rs("CRYRESCS")       ' 結晶検査実績（Cs)
            .CRYRESGD = rs("CRYRESGD")       ' 結晶検査実績（GD)
            .CRYREST = rs("CRYREST")         ' 結晶検査実績（T)
            .CRYRESEP = rs("CRYRESEP")       ' 結晶検査実績（EPD)
            .SMPLNUM = rs("SMPLNUM")         ' サンプル枚数
            .SMPLPAT = rs("SMPLPAT")         ' サンプルパターン
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME043 = FUNCTION_RETURN_SUCCESS
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
'概要      :結晶加工払出用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:pHinSpec　　　,IO ,typ_HinSpec1   　,製品仕様
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function fcmic001b_GetSpec(pHinSpec As typ_CutSpec1) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim ctcen As Double
    Dim cycen As Double

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function fcmic001b_GetSpec"

    '' 製品仕様の取得
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP,"     '4/2 Yam
    sql = sql & "HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXSDSLP,"   '3/7 Yam
    sql = sql & "HSXCTCEN, HSXCYCEN "  '4/2 Yam
    sql = sql & " from TBCME018 A,TBCME020 B"
    sql = sql & " where A.HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and A.MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and A.FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and A.OPECOND='" & pHinSpec.HIN.opecond & "'"
    sql = sql & " and B.HINBAN='" & pHinSpec.HIN.hinban & "'"
    sql = sql & " and B.MNOREVNO=" & pHinSpec.HIN.mnorevno
    sql = sql & " and B.FACTORY='" & pHinSpec.HIN.factory & "'"
    sql = sql & " and B.OPECOND='" & pHinSpec.HIN.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")       ' タイプ
        .HSXCDIR = rs("HSXCDIR")       ' 方位
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))     ' 直径
        .HSXCDOP = rs("HSXCDOP")       ' 結晶ドープ  4/2 Yam
        .HSXDPDIR = rs("HSXDPDIR")     ' ノッチ位置
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))     ' ノッチ深さ（ＭＩＮ）3/7 Yam
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))     ' ノッチ深さ（ＭＡＸ）
        .HSXSDSLP = rs("HSXSDSLP")     ' シード傾き
        .HSXCTCEN = fncNullCheck(rs("HSXCTCEN"))     ' シード傾き用（傾縦中心）4/2 Yam
        .HSXCYCEN = fncNullCheck(rs("HSXCYCEN"))     ' シード傾き用（傾縦中心）4/2 Yam
    End With
    rs.Close

    fcmic001b_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMI002」から該当する結晶番号のレコードを検索
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :crynum        ,I  ,string           ,結晶番号
'          :recCount      ,O  ,Integer          ,レコード数
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :
'履歴      :2002/08/09 H.FURUYA
Public Function DBDRV_GetTBCMI002(CRYNUM As String, recCount As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim rs As OraDynaset    'RecordSet

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetTBCMI002"
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる
    sql = "Select * From TBCMI002 where CRYNUM ='" & Left(CRYNUM, 9) & "000'"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    'レコード数をセット
    recCount = rs.RecordCount

    '成功をセット
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_GetTBCMI002 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMI003」から該当する結晶番号のレコードを検索
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型               ,説明
'          :crynum        ,I  ,string           ,結晶番号
'          :recCount      ,O  ,Integer          ,レコード数
'          :戻り値        ,O  ,FUNCTION_RETURN  ,抽出の成否
'説明      :
'履歴      :2002/08/09 H.FURUYA
Public Function DBDRV_GetTBCMI003(CRYNUM As String, recCount As Integer) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim rs As OraDynaset    'RecordSet

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc018_SQL.bas -- Function DBDRV_GetTBCMI003"
    '初期値セット
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_FAILURE

    ''SQLを組み立てる
    sql = "Select * From TBCMI003 where CRYNUM ='" & Left(CRYNUM, 9) & "000'"

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    'レコード数をセット
    recCount = rs.RecordCount


    '成功をセット
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_GetTBCMI003 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

