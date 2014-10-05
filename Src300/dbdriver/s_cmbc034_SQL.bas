Attribute VB_Name = "s_cmbc034_SQL"
Option Explicit

' ＷＦセンター払出

' ブロック一覧
Public Type typ_BlkMap
    BLOCKID             As String * 12      ' ブロックID
    HIN(1 To 5)         As tFullHinban      ' 品番
    WFINDDATE           As String * 10      ' 最終抜試日付
    CRYNUM              As String * 12      ' 結晶番号
    INGOTPOS            As Integer          ' インゴット内位置
    LENGTH              As Integer          ' ブロック長さ
    REALLEN             As Integer          ' ブロック実長さ
    HINREALLEN(1 To 5)  As Integer          ' 品番実長さ
    HinLen(1 To 5)      As Integer          ' 品番長さ
    DIAMETER            As Double           ' 直径 2002/05/01 S.Sano
    sBlockID            As String * 12      ' 先頭ブロックID
    BLOCKORDER          As Integer          ' ブロック順序
    HOLDCLS             As String * 1       ' ホールド状態  --- 2001/09/19 kuramoto 追加 ---
    PASSFLAG            As String * 1       ' 通過フラグ　　--- 200/04/16 Yam
    AGRSTATUS           As String           ' 承認確認区分      add SETkimizuka
    STOP                As String           ' 停止      add SETkimizuka
    CAUSE               As String           ' 停止理由  add SETkimizuka
    PRINTNO             As String           ' 先行評価  add SETkimizuka
End Type

''ブロック内品番情報(構成品番取得用)　　--- 2007/07/17 マルチブロック対応 shindo
Public Type typ_WkBlkMap
    BLOCKID             As String * 12      ' ブロックID
    HINCNT As Integer
    HIN()         As tFullHinban      ' 品番
    HINREALLEN()  As Integer          ' 品番実長さ
    HinLen()      As Integer          ' 品番長さ
    INPOSCA() As Integer '結晶内開始位置
End Type

'品番情報--- 2007/07/17 マルチブロック対応 shindo
Public Wk_tblBlkMap() As typ_WkBlkMap

'ブロック内品番情報
Public Type typ_BlkHinMap
    BLOCKID             As String * 12      ' ブロックID
    HIN                 As tFullHinban      ' 品番
    REALLEN             As Integer          ' 品番実長さ
    HinLen              As Integer          ' 製品長
    PASSFLAG            As String * 1       ' 通過フラグ
    INPOSCA             As Integer          ' 結晶内開始位置　--- 2007/07/17 shindo 追加 ---
    PLANTCATCA          As String           ' 向先 2007/09/12 SPK Tsutsumi Add
End Type

'ブロックの情報
Public Type typ_BlkData
    CRYNUM              As String * 12      ' 結晶番号
    BLOCKID             As String * 12      ' ブロックID
    INGOTPOS            As Integer          ' インゴット内位置
    LENGTH              As Integer          ' ブロック長さ
    REALLEN             As Integer          ' ブロック実長さ
    sBlockID            As String * 12      ' 払出先頭ブロックID
    BLOCKORDER          As Integer          ' ブロック順序
    DIAMETER            As Double           ' 直径 2002/05/01 S.Sano
    WFINDDATE           As String * 10      ' 最終抜試日付
    HOLDCLS             As String * 1       ' ホールド状態
    AGRSTATUS           As String           ' 承認確認区分      add SETkimizuka
    STOP                As String           ' 停止      add SETkimizuka
    CAUSE               As String           ' 停止理由  add SETkimizuka
    PRINTNO             As String           ' 先行評価  add SETkimizuka
End Type


Public Type typ_KOSEIHIN
    KHINBAN As String * 10                  '構成品番
    KHINPOS As Integer                      '構成品番_結晶開始位置
    KHINLEN As Integer                       '構成品番_現在長さ
End Type

''''''概要      :ＷＦセンター払出 画面表示時ＤＢドライバ
''''''ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
''''''      　　:pCryInf　　　,O  ,typ_TBCME037   　,結晶情報
''''''      　　:pBlkMng　　　,O  ,typ_TBCME040   　,ブロック管理
''''''      　　:pSXLMng　　　,O  ,typ_TBCME042   　,SXL管理
''''''      　　:pBsInd 　　　,O  ,typ_TBCMW001   　,抜試指示実績
''''''      　　:pBlkForm　 　,O  ,typ_BlkForm    　,ブロック外形情報
''''''      　　:pBlkBad　　　,O  ,typ_BlkBadPos  　,不良位置
''''''      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
''''''説明      :
''''''履歴      :2001/07/12 作成 蔵本
'''''Public Function DBDRV_scmzc_fcmkc001h_Disp(pCryInf() As typ_TBCME037, pBlkMng() As typ_TBCME040, _
'''''                                           pSXLMng() As typ_TBCME042, pBsInd() As typ_TBCMW001) As FUNCTION_RETURN
'''''
'''''    Dim sql As String
'''''
'''''    '' エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp"
'''''
'''''    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
'''''    If DBDRV_GetTBCME037(pCryInf()) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' ブロック管理の取得(s_cmzcTBCME040_SQL.bas が必要)
'''''    sql = " where NOWPROC='" & PROCD_WFC_HARAIDASI & "'"
'''''    sql = sql & " and LSTATCLS='T' order by CRYNUM, INGOTPOS"
'''''    If DBDRV_GetTBCME040(pBlkMng(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' SXL管理の取得(s_cmzcTBCME042_SQL.bas が必要)
'''''    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''    '' 抜試指示実績の取得(s_cmzcTBCMW001_SQL.bas が必要)
'''''    sql = " where TRANCNT=" & "any(select max(TRANCNT)"
'''''    sql = sql & " from TBCMW001 group by CRYNUM) order by CRYNUM, INGOTPOS"
'''''    If DBDRV_GetTBCMW001(pBsInd(), sql) = FUNCTION_RETURN_FAILURE Then
'''''        DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''        GoTo proc_exit
'''''    End If
'''''
'''''
'''''    DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' 終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001h_Disp = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function


''''''
'''''' 抜試との統合により33_SQLの「Function DBDRV_scmzc_fcmkc001h_Disp22」に変更移行
''''''
''''''概要      :ＷＦセンター払出 画面表示時ＤＢドライバ (Step3.3版)
''''''ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
''''''      　　:pBlkData()   ,O  ,typ_BlkData      ,ブロック情報
''''''      　　:pBlkHinMap() ,O  ,typ_BlkHinMap    ,ブロック品番情報
''''''      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
''''''説明      :
''''''履歴      :2002/04/22 作成 野村
'''''Public Function DBDRV_scmzc_fcmkc001h_Disp2(pBlkData() As typ_BlkData, pBlkHinMap() As typ_BlkHinMap) As FUNCTION_RETURN
'''''Dim sql As String
'''''Dim rs As OraDynaset
'''''Dim recCnt As Long
'''''Dim i As Long
'''''Dim sBlkId As String
'''''Dim blkOrder As Integer
'''''    Dim Jiltuseki As Judg_Kakou '2002/05/01 S.Sano
'''''
'''''    '' エラーハンドラの設定
'''''    On Error GoTo proc_err
'''''    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Disp2"
'''''
'''''    ''ブロックの情報を取得する
'''''    sql = "select B.CRYNUM, B.BLOCKID, B.INGOTPOS, B.LENGTH, B.REALLEN, B2.BLOCKID as SBLOCKID"
'''''    sql = sql & ", nvl("
'''''    sql = sql & "    (select DMTOP1 from TBCMI002 I2"
'''''    sql = sql & "     where CRYNUM=B.CRYNUM"
'''''    sql = sql & "       and INGOTPOS=(select max(INGOTPOS) from TBCMI002 where CRYNUM=B.CRYNUM and INGOTPOS<=B.INGOTPOS)"
'''''    sql = sql & "       and TRANCNT=(select max(TRANCNT) from TBCMI002 where CRYNUM=I2.CRYNUM and INGOTPOS=I2.INGOTPOS)"
'''''    sql = sql & "    )"
'''''    sql = sql & "    , (select DIAMETER from TBCME037 where CRYNUM=B.CRYNUM)"
'''''    sql = sql & "  ) as DIAM"
'''''    sql = sql & ", (select max(UPDDATE) from TBCMW001 where CRYNUM=B2.CRYNUM and INGOTPOS=B2.INGOTPOS) as NUKISHI_AT"
'''''    sql = sql & ", nvl((select HLDTRCLS from TBCMJ012 J12"
'''''    sql = sql & "       where CRYNUM=B.CRYNUM and INGOTPOS=B.INGOTPOS"
'''''    sql = sql & "         and TRANCNT=(select max(TRANCNT) from TBCMJ012 where CRYNUM=J12.CRYNUM and INGOTPOS=J12.INGOTPOS)"
'''''    sql = sql & "      ), '0'"
'''''    sql = sql & "  ) as HOLDCLS "
'''''    sql = sql & "from TBCME040 B, TBCME040 B2 "
'''''    sql = sql & "where B.DELCLS='0' and B.NOWPROC='CC720'"
'''''    sql = sql & "  and B2.CRYNUM=B.CRYNUM"
'''''    sql = sql & "  and B2.INGOTPOS=nvl("
'''''    sql = sql & "        (select max(BLK.INGOTPOS) from TBCME040 BLK, TBCME042 SXL"
'''''    sql = sql & "         where BLK.CRYNUM=B.CRYNUM and BLK.INGOTPOS<=B.INGOTPOS"
'''''    sql = sql & "           and SXL.CRYNUM=BLK.CRYNUM and SXL.INGOTPOS=BLK.INGOTPOS"
'''''    sql = sql & "        ), B.INGOTPOS) "
'''''    sql = sql & "order by B.CRYNUM, B.INGOTPOS"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''    recCnt = rs.RecordCount
'''''    If recCnt <= 0 Then
'''''        ReDim pBlkData(0)
'''''    Else
'''''        ReDim pBlkData(1 To recCnt)
'''''        sBlkId = vbNullString
'''''        blkOrder = 0
'''''        For i = 1 To recCnt
'''''            With pBlkData(i)
'''''                .Crynum = rs("CRYNUM")
'''''                .BLOCKID = rs("BLOCKID")
'''''                .INGOTPOS = rs("INGOTPOS")
'''''                .LENGTH = rs("LENGTH")
'''''                .REALLEN = rs("REALLEN")
'''''                .sBlockID = rs("SBLOCKID")
'''''                If sBlkId <> .sBlockID Then
'''''                    sBlkId = .sBlockID
'''''                    blkOrder = 1
'''''                Else
'''''                    blkOrder = blkOrder + 1
'''''                End If
'''''                .BLOCKORDER = blkOrder
'''''                .DIAMETER = rs("DIAM")
'''''                If (vbNullString & rs("NUKISHI_AT")) = vbNullString Then
'''''                    .WFINDDATE = vbNullString
'''''                Else
'''''                    .WFINDDATE = Format$(rs("NUKISHI_AT"), "yyyy/mm/dd")
'''''                End If
'''''                .HOLDCLS = rs("HOLDCLS")
'''''            End With
'''''            rs.MoveNext
'''''        Next
''''''2002/05/01 S.Sano Start
'''''        rs.Close
'''''        For i = 1 To recCnt
'''''            With pBlkData(i)
'''''            If scmzc_getKakouJiltuseki(.BLOCKID, Jiltuseki) = FUNCTION_RETURN_SUCCESS Then
'''''                .DIAMETER = (Jiltuseki.TAIL(1) + Jiltuseki.TAIL(2) + Jiltuseki.TOP(1) + Jiltuseki.TOP(2)) / 4
'''''            End If
'''''            End With
'''''        Next
''''''2002/05/01 S.Sano End
'''''    End If
'''''
'''''
'''''
'''''    ''ブロック内の品番構成を取得する (ブロックID, 品番, 実長さ, 製品長)
'''''    sql = "select BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, sum(REALLEN) as REALLEN, sum(HINLEN) as HINLEN "
'''''    sql = sql & ", PASSFLAG "
'''''    sql = sql & "from ("
'''''    sql = sql & "  select BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, REALLEN, HINFROM"
'''''    sql = sql & "  , REALLEN"
'''''    sql = sql & "    - case when BD1FROM<HINTO and BD1TO>HINFROM then least(HINTO,BD1TO)-greatest(HINFROM,BD1FROM) else 0 end"
'''''    sql = sql & "    - case when BD2FROM<HINTO and BD2TO>HINFROM then least(HINTO,BD2TO)-greatest(HINFROM,BD2FROM) else 0 end"
'''''    sql = sql & "    - case when BD3FROM<HINTO and BD3TO>HINFROM then least(HINTO,BD3TO)-greatest(HINFROM,BD3FROM) else 0 end"
'''''    sql = sql & "    - case when BD4FROM<HINTO and BD4TO>HINFROM then least(HINTO,BD4TO)-greatest(HINFROM,BD4FROM) else 0 end"
'''''    sql = sql & "    - case when BD5FROM<HINTO and BD5TO>HINFROM then least(HINTO,BD5TO)-greatest(HINFROM,BD5FROM) else 0 end"
'''''    sql = sql & "    as HINLEN"
'''''    sql = sql & "  , PASSFLAG"
'''''    sql = sql & "  from"
'''''    sql = sql & "  ("
'''''    sql = sql & "    select HINS.BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND, HINS.HINFROM, HINS.HINTO, HINS.REALLEN"
'''''    sql = sql & "    , BD1FROM, BD1TO, BD2FROM, BD2TO, BD3FROM, BD3TO, BD4FROM, BD4TO, BD5FROM, BD5TO"
'''''    sql = sql & "    , HINS.PASSFLAG"
'''''    sql = sql & "    from"
'''''    sql = sql & "    (select BLK.CRYNUM, BLK.INGOTPOS, BLK.BLOCKID, SXL.HINBAN, REVNUM, FACTORY, OPECOND"
'''''    sql = sql & "      , greatest(BLK.INGOTPOS,SXL.INGOTPOS) as HINFROM"
'''''    sql = sql & "      , least(BLK.INGOTPOS+BLK.REALLEN,SXL.INGOTPOS+SXL.LENGTH) as HINTO"
'''''    sql = sql & "      , greatest(0,least(BLK.INGOTPOS+BLK.REALLEN,SXL.INGOTPOS+SXL.LENGTH) - greatest(BLK.INGOTPOS,SXL.INGOTPOS)) as REALLEN"
'''''    sql = sql & "      , BLK.PASSFLAG"
'''''    sql = sql & "      from TBCME040 BLK, TBCME042 SXL"
'''''    sql = sql & "      where BLK.DELCLS='0' and BLK.NOWPROC='CC720'"
'''''    sql = sql & "        and SXL.CRYNUM=BLK.CRYNUM"
'''''    sql = sql & "        and SXL.INGOTPOS<BLK.INGOTPOS+BLK.LENGTH"
'''''    sql = sql & "        and SXL.INGOTPOS+SXL.LENGTH>BLK.INGOTPOS"
'''''    sql = sql & "    ) HINS,"
'''''    sql = sql & "    (select B.CRYNUM, B.INGOTPOS"
'''''    sql = sql & "      , B.INGOTPOS + case when PART1=9999 then B.REALLEN-J.P1BDLEN else PART1 end as BD1FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART1=9999 then B.REALLEN-J.P1BDLEN else PART1 end + P1BDLEN as BD1TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART2=9999 then B.REALLEN-J.P2BDLEN else PART2 end as BD2FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART2=9999 then B.REALLEN-J.P2BDLEN else PART2 end + P2BDLEN as BD2TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART3=9999 then B.REALLEN-J.P3BDLEN else PART3 end as BD3FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART3=9999 then B.REALLEN-J.P3BDLEN else PART3 end + P3BDLEN as BD3TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART4=9999 then B.REALLEN-J.P4BDLEN else PART4 end as BD4FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART4=9999 then B.REALLEN-J.P4BDLEN else PART4 end + P4BDLEN as BD4TO"
'''''    sql = sql & "      , B.INGOTPOS + case when PART5=9999 then B.REALLEN-J.P5BDLEN else PART5 end as BD5FROM"
'''''    sql = sql & "      , B.INGOTPOS + case when PART5=9999 then B.REALLEN-J.P5BDLEN else PART5 end + P5BDLEN as BD5TO"
'''''    sql = sql & "      from TBCMJ010 J, TBCME040 B"
'''''    sql = sql & "      where B.DELCLS='0' and B.NOWPROC='CC720'"
'''''    sql = sql & "        and J.CRYNUM=B.CRYNUM and J.INGOTPOS=B.INGOTPOS"
'''''    sql = sql & "        and J.TRANCNT=(select max(TRANCNT) from TBCMJ010 where CRYNUM=J.CRYNUM and INGOTPOS=J.INGOTPOS)"
'''''    sql = sql & "    ) BADS"
'''''    sql = sql & "    where HINS.CRYNUM=BADS.CRYNUM"
'''''    sql = sql & "      and HINS.INGOTPOS=BADS.INGOTPOS"
'''''    sql = sql & "  )"
'''''    sql = sql & ")"
'''''    sql = sql & "group by BLOCKID, HINBAN, REVNUM, FACTORY, OPECOND "
'''''    sql = sql & ", PASSFLAG "
'''''    sql = sql & "order by BLOCKID, min(HINFROM)"
'''''    Set rs = OraDB.CreateDynaset(sql, ORADB_DEFAULT)
'''''    recCnt = rs.RecordCount
'''''    If recCnt <= 0 Then
'''''        ReDim pBlkHinMap(0)
'''''    Else
'''''        ReDim pBlkHinMap(1 To recCnt)
'''''        For i = 1 To recCnt
'''''            With pBlkHinMap(i)
'''''                .BLOCKID = rs("BLOCKID")
'''''                .HIN.hinban = rs("HINBAN")
'''''                .HIN.mnorevno = rs("REVNUM")
'''''                .HIN.factory = rs("FACTORY")
'''''                .HIN.opecond = rs("OPECOND")
'''''                .REALLEN = rs("REALLEN")
'''''                .HINLEN = rs("HINLEN")
'''''                .PASSFLAG = vbNullString & rs("PASSFLAG")
'''''            End With
'''''            rs.MoveNext
'''''        Next
'''''    End If
'''''    rs.Close '2002/05/01 S.Sano
'''''
'''''    DBDRV_scmzc_fcmkc001h_Disp2 = FUNCTION_RETURN_SUCCESS
'''''
'''''proc_exit:
'''''    '' 終了
'''''    gErr.Pop
'''''    Exit Function
'''''
'''''proc_err:
'''''    '' エラーハンドラ
'''''    Debug.Print "====== Error SQL ======"
'''''    Debug.Print sql
'''''    gErr.HandleError
'''''    DBDRV_scmzc_fcmkc001h_Disp2 = FUNCTION_RETURN_FAILURE
'''''    Resume proc_exit
'''''
'''''End Function



'概要      :ＷＦセンター払出 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sStaffID　　　,I  ,String         　,社員ID
'      　　:pBlkMap 　　　,I  ,typ_BlkMap     　,ブロック一覧
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Public Function DBDRV_scmzc_fcmkc001h_Exec(ByVal sStaffID As String, pBlkMap() As typ_BlkMap, sErrMsg As String) As FUNCTION_RETURN

    Dim sql     As String
    Dim sDbName As String
    Dim recCnt  As Long
    Dim iPos    As Integer
    Dim i       As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_scmzc_fcmkc001h_Exec"
    sErrMsg = ""

    recCnt = UBound(pBlkMap)
    For i = 1 To recCnt
        '' ブロック新規情報の挿入
        If DBDRV_BlockNewInf_Ins(pBlkMap(i), sDbName) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

        With pBlkMap(i)
            '' ブロック管理の更新
            sDbName = "E040"
            sql = "update TBCME040 set "
            sql = sql & "LPKRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
            sql = sql & "LASTPASS  ='" & PROCD_WFC_HARAIDASI & "', "
            sql = sql & "DELCLS    ='1', "
            sql = sql & "LSTATCLS  ='W', "
            sql = sql & "UPDDATE   =sysdate, "
            sql = sql & "SENDFLAG  ='0'"
            sql = sql & " where CRYNUM='" & .CRYNUM & "' and INGOTPOS=" & .INGOTPOS

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

''↓削除START SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本
'            '' SXL管理の更新
'            sDbName = "E042"
'            iPos = .INGOTPOS + .LENGTH
'            sql = "update TBCME042 set "
'            sql = sql & "KRPROCCD  ='" & MGPRCD_WFC_SOUGOUHANTEI & "', "
'            sql = sql & "NOWPROC   ='" & PROCD_WFC_SOUGOUHANTEI & "', "
'            sql = sql & "LPKRPROCCD='" & MGPRCD_WFC_HARAIDASI & "', "
'            sql = sql & "LASTPASS  ='" & PROCD_WFC_HARAIDASI & "', "
'            sql = sql & "UPDDATE   =sysdate, "
'            sql = sql & "SENDFLAG  ='0'"
'            sql = sql & " where CRYNUM='" & .CRYNUM & "'"
'            sql = sql & " and INGOTPOS>=" & .INGOTPOS
'            sql = sql & " and INGOTPOS<" & iPos
'            If OraDB.ExecuteSQL(sql) < 0 Then
'                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
'                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
'                GoTo proc_exit
'            End If
''↑削除END   SXL管理（E042）→XSDCB機能移行 '05/12/19 SMP岡本

            '' WF払出実績の挿入
            sDbName = "J011"
            sql = "insert into TBCMJ011 "
            sql = sql & "(CRYNUM,  INGOTPOS, LENGTH,         KRPROCCD, PROCCODE, "
            sql = sql & " BLOCKID, SBLOCKID, BLOCKORDER,     TSTAFFID, REGDATE, "

            '2007/08/31 SPK Tsutsumi Add Start
            sql = sql & " KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE,PLANTCAT)"
'            sql = sql & " KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            '2007/08/31 SPK Tsutsumi Add End

            sql = sql & " values ('"
            sql = sql & .CRYNUM & "', "                 ' 結晶番号
            sql = sql & .INGOTPOS & ", "                ' インゴット内位置
            sql = sql & .LENGTH & ", '"                 ' 長さ
            sql = sql & MGPRCD_WFC_HARAIDASI & "', '"   ' 管理工程コード
            sql = sql & PROCD_WFC_HARAIDASI & "', '"    ' 工程コード
            sql = sql & .BLOCKID & "', '"               ' ブロックID
            sql = sql & .sBlockID & "', "               ' 先頭ブロックID
            sql = sql & .BLOCKORDER & ", '"             ' ブロック順序
            sql = sql & sStaffID & "', "                ' 登録社員ID
            sql = sql & "sysdate, '"                    ' 登録日付
            sql = sql & sStaffID & "', "                ' 更新社員ID
            sql = sql & "sysdate, "                     ' 更新日付
            sql = sql & "'0', "                         ' SUMMIT送信フラグ
            sql = sql & "'0', "                         ' 送信フラグ

            '2007/08/31 SPK Tsutsumi Add Start
            sql = sql & "sysdate,'"                      ' 送信日付
            sql = sql & sCmbMukesaki & "')"              ' 向先
'            sql = sql & "sysdate)"                      ' 送信日付
            '2007/08/31 SPK Tsutsumi Add End

            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    '' 2003/04/22 ooba  WFセンター払出しでの送信処理を停止する
    '' 測定評価指示の送信を予約する
'    sDBName = "Y003"
'    sql = "update TBCMY003 set SENDFLAG='0' "
'    sql = sql & "where substr(SAMPLEID,1,12) in ("
'    recCnt = UBound(pBlkMap)
'    For i = 1 To recCnt
'        If i = recCnt Then
'            sql = sql & "'" & pBlkMap(i).BLOCKID & "'"
'        Else
'            sql = sql & "'" & pBlkMap(i).BLOCKID & "',"
'        End If
'    Next i
'    sql = sql & ")"
'    If OraDB.ExecuteSQL(sql) < 0 Then   '0件はエラーとしない
'        sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'        DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
'        GoTo proc_exit
'    End If

    '関連ﾌﾞﾛｯｸ情報登録　07/12/21 ooba START =====================================>
    If recCnt > 1 Then
        sDbName = "Y023"
        If DBDRV_KanrenBlk(pBlkMap()) = FUNCTION_RETURN_FAILURE Then

            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    End If
    '関連ﾌﾞﾛｯｸ情報登録　07/12/21 ooba END =======================================>
    
    DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001h_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :内部関数：ブロック新規情報の作成（抜試指示付）
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pBlkMap　　　,I  ,typ_BlkMap     　,ブロック一覧
'      　　:sDBName　　　,O  ,String         　,DB名称
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/12  作成 蔵本
Private Function DBDRV_BlockNewInf_Ins(pBlkMap As typ_BlkMap, sDbName As String) As FUNCTION_RETURN

    Dim rs          As OraDynaset
    Dim sql         As String
    Dim CRYSTALMEN  As String
    Dim SEED        As Integer
    Dim TANMEN      As String * 3
    Dim WARPRANK    As String * 1
    Dim Ans         As String
    Dim MainHin     As tFullHinban      '代表品番　05/11/25 ooba
    Dim SubHin      As tFullHinban      'ｻﾌﾞ代表品番　05/11/25 ooba
    Dim c0 As Integer
    Dim c1 As Integer
    Dim KOSEIHIN() As typ_KOSEIHIN
    Dim LENKEI As Integer
    Dim LENSA As Integer
    Dim KOSCNT As Integer               '構成品番数
    Dim FRKOSCNT As Integer             'ブロックに紐付く品番数
    Dim GNLC2 As Integer                'ブロック現在長さ 07/08/01 shindo
    Dim REALLC2 As Integer              'ブロック実長さ 07/08/01 shindo



    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_BlockNewInf_Ins"

    '' シード傾きの取得
'頭8を購入単結晶扱いしない 2007/10/01 SETsw kubota
'    If Left(pBlkMap.CRYNUM, 1) = "8" Then
'        '' 購入単結晶の場合
'        'ブロック新規情報、シード傾きの求め方を変更
'        sDbName = "G002"
'        If DBDRV_getSEEDDEG(Trim(pBlkMap.BLOCKID), SEED) = FUNCTION_RETURN_FAILURE Then
'            DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    Else
        '' 購入単結晶以外の場合
        sDbName = "H004"
        If DBDRV_getSEED(pBlkMap.CRYNUM, SEED) = FUNCTION_RETURN_FAILURE Then
            DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
'    End If

''    '' 結晶面の取得
''    sDbName = "E022"
''    sql = "select HWFCDIR from TBCME022"
''    sql = sql & " where HINBAN='" & pBlkMap.HIN(1).hinban & "'"
''    sql = sql & " and MNOREVNO=" & pBlkMap.HIN(1).mnorevno
''    sql = sql & " and FACTORY='" & pBlkMap.HIN(1).factory & "'"
''    sql = sql & " and OPECOND='" & pBlkMap.HIN(1).opecond & "'"
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    If rs.RecordCount <= 0 Then
''        rs.Close
''        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
''        GoTo proc_exit
''    End If
''    CRYSTALMEN = rs("HWFCDIR")
''    rs.Close
''
''    If CRYSTALMEN = "B" Then
''        CRYSTALMEN = "100"
''    ElseIf CRYSTALMEN = "C" Then
''        CRYSTALMEN = "511"
''    ElseIf CRYSTALMEN = "D" Then
''        CRYSTALMEN = "110"
''    Else
''        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
''        GoTo proc_exit
''    End If

    'ブロック新規情報
    '' 端面角度を求める
'    If DBDRV_getTANMEN(pBlkMap, ans) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_getTANMEN(pBlkMap, SubHin, Ans) = FUNCTION_RETURN_FAILURE Then     '05/11/25 ooba
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    TANMEN = Ans
    '' ワープランクを求める
'    If DBDRV_getWARPRANK(pBlkMap, ans) = FUNCTION_RETURN_FAILURE Then
    If DBDRV_getWARPRANK(pBlkMap, MainHin, Ans) = FUNCTION_RETURN_FAILURE Then  '05/11/25 ooba
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    WARPRANK = Ans

    'ｻﾌﾞ代表品番で結晶面を取得　05/11/25 ooba START ================================>
    '' 結晶面の取得
    sDbName = "E022"
    sql = "select HWFCDIR from TBCME022"
    sql = sql & " where HINBAN='" & SubHin.hinban & "'"
    sql = sql & " and MNOREVNO=" & SubHin.mnorevno
    sql = sql & " and FACTORY='" & SubHin.factory & "'"
    sql = sql & " and OPECOND='" & SubHin.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    CRYSTALMEN = rs("HWFCDIR")
    rs.Close

    If CRYSTALMEN = "B" Then
        CRYSTALMEN = "100"
    ElseIf CRYSTALMEN = "C" Then
        CRYSTALMEN = "511"
    ElseIf CRYSTALMEN = "D" Then
        CRYSTALMEN = "110"
    Else
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    'ｻﾌﾞ代表品番で結晶面を取得　05/11/25 ooba END ==================================>


    'Null 対応に伴なう修正でNullの場合は”０”と見なす。　濱　平成16年10月8日
    'If DBDRV_NULLChk(pBlkMap) = FUNCTION_RETURN_FAILURE Then
    '    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
    '    GoTo PROC_EXIT
    'End If

 '構成品番の情報を取得 07/07/19 SHINDO STR=======================================>
    For c0 = 0 To UBound(Wk_tblBlkMap())

    If pBlkMap.BLOCKID = Wk_tblBlkMap(c0).BLOCKID Then
        'ブロックに紐付く品番数を取得
        FRKOSCNT = UBound(Wk_tblBlkMap(c0).HIN)
        '品番情報を取得
        For c1 = 1 To UBound(Wk_tblBlkMap(c0).HIN)
    ReDim Preserve KOSEIHIN(c1)
            With KOSEIHIN(c1)
                .KHINBAN = Wk_tblBlkMap(c0).HIN(c1).hinban + Format(Wk_tblBlkMap(c0).HIN(c1).mnorevno, "00")
                .KHINPOS = Wk_tblBlkMap(c0).INPOSCA(c1)
                .KHINLEN = Wk_tblBlkMap(c0).HinLen(c1)
            End With
        Next c1
    End If
    Next c0
  'ブロックの実長さ、現在長さを取得  07/08/01 shindo
       sDbName = "XSDC2"
       sql = "select GNLC2,REALLC2"
       sql = sql & " from XSDC2"
       sql = sql & " where CRYNUMC2='" & pBlkMap.BLOCKID & "'"
       Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
       If rs.RecordCount = 0 Then
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
       End If
       GNLC2 = rs("GNLC2")
       REALLC2 = rs("REALLC2")
       rs.Close

    '結晶内開始位置をブロック内開始位置に変更

'07/08/01 shindo DELL_STR
    '長さの合計算出
'        LENKEI = 0
'        LENSA = 0
'07/08/01 shindo DELL_END

        For c1 = 1 To FRKOSCNT
            With KOSEIHIN(c1)
                .KHINPOS = .KHINPOS - pBlkMap.INGOTPOS
'07/08/01 shindo DELL
'               LENKEI = LENKEI + .KHINLEN
            End With
        Next c1
'07/08/01 shindo DELL
'        'ブックの製品長さと実長さの差を算出
'        LENSA = pBlkMap.REALLEN - LENKEI

        If FRKOSCNT <= 5 Then
            KOSCNT = FRKOSCNT
            With KOSEIHIN(KOSCNT)
'07/08/01 shindo UPDATE_STR
'                .KHINLEN = .KHINLEN + LENSA
                .KHINLEN = .KHINLEN + (REALLC2 - GNLC2)
'07/08/01 shindo UPDATE_END
            End With
        Else
            KOSCNT = 5
            For c1 = 6 To FRKOSCNT
                    KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + KOSEIHIN(c1).KHINLEN
            Next c1
'07/08/01 shindo UPDATE_STR
'            KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + LENSA
            KOSEIHIN(5).KHINLEN = KOSEIHIN(5).KHINLEN + (REALLC2 - GNLC2)
'07/08/01 shindo UPDATE_END
        End If


 '構成品番の情報を取得 07/07/19 SHINDO END=======================================<

    '' ブロック新規情報の挿入
    sDbName = "Y001"
    With pBlkMap
        sql = "insert into TBCMY001 ("
        sql = sql & "BLOCKID, "         ' ブロックID
        sql = sql & "BLOCKLEN, "        ' ブロックの長さ
        sql = sql & "MAINHINBAN, "      ' 代表品番
        sql = sql & "PNTYPE, "          ' タイプ
        sql = sql & "ROUP, "            ' 比抵抗上限値
        sql = sql & "ROLOW, "           ' 比抵抗下限値
        sql = sql & "OIUP, "            ' 酸素濃度上限値
        sql = sql & "OILOW, "           ' 酸素濃度下限値
        sql = sql & "TANMEN, "          ' 端面角度
        sql = sql & "WARPRANK, "        ' ワープランク
        sql = sql & "CRYSTALMEN, "      ' 結晶面
        sql = sql & "SLPCEN, "          ' 傾中心
        sql = sql & "SLPLOW, "          ' 傾下限
        sql = sql & "SLPUP, "           ' 傾上限
        sql = sql & "INSPMETH, "        ' 検査方法
        sql = sql & "INSPFREQ, "        ' 検査頻度
        sql = sql & "SLPDRC, "          ' 傾方位
        sql = sql & "SLPDRCAPP, "       ' 傾方位指定
        sql = sql & "SLPHEIDRC, "       ' 傾縦方位
        sql = sql & "SLPHEICEN, "       ' 傾縦中心
        sql = sql & "SLPHEILOW, "       ' 傾縦下限
        sql = sql & "SLPHEIUP, "        ' 傾縦上限
        sql = sql & "SLPWIDDRC, "       ' 傾横方位
        sql = sql & "SLPWIDCEN, "       ' 傾横中心
        sql = sql & "SLPWIDLOW, "       ' 傾横下限
        sql = sql & "SLPWIDUP, "        ' 傾横上限
        sql = sql & "SEED, "            ' 引上時使用したシ－ド傾き
        sql = sql & "TXID, "            ' トランザクションID
        sql = sql & "SBLOCKID, "        ' 先頭ブロックID
        sql = sql & "BLOCKORDER, "      ' ブロック順序
        sql = sql & "REGDATE, "         ' 登録日付
        sql = sql & "SENDFLAG, "        ' 送信フラグ
'2007/07/17 UPDATE_STR マルチブロック対応　SHINDO
'        sql = sql & "SENDDATE)"         ' 送信日付
'****************
        sql = sql & "SENDDATE,"         ' 送信日付
        sql = sql & "PLANTCAT, "        ' 向先  2007/08/31 SPK Tsutsumi Add
        sql = sql & "HINCNT, "          ' 構成品番数"
        sql = sql & "MULUTIHINBAN1, "   ' 構成品番その１品番"
        sql = sql & "TOPICHI1, "        ' 構成品番その１Top位置(mm)"
        sql = sql & "TAILICHI1, "       ' 構成品番その１Tail位置(mm)"
        sql = sql & "HINBANLEN1, "      ' 構成品番その１長さ(mm)"
        sql = sql & "MULUTIHINBAN2, "   ' 構成品番その２品番"
        sql = sql & "TOPICHI2, "        ' 構成品番その２Top位置(mm)"
        sql = sql & "TAILICHI2, "       ' 構成品番その２Tail位置(mm)"
        sql = sql & "HINBANLEN2, "      ' 構成品番その２長さ(mm)"
        sql = sql & "MULUTIHINBAN3, "   ' 構成品番その３品番"
        sql = sql & "TOPICHI3, "        ' 構成品番その３Top位置(mm)"
        sql = sql & "TAILICHI3, "       ' 構成品番その３Tail位置(mm)"
        sql = sql & "HINBANLEN3, "      ' 構成品番その３長さ(mm)"
        sql = sql & "MULUTIHINBAN4, "   ' 構成品番その４品番"
        sql = sql & "TOPICHI4, "        ' 構成品番その４Top位置(mm)"
        sql = sql & "TAILICHI4, "       ' 構成品番その４Tail位置(mm)"
        sql = sql & "HINBANLEN4, "      ' 構成品番その４長さ(mm)"
        sql = sql & "MULUTIHINBAN5, "   ' 構成品番その５品番"
        sql = sql & "TOPICHI5, "        ' 構成品番その５Top位置(mm)"
        sql = sql & "TAILICHI5, "       ' 構成品番その５Tail位置(mm)"
        sql = sql & "HINBANLEN5)"       ' 構成品番その５長さ(mm)"

'2007/07/17 UPDATE_END マルチブロック対応　SHINDO

        sql = sql & " select '"
        sql = sql & .BLOCKID & "', "                                            ' ブロックID
        sql = sql & .REALLEN & ", '"                                            ' ブロックの長さ
''        sql = sql & .HIN(1).hinban & Format(.HIN(1).mnorevno, "00") & "', "     ' 代表品番
''        sql = sql & "E021HWFTYPE, "                                             ' タイプ
''        sql = sql & "case when E021HWFRMAX>=99999.9 then '99999.9'"
''        sql = sql & " when E021HWFRMAX>=9999.995 then to_char(round(E021HWFRMAX,2),'fm99990.0')"
''        sql = sql & " when E021HWFRMAX>=999.9995 then to_char(round(E021HWFRMAX,3),'fm9990.00')"
''        sql = sql & " when E021HWFRMAX>=99.99995 then to_char(round(E021HWFRMAX,4),'fm990.000')"
''        sql = sql & " when E021HWFRMAX>=10.00000 then to_char(round(E021HWFRMAX,5),'fm90.0000')"
''        sql = sql & " when E021HWFRMAX>=0.0 then to_char(E021HWFRMAX,'fm0.00000')"
''        sql = sql & " when nvl(E021HWFRMAX,0) = 0 then '0.0000'"
''        sql = sql & " else '-1.0000'"
''        sql = sql & "end as RMAX,"                                              ' 比抵抗上限値
''        sql = sql & "case when E021HWFRMIN>=99999.9 then '99999.9'"
''        sql = sql & " when E021HWFRMIN>=9999.995 then to_char(round(E021HWFRMIN,2),'fm99990.0')"
''        sql = sql & " when E021HWFRMIN>=999.9995 then to_char(round(E021HWFRMIN,3),'fm9990.00')"
''        sql = sql & " when E021HWFRMIN>=99.99995 then to_char(round(E021HWFRMIN,4),'fm990.000')"
''        sql = sql & " when E021HWFRMIN>=10.00000 then to_char(round(E021HWFRMIN,5),'fm90.0000')"
''        sql = sql & " when E021HWFRMIN>=0.0 then to_char(E021HWFRMIN,'fm0.00000')"
''        sql = sql & " when nvl(E021HWFRMIN,0) = 0 then '0.0000'"
''        sql = sql & " else '-1.0000'"
''        sql = sql & "end as RMIN,"                                              ' 比抵抗下限値
''        sql = sql & "nvl(to_char(abs(E025HWFONMAX),'fm90.00'),'0.00'), "        ' 酸素濃度上限値"
''        sql = sql & "nvl(to_char(abs(E025HWFONMIN),'fm90.00'),'0.00'), "        ' 酸素濃度下限値
''        sql = sql & "'" & TANMEN & "', "                                        ' 端面角度
''        sql = sql & "'" & WARPRANK & "', '"                                     ' ワープランク
''        sql = sql & CRYSTALMEN & "', "                                          ' 結晶面
''        sql = sql & "nvl(to_char(abs(E022HWFCSCEN),'fm0.00'),'0.00'), "         ' 傾中心
''        sql = sql & "nvl(to_char(E022HWFCSMIN,'fm0.00'),'0.00'), "              ' 傾下限
''        sql = sql & "nvl(to_char(E022HWFCSMAX,'fm0.00'),'0.00'), "              ' 傾上限
''        sql = sql & "E022HWFCKWAY, "                                            ' 検査方法
''        sql = sql & "E022HWFCKHNM || E022HWFCKHNN || E022HWFCKHNH || E022HWFCKHNU, "    ' 検査頻度（枚、抜、保、ウの順で足す）
''        sql = sql & "E022HWFCSDIR, "                                            ' 傾方位
''        sql = sql & "E022HWFCSDIS, "                                            ' 傾方位指定
''        sql = sql & "E022HWFCTDIR, "                                            ' 傾縦方位
''        sql = sql & "nvl(to_char(E022HWFCTCEN,'fm0.00'),'0.00'), "              ' 傾縦中心
''        sql = sql & "nvl(to_char(E022HWFCTMIN,'fm0.00'),'0.00'), "              ' 傾縦下限
''        sql = sql & "nvl(to_char(E022HWFCTMAX,'fm0.00'),'0.00'), "              ' 傾縦上限
''        sql = sql & "E022HWFCYDIR, "                                            ' 傾横方位
''        sql = sql & "nvl(to_char(E022HWFCYCEN,'fm0.00'),'0.00'), "              ' 傾横中心
''        sql = sql & "nvl(to_char(E022HWFCYMIN,'fm0.00'),'0.00'), "              ' 傾横下限
''        sql = sql & "nvl(to_char(E022HWFCYMAX,'fm0.00'),'0.00'), "              ' 傾横上限
''        sql = sql & SEED & ", "                                                 ' 引上時使用したシ－ド傾き
''        sql = sql & "'TX850I', '"                                               ' トランザクションID
''        sql = sql & .sBlockID & "', "                                           ' 先頭ブロックID
''        sql = sql & .BLOCKORDER & ", "                                          ' ブロック順序
''        sql = sql & "sysdate, "                                                 ' 登録日付
''        sql = sql & "'0', "                                                     ' 送信フラグ
''        sql = sql & "sysdate"                                                   ' 送信日付
''        sql = sql & " from VECME001"
''        sql = sql & " where E018HINBAN='" & .HIN(1).hinban & "'"
''        sql = sql & " and E018MNOREVNO=" & .HIN(1).mnorevno
''        sql = sql & " and E018FACTORY='" & .HIN(1).factory & "'"
''        sql = sql & " and E018OPECOND='" & .HIN(1).opecond & "'"

        '代表品番、ｻﾌﾞ代表品番の仕様を取得　05/11/25 ooba START ==============================>
        sql = sql & MainHin.hinban & Format(MainHin.mnorevno, "00") & "', "     ' 代表品番
        sql = sql & "MAIN.E021HWFTYPE, "                                        ' タイプ
        sql = sql & "case when MAIN.E021HWFRMAX>=99999.9 then '99999.9'"
        sql = sql & " when MAIN.E021HWFRMAX>=9999.995 then to_char(round(MAIN.E021HWFRMAX,2),'fm99990.0')"
        sql = sql & " when MAIN.E021HWFRMAX>=999.9995 then to_char(round(MAIN.E021HWFRMAX,3),'fm9990.00')"
        sql = sql & " when MAIN.E021HWFRMAX>=99.99995 then to_char(round(MAIN.E021HWFRMAX,4),'fm990.000')"
        sql = sql & " when MAIN.E021HWFRMAX>=10.00000 then to_char(round(MAIN.E021HWFRMAX,5),'fm90.0000')"
        sql = sql & " when MAIN.E021HWFRMAX>=0.0 then to_char(MAIN.E021HWFRMAX,'fm0.00000')"
        sql = sql & " when nvl(MAIN.E021HWFRMAX,0) = 0 then '0.0000'"
        sql = sql & " else '-1.0000'"
        sql = sql & "end as RMAX,"                                              ' 比抵抗上限値
        sql = sql & "case when MAIN.E021HWFRMIN>=99999.9 then '99999.9'"
        sql = sql & " when MAIN.E021HWFRMIN>=9999.995 then to_char(round(MAIN.E021HWFRMIN,2),'fm99990.0')"
        sql = sql & " when MAIN.E021HWFRMIN>=999.9995 then to_char(round(MAIN.E021HWFRMIN,3),'fm9990.00')"
        sql = sql & " when MAIN.E021HWFRMIN>=99.99995 then to_char(round(MAIN.E021HWFRMIN,4),'fm990.000')"
        sql = sql & " when MAIN.E021HWFRMIN>=10.00000 then to_char(round(MAIN.E021HWFRMIN,5),'fm90.0000')"
        sql = sql & " when MAIN.E021HWFRMIN>=0.0 then to_char(MAIN.E021HWFRMIN,'fm0.00000')"
        sql = sql & " when nvl(MAIN.E021HWFRMIN,0) = 0 then '0.0000'"
        sql = sql & " else '-1.0000'"
        sql = sql & "end as RMIN,"                                              ' 比抵抗下限値
        sql = sql & "nvl(to_char(abs(MAIN.E025HWFONMAX),'fm90.00'),'0.00'), "   ' 酸素濃度上限値"
        sql = sql & "nvl(to_char(abs(MAIN.E025HWFONMIN),'fm90.00'),'0.00'), "   ' 酸素濃度下限値
        sql = sql & "'" & TANMEN & "', "                                        ' 端面角度
        sql = sql & "'" & WARPRANK & "', '"                                     ' ワープランク
        sql = sql & CRYSTALMEN & "', "                                          ' 結晶面
        sql = sql & "nvl(to_char(abs(SUB.E022HWFCSCEN),'fm0.00'),'0.00'), "     ' 傾中心
        sql = sql & "nvl(to_char(SUB.E022HWFCSMIN,'fm0.00'),'0.00'), "          ' 傾下限
        sql = sql & "nvl(to_char(SUB.E022HWFCSMAX,'fm0.00'),'0.00'), "          ' 傾上限
        sql = sql & "SUB.E022HWFCKWAY, "                                        ' 検査方法
        sql = sql & "SUB.E022HWFCKHNM || SUB.E022HWFCKHNN || SUB.E022HWFCKHNH || SUB.E022HWFCKHNU, "    ' 検査頻度（枚、抜、保、ウの順で足す）
        sql = sql & "SUB.E022HWFCSDIR, "                                        ' 傾方位
        sql = sql & "SUB.E022HWFCSDIS, "                                        ' 傾方位指定
        sql = sql & "SUB.E022HWFCTDIR, "                                        ' 傾縦方位
        sql = sql & "nvl(to_char(SUB.E022HWFCTCEN,'fm0.00'),'0.00'), "          ' 傾縦中心
        sql = sql & "nvl(to_char(SUB.E022HWFCTMIN,'fm0.00'),'0.00'), "          ' 傾縦下限
        sql = sql & "nvl(to_char(SUB.E022HWFCTMAX,'fm0.00'),'0.00'), "          ' 傾縦上限
        sql = sql & "SUB.E022HWFCYDIR, "                                        ' 傾横方位
        sql = sql & "nvl(to_char(SUB.E022HWFCYCEN,'fm0.00'),'0.00'), "          ' 傾横中心
        sql = sql & "nvl(to_char(SUB.E022HWFCYMIN,'fm0.00'),'0.00'), "          ' 傾横下限
        sql = sql & "nvl(to_char(SUB.E022HWFCYMAX,'fm0.00'),'0.00'), "          ' 傾横上限
        sql = sql & SEED & ", "                                                 ' 引上時使用したシ－ド傾き
        sql = sql & "'TX850I', '"                                               ' トランザクションID
        sql = sql & .sBlockID & "', "                                           ' 先頭ブロックID
        sql = sql & .BLOCKORDER & ", "                                          ' ブロック順序
        sql = sql & "sysdate, "                                                 ' 登録日付
        sql = sql & "'0', "                                                     ' 送信フラグ
        sql = sql & "sysdate,"                                                  ' 送信日付
        sql = sql & "'" & sCmbMukesaki & "', "                                  ' 向先 2007/08/31 SPK Tsutsumi Add

    '構成品番を追加 07/07/19 SHINDO STR=======================================>
        sql = sql & KOSCNT & ""
       For c0 = 1 To 5
        If c0 <= KOSCNT Then
            sql = sql & ",'" & KOSEIHIN(c0).KHINBAN & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINPOS & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINPOS + KOSEIHIN(c0).KHINLEN & "'"
            sql = sql & ",'" & KOSEIHIN(c0).KHINLEN & "'"
        Else
            sql = sql & ",NULL"
            sql = sql & ",NULL"
            sql = sql & ",NULL"
            sql = sql & ",NULL"
        End If
       Next c0

   '構成品番を追加 07/07/19 SHINDO STR=======================================>


        sql = sql & " from ("
        sql = sql & "select "
        sql = sql & "E021HWFTYPE, "
        sql = sql & "E021HWFRMAX, "
        sql = sql & "E021HWFRMIN, "
        sql = sql & "E025HWFONMAX, "
        sql = sql & "E025HWFONMIN "
        sql = sql & "from VECME001 "
        sql = sql & "where E018HINBAN='" & MainHin.hinban & "' "
        sql = sql & "and E018MNOREVNO=" & MainHin.mnorevno & " "
        sql = sql & "and E018FACTORY='" & MainHin.factory & "' "
        sql = sql & "and E018OPECOND='" & MainHin.opecond & "' "
        sql = sql & ") MAIN, "
        sql = sql & "("
        sql = sql & "select "
        sql = sql & "E022HWFCSCEN, "
        sql = sql & "E022HWFCSMIN, "
        sql = sql & "E022HWFCSMAX, "
        sql = sql & "E022HWFCKWAY, "
        sql = sql & "E022HWFCKHNM, "
        sql = sql & "E022HWFCKHNN, "
        sql = sql & "E022HWFCKHNH, "
        sql = sql & "E022HWFCKHNU, "
        sql = sql & "E022HWFCSDIR, "
        sql = sql & "E022HWFCSDIS, "
        sql = sql & "E022HWFCTDIR, "
        sql = sql & "E022HWFCTCEN, "
        sql = sql & "E022HWFCTMIN, "
        sql = sql & "E022HWFCTMAX, "
        sql = sql & "E022HWFCYDIR, "
        sql = sql & "E022HWFCYCEN, "
        sql = sql & "E022HWFCYMIN, "
        sql = sql & "E022HWFCYMAX "
        sql = sql & "from VECME001 "
        sql = sql & "where E018HINBAN='" & SubHin.hinban & "' "
        sql = sql & "and E018MNOREVNO=" & SubHin.mnorevno & " "
        sql = sql & "and E018FACTORY='" & SubHin.factory & "' "
        sql = sql & "and E018OPECOND='" & SubHin.opecond & "' "
        sql = sql & ") SUB "
        '代表品番、ｻﾌﾞ代表品番の仕様を取得　05/11/25 ooba END ================================>





    End With
    If OraDB.ExecuteSQL(sql) <= 0 Then
        DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    sDbName = ""
    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_BlockNewInf_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :内部関数：端面角度を求める
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pBlkMap　　　,I  ,typ_BlkMap     　,ブロック一覧
'      　　:SubHinban　　,O  ,tFullHinban      ,ｻﾌﾞ代表品番　05/11/25 ooba
'      　　:ans    　　　,O  ,String         　,端面角度
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2002/04/17  佐野 信哉 作成
Public Function DBDRV_getTANMEN(pBlkMap As typ_BlkMap, SubHinban As tFullHinban, Ans As String) As FUNCTION_RETURN
    Dim rs              As OraDynaset
    Dim sql             As String
    Dim SQLHIN          As String
    Dim tHin(5)         As tFullHinban       ' 品番
    Dim tHSXCSCEN(5)    As String * 3
    Dim c0              As Integer
    Dim c1              As Integer
    Dim RecCount        As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_getTANMEN"
    DBDRV_getTANMEN = FUNCTION_RETURN_FAILURE

    SQLHIN = SQLMake_HINBAN(pBlkMap.HIN())
    'NULL対応のため、HSXCSMAX・HSXCSMINの項目を追加
    sql = "select HSXCSCEN, HSXCSMIN, HSXCSMAX, HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME018 where "
    sql = sql & "(ABS(HSXCSMAX - HSXCSMIN) = (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 where " & SQLHIN & ")) and "
    sql = sql & SQLHIN
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    RecCount = rs.RecordCount
    If RecCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If

    For c0 = 1 To RecCount
        If IsNull(rs("HSXCSCEN")) Or IsNull(rs("HSXCSMIN")) Or IsNull(rs("HSXCSMAX")) Then
            DBDRV_getTANMEN = FUNCTION_RETURN_FAILURE
            GoTo proc_err
        End If
        tHSXCSCEN(c0) = fncNullCheck(rs("HSXCSCEN"))
        tHin(c0).factory = rs("FACTORY")
        tHin(c0).hinban = rs("HINBAN")
        tHin(c0).mnorevno = rs("MNOREVNO")
        tHin(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close
    '複数存在した場合、最も上側の品番を採用し端面角度を求める。
    Ans = ""
'    For c0 = 1 To RecCount
    For c0 = 1 To 5     '06/01/19 ooba
        If Trim(pBlkMap.HIN(c0).hinban) <> "" Then
            For c1 = 1 To RecCount
                If (pBlkMap.HIN(c0).factory = tHin(c1).factory) And _
                   (pBlkMap.HIN(c0).hinban = tHin(c1).hinban) And _
                   (pBlkMap.HIN(c0).mnorevno = tHin(c1).mnorevno) And _
                   (pBlkMap.HIN(c0).opecond = tHin(c1).opecond) Then
                    Ans = tHSXCSCEN(c1)
                    SubHinban.hinban = tHin(c1).hinban          '05/11/25 ooba START =====>
                    SubHinban.mnorevno = tHin(c1).mnorevno
                    SubHinban.factory = tHin(c1).factory
                    SubHinban.opecond = tHin(c1).opecond        '05/11/25 ooba END =======>
                    Exit For
                End If
            Next
        End If
        If Ans <> "" Then Exit For
    Next
    DBDRV_getTANMEN = FUNCTION_RETURN_SUCCESS
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

'概要      :内部関数：ワープランクを求める
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pBlkMap　　　,I  ,typ_BlkMap     　,ブロック一覧
'      　　:MainHinban　 ,O  ,tFullHinban      ,代表品番　05/11/25 ooba
'      　　:ans    　　　,O  ,String         　,端面角度
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2002/04/17  佐野 信哉 作成
Public Function DBDRV_getWARPRANK(pBlkMap As typ_BlkMap, MainHinban As tFullHinban, Ans As String) As FUNCTION_RETURN
    Dim rs  As OraDynaset
    Dim sql As String
    Dim SQLHIN          As String           '05/11/25 ooba START ========>
    Dim tHin(5)         As tFullHinban
    Dim c0              As Integer
    Dim c1              As Integer
    Dim RecCount        As Integer          '05/11/25 ooba END ==========>

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_getWARPRANK"
    DBDRV_getWARPRANK = FUNCTION_RETURN_FAILURE

    '初期化　06/04/28 ooba
    For c0 = 1 To 5
        tHin(c0).hinban = ""
        tHin(c0).mnorevno = 0
        tHin(c0).factory = ""
        tHin(c0).opecond = ""
    Next c0

''    sql = "select max(HWFWARPR) as maxHWFWARPR from TBCME027"
''    sql = sql & " where " & SQLMake_HINBAN(pBlkMap.HIN())
''
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    If rs.RecordCount <= 0 Then
''        rs.Close
''        GoTo proc_exit
''    End If
''    ans = rs("maxHWFWARPR")
''    rs.Close

    'ﾜｰﾌﾟﾗﾝｸが最大の品番を代表品番とする　05/11/25 ooba START ==============================>
    SQLHIN = SQLMake_HINBAN(pBlkMap.HIN())

    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFWARPR from TBCME027 where "
''    sql = sql & "HWFWARPR = (select MAX(HWFWARPR) from TBCME027 where " & SQLHIN & ") and "
''    sql = sql & SQLHIN
    '①ﾅﾉﾄﾎﾟﾌﾗｸﾞ(0:ｶﾞﾗｽ接着無し,1:ｶﾞﾗｽ接着有り)が最大な品番の中で           06/01/19 ooba
    '合成角の規格幅(結晶面傾上限-結晶面傾下限)が最小の品番の中で            06/07/19 kondoh Add
    'ﾜｰﾌﾟﾗﾝｸが最大の品番                                                    06/01/19 ooba
    sql = sql & "HWFWARPR = (select MAX(HWFWARPR) from TBCME027 "
    sql = sql & "            where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "

''06/07/19 SMP)kondoh START Add =========================================================>
    sql = sql & "               ("
    sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
    sql = sql & "               from TBCME018 "
    sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
    sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
    sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                               from TBCME036 where " & SQLHIN
    sql = sql & "                               ) "
    sql = sql & "                           and " & SQLHIN
    sql = sql & "                           ) "
    sql = sql & "                       ) "
    sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                          from TBCME036 where " & SQLHIN
    sql = sql & "                         )"
    sql = sql & "                   and " & SQLHIN
    sql = sql & "                  ) "
    sql = sql & "               ) "
''06/07/19 SMP)kondoh END Add =========================================================>

''06/07/19 SMP)kondoh START Del =========================================================>
''    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
''    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
''    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
''    sql = sql & "                          from TBCME036 where " & SQLHIN
''    sql = sql & "                         ) "
''    sql = sql & "                   and " & SQLHIN
''    sql = sql & "                  ) "
''06/07/19 SMP)kondoh END Del =========================================================>

    sql = sql & "           ) "
    sql = sql & "and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "

''06/07/19 SMP)kondoh START Add =========================================================>
    sql = sql & "               ("
    sql = sql & "               select HINBAN, MNOREVNO, FACTORY, OPECOND "
    sql = sql & "               from TBCME018 "
    sql = sql & "               where ABS(HSXCSMAX - HSXCSMIN) = "
    sql = sql & "                       (select MIN(ABS(HSXCSMAX - HSXCSMIN)) from TBCME018 "
    sql = sql & "                      where (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                           (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                           where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                               (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                               from TBCME036 where " & SQLHIN
    sql = sql & "                               ) "
    sql = sql & "                           and " & SQLHIN
    sql = sql & "                           ) "
    sql = sql & "                       ) "
    sql = sql & "                and (HINBAN, MNOREVNO, FACTORY, OPECOND) in "
    sql = sql & "                  (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
    sql = sql & "                   where decode(GLASS,null,'0',' ','0',GLASS) = "
    sql = sql & "                         (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
    sql = sql & "                          from TBCME036 where " & SQLHIN
    sql = sql & "                         )"
    sql = sql & "                   and " & SQLHIN
    sql = sql & "                  ) "
    sql = sql & "               ) "
''06/07/19 SMP)kondoh END Add =========================================================>

''06/07/19 SMP)kondoh START Del =========================================================>
''    sql = sql & "    (select HINBAN, MNOREVNO, FACTORY, OPECOND from TBCME036 "
''    sql = sql & "     where decode(GLASS,null,'0',' ','0',GLASS) = "
''    sql = sql & "           (select MAX(decode(GLASS,null,'0',' ','0',GLASS)) "
''    sql = sql & "            from TBCME036 where " & SQLHIN
''    sql = sql & "           ) "
''    sql = sql & "     and " & SQLHIN
''    sql = sql & "    )"
''06/07/19 SMP)kondoh END Del =========================================================>

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    RecCount = rs.RecordCount
    If RecCount <= 0 Then
        rs.Close
        GoTo proc_exit
    End If

    Ans = rs("HWFWARPR")

    For c0 = 1 To RecCount
        tHin(c0).hinban = rs("HINBAN")
        tHin(c0).factory = rs("FACTORY")
        tHin(c0).mnorevno = rs("MNOREVNO")
        tHin(c0).opecond = rs("OPECOND")
        rs.MoveNext
    Next
    rs.Close


''06/07/19 SMP)kondoh START Del =========================================================>
''    ''06/04/28 ooba START =========================================================>
''    '①を満たす中でﾅﾉﾄﾎﾟ規格が一番厳しい(品WFﾅﾉﾄﾎﾟ2上限が一番小さい)品番
''    SQLHIN = SQLMake_HINBAN(tHin())
''
''    sql = "select HINBAN, MNOREVNO, FACTORY, OPECOND, HWFNP2MAX "
''    sql = sql & "from TBCME026 "
''    sql = sql & "where nvl(HWFNP2MAX,999.99) = (select min(nvl(HWFNP2MAX,999.99)) "
''    sql = sql & "                               from TBCME026 "
''    sql = sql & "                               where " & SQLHIN & ") "
''    sql = sql & "and " & SQLHIN
''
''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
''    RecCount = rs.RecordCount
''    If RecCount <= 0 Then
''        rs.Close
''        GoTo proc_exit
''    End If
''
''    For c0 = 1 To RecCount
''        tHin(c0).hinban = rs("HINBAN")
''        tHin(c0).factory = rs("FACTORY")
''        tHin(c0).mnorevno = rs("MNOREVNO")
''        tHin(c0).opecond = rs("OPECOND")
''        rs.MoveNext
''    Next
''    rs.Close
''    ''06/04/28 ooba END ===========================================================>
''06/07/19 SMP)kondoh END Del =========================================================>

    MainHinban.hinban = ""
    '複数存在した場合、最も上側の品番を代表品番とする。
'    For c0 = 1 To RecCount
    For c0 = 1 To 5     '06/01/19 ooba
        If Trim(pBlkMap.HIN(c0).hinban) <> "" Then
            For c1 = 1 To RecCount
                If (pBlkMap.HIN(c0).hinban = tHin(c1).hinban) And _
                   (pBlkMap.HIN(c0).mnorevno = tHin(c1).mnorevno) And _
                   (pBlkMap.HIN(c0).factory = tHin(c1).factory) And _
                   (pBlkMap.HIN(c0).opecond = tHin(c1).opecond) Then

                    MainHinban.hinban = tHin(c1).hinban
                    MainHinban.mnorevno = tHin(c1).mnorevno
                    MainHinban.factory = tHin(c1).factory
                    MainHinban.opecond = tHin(c1).opecond
                    Exit For
                End If
            Next
        End If
        If Trim(MainHinban.hinban) <> "" Then Exit For
    Next
    'ﾜｰﾌﾟﾗﾝｸが最大の品番を代表品番とする　05/11/25 ooba END ================================>

    DBDRV_getWARPRANK = FUNCTION_RETURN_SUCCESS
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

'概要      :内部関数：仕様値がNULLだった場合、Insert文が発行できないようにエラーで終了する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pBlkMap　　　,I  ,typ_BlkMap     　,ブロック一覧
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2003/12/12  システムブレイン 作成
Public Function DBDRV_NULLChk(pBlkMap As typ_BlkMap) As FUNCTION_RETURN
    Dim rs  As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_NULLChk"
    DBDRV_NULLChk = FUNCTION_RETURN_FAILURE

    'NUMBER型のデータを読み込み、NULLだった場合はエラーとする
    sql = ""
    With pBlkMap
        sql = sql & "select E021HWFRMAX, E021HWFRMIN,"
        sql = sql & "       E025HWFONMAX, E025HWFONMIN,"
        sql = sql & "       E022HWFCSCEN, E022HWFCSMIN, E022HWFCSMAX,"
        sql = sql & "       E022HWFCTCEN, E022HWFCTMIN, E022HWFCTMAX,"
        sql = sql & "       E022HWFCYCEN, E022HWFCYMIN, E022HWFCYMAX from VECME001"
        sql = sql & " where E018HINBAN='" & .HIN(1).hinban & "'"
        sql = sql & " and E018MNOREVNO=" & .HIN(1).mnorevno
        sql = sql & " and E018FACTORY='" & .HIN(1).factory & "'"
        sql = sql & " and E018OPECOND='" & .HIN(1).opecond & "'"
    End With

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    '取得エラー、または１つでもNULLがあったらエラーとする
    If rs.RecordCount <= 0 Or _
       IsNull(rs("E021HWFRMAX")) Or IsNull(rs("E021HWFRMIN")) Or _
       IsNull(rs("E025HWFONMAX")) Or IsNull(rs("E025HWFONMIN")) Or _
       IsNull(rs("E022HWFCSCEN")) Or IsNull(rs("E022HWFCSMIN")) Or IsNull(rs("E022HWFCSMAX")) Or _
       IsNull(rs("E022HWFCTCEN")) Or IsNull(rs("E022HWFCTMIN")) Or IsNull(rs("E022HWFCTMAX")) Or _
       IsNull(rs("E022HWFCYCEN")) Or IsNull(rs("E022HWFCYMIN")) Or IsNull(rs("E022HWFCYMAX")) Then

        DBDRV_NULLChk = FUNCTION_RETURN_FAILURE
        rs.Close
        GoTo proc_exit
    End If

    DBDRV_NULLChk = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    rs.Close
    gErr.Pop
    Exit Function
proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function
''''''------------------------------------------------
'''''' DBアクセス関数
''''''------------------------------------------------
'''''
''''''概要      :テーブル「TBCME037」から条件にあったレコードを抽出する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :records()     ,O  ,typ_TBCME037 ,抽出レコード
''''''          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
''''''          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
''''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''''''説明      :
''''''履歴      :2001/08/24作成　野村 (2002/07 s_cmzcF_TBCME037_SQL.basより移動)
'''''Public Function DBDRV_GetTBCME037(records() As typ_TBCME037, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL全体
'''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      'レコード数
'''''Dim i As Long
'''''
'''''    ''SQLを組み立てる
'''''    sqlBase = "Select CRYNUM, DELCLS, KRPROCCD, PROCCD, LPKRPROCCD, LASTPASS, RPHINBAN, RPREVNUM, RPFACT, RPOPCOND, PRODCOND," & _
'''''              " PGID, UPLENGTH, TOPLENG, BODYLENG, BOTLENG, FREELENG, DIAMETER, CHARGE, SEED, ADDDPCLS, ADDDPPOS, ADDDPVAL," & _
'''''              " REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCME037"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''データを抽出する
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME037 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''抽出結果を格納する
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' 結晶番号
'''''            .DELCLS = rs("DELCLS")           ' 削除区分
'''''            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
'''''            .PROCCD = rs("PROCCD")           ' 工程コード
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
'''''            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
'''''            .RPHINBAN = rs("RPHINBAN")       ' ねらい品番
'''''            .RPREVNUM = rs("RPREVNUM")       ' ねらい品番製品番号改訂番号
'''''            .RPFACT = rs("RPFACT")           ' ねらい品番工場
'''''            .RPOPCOND = rs("RPOPCOND")       ' ねらい品番操業条件
'''''            .PRODCOND = rs("PRODCOND")       ' 製作条件
'''''            .PGID = rs("PGID")               ' ＰＧ－ＩＤ
'''''            .UPLENGTH = rs("UPLENGTH")       ' 引上げ長さ
'''''            .TOPLENG = rs("TOPLENG")         ' ＴＯＰ長さ
'''''            .BODYLENG = rs("BODYLENG")       ' 直胴長さ
'''''            .BOTLENG = rs("BOTLENG")         ' ＢＯＴ長さ
'''''            .FREELENG = rs("FREELENG")       ' フリー長
'''''            .DIAMETER = rs("DIAMETER")       ' 直径
'''''            .CHARGE = rs("CHARGE")           ' チャージ量
'''''            .SEED = rs("SEED")               ' シード
'''''            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドープ種類
'''''            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
'''''            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
'''''            .REGDATE = rs("REGDATE")         ' 登録日付
'''''            .UPDDATE = rs("UPDDATE")         ' 更新日付
'''''            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'''''            .SENDDATE = rs("SENDDATE")       ' 送信日付
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME037 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DBアクセス関数
''''''------------------------------------------------
'''''
''''''概要      :テーブル「TBCME040」から条件にあったレコードを抽出する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :records()     ,O  ,typ_TBCME040 ,抽出レコード
''''''          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
''''''          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
''''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''''''説明      :
''''''履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCME040_SQL.basより移動)
'''''Public Function DBDRV_GetTBCME040(records() As typ_TBCME040, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL全体
'''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      'レコード数
'''''Dim i As Long
'''''
'''''    ''SQLを組み立てる
'''''    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, REALLEN, BLOCKID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS," & _
'''''              " RSTATCLS, HOLDCLS, BDCAUS, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE," & _
'''''              " PASSFLAG "   '02/07/05 hama
'''''
'''''    sqlBase = sqlBase & "From TBCME040"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''データを抽出する
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME040 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''抽出結果を格納する
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' 結晶番号
'''''            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
'''''            .LENGTH = rs("LENGTH")           ' 長さ
'''''            .REALLEN = rs("REALLEN")         ' 実長さ
'''''            .BLOCKID = rs("BLOCKID")         ' ブロックID
'''''            .KRPROCCD = rs("KRPROCCD")       ' 現在管理工程
'''''            .NOWPROC = rs("NOWPROC")         ' 現在工程
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
'''''            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
'''''            .DELCLS = rs("DELCLS")           ' 削除区分
'''''            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
'''''            .RSTATCLS = rs("RSTATCLS")       ' 流動状態区分
'''''            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
'''''            .BDCAUS = rs("BDCAUS")           ' 不良理由
'''''            .REGDATE = rs("REGDATE")         ' 登録日付
'''''            .UPDDATE = rs("UPDDATE")         ' 更新日付
'''''            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'''''            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'''''            .SENDDATE = rs("SENDDATE")       ' 送信日付
'''''            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/07/05 hama
'''''             If rs("PASSFLAG") = "1" Then
'''''                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/07/05 hama
'''''            End If
'''''
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME040 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DBアクセス関数
''''''------------------------------------------------
'''''
''''''概要      :テーブル「TBCME042」から条件にあったレコードを抽出する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :records()     ,O  ,typ_TBCME042 ,抽出レコード
''''''          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
''''''          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
''''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''''''説明      :
''''''履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME042_SQL.basより移動)
'''''Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL全体
'''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      'レコード数
'''''Dim i As Long
'''''
'''''    ''SQLを組み立てる
'''''    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
'''''              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
'''''              " PASSFLAG "   '02/04/16 Yam
'''''    sqlBase = sqlBase & "From TBCME042"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''データを抽出する
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCME042 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''抽出結果を格納する
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' 結晶番号
'''''            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
'''''            .LENGTH = rs("LENGTH")           ' 長さ
'''''            .SXLID = rs("SXLID")             ' SXLID
'''''            .KRPROCCD = rs("KRPROCCD")       ' 管理工程
'''''            .NOWPROC = rs("NOWPROC")         ' 現在工程
'''''            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
'''''            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
'''''            .DELCLS = rs("DELCLS")           ' 削除区分
'''''            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
'''''            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
'''''            .hinban = rs("HINBAN")           ' 品番
'''''            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
'''''            .factory = rs("FACTORY")         ' 工場
'''''            .opecond = rs("OPECOND")         ' 操業条件
'''''            .BDCAUS = rs("BDCAUS")           ' 不良理由
'''''            .COUNT = rs("COUNT")             ' 枚数
'''''            .REGDATE = rs("REGDATE")         ' 登録日付
'''''            .UPDDATE = rs("UPDDATE")         ' 更新日付
'''''            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
'''''            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'''''            .SENDDATE = rs("SENDDATE")       ' 送信日付
'''''            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/04/16 Yam
'''''            If rs("PASSFLAG") = "1" Then
'''''                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/04/05 Yam
'''''            End If
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCME042 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''
'''''
''''''------------------------------------------------
'''''' DBアクセス関数
''''''------------------------------------------------
'''''
''''''概要      :テーブル「TBCMW001」から条件にあったレコードを抽出する
''''''ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
''''''          :records()     ,O  ,typ_TBCMW001 ,抽出レコード
''''''          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
''''''          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
''''''          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
''''''説明      :
''''''履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCMW001_SQL.basより移動)
'''''Public Function DBDRV_GetTBCMW001(records() As typ_TBCMW001, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
'''''Dim sql As String       'SQL全体
'''''Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'''''Dim rs As OraDynaset    'RecordSet
'''''Dim recCnt As Long      'レコード数
'''''Dim i As Long
'''''
'''''    ''SQLを組み立てる
'''''    sqlBase = "Select CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, PROCCODE, BLOCKID, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
'''''              " SENDFLAG, SENDDATE "
'''''    sqlBase = sqlBase & "From TBCMW001"
'''''    sql = sqlBase
'''''    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
'''''        sql = sql & " " & sqlWhere & " " & sqlOrder
'''''    End If
'''''
'''''    ''データを抽出する
'''''    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'''''    If rs Is Nothing Then
'''''        ReDim records(0)
'''''        DBDRV_GetTBCMW001 = FUNCTION_RETURN_FAILURE
'''''        Exit Function
'''''    End If
'''''
'''''    ''抽出結果を格納する
'''''    recCnt = rs.RecordCount
'''''    ReDim records(recCnt)
'''''    For i = 1 To recCnt
'''''        With records(i)
'''''            .Crynum = rs("CRYNUM")           ' 結晶番号
'''''            .INGOTPOS = rs("INGOTPOS")       ' インゴット位置
'''''            .TRANCNT = rs("TRANCNT")         ' 処理回数
'''''            .CRYLEN = rs("CRYLEN")           ' 長さ
'''''            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
'''''            .PROCCODE = rs("PROCCODE")       ' 工程コード
'''''            .BLOCKID = rs("BLOCKID")         ' ブロックID
'''''            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
'''''            .REGDATE = rs("REGDATE")         ' 登録日付
'''''            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
'''''            .UPDDATE = rs("UPDDATE")         ' 更新日付
'''''            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
'''''            .SENDDATE = rs("SENDDATE")       ' 送信日付
'''''        End With
'''''        rs.MoveNext
'''''    Next
'''''    rs.Close
'''''
'''''    DBDRV_GetTBCMW001 = FUNCTION_RETURN_SUCCESS
'''''End Function
'''''

'概要      :関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)登録
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型                 ,説明
'      　　:KblkData()  ,I  ,typ_BlkMap         ,関連ﾌﾞﾛｯｸﾃﾞｰﾀ
'      　　:戻り値      ,O  ,FUNCTION_RETURN　  ,書き込みの成否
'説明      :
'履歴      :07/12/21 ooba
Public Function DBDRV_KanrenBlk(KblkData() As typ_BlkMap) As FUNCTION_RETURN

    Dim sql             As String
    Dim i               As Long
    Dim rs              As OraDynaset
    Dim lRecCnt         As Long             'ﾚｺｰﾄﾞ数
    Dim KanrenData()    As typ_TBCMY023     '関連ﾌﾞﾛｯｸ紐付紐切ﾃﾞｰﾀ
    Dim iTrnCnt         As Integer          '処理回数


    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc034_SQL.bas -- Function DBDRV_KanrenBlk"

    DBDRV_KanrenBlk = FUNCTION_RETURN_FAILURE

    '処理回数取得
    sql = "SELECT NVL(MAX(TRANCNT),0) MAXCNT FROM TBCMY023"
    sql = sql & " WHERE CRYNUM = '" & KblkData(1).CRYNUM & "'"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)

    If rs.RecordCount = 0 Then
        iTrnCnt = 1
    Else
        iTrnCnt = rs("MAXCNT") + 1          '処理回数(最大) + 1
    End If
    rs.Close


    lRecCnt = UBound(KblkData)              '登録ﾚｺｰﾄﾞ数
    ReDim KanrenData(lRecCnt)

    '関連ﾌﾞﾛｯｸ情報を関連ﾌﾞﾛｯｸ紐付紐切(TBCMY023)に登録
    For i = 1 To lRecCnt
        With KanrenData(i)
            .CRYNUM = KblkData(i).CRYNUM        '結晶番号
            .TRANCNT = iTrnCnt                  '処理回数
            .BLOCKID = KblkData(i).BLOCKID      'ﾌﾞﾛｯｸID
            .PROCCAT = "N"                      '処理区分(N:新規)
            .TXID = "TX879I"                    'ﾄﾗﾝｻﾞｸｼｮﾝID
            
            sql = "INSERT INTO TBCMY023"
            sql = sql & " (CRYNUM,"
            sql = sql & " TRANCNT,"
            sql = sql & " BLOCKID,"
            sql = sql & " PROCCAT,"
            sql = sql & " TXID,"
            sql = sql & " REGDATE,"
            sql = sql & " SENDFLAG,"
            sql = sql & " SENDDATE,"
            sql = sql & " PLANTCAT,"
            sql = sql & " SUMITFLAG,"
            sql = sql & " SUMITSND,"
            sql = sql & " SSENDNO) "
            sql = sql & " VALUES"
            sql = sql & " ('" & .CRYNUM & "',"          '結晶番号
            sql = sql & .TRANCNT & ","                  '処理回数
            sql = sql & " '" & .BLOCKID & "',"          'ﾌﾞﾛｯｸID
            sql = sql & " '" & .PROCCAT & "',"          '処理区分
            sql = sql & " '" & .TXID & "',"             'ﾄﾗﾝｻﾞｸｼｮﾝID
            sql = sql & " SYSDATE,"                     '登録日付
            sql = sql & " '5',"                         '送信ﾌﾗｸﾞ(5:WF送信対象外)
            sql = sql & " NULL, "                       '送信日付
            sql = sql & "  '" & sCmbMukesaki & "', "    '向先
            sql = sql & " '0',"                         'SUMIT送信ﾌﾗｸﾞ
            sql = sql & " NULL,"                        'SUMIT送信日付
            sql = sql & " NULL) "                       '送信順連番
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
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit

End Function

