Attribute VB_Name = "s_cmbc016_SQL"
Option Explicit

' 結晶加工払出

' 切断仕様 (by SUMCO)
Public Type typ_CutSpec1
    hin As tFullHinban          ' 品番
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
    hin As tFullHinban          ' 品番
    HSXTYPE As String * 1       ' タイプ
    HSXCDIR As String * 1       ' 方位
    HSXD1CEN As Double          ' 直径
    HSXDOP As String * 1        ' 結晶ドープ
    HSXDPDIR As String * 2      ' ノッチ位置
    HSXDDMIN As Double          ' ノッチ深さ（ＭＩＮ）
    HSXDDMAX As Double          ' ノッチ深さ（ＭＡＸ）
    HSXSDSLP As Integer         ' シード傾き
' 払出規制項目追加対応 yakimura 2002.12.01 start
    TOPREG As Integer           ' TOP規制
    TAILREG As Double           ' TAIL規制
    BTMSPRT As Integer          ' ボトム析出規制
' 払出規制項目追加対応 yakimura 2002.12.01 end
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

'お勧め検査の各メンバの意味
Public Enum RCMD_COL
    RCMD_RS = 1
    RCMD_OI
    RCMD_B1
    RCMD_B2
    RCMD_B3
    RCMD_L1
    RCMD_L2
    RCMD_L3
    RCMD_L4
    RCMD_CS
    RCMD_GD
    RCMD_LT
    RCMD_EPD
End Enum

'優先順位格納用変数
Public CUT_PRIORITY As String * 1


'概要      :結晶加工払出用 結晶番号入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sCryNum 　　　,I  ,String         　,結晶番号
'      　　:pCryInf 　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'      　　:pPupEnd 　　　,O  ,typ_TBCMH004   　,引上げ終了実績
'      　　:pHinSpec　　　,O  ,typ_HinSpec1   　,製品仕様
'      　　:pCutInd 　　　,O  ,typ_CutInd     　,切断指示
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001b_Disp(sCryNum As String, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, pHinSpec() As typ_HinSpec1, _
                                           pCutInd() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim ctcen As Double
    Dim cycen As Double

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Disp"
    sErrMsg = ""

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' 工程チェック
    If pCryInf.PROCCD <> PROCD_KAKOU_HARAIDASI Then
        sErrMsg = GetMsgStr("EPRC0")
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 引上げ終了実績の取得(s_cmzcTBCMH004_SQL.bas が必要)
    sDbName = "H004"
    If DBDRV_GetTBCMH004(tmpPupEnd(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpPupEnd) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pPupEnd = tmpPupEnd(1)

    '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    recCnt = UBound(pHinDsn)
    If recCnt = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 製品仕様の取得
' 払出規制項目追加対応 yakimura 2002.12.01 start
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).HINBAN)
        If sHin <> "G" And sHin <> "Z" Then
            sql = "select "
            sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
            sql = sql & " ,NVL(TOPREG,0) TOPREG, NVL(TAILREG,0) TAILREG, NVL(BTMSPRT,0) BTMSPRT "
            sql = sql & " from TBCME018 E018,TBCME036 E036"
            sql = sql & " where E018.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and E018.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and E018.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and E018.OPECOND='" & pHinDsn(i).OPCOND & "'"
            sql = sql & " and E036.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and E036.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and E036.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and E036.OPECOND='" & pHinDsn(i).OPCOND & "'"
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDbName)
                DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            j = j + 1
            With pHinSpec(j)
                .hin.HINBAN = pHinDsn(i).HINBAN
                .hin.mnorevno = pHinDsn(i).REVNUM
                .hin.factory = pHinDsn(i).FACT
                .hin.opecond = pHinDsn(i).OPCOND
                .HSXTYPE = rs("HSXTYPE")    ' タイプ
                .HSXCDIR = rs("HSXCDIR")    ' 方位
                .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' 直径
                .HSXDOP = rs("HSXDOP")      ' 結晶ドープ
                .HSXDPDIR = rs("HSXDPDIR")          ' 品ＳＸ溝位置方位
                .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' 品ＳＸ溝深下限
                .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' 品ＳＸ溝深上限
                ctcen = Abs(fncNullCheck(rs("HSXCTCEN")))
                cycen = Abs(fncNullCheck(rs("HSXCYCEN")))
                .TOPREG = rs("TOPREG")              ' TOP規制
                .TAILREG = rs("TAILREG")            ' TAIL規制
                .BTMSPRT = rs("BTMSPRT")            ' ボトム析出規制
                If ((ctcen = 2.83) And (cycen = 2.83)) _
                Or ((ctcen = 4) And (cycen = 0)) _
                Or ((ctcen = 0) And (cycen = 4)) Then
                    .HSXSDSLP = 4
                Else
                    .HSXSDSLP = 0
                End If
            End With
            rs.Close
        End If
    Next i
    ReDim Preserve pHinSpec(j)
' 払出規制項目追加対応 yakimura 2002.12.01 end

    '' ブロック設計の取得
    sDbName = "E038"
    sql = "select INGOTPOS, LENGTH from TBCME038"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' カット位置
            .LENGTH = rs("LENGTH")          ' 長さ
        End With
        rs.MoveNext
    Next i
    rs.Close

    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :結晶加工払出用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:pHinSpec　　　,IO ,typ_HinSpec1   　,製品仕様
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmic001b_GetSpec(pHinSpec As typ_HinSpec1) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim ctcen As Double
    Dim cycen As Double

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_GetSpec"

    '' 製品仕様の取得
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP, HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXCTCEN, HSXCYCEN"
    sql = sql & " from TBCME018"
    sql = sql & " where HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and OPECOND='" & pHinSpec.hin.opecond & "'"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    With pHinSpec
        .HSXTYPE = rs("HSXTYPE")            ' タイプ
        .HSXCDIR = rs("HSXCDIR")            ' 方位
        .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))          ' 直径
        .HSXDOP = rs("HSXDOP")              ' 結晶ドープ
        .HSXDPDIR = rs("HSXDPDIR")          ' 品ＳＸ溝位置方位
        .HSXDDMIN = fncNullCheck(rs("HSXDDMIN"))          ' 品ＳＸ溝深下限
        .HSXDDMAX = fncNullCheck(rs("HSXDDMAX"))          ' 品ＳＸ溝深上限
        ctcen = Abs(fncNullCheck(rs("HSXCTCEN")))
        cycen = Abs(fncNullCheck(rs("HSXCYCEN")))
        If ((ctcen = 2.83) And (cycen = 2.83)) _
        Or ((ctcen = 4) And (cycen = 0)) _
        Or ((ctcen = 0) And (cycen = 4)) Then
            .HSXSDSLP = 4
        Else
            .HSXSDSLP = 0
        End If
    End With
    rs.Close

    DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_scmzc_fcmic001b_GetSpec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :結晶加工払出用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pCryInf　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:pProcBR　　　,I  ,typ_TBCMI001   　,加工払出実績
'      　　:pCutInd　　　,I  ,typ_CutInd     　,切断指示
'      　　:pNotCut　　　,I  ,typ_CutInd     　,切断指示（無切断部）
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DBDRV_scmzc_fcmic001b_Exec(pCryInf As typ_TBCME037, pProcBR As typ_TBCMI001, _
                                           pCutInd() As typ_CutInd, pNotCut() As typ_CutInd, _
                                           newLength As Integer, sErrMsg As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim tmpBlkMng(3) As typ_TBCME040
    Dim sql As String
    Dim sDbName As String
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function DBDRV_scmzc_fcmic001b_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"
    
    '' 結晶情報の更新
    sDbName = "E037"
    With pCryInf

''''''''' pCryInf に入っている内容を使用する
'        sql = "update TBCME037 set "
'        sql = sql & "KRPROCCD='" & MGPRCD_KENNSAKU_KAKOU & "', "
'        sql = sql & "PROCCD='" & PROCD_KENNSAKU_KAKOU & "', "
'        sql = sql & "LPKRPROCCD='" & MGPRCD_KAKOU_HARAIDASI & "', "
'        sql = sql & "LASTPASS='" & PROCD_KAKOU_HARAIDASI & "', "
'        sql = sql & "BODYLENG=" & .BODYLENG & ", "
'        sql = sql & "FREELENG=" & .FREELENG & ", "
'        sql = sql & "SEED='" & .SEED & "', "
'        sql = sql & "UPDDATE=sysdate, "
'        sql = sql & "SENDFLAG='0'"
'        sql = sql & " where CRYNUM='" & .CRYNUM & "'"
        
        sql = "update TBCME037 set "
        sql = sql & "KRPROCCD='" & .KRPROCCD & "', "
        sql = sql & "PROCCD='" & .PROCCD & "', "
        sql = sql & "LPKRPROCCD='" & .LPKRPROCCD & "', "
        sql = sql & "LASTPASS='" & .LASTPASS & "', "
        sql = sql & "BODYLENG=" & .BODYLENG & ", "
        sql = sql & "FREELENG=" & .FREELENG & ", "
        sql = sql & "SEED='" & .SEED & "', "
        sql = sql & "UPDDATE=sysdate, "
        sql = sql & "SENDFLAG='0'"
        sql = sql & " where CRYNUM='" & .Crynum & "'"
    End With
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ブロック設計の更新
    sDbName = "E038"
    sql = "update TBCME038 set "
    sql = sql & "USECLASS='1', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.Crynum, 7) & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 品番設計の更新
    sDbName = "E039"
    sql = "update TBCME039 set "
    sql = sql & "USECLASS='1', "
    sql = sql & "UPDDATE=sysdate, "
    sql = sql & "SENDFLAG='0'"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.Crynum, 7) & "'"
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    '' 品番設計および引上げ指示終了の更新（設計時と最下切断位置が違っていたら更新する）
'    If newLength > 0 Then
        '品番設計の最下品番を、最下切断位置までに伸ばす
'        sDBName = "E039"
'        sql = "update TBCME039 set "
'        sql = sql & "LENGTH=" & newLength & ", "
'        sql = sql & "UPDDATE=sysdate, "
'        sql = sql & "SENDFLAG='0'"
'        sql = sql & " where substr(CRYNUM,1,7)='" & Left(pCryInf.CRYNUM, 7) & "'"
'        sql = sql & "and INGOTPOS=(select max(INGOTPOS) from TBCME039 where "
'        sql = sql & "substr(CRYNUM,1,7)='" & Left(pCryInf.CRYNUM, 7) & "')"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) <= 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
        
        '引上終了の SUMMITSENDFLAG をクリアする
'        sDBName = "H004"
'        sql = "update TBCMH004 set SUMMITSENDFLAG='0' "
'        sql = sql & "where CRYNUM='" & pCryInf.CRYNUM & "'"
'        WriteDBLog sql, sDBName
'        If OraDB.ExecuteSQL(sql) <= 0 Then
'            sErrMsg = GetMsgStr("ENG11", vbNullString, sDBName)
'            DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
'            GoTo proc_exit
'        End If
'    End If

    '' 失敗引上（切断不可）ならフラグを立てる
    bFlag = False
    If UBound(pNotCut) = 1 Then
        If pNotCut(1).INGOTPOS = 0 Then
            bFlag = True
        End If
    End If

    If bFlag = False Then
        '' 切断指示の挿入
        sDbName = "E045"
        recCnt = UBound(pCutInd)
        For i = 1 To recCnt
            With pCutInd(i)
                sql = "insert into TBCME045 "
                sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, STAFFID, "
                sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, STATCLS, BLOCKID, "
                sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, "
                sql = sql & "CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, "
                sql = sql & "CRYINDEP, PRIORITY, PALTNUM, REGDATE, UPDDATE, SENDFLAG, SENDDATE)"
                sql = sql & " select '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & .INGOTPOS & ", "
                sql = sql & "nvl(max(TRANCNT),0)+1, "
                sql = sql & .LENGTH & ", '"
                sql = sql & MGPRCD_KAKOU_HARAIDASI & "', '"
                sql = sql & PROCD_KAKOU_HARAIDASI & "', '"
                sql = sql & pProcBR.TSTAFFID & "', '"
                sql = sql & .HINDN.HINBAN & "', "
                sql = sql & .HINDN.mnorevno & ", '"
                sql = sql & .HINDN.factory & "', '"
                sql = sql & .HINDN.opecond & "', '"
                sql = sql & .BDCAUS & "', "
                sql = sql & "'0', '"
                sql = sql & pCryInf.Crynum & "', '"
                sql = sql & .SMP.CRYINDRS & "', '"
                sql = sql & .SMP.CRYINDOI & "', '"
                sql = sql & .SMP.CRYINDB1 & "', '"
                sql = sql & .SMP.CRYINDB2 & "', '"
                sql = sql & .SMP.CRYINDB3 & "', '"
                sql = sql & .SMP.CRYINDL1 & "', '"
                sql = sql & .SMP.CRYINDL2 & "', '"
                sql = sql & .SMP.CRYINDL3 & "', '"
                sql = sql & .SMP.CRYINDL4 & "', '"
                sql = sql & .SMP.CRYINDCS & "', '"
                sql = sql & .SMP.CRYINDGD & "', '"
                sql = sql & .SMP.CRYINDT & "', '"
                sql = sql & .SMP.CRYINDEP & "', "
                '切断優先順位の格納
                sql = sql & "'" & CUT_PRIORITY & "', '"
                sql = sql & .PALTNUM & "', "
                sql = sql & "sysdate, "
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate"
                sql = sql & " from TBCME045"
                sql = sql & " where CRYNUM='" & pCryInf.Crynum & "'"
                sql = sql & " and INGOTPOS=" & .INGOTPOS
            End With
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i

        '' 切断指示の挿入（無切断部）
        recCnt = UBound(pNotCut)
        For i = 1 To recCnt
            With pNotCut(i)
                sql = "insert into TBCME045 "
                sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, LENGTH, KRPROCCD, PROCCODE, STAFFID, "
                sql = sql & "HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, STATCLS, BLOCKID, "
                sql = sql & "CRYINDRS, CRYINDOI, CRYINDB1, CRYINDB2, CRYINDB3, CRYINDL1, "
                sql = sql & "CRYINDL2, CRYINDL3, CRYINDL4, CRYINDCS, CRYINDGD, CRYINDT, "
                sql = sql & "CRYINDEP, PRIORITY, PALTNUM, REGDATE, UPDDATE, SENDFLAG, SENDDATE)"
                sql = sql & " select '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & .INGOTPOS & ", "
                sql = sql & "nvl(max(TRANCNT),0)+1, "
                sql = sql & .LENGTH & ", '"
                sql = sql & MGPRCD_KAKOU_HARAIDASI & "', '"
                sql = sql & PROCD_KAKOU_HARAIDASI & "', '"
                sql = sql & pProcBR.TSTAFFID & "', "
                sql = sql & "'        ', "
                sql = sql & "0, "
                sql = sql & "' ', "
                sql = sql & "' ', '"
                sql = sql & .BDCAUS & "', "
                sql = sql & "'0', '"
                sql = sql & pCryInf.Crynum & "', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                sql = sql & "'0', "
                '切断優先順位の格納
                sql = sql & "'" & CUT_PRIORITY & "', "
                sql = sql & "'    ', "
                sql = sql & "sysdate, "
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate"
                sql = sql & " from TBCME045"
                sql = sql & " where CRYNUM='" & pCryInf.Crynum & "'"
                sql = sql & " and INGOTPOS=" & .INGOTPOS
            End With
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i
    Else
        '' ブロック管理の挿入
        sDbName = "E040"
        With tmpBlkMng(1)
            '' TOP
            .Crynum = pCryInf.Crynum
            .INGOTPOS = -99
            .LENGTH = pCryInf.TOPLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "TOP"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = "TOP"
        End With
        With tmpBlkMng(2)
            '' BOT
            .Crynum = pCryInf.Crynum
            .INGOTPOS = -100
            .LENGTH = pCryInf.BOTLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "BOT"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = "BOT"
        End With
        With tmpBlkMng(3)
            '' TOP側無切断部
            .Crynum = pCryInf.Crynum
            .INGOTPOS = 0
            .LENGTH = pCryInf.BODYLENG
            .REALLEN = .LENGTH
            .BLOCKID = Left(pCryInf.Crynum, 9) & "0$1"
            .KRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .NOWPROC = nextCd
            .LPKRPROCCD = MGPRCD_KAKOU_HARAIDASI
            .LASTPASS = nowCd
            .DELCLS = "1"
            .LSTATCLS = "H"
            .RSTATCLS = "T"
            .HOLDCLS = "0"
            .BDCAUS = pNotCut(1).BDCAUS
        End With
        For i = 1 To 3
            If DBDRV_BlockMng_Ins(tmpBlkMng(i)) = FUNCTION_RETURN_FAILURE Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        Next i
    End If

    '' 加工払出実績の挿入
    sDbName = "I001"
    With pProcBR
        sql = "insert into TBCMI001 "
        sql = sql & "(CRYNUM, TRANCNT, KRPROCCD, PROCCODE, "
        sql = sql & "UPLENGTH, FREELENG, UPWEIGHT, SEED, PRCMCN, "
        sql = sql & "TSTAFFID, REGDATE, KSTAFFID, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
        sql = sql & " select '"
        sql = sql & .Crynum & "', "
        sql = sql & "nvl(max(TRANCNT),0)+1, '"
        sql = sql & .KRPROCCD & "', '"
        sql = sql & .PROCCODE & "', "
        sql = sql & .UPLENGTH & ", "
        sql = sql & .FREELENG & ", "
        sql = sql & .UPWEIGHT & ", '"
        sql = sql & .SEED & "', '"
        sql = sql & .PRCMCN & "', '"
        sql = sql & .TSTAFFID & "', "
        sql = sql & "sysdate, '"
        sql = sql & .KSTAFFID & "', "
        sql = sql & "'0', "
        sql = sql & "'0', "
        sql = sql & "sysdate"
        sql = sql & " from TBCMI001"
        sql = sql & " where CRYNUM='" & .Crynum & "'"
    End With
    '' WriteDBLog sql, sDbName
    If OraDB.ExecuteSQL(sql) <= 0 Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmic001b_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv SUMCO作成部分 vvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'概要      :結晶加工払出用 結晶番号入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sCryNum 　　　,I  ,String         　,結晶番号
'      　　:pCryInf 　　　,I  ,typ_TBCME037   　,結晶情報
'      　　:pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'      　　:pPupEnd 　　　,O  ,typ_TBCMH004   　,引上げ終了実績
'      　　:pHinSpec　　　,O  ,typ_HinSpec1   　,製品仕様
'      　　:pCutInd 　　　,O  ,typ_CutInd     　,切断指示
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function fcmic001b_Disp(sCryNum As String, _
                                           pCryInf As typ_TBCME037, pHinDsn() As typ_TBCME039, _
                                           pPupEnd As typ_TBCMH004, pHinSpec() As typ_CutSpec1, _
                                           pCutInd() As typ_CutInd, sErrMsg As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpPupEnd() As typ_TBCMH004
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sHin As String
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function fcmic001b_Disp"
    sErrMsg = ""

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ECRY0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("ECRY0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pCryInf = tmpCryInf(1)

    '' 工程チェック
    If pCryInf.PROCCD <> PROCD_KAKOU_HARAIDASI Then
        sErrMsg = GetMsgStr("EPRC0")
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 引上げ終了実績の取得(s_cmzcTBCMH004_SQL.bas が必要)
    sDbName = "H004"
    If DBDRV_GetTBCMH004(tmpPupEnd(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpPupEnd) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    pPupEnd = tmpPupEnd(1)

    '' 品番設計の取得(s_cmzcTBCME039_SQL.bas が必要)
    sDbName = "E039"
    sql = " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    If DBDRV_GetTBCME039(pHinDsn(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    recCnt = UBound(pHinDsn)
    If recCnt = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 製品仕様の取得
    sDbName = "E018"
    j = 0
    ReDim pHinSpec(recCnt)
    For i = 1 To recCnt
        sHin = Trim(pHinDsn(i).HINBAN)
        If sHin <> "G" And sHin <> "Z" Then
            sql = "select "
            'sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXDOP"
            sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP "  '4/3 Yam
            sql = sql & " from TBCME018 A,TBCME020 B"
            'sql = sql & " from TBCME018"
            sql = sql & " where A.HINBAN='" & pHinDsn(i).HINBAN & "'"
            sql = sql & " and A.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and A.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and A.OPECOND='" & pHinDsn(i).OPCOND & "'"
            sql = sql & " and B.HINBAN='" & pHinDsn(i).HINBAN & "'"    '4/3 Yam
            sql = sql & " and B.MNOREVNO=" & pHinDsn(i).REVNUM
            sql = sql & " and B.FACTORY='" & pHinDsn(i).FACT & "'"
            sql = sql & " and B.OPECOND='" & pHinDsn(i).OPCOND & "'"
            
            Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sErrMsg = GetMsgStr("EGET2", sDbName)
                fcmic001b_Disp = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
            j = j + 1
            With pHinSpec(j)
                .hin.HINBAN = pHinDsn(i).HINBAN
                .hin.mnorevno = pHinDsn(i).REVNUM
                .hin.factory = pHinDsn(i).FACT
                .hin.opecond = pHinDsn(i).OPCOND
                .HSXTYPE = rs("HSXTYPE")    ' タイプ
                .HSXCDIR = rs("HSXCDIR")    ' 方位
                .HSXD1CEN = fncNullCheck(rs("HSXD1CEN"))  ' 直径
'                .HSXDOP = rs("HSXDOP")      ' 結晶ドープ
                .HSXCDOP = rs("HSXCDOP")     ' 結晶ドープ  4/2 Yam
            End With
            rs.Close
        End If
    Next i
    ReDim Preserve pHinSpec(j)

    '' ブロック設計の取得
    sDbName = "E038"
    sql = "select INGOTPOS, LENGTH from TBCME038"
    sql = sql & " where substr(CRYNUM,1,7)='" & Left(sCryNum, 7) & "' order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        fcmic001b_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pCutInd(recCnt)
    For i = 1 To recCnt
        With pCutInd(i)
            .INGOTPOS = rs("INGOTPOS")      ' カット位置
            .LENGTH = rs("LENGTH")          ' 長さ
        End With
        rs.MoveNext
    Next i
    rs.Close

    fcmic001b_Disp = FUNCTION_RETURN_SUCCESS

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
    fcmic001b_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

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
    gErr.Push "s_cmzcF_cmbc016_SQL.bas -- Function fcmic001b_GetSpec"

    '' 製品仕様の取得
    sql = "select "
    sql = sql & "HSXTYPE, HSXCDIR, HSXD1CEN, HSXCDOP,"     '4/2 Yam
    sql = sql & "HSXDPDIR, HSXDDMIN, HSXDDMAX, HSXSDSLP,"   '3/7 Yam
    sql = sql & "HSXCTCEN, HSXCYCEN "  '4/2 Yam
    sql = sql & " from TBCME018 A,TBCME020 B"
    sql = sql & " where A.HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and A.MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and A.FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and A.OPECOND='" & pHinSpec.hin.opecond & "'"
    sql = sql & " and B.HINBAN='" & pHinSpec.hin.HINBAN & "'"
    sql = sql & " and B.MNOREVNO=" & pHinSpec.hin.mnorevno
    sql = sql & " and B.FACTORY='" & pHinSpec.hin.factory & "'"
    sql = sql & " and B.OPECOND='" & pHinSpec.hin.opecond & "'"
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

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ SUMCO作成部分 ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^


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
            .Crynum = rs("CRYNUM")           ' 結晶番号
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
            .Crynum = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .HINBAN = rs("HINBAN")           ' 品番
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

'概要      :テーブル「TBCMH004」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMH004 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMH004_SQL.basより移動)
Public Function DBDRV_GetTBCMH004(records() As typ_TBCMH004, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, KRPROCCD, PROCCODE, LENGTOP, LENGTKDO, LENGTAIL, LENGFREE, DM1, DM2, DM3, WGHTTOP, WGHTTKDO," & _
              " WGHTTAIL, WGHTFREE, WGTOPCUT, UPWEIGHT, CHARGE, SEED, STATCLS, JDGECODE, PWTIME, ADDDPPOS, ADDDPCLS, ADDDPVAL," & _
              " ADDDPNAM, TSTAFFID, REGDATE, KSTAFFID, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE "
    sqlBase = sqlBase & "From TBCMH004"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMH004 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .Crynum = rs("CRYNUM")           ' 結晶番号
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
            .PROCCODE = rs("PROCCODE")       ' 工程コード
            .LENGTOP = rs("LENGTOP")         ' 長さ（TOP）
            .LENGTKDO = rs("LENGTKDO")       ' 長さ（直胴）
            .LENGTAIL = rs("LENGTAIL")       ' 長さ（TAIL）
            .LENGFREE = rs("LENGFREE")       ' フリー長さ
            .DM1 = rs("DM1")                 ' 直胴直径１
            .DM2 = rs("DM2")                 ' 直胴直径２
            .DM3 = rs("DM3")                 ' 直胴直径３
            .WGHTTOP = rs("WGHTTOP")         ' 重量（TOP）
            .WGHTTKDO = rs("WGHTTKDO")       ' 重量（直胴）
            .WGHTTAIL = rs("WGHTTAIL")       ' 重量（TAIL)
            .WGHTFREE = rs("WGHTFREE")       ' 重量（フリー長さ）
            .WGTOPCUT = rs("WGTOPCUT")       ' トップカット重量
            .UPWEIGHT = rs("UPWEIGHT")       ' 引上げ重量
            .CHARGE = rs("CHARGE")           ' チャージ量
            .SEED = rs("SEED")               ' シード
            .STATCLS = rs("STATCLS")         ' BOT状況区分
            .JDGECODE = rs("JDGECODE")       ' 判定コード
            .PWTIME = rs("PWTIME")           ' パワー時間
            .ADDDPPOS = rs("ADDDPPOS")       ' 追加ドープ位置
            .ADDDPCLS = rs("ADDDPCLS")       ' 追加ドーパント種類
            .ADDDPVAL = rs("ADDDPVAL")       ' 追加ドープ量
            .ADDDPNAM = rs("ADDDPNAM")       ' 追加ドープ名
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMH004 = FUNCTION_RETURN_SUCCESS
End Function



