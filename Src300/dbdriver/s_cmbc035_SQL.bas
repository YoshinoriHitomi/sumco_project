Attribute VB_Name = "s_cmbc035_SQL"
Option Explicit

' 結晶情報変更

' ブロック情報
Public Type typ_BlkInf2
    BLOCKID As String * 12      ' ブロックID
    LENGTH As Integer           ' 長さ
    REALLEN As Integer          ' 実長さ
    NOWPROC As String * 5       ' 現在工程
    DELFLG As String * 1        ' 削除区分
    TOPBDLN As Integer          ' TOP不良長さ
    TOPBDCS As String * 3       ' TOP不良理由
    TAILBDLN As Integer         ' TAIL不良長さ
    TAILBDCS As String * 3      ' TAIL不良理由
    COF As type_Coefficient     ' 偏析係数計算
End Type

'概要      :結晶情報変更用 ブロックＩＤ入力時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sBlockID　　　,I  ,String         　,ブロックID
'　　      :pCryInf 　　　,O  ,typ_TBCME037   　,結晶情報
'　　      :pHinDsn 　　　,O  ,typ_TBCME039   　,品番設計
'　　      :pHinMng 　　　,O  ,typ_TBCME041   　,品番管理
'      　　:pSXLMng 　　　,O  ,typ_TBCME042   　,SXL管理
'      　　:pWafSmp 　　　,O  ,typ_XSDCW   　   ,新サンプル管理（SXL）
'　　      :pBlkInf 　　　,O  ,typ_BlkInf2    　,ブロック情報
'　　      :pHinSpec　　　,O  ,typ_HinSpec    　,製品仕様
'　　      :pBlkID  　　　,O  ,String         　,払出単位ブロックID
'      　　:dNeraiRes 　　,O  ,Double         　,ねらい品番の比抵抗上限値（P+の判断用）
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:sPuptn　　    ,O  ,String         　,引上ﾊﾟﾀｰﾝ  2004/12/08 追加
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
'説明      :
'履歴      :2001/07/10 小林 作成
Public Function DBDRV_scmzc_fcmkc001i_Disp(ByVal sBlockId As String, pCryInf As typ_TBCME037, _
                                           pHinDsn() As typ_TBCME039, pHinMng() As typ_TBCME041, _
                                           pSXLMng() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                                           pBlkInf() As typ_BlkInf2, pHinSpec() As typ_HinSpec, _
                                           pBlkID() As String, dNeraiRes As Double, sErrMsg As String, sPuptn As String) As FUNCTION_RETURN

    Dim tmpCryInf() As typ_TBCME037
    Dim tmpStrRslt() As typ_TBCMY009
    Dim rs As OraDynaset
    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sHin As String
    Dim sBlk As String
    Dim dMenseki As Double
    Dim dTopWght As Double
    Dim dCharge As Double
    Dim dMeas(4) As Double
    Dim bFlag As Boolean
    Dim recCnt As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_scmzc_fcmkc001i_Disp"
    sErrMsg = ""

'↓引上ﾊﾟﾀｰﾝ追加対応(2004/12/08) kubota
    sPuptn = ""
    sDbName = "XSDC1"
    sql = "select PUPTNC1"
    sql = sql & "  from XSDC1,XSDC2"
    sql = sql & " where CRYNUMC2 = '" & Trim$(sBlockId) & "'"
    sql = sql & "   and XTALC1   = XTALC2"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt <> 0 Then
        sPuptn = rs("PUPTNC1")
    End If
    rs.Close
'↑引上ﾊﾟﾀｰﾝ追加対応(2004/12/08) kubota

    '' ブロック管理の取得
    sDbName = "E040"
    sCryNum = Left(sBlockId, 9) & "000"
    sql = "select INGOTPOS, LENGTH, REALLEN, BLOCKID, "
    sql = sql & "KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, LSTATCLS"
    sql = sql & " from TBCME040 where CRYNUM='" & sCryNum & "'"
    sql = sql & " and INGOTPOS>=0 and LENGTH>0 order by INGOTPOS"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
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
            .TOPBDLN = 0
            .TOPBDCS = ""
            .TAILBDLN = 0
            .TAILBDCS = ""
            If .BLOCKID = sBlockId Then
                '' 工程チェック
                If rs("LSTATCLS") <> "W" Then
                    sErrMsg = GetMsgStr("EPRC2")
                    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
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
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 結晶情報変更の取得
    sDbName = "W002"
    For i = 1 To recCnt
        With pBlkInf(i)
            sql = "select CRYLEN, TOPBDLN, TOPBDCS, TAILBDLN, TAILBDCS from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .COF.TOPSMPLPOS
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and INGOTPOS=" & .COF.TOPSMPLPOS & ")"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount > 0 Then
                .REALLEN = rs("CRYLEN")
                .TOPBDLN = rs("TOPBDLN")
                .TOPBDCS = rs("TOPBDCS")
                .TAILBDLN = rs("TAILBDLN")
                .TAILBDCS = rs("TAILBDCS")
            End If
            rs.Close
        End With
    Next i

    '' 結晶情報の取得(s_cmzcTBCME037_SQL.bas が必要)
    sDbName = "E037"
    sql = " where CRYNUM='" & sCryNum & "'"
    If DBDRV_GetTBCME037(tmpCryInf(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(tmpCryInf) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
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
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinDsn) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 品番管理の取得(s_cmzcTBCME041_SQL.bas が必要)
    sDbName = "E041"
    sql = " where CRYNUM='" & sCryNum & "' order by INGOTPOS"
    If DBDRV_GetTBCME041(pHinMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pHinMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' SXL管理の取得(s_cmzcTBCME042_SQL.bas が必要)
    sDbName = "E042"
    If DBDRV_GetTBCME042(pSXLMng(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    If UBound(pSXLMng) = 0 Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WFサンプル管理の取得(s_cmzcTBCME044_SQL.bas が必要)
    sDbName = "E044"
' 新サンプル管理(ブロック)追加による修正  2003/10/06 Takada ===================> START
    sql = " where XTALCW='" & sCryNum & "' and LIVKCW='0' order by INPOSCW, TBKBNCW"
''    sql = " where XTALCW='" & sCryNum & "' order by INPOSCW, TBKBNCW"
' 新サンプル管理(ブロック)追加による修正  2003/10/06 Takada ===================> END
    If DBDRV_GetTBCME044(pWafSmp(), sql) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' 引上げ終了実績の取得
    sDbName = "H004"
    sql = "select (DM1+DM2+DM3)/3.0 as DM, WGHTTOP, CHARGE from TBCMH004 where CRYNUM='" & sCryNum & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    dMenseki = AreaOfCircle(rs("DM"))
    dTopWght = rs("WGHTTOP")
    dCharge = rs("CHARGE")
    rs.Close

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
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.TOPSMPLPOS & " and SMPKBN='T')"
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
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B'"
            sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
            sql = sql & " where CRYNUM='" & sCryNum & "'"
            sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='B')"
            Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
            If rs.RecordCount = 0 Then
                rs.Close
                sql = "select MEAS1, MEAS2, MEAS3, MEAS4, MEAS5 from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T'"
                sql = sql & " and TRANCNT=ANY(select MAX(TRANCNT) from TBCMJ002"
                sql = sql & " where CRYNUM='" & sCryNum & "'"
                sql = sql & " and POSITION=" & .COF.BOTSMPLPOS & " and SMPKBN='T')"
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

    '' ブロック新規情報の取得
    sDbName = "Y001"
    sql = "select SBLOCKID from TBCMY001 where BLOCKID='" & sBlockId & "'"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount <= 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    sBlk = rs("SBLOCKID")
    rs.Close

    sql = "select BLOCKID from TBCMY001"
    sql = sql & " where SBLOCKID='" & sBlk & "'"
    sql = sql & " order by SBLOCKID, BLOCKORDER"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    recCnt = rs.RecordCount
    If recCnt <= 0 Then
        rs.Close
        sErrMsg = GetMsgStr("EGET2", sDbName)
        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    ReDim pBlkID(recCnt)
    For i = 1 To recCnt
        pBlkID(i) = rs("BLOCKID")
        rs.MoveNext
    Next i
    rs.Close
    
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
                    If pHinSpec(j).HIN.hinban = .hinban Then
                        pHinSpec(j).LENGTH = pHinSpec(j).LENGTH + .LENGTH
                        Exit For
                    End If
                Next j
                If j > k Then
                    k = k + 1
                    pHinSpec(k).INGOTPOS = .INGOTPOS
                    pHinSpec(k).HIN.hinban = .hinban
                    pHinSpec(k).HIN.mnorevno = .REVNUM
                    pHinSpec(k).HIN.factory = .factory
                    pHinSpec(k).HIN.opecond = .opecond
                    pHinSpec(k).LENGTH = .LENGTH
                    
                    ''残存酸素仕様チェック　03/12/09 ooba START ==============================>
                    iChkAoi = ChkAoiSiyou(pHinSpec(k).HIN)
                    If iChkAoi < 0 Then
                        sErrMsg = "残存酸素(AOi)仕様エラー"
                        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                    ''残存酸素仕様チェック　03/12/09 ooba END ================================>
                    
                    If DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec(k)) = FUNCTION_RETURN_FAILURE Then
                        sErrMsg = GetMsgStr("EGET") & sDbName
                        DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i
    ReDim Preserve pHinSpec(k)

    '' ねらい品番の比抵抗上限値を取得
    sql = "select HSXRMAX"
    sql = sql & " from TBCME037 E37, TBCME018 E18"
    sql = sql & " where (E37.CRYNUM='" & Left$(sBlockId, 9) & "000')"
    sql = sql & " and (E37.RPHINBAN=E18.HINBAN) and (E37.RPREVNUM=E18.MNOREVNO)"
    sql = sql & " and (E37.RPFACT=E18.FACTORY) and (E37.RPOPCOND=E18.OPECOND)"
    Set rs = OraDB.CreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        dNeraiRes = rs("HSXRMAX")
    Else
        dNeraiRes = 0#      'ここまではこないはず
    End If
    rs.Close

    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001i_Disp = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :結晶情報変更用 実行時ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:sStaffID　　　,I  ,String         　,社員ID
'　　      :pBlkInf 　　　,I  ,typ_BlkInf2    　,ブロック情報
'　　      :pHinMng 　　　,I  ,typ_TBCME041   　,品番管理
'      　　:pSXLMng 　　　,I  ,typ_TBCME042   　,SXL管理
'      　　:pSXLOld 　　　,I  ,typ_TBCME042   　,変更前SXL管理
'      　　:pWafSmp 　　　,I  ,typ_XSDCW   　   ,新サンプル管理（SXL）
'      　　:pWafOld 　　　,I  ,typ_XSDCW   　   ,変更前新サンプル管理（SXL）
'      　　:pTrnScr 　　　,I  ,typ_TBCMW006   　,振替廃棄実績
'      　　:pMesInd 　　　,I  ,typ_TBCMY003   　,測定評価方法指示
'      　　:pSXLDcd 　　　,I  ,typ_TBCMY007   　,SXL確定指示
'      　　:sErrMsg 　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/10 小林 新規作成
'      　　:2001/07/11 蔵本 変更
'      　　:2003/04/12 HITEC)会田：TBCMY003,TBCMY007の送信フラグを'0'=>'3'に変更

Public Function DBDRV_scmzc_fcmkc001i_Exec(SSTAFFID As String, pBlkInf() As typ_BlkInf2, _
                                           pHinMng() As typ_TBCME041, pSXLMng() As typ_TBCME042, _
                                           pSXLOld() As typ_TBCME042, pWafSmp() As typ_XSDCW, _
                                           pWafOld() As typ_XSDCW, pTrnScr() As typ_TBCMW006, _
                                           pMesInd() As typ_TBCMY003, pSXLDcd() As typ_TBCMY007, sErrMsg As String) As FUNCTION_RETURN

    Dim sql As String
    Dim sDbName As String
    Dim sCryNum As String
    Dim sAllScrap As String
    Dim recCnt As Long
    Dim i As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_scmzc_fcmkc001i_Exec"
    sErrMsg = ""

    '' WriteDBLog " ", "Start"

    '' SXL管理の挿入／更新(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "E042"
    If DBDRV_SXL_UpdIns(pSXLOld(), pSXLMng()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' WFサンプル管理の挿入／更新(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "E044"
    If DBDRV_WfSmp_UpdIns(pWafOld(), pWafSmp()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    recCnt = UBound(pBlkInf)
    For i = 1 To recCnt
        With pBlkInf(i)
            sCryNum = Left(.BLOCKID, 9) & "000"
            '' ブロック管理の更新
            sDbName = "E040"
            sql = "update TBCME040 set "
            sql = sql & "REALLEN='" & .REALLEN & "', "
            sql = sql & "UPDDATE=sysdate, "
            sql = sql & "SENDFLAG='0'"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            '' 結晶情報変更実績の挿入
            sDbName = "W002"
            sql = "insert into TBCMW002 "
            sql = sql & "(CRYNUM, INGOTPOS, TRANCNT, CRYLEN, KRPROCCD, "
            sql = sql & "PROCCODE, BLOCKID, DELFLG, TOPBDLN, TOPBDCS, "
            sql = sql & "TAILBDLN, TAILBDCS, TSTAFFID, REGDATE, KSTAFFID, "
            sql = sql & "UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            sql = sql & " select '"
            sql = sql & sCryNum & "', "
            sql = sql & .COF.TOPSMPLPOS & ", "
            sql = sql & "nvl(max(TRANCNT),0)+1, "
            sql = sql & .REALLEN & ", '"
            sql = sql & MGPRCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', '"
            sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', '"
            sql = sql & .BLOCKID & "', '"
            sql = sql & .DELFLG & "', "
            sql = sql & .TOPBDLN & ", '"
            sql = sql & .TOPBDCS & "', "
            sql = sql & .TAILBDLN & ", '"
            sql = sql & .TAILBDCS & "', '"
            sql = sql & SSTAFFID & "', "
            sql = sql & "sysdate, '"
            sql = sql & SSTAFFID & "', "
            sql = sql & "sysdate, "
            sql = sql & "'0', "
            sql = sql & "'0', "
            sql = sql & "sysdate"
            sql = sql & " from TBCMW002"
            sql = sql & " where CRYNUM='" & sCryNum & "' and INGOTPOS=" & .COF.TOPSMPLPOS
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If

            '' 欠落情報の挿入
            sDbName = "Y012"
            sql = "delete TBCMY012 where LOTID='" & .BLOCKID & "' and BLOCKSEQ<0"
            Call OraDB.ExecuteSQL(sql)
            If .TOPBDLN >= .REALLEN Or .TAILBDLN >= .REALLEN Or _
               .TOPBDLN + .TAILBDLN >= .REALLEN Then
                sql = "insert into TBCMY012 "
                sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                sql = sql & " values ('"
                sql = sql & .BLOCKID & "', "
                sql = sql & "-1, "
                sql = sql & "-1, "
                sql = sql & "0, "
                sql = sql & "'A', "
                sql = sql & "sysdate, '"
                sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                sql = sql & "'Y', "
                sql = sql & "0, "
                sql = sql & .REALLEN & ", "
                sql = sql & "'      ', "
                sql = sql & "'1', "         ' 1129 チェックフラグはチェック済み
                sql = sql & "sysdate, "
                sql = sql & "'0', "
                sql = sql & "sysdate)"
                '' WriteDBLog sql, sDbName
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            Else
                If .TOPBDLN > 0 Then
                    sql = "insert into TBCMY012 "
                    sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                    sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                    sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                    sql = sql & " values ('"
                    sql = sql & .BLOCKID & "', "
                    sql = sql & "-1, "
                    sql = sql & "1, "
                    sql = sql & "0, "
                    sql = sql & "'A', "
                    sql = sql & "sysdate, '"
                    sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                    sql = sql & "'N', "
                    sql = sql & "0, "
                    sql = sql & .TOPBDLN & ", "
                    sql = sql & "'      ', "
                    sql = sql & "'1', "         ' 1129 チェックフラグはチェック済み
                    sql = sql & "sysdate, "
                    sql = sql & "'0', "
                    sql = sql & "sysdate)"
                    '' WriteDBLog sql, sDbName
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
                If .TAILBDLN > 0 Then
                    sql = "insert into TBCMY012 "
                    sql = sql & "(LOTID, BLOCKSEQ, REJPCS, TOP_POS, REJCAT, "
                    sql = sql & "REJDTTM, REJPROC, ALLSCRAP, LENFROM, LENTO, "
                    sql = sql & "TXID, CHKFLG, REGDATE, SENDFLAG, SENDDATE)"
                    sql = sql & " values ('"
                    sql = sql & .BLOCKID & "', "
                    sql = sql & "-2, "
                    sql = sql & "1, "
                    sql = sql & .REALLEN - .TAILBDLN & ", "
                    sql = sql & "'A', "
                    sql = sql & "sysdate, '"
                    sql = sql & PROCD_KESSYOU_SIYOUJOUHOU_HENKOU & "', "
                    sql = sql & "'N', "
                    sql = sql & .REALLEN - .TAILBDLN & ", "
                    sql = sql & .REALLEN & ", "
                    sql = sql & "'      ', "
                    sql = sql & "'1', "         ' 1129 チェックフラグはチェック済み
                    sql = sql & "sysdate, "
                    sql = sql & "'0', "
                    sql = sql & "sysdate)"
                    '' WriteDBLog sql, sDbName
                    If OraDB.ExecuteSQL(sql) <= 0 Then
                        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                        GoTo proc_exit
                    End If
                End If
            End If
        End With
    Next i

    '' 振替廃棄実績の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "W006"
    recCnt = UBound(pTrnScr)
    For i = 1 To recCnt
        If DBDRV_Furikae_Ins(pTrnScr(i)) = FUNCTION_RETURN_FAILURE Then
            sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
            DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
    Next i

    '' 測定評価方法指示の挿入(s_cmzcDBdriverCOM_SQL.bas が必要)
    sDbName = "Y003"
    If DBDRV_SokuSizi_Ins(pMesInd()) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' ブロック変更情報の挿入
    If DBDRV_BlkChg_Ins(pBlkInf(), pHinMng(), sDbName) = FUNCTION_RETURN_FAILURE Then
        sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
        DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    '' SXL確定指示の挿入
    sDbName = "Y007"
    recCnt = UBound(pSXLDcd)
    For i = 1 To recCnt
        With pSXLDcd(i)
            sql = "insert into TBCMY007 "
            sql = sql & "(SXL_ID, SAMPLE_FROM, SAMPLE_TO, BLOCKID, HINBAN, KUBUN, "
'            sql = sql & "TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE)"
            '比抵抗ﾃﾞｰﾀ登録追加　04/04/09 ooba START =======================================>
            sql = sql & "TXID, REGDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, "
            sql = sql & "MESDATA1TOP, "      ' 測定値１(Top)  center
            sql = sql & "MESDATA2TOP, "      ' 測定値２(Top)  R/2
            sql = sql & "MESDATA3TOP, "      ' 測定値３(Top)  Inside 10mm
            sql = sql & "MESDATA4TOP, "      ' 測定値４(Top)  Inside   6mm
            sql = sql & "MESDATA5TOP, "      ' 測定値５(Top)  Inside   3mm
            sql = sql & "MESDATA1BOT, "      ' 測定値１(Tail)  center
            sql = sql & "MESDATA2BOT, "      ' 測定値２(Tail)  R/2
            sql = sql & "MESDATA3BOT, "      ' 測定値３(Tail)  Inside 10mm
            sql = sql & "MESDATA4BOT, "      ' 測定値４(Tail)  Inside   6mm
            sql = sql & "MESDATA5BOT )"      ' 測定値５(Tail)  Inside   3mm
            '比抵抗ﾃﾞｰﾀ登録追加　04/04/09 ooba END =========================================>
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
            
' vvvvv 2003.04.12 ALT BY HITEC)会田：送信フラグ'0'=>'3'に変更
'''''            sql = sql & "'0', "                 ' 送信フラグ
            sql = sql & "'3', "                 ' 送信フラグ
' ^^^^^ 2003.04.12 ALT BY HITEC)会田  END
'            sql = sql & "sysdate)"              ' 送信日付
            '比抵抗ﾃﾞｰﾀ登録追加　04/04/09 ooba START =======================================>
            sql = sql & "sysdate, "              ' 送信日付
            sql = sql & " '" & .MESDATA1TOP & "', "      ' 測定値１(Top)  center
            sql = sql & " '" & .MESDATA2TOP & "', "      ' 測定値２(Top)  R/2
            sql = sql & " '" & .MESDATA3TOP & "', "      ' 測定値３(Top)  Inside 10mm
            sql = sql & " '" & .MESDATA4TOP & "', "      ' 測定値４(Top)  Inside   6mm
            sql = sql & " '" & .MESDATA5TOP & "', "      ' 測定値５(Top)  Inside   3mm
            sql = sql & " '" & .MESDATA1BOT & "', "      ' 測定値１(Tail)  center
            sql = sql & " '" & .MESDATA2BOT & "', "      ' 測定値２(Tail)  R/2
            sql = sql & " '" & .MESDATA3BOT & "', "      ' 測定値３(Tail)  Inside 10mm
            sql = sql & " '" & .MESDATA4BOT & "', "      ' 測定値４(Tail)  Inside   6mm
            sql = sql & " '" & .MESDATA5BOT & "' ) "     ' 測定値５(Tail)  Inside   3mm
            '比抵抗ﾃﾞｰﾀ登録追加　04/04/09 ooba END =========================================>
            '' WriteDBLog sql, sDbName
            If OraDB.ExecuteSQL(sql) <= 0 Then
                sErrMsg = GetMsgStr("ENG11", vbNullString, sDbName)
                DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
                GoTo proc_exit
            End If
        End With
    Next i

    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_SUCCESS

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
    DBDRV_scmzc_fcmkc001i_Exec = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :内部関数：ブロック変更情報の作成
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'　　      :pBlkInf　　　,I  ,typ_BlkInf2    　,ブロック情報
'　　      :pHinMng　　　,I  ,typ_TBCME041   　,品番管理
'      　　:sDBName　　　,O  ,String         　,DB名称
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
'説明      :
'履歴      :2001/07/25  作成 蔵本
Private Function DBDRV_BlkChg_Ins(pBlkInf() As typ_BlkInf2, pHinMng() As typ_TBCME041, sDbName As String) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim CRYSTALMEN As String
    Dim SEED As Integer
    Dim TRANCNT As Long
    Dim m As Long
    Dim n As Long
    Dim i As Long
    Dim j As Long

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function DBDRV_BlkChg_Ins"

    m = UBound(pBlkInf)
    n = UBound(pHinMng)
    For i = 1 To m
        With pBlkInf(i)
            '' 品番の検索
            For j = 1 To n
                If .COF.TOPSMPLPOS >= pHinMng(j).INGOTPOS And _
                   .COF.TOPSMPLPOS < pHinMng(j).INGOTPOS + pHinMng(j).LENGTH Then
                    Exit For
                End If
            Next j
            If RTrim$(pHinMng(j).hinban) <> "Z" Then
                '' シード傾きの取得(s_cmzcDBdriverCOM_SQL.bas が必要)
                sDbName = "H004"
                If DBDRV_getSEED(Left(pBlkInf(i).BLOCKID, 9) & "000", SEED) = FUNCTION_RETURN_FAILURE Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                '' 結晶面の取得
                sDbName = "E022"
                sql = "select HWFCDIR from TBCME022"
                sql = sql & " where HINBAN='" & pHinMng(j).hinban & "'"
                sql = sql & " and MNOREVNO=" & pHinMng(j).REVNUM
                sql = sql & " and FACTORY='" & pHinMng(j).factory & "'"
                sql = sql & " and OPECOND='" & pHinMng(j).opecond & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount <= 0 Then
                    rs.Close
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
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
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If

                '' 処理回数最大値の取得
                sDbName = "Y005"
                sql = "select nvl(max(TRANCNT),0)+1 as M"
                sql = sql & " from TBCMY005"
                sql = sql & " where BLOCKID='" & .BLOCKID & "'"
                Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
                If rs.RecordCount = 0 Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
                TRANCNT = rs("M")
                rs.Close

                '' ブロック変更情報の挿入
                sql = "insert into TBCMY005 ("
                sql = sql & "BLOCKID, "           ' ブロックID
                sql = sql & "TRANCNT, "           ' 処理回数
                sql = sql & "DELFLG, "            ' 削除指示
                sql = sql & "BLOCKLEN, "          ' ブロックの長さ
                sql = sql & "MAINHINBAN, "        ' 代表品番
                sql = sql & "PNTYPE, "            ' タイプ
                sql = sql & "ROUP, "              ' 比抵抗上限値
                sql = sql & "ROLOW, "             ' 比抵抗下限値
                sql = sql & "OIUP, "              ' 酸素濃度上限値
                sql = sql & "OILOW, "             ' 酸素濃度下限値
                sql = sql & "TANMEN, "            ' 端面角度
                sql = sql & "WARPRANK, "          ' ワープランク
                sql = sql & "CRYSTALMEN, "        ' 結晶面
                sql = sql & "SLPCEN, "            ' 傾中心
                sql = sql & "SLPLOW, "            ' 傾下限
                sql = sql & "SLPUP, "             ' 傾上限
                sql = sql & "INSPMETH, "          ' 検査方法
                sql = sql & "INSPFREQ, "          ' 検査頻度
                sql = sql & "SLPDRC, "            ' 傾方位
                sql = sql & "SLPDRCAPP, "         ' 傾方位指定
                sql = sql & "SLPHEIDRC, "         ' 傾縦方位
                sql = sql & "SLPHEICEN, "         ' 傾縦中心
                sql = sql & "SLPHEILOW, "         ' 傾縦下限
                sql = sql & "SLPHEIUP, "          ' 傾縦上限
                sql = sql & "SLPWIDDRC, "         ' 傾横方位
                sql = sql & "SLPWIDCEN, "         ' 傾横中心
                sql = sql & "SLPWIDLOW, "         ' 傾横下限
                sql = sql & "SLPWIDUP, "          ' 傾横上限
                sql = sql & "SEED, "              ' 引上時使用したシ－ド傾き
                sql = sql & "TXID, "              ' トランザクションID
                sql = sql & "REGDATE, "           ' 登録日付
                sql = sql & "SENDFLAG, "          ' 送信フラグ
                sql = sql & "SENDDATE)"           ' 送信日付
                sql = sql & " select '"
                sql = sql & .BLOCKID & "', "                            ' ブロックID
                sql = sql & TRANCNT & ", '"                             ' 処理回数
                sql = sql & .DELFLG & "', '"                            ' 削除指示
                sql = sql & .REALLEN & "', '"                           ' ブロックの長さ
                                                                        ' 代表品番
                sql = sql & pHinMng(j).hinban & Format(pHinMng(j).REVNUM, "00") & "', "
                sql = sql & "E021HWFTYPE, "                             ' タイプ
                sql = sql & "case when E021HWFRMAX>=99999.9 then '99999.9'"
                sql = sql & " when E021HWFRMAX>=9999.995 then to_char(round(E021HWFRMAX,2),'fm99990.0')"
                sql = sql & " when E021HWFRMAX>=999.9995 then to_char(round(E021HWFRMAX,3),'fm9990.00')"
                sql = sql & " when E021HWFRMAX>=99.99995 then to_char(round(E021HWFRMAX,4),'fm990.000')"
                sql = sql & " when E021HWFRMAX>=10.00000 then to_char(round(E021HWFRMAX,5),'fm90.0000')"
                sql = sql & " when E021HWFRMAX>=0.0 then to_char(E021HWFRMAX,'fm0.00000')"
                sql = sql & " else '-1.0000'"
                sql = sql & "end as RMAX,"                              ' 比抵抗上限値
                sql = sql & "case when E021HWFRMIN>=99999.9 then '99999.9'"
                sql = sql & " when E021HWFRMIN>=9999.995 then to_char(round(E021HWFRMIN,2),'fm99990.0')"
                sql = sql & " when E021HWFRMIN>=999.9995 then to_char(round(E021HWFRMIN,3),'fm9990.00')"
                sql = sql & " when E021HWFRMIN>=99.99995 then to_char(round(E021HWFRMIN,4),'fm990.000')"
                sql = sql & " when E021HWFRMIN>=10.00000 then to_char(round(E021HWFRMIN,5),'fm90.0000')"
                sql = sql & " when E021HWFRMIN>=0.0 then to_char(E021HWFRMIN,'fm0.00000')"
                sql = sql & " else '-1.0000'"
                sql = sql & "end as RMIN,"                              ' 比抵抗下限値
                sql = sql & "to_char(abs(E025HWFONMAX),'fm90.00'), "    ' 酸素濃度上限値
                sql = sql & "to_char(abs(E025HWFONMIN),'fm90.00'), "    ' 酸素濃度下限値
                sql = sql & "'0', "                                     ' 端面角度
                sql = sql & "E027HWFWARPR, '"                           ' ワープランク
                sql = sql & CRYSTALMEN & "', "                          ' 結晶面
                sql = sql & "to_char(abs(E022HWFCSCEN),'fm0.00'), "     ' 傾中心
                sql = sql & "to_char(E022HWFCSMIN,'fm0.00'), "          ' 傾下限
                sql = sql & "to_char(E022HWFCSMAX,'fm0.00'), "          ' 傾上限
                sql = sql & "E022HWFCKWAY, "                            ' 検査方法
                                                                        ' 検査頻度（枚、抜、保、ウの順で足す）
                sql = sql & "E022HWFCKHNM || E022HWFCKHNN || E022HWFCKHNH || E022HWFCKHNU, "
                sql = sql & "E022HWFCSDIR, "                            ' 傾方位
                sql = sql & "E022HWFCSDIS, "                            ' 傾方位指定
                sql = sql & "E022HWFCTDIR, "                            ' 傾縦方位
'''                sql = sql & "to_char(E022HWFCTCEN,'fm0.00'), "          ' 傾縦中心
'''                sql = sql & "to_char(E022HWFCTMIN,'fm0.00'), "          ' 傾縦下限
'''                sql = sql & "to_char(E022HWFCTMAX,'fm0.00'),"           ' 傾縦上限
                sql = sql & "to_char(nvl(E022HWFCTCEN,0),'fm0.00'), "   ' 傾縦中心      '05/03/29 ooba NULL対応
                sql = sql & "to_char(nvl(E022HWFCTMIN,0),'fm0.00'), "   ' 傾縦下限      '05/03/29 ooba NULL対応
                sql = sql & "to_char(nvl(E022HWFCTMAX,0),'fm0.00'),"    ' 傾縦上限      '05/03/29 ooba NULL対応
                sql = sql & "E022HWFCYDIR, "                            ' 傾横方位
'''                sql = sql & "to_char(E022HWFCYCEN,'fm0.00'), "          ' 傾横中心
'''                sql = sql & "to_char(E022HWFCYMIN,'fm0.00'), "          ' 傾横下限
'''                sql = sql & "to_char(E022HWFCYMAX,'fm0.00'), '"         ' 傾横上限
                sql = sql & "to_char(nvl(E022HWFCYCEN,0),'fm0.00'), "   ' 傾横中心      '05/03/29 ooba NULL対応
                sql = sql & "to_char(nvl(E022HWFCYMIN,0),'fm0.00'), "   ' 傾横下限      '05/03/29 ooba NULL対応
                sql = sql & "to_char(nvl(E022HWFCYMAX,0),'fm0.00'), '"  ' 傾横上限      '05/03/29 ooba NULL対応
                sql = sql & SEED & "', "                                ' 引上時使用したシ－ド傾き
                sql = sql & "'TX852I', "                                ' トランザクションID
                sql = sql & "sysdate, "                                 ' 登録日付
                sql = sql & "'0', "                                     ' 送信フラグ
                sql = sql & "sysdate "                                  ' 送信日付
                sql = sql & " from VECME001"
                sql = sql & " where E018HINBAN='" & pHinMng(j).hinban & "'"
                sql = sql & " and E018MNOREVNO=" & pHinMng(j).REVNUM
                sql = sql & " and E018FACTORY='" & pHinMng(j).factory & "'"
                sql = sql & " and E018OPECOND='" & pHinMng(j).opecond & "'"
                '' WriteDBLog sql
                If OraDB.ExecuteSQL(sql) <= 0 Then
                    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
                    GoTo proc_exit
                End If
            End If
        End With
    Next i

    sDbName = ""
    DBDRV_BlkChg_Ins = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    DBDRV_BlkChg_Ins = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :内部関数：受入実績のチェック
'ﾊﾟﾗﾒｰﾀ　　:変数名      ,IO ,型       ,説明
'　　      :blkID 　　　,I  ,String 　,ブロックID
'      　　:戻り値      ,O  ,Boolean　,登録の有無
'説明      :
'履歴      :2001/08/30  作成 野村
Public Function wasUkeire(ByVal blkID$) As Boolean

    Dim rs As OraDynaset
    Dim sql As String

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmbc035_SQL.bas -- Function wasUkeire"

    '' 同時払出ブロックのいずれかが受入実績(TBCMY009)に登録されているかをチェックする
    wasUkeire = False
    sql = "select Y009.LOTID "
    sql = sql & "from TBCMY009 Y009, TBCMY001 Y001 "
    sql = sql & "Where Y009.LOTID = Y001.BLOCKID "
    sql = sql & "and Y001.SBLOCKID=(select SBLOCKID from TBCMY001 where BLOCKID='" & blkID & "')"
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount > 0 Then
        wasUkeire = True
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

'(2002/07 s_cmzcF_cmkc001g_SQL.basよりコピー)
'概要      :抜試指示用 製品仕様専用ＤＢドライバ
'ﾊﾟﾗﾒｰﾀ　　:変数名        ,IO ,型               ,説明
'      　　:pHinSpec　　　,IO ,typ_HinSpec    　,製品仕様
'      　　:戻り値        ,O  ,FUNCTION_RETURN　,読み込みの成否
Public Function DBDRV_scmzc_fcmkc001g_GetSpec(pHinSpec As typ_HinSpec) As FUNCTION_RETURN

    Dim rs As OraDynaset
    Dim sql As String
    Dim sOT1    As String   '03/05/24 後藤
    Dim sOT2    As String
    Dim sMAI1    As String   '04/06/25
    Dim sMAI2    As String
    Dim rtn     As FUNCTION_RETURN

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_cmzcF_cmkc001g_SQL.bas -- Function DBDRV_scmzc_fcmkc001g_GetSpec"

    '' 製品仕様の取得
    With pHinSpec
        sql = "select "
        sql = sql & "E021HWFRMIN, E021HWFRMAX, E021HWFRHWYS, "
        sql = sql & "E024HWFMKHWS, E025HWFONHWS, E025HWFOS1HS, E025HWFOS2HS, E025HWFOS3HS, "
        sql = sql & "E026HWFDSOHS, E028HWFSPVHS, E028HWFDLHWS, E029HWFOF1HS, E029HWFOF2HS, "
        sql = sql & "E029HWFOF3HS, E029HWFOF4HS, E029HWFBM1HS, E029HWFBM2HS, E029HWFBM3HS"
        sql = sql & " from VECME004"
        sql = sql & " where E018HINBAN='" & .HIN.hinban & "'"
        sql = sql & " and E018MNOREVNO=" & .HIN.mnorevno
        sql = sql & " and E018FACTORY='" & .HIN.factory & "'"
        sql = sql & " and E018OPECOND='" & .HIN.opecond & "'"
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If

'        .HWFRMIN = rs("E021HWFRMIN")
'        .HWFRMAX = rs("E021HWFRMAX")
        .HWFRMIN = fncNullCheck(rs("E021HWFRMIN"))      'Null対応 2003/12/10
        .HWFRMAX = fncNullCheck(rs("E021HWFRMAX"))      'Null対応 2003/12/10
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
        rtn = scmzc_getE036(pHinSpec.HIN, sOT1, sOT2, sMAI1, sMAI2)   '03/05/24
        If rtn = FUNCTION_RETURN_FAILURE Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        .HWFOTHER1 = sOT1 '### 03/05/24
        .HWFOTHER2 = sOT2
        .HWFOTHER1MAI = sMAI1   '04/06/25
        .HWFOTHER2MAI = sMAI2   '04/06/25
        
        rs.Close
        
        ''残存酸素仕様取得　03/12/09 ooba START ==============================>
        sql = "select HWFZOHWS from TBCME025 "
        sql = sql & "where HINBAN  ='" & .HIN.hinban & "' "
        sql = sql & "and MNOREVNO= " & .HIN.mnorevno & " "
        sql = sql & "and FACTORY ='" & .HIN.factory & "' "
        sql = sql & "and OPECOND ='" & .HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFZOHWS")) = False Then .HWFZOHWS = rs("HWFZOHWS") '品WF残存酸素保証方法_処
        rs.Close
        ''残存酸素仕様取得　03/12/09 ooba END ================================>
        
        '' GD仕様取得　05/01/25 ooba START ==================================>
        sql = "select "
        sql = sql & "HWFDENHS, "        '品WFDen保証方法_処
        sql = sql & "HWFLDLHS, "        '品WFL/DL保証方法_処
        sql = sql & "HWFDVDHS "         '品WFDVD2保証方法_処
        sql = sql & "from TBCME026 "
        sql = sql & "where HINBAN = '" & .HIN.hinban & "' "
        sql = sql & "and MNOREVNO = " & .HIN.mnorevno & " "
        sql = sql & "and FACTORY = '" & .HIN.factory & "' "
        sql = sql & "and OPECOND = '" & .HIN.opecond & "' "
        
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs.RecordCount = 0 Then
            rs.Close
            DBDRV_scmzc_fcmkc001g_GetSpec = FUNCTION_RETURN_FAILURE
            GoTo proc_exit
        End If
        
        If IsNull(rs("HWFDENHS")) = False Then .HWFDENHS = rs("HWFDENHS")   '品WFDen保証方法_処
        If IsNull(rs("HWFLDLHS")) = False Then .HWFLDLHS = rs("HWFLDLHS")   '品WFL/DL保証方法_処
        If IsNull(rs("HWFDVDHS")) = False Then .HWFDVDHS = rs("HWFDVDHS")   '品WFDVD2保証方法_処
        
        rs.Close
        '' GD仕様取得　05/01/25 ooba END ====================================>
        
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

' vvvvv 2003.04.11 ADD BY HITEC)会田：cmbc036_SQLを元に作成
'基本処理パラメータ作成
'引数：frmFormID=処理画面の判定（2:結晶情報変更）
Public Function MakeParameter(ByVal strCryNum As String) As FUNCTION_RETURN

    Dim lng     As Long
    Dim dat     As Variant
    Dim lRowCnt As Long
    Dim rsMain      As OraDynaset
    Dim sql     As String
    Dim intCnt  As Integer
    Dim errTbl  As String
    Dim sErrMsg As String
    Dim lngBeginIngotpos    As Long
    Dim lngEndIngotpos      As Long
    Dim strIngotpos As String
    Dim varIngotpos As Variant
    
    With f_cmbc035_1.sprExamine
        .GetText 3, 1, varIngotpos
        lngBeginIngotpos = CInt(Trim(varIngotpos))
        .GetText 3, .MaxRows, varIngotpos
        lngEndIngotpos = CInt(Trim(varIngotpos))
    End With
    
    '構造体作成
    If cmbc035_1_CreateTable(strCryNum, lngBeginIngotpos, lngEndIngotpos, sErrMsg) = FUNCTION_RETURN_FAILURE Then
        MakeParameter = FUNCTION_RETURN_FAILURE
        f_cmbc035_1.lblMsg.Caption = sErrMsg
        Exit Function
    End If
    MakeParameter = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

End Function

Public Function cmbc035_1_CreateTable(ByVal strCryNum As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long, ByRef strErrMsg As String) As FUNCTION_RETURN

    Dim rs  As OraDynaset
    Dim errTbl  As String
    Dim strBlockID()  As String
    Dim strDBName   As String
    Dim bNoData     As Boolean
    Dim intLoopCnt  As Integer
    Dim sql     As String
    
    bNoData = False

    'ブロック管理からブロックＩＤを取得
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
    sql = sql & "   AND INGOTPOS>=" & lngBeginIngotpos & " AND (INGOTPOS + LENGTH) <=" & lngEndIngotpos
    Debug.Print sql

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    'ブロックIDを取得
    giInpos = 9000
    intLoopCnt = 0
    Do While Not rs.EOF
        ReDim Preserve strBlockID(intLoopCnt) As String
        If IsNull(rs("BLOCKID")) = True Then
            strBlockID(intLoopCnt) = ""
        Else
            strBlockID(intLoopCnt) = rs("BLOCKID")            'ブロックID
        End If
        
        '基本情報構造体
        With Kihon
            .STAFFID = Trim(f_cmbc035_1.txtStaffID.Text)
''''            .NEWPROC = PROCD_WFC_SOUGOUHANTEI
            .NEWPROC = "CRV01"  'upd 2003/05/31 hitec)matsumoto
            '---------------------------2003/04/13 okazaki
            .NOWPROC = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU
            .DIAMETER = 0      '--------------保留
            .ALLSCRAP = "N" '全数スクラップ
        End With
        
        '分割結晶（ブロック）から前工程実績取得
        strDBName = "XSDC2"
        If cmbc035_1_CreateXSDC2(strBlockID(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc035_1_CreateTable = FUNCTION_RETURN_SUCCESS '処理は行わないが、正常で返す
                Debug.Print "cmbc035_1_CreateXSDC2(" & strBlockID(intLoopCnt) & "," & bNoData & "):XSDC2前工程実績無し"
                Exit Function
            Else
                cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "cmbc035_1_CreateXSDC2(" & strBlockID(intLoopCnt) & "," & bNoData & "):XSDC2前工程実績読込みエラー"
                Exit Function
            End If
        End If
        
        '分割結晶（品番）から前工程実績取得
        strDBName = "XSDCA"
        If cmbc035_1_CreateXSDCA(strBlockID(intLoopCnt), bNoData) = FUNCTION_RETURN_FAILURE Then
            If bNoData = True Then
                cmbc035_1_CreateTable = FUNCTION_RETURN_SUCCESS '処理は行わないが、正常で返す
                Debug.Print "XSDCA：前工程実績無し"
                Exit Function
            Else
                cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
                strErrMsg = GetMsgStr("EAPLY") & strDBName
                Debug.Print "XSDCA：前工程実績読込みエラー"
                Exit Function
            End If
        End If
        
        '現在工程実績作成
        If cmbc035_1_CreateNowProc(strBlockID(intLoopCnt), lngBeginIngotpos, lngEndIngotpos) = FUNCTION_RETURN_FAILURE Then
            cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
            strErrMsg = GetMsgStr("EAPLY")
            Debug.Print "XSDC2,XSDCA：現在工程実績作成エラー"
            Exit Function
        End If
        
        '基本処理
''''        giInpos = 900   'del 2003/05/27 hitec)matsumoto
        If KihonProc = FUNCTION_RETURN_FAILURE Then
            cmbc035_1_CreateTable = FUNCTION_RETURN_FAILURE
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


'分割結晶（品番）前工程実績取得＆構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateXSDCA(ByVal strBlockID As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim iLoopCnt    As Integer
    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0

    'ブロックIDを得る
    sql = "SELECT * from XSDCA"
    sql = sql & " WHERE CRYNUMCA='" & strBlockID & "'"
    sql = sql & "   AND LIVKCA= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateXSDCA = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If
    
    rs.MoveFirst
    iLoopCnt = 0
    
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
        End With
        '良品件数セット
        With Kihon
            .CNTHINOLD = iLoopCnt + 1
        End With
        iLoopCnt = iLoopCnt + 1
        rs.MoveNext
    Loop
    
    rs.Close
    cmbc035_1_CreateXSDCA = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateXSDCA = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'分割結晶（ブロック）前工程実績取得＆構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateXSDC2(ByVal strBlockID As String, ByRef bNoData As Boolean) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    bNoData = False
    
    'ブロックIDを得る
    sql = "SELECT * from XSDC2 "
    sql = sql & " WHERE CRYNUMC2='" & strBlockID & "'"
    sql = sql & "   AND LIVKC2= '0'"   '生死区分

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        bNoData = True
        cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_FAILURE
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
        End With
    End If
    
    rs.Close
    cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
'2002/09/10 ADD hitec)N.MATSUMOTO

'現在工程構造体作成 2002/09/10 ADD hitec)N.MATSUMOTO
Public Function cmbc035_1_CreateNowProc(ByVal strBlockID As String, ByVal lngBeginIngotpos As Long, ByVal lngEndIngotpos As Long) As FUNCTION_RETURN

    Dim rs      As OraDynaset
    Dim sql     As String
    Dim intProcNo   As Integer
    Dim intHinOldCnt    As Integer
    Dim intLengthCnt    As Integer
    Dim intLoopCnt      As Integer
    Dim dblDiameter     As Double
    Dim intNum          As Integer
    Dim strCryNum       As String
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
    Dim lngSumGNWCA     As Long     'add 2003/05/20 hitec)matsumoto
    Dim lngSumGNMCA     As Long     'add 2003/05/20 hitec)matsumoto
    Dim bChgFlg         As Boolean  'add 2003/05/20 hitec)matsumoto
    Dim i               As Integer  'add 2003/05/20 hitec)matsumoto

    '' エラーハンドラの設定
    On Error GoTo proc_err

    intProcNo = 0
    
    intBlkLength = 0
    intBlkIngotPos = 0
    intSxlLength = 0
    intSxlIngotPos = 0
    strCryNum = ""

    'ブロック管理から長さを取得
    sql = "SELECT * from TBCME040 "
    sql = sql & " WHERE BLOCKID='" & strBlockID & "'"
''''    sql = sql & "   AND INGOTPOS=0"

    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
    If rs.RecordCount = 0 Then
        rs.Close
        cmbc035_1_CreateNowProc = FUNCTION_RETURN_FAILURE
        GoTo proc_exit
    End If

    intLoopCnt = 0
    If rs.EOF = False Then
        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '結晶番号
        If IsNull(rs("LENGTH")) = False Then intBlkLength = rs("LENGTH")            '長さ
        If IsNull(rs("INGOTPOS")) = False Then intBlkIngotPos = rs("INGOTPOS")      '位置
    End If

    rs.Close

    'ブロック管理で取得した長さをもとにシングル管理からデータを取得
    sql = "SELECT * from TBCME042 "
    sql = sql & " WHERE CRYNUM='" & strCryNum & "'"
    '↓ループ内で判定
    sql = sql & "   AND ((INGOTPOS >=" & lngBeginIngotpos & ") And (INGOTPOS + LENGTH <= " & lngEndIngotpos & "))"
    sql = sql & "   AND LSTATCLS<>'H'"

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
        If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '不良なし
            '基本情報構造体
            With Kihon
                .FURYOUMU = "N"
            End With
        Else                                            '不良あり
            '基本情報構造体
            With Kihon
                .FURYOUMU = "Y"
            End With
            '不良構造体を作成
            With Furyou
                .XTALC4 = BlkNow.CRYNUMC2   'ブロックID
                .INPOSC4 = BlkNow.INPOSC2   '結晶内開始位置
                .KCKNTC4 = BlkNow.KCNTC2    '工程連番
                .HINBC4 = "Z"               '品番
    '            .REVNUMC4                   '製品番号改訂番号
    '            .FACTORYC4                  '工場
    '            .OPEC4                      '操業条件
     '           .WKKTC4 = PROCD_NUKISI_HENKOU
                .WKKTC4 = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU      ' 2003/04/12 okazaki
                '「仮処理」。登録前にもう一度不良長さ・重量・枚数を求めなおす start -------------
                .PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2) '不良長さ(前工程-現在工程（良品）)
                .PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)    'upd 2003/05/31 hitec)matsumoto 重量は再計算しない
                .PUCUTMC4 = 0 'upd 2003/05/31 hitec)matsumoto 枚数は再計算しない
                .SUMITBC3 = "0"
            End With
        End If
        rs.Close
        cmbc035_1_CreateNowProc = FUNCTION_RETURN_SUCCESS
        GoTo proc_exit
    End If
    
    '前工程の構造体を現在工程の構造体へコピー
    BlkNow = BlkOld
    '工程連番に＋１する
    With BlkNow
        .KCNTC2 = CInt(.KCNTC2) + 1         '工程連番
        .NEWKNTC2 = Kihon.NOWPROC         '前工程
        .GNWKNTC2 = Kihon.NEWPROC           '現在工程
        .SUMITLC2 = "0"                     'SUMMIT長さ
        .SUMITMC2 = "0"                   'SUMMIT枚数
        .SUMITWC2 = "0"                   'SUMMIT重量
        .SUMITBC2 = "0"
    End With
    
    intLoopCnt = 0
    BlkNow.GNLC2 = 0    '現在工程（ブロック）の長さをクリアしておく
    BlkNow.GNWC2 = 0    '現在工程（ブロック）の長さをクリアしておく
    BlkNow.GNMC2 = 0    '現在工程（ブロック）の長さをクリアしておく
    
    Do While Not rs.EOF
        ReDim Preserve HinNow(intLoopCnt) As typ_XSDCA_Update
        '前工程の構造体を現在工程の構造体へコピー
''''        HinNow(intLoopCnt) = HinOld(intHinOldCnt)
        
        If IsNull(rs("CRYNUM")) = False Then strCryNum = rs("CRYNUM")               '結晶番号
        If IsNull(rs("LENGTH")) = False Then intSxlLength = rs("LENGTH")            '長さ
        If IsNull(rs("INGOTPOS")) = False Then intSxlIngotPos = rs("INGOTPOS")      '位置
        
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
        
            GoTo LoopNext

        End If
        '----------------------------------------------------
        
        '現在工程編集
        With HinNow(intLoopCnt)
            If IsNull(rs("CRYNUM")) = False Then .XTALCA = rs("CRYNUM")
            .CRYNUMCA = strBlockID         'ブロックID
            If IsNull(rs("HINBAN")) = False Then .HINBCA = rs("HINBAN")         '品番
            If IsNull(rs("REVNUM")) = False Then .REVNUMCA = rs("REVNUM")       '製品番号改訂番号
            If IsNull(rs("FACTORY")) = False Then .FACTORYCA = rs("FACTORY")    '工場
            If IsNull(rs("OPECOND")) = False Then .OPECA = rs("OPECOND")        '操業条件
            .INPOSCA = intIngotPos    '結晶内開始位置
            .GNLCA = intLength          '長さ
            BlkNow.GNLC2 = CStr(CLng(BlkNow.GNLC2) + CLng(HinNow(intLoopCnt).GNLCA))  '長さ
            If IsNull(rs("SXLID")) = False Then .SXLIDCA = rs("SXLID")          'シングルID
            .SUMITBCA = 0
            .SUMITLCA = 0
            .SUMITMCA = 0
            .SUMITWCA = 0
            .NEWKNTCA = Kihon.NOWPROC   '前工程
            .GNWKNTCA = Kihon.NEWPROC   '現在工程
            .KCKNTCA = BlkNow.KCNTC2    '工程連番
            .NEMACOCA = BlkNow.NEMACOC2 '最終通過処理回数
            .GNMACOCA = BlkNow.GNMACOC2 '現在処理回数
''''            .XTALCA = strCryNum         '結晶番号
            '現在重量を求める
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
''''                GoTo proc_wxit
            End If
            '基本情報の直径セット
            Kihon.DIAMETER = dblDiameter
            
            '取得した直径を元に重量を求める
            .GNWCA = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLCA))))
            
            '現在枚数を求める
            If WfCount(strBlockID, CLng(.GNLCA), intNum) = FUNCTION_RETURN_FAILURE Then
                .GNMCA = 0
''''                GoTo proc_wxit
            Else
                .GNMCA = intNum
            End If
        End With
        
        With BlkNow
            '現在重量を求める
            If GetDiameter(strBlockID, dblDiameter) = FUNCTION_RETURN_FAILURE Then  '直径を求める
                dblDiameter = 0
    ''''                GoTo proc_wxit
            End If
            '基本情報の直径セット
            Kihon.DIAMETER = dblDiameter
            '取得した直径を元に重量を求める
            .GNWC2 = CStr(CLng(WeightOfCylinder(dblDiameter, CDbl(.GNLC2))))
            '現在枚数を求める
            If WfCount(strBlockID, CLng(.GNLC2), intNum) = FUNCTION_RETURN_FAILURE Then
                .GNMC2 = 0
''''                GoTo proc_wxit
            Else
                .GNMC2 = intNum
            End If
            
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
    
    '前工程の長さと現在工程の長さをくらべ、不良が存在するか判定
    If CInt(BlkNow.GNLC2) = CInt(BlkOld.GNLC2) Then '不良なし
        '基本情報構造体
        With Kihon
            .FURYOUMU = "N"
        End With
    Else                                            '不良あり
        '基本情報構造体
        With Kihon
            .FURYOUMU = "Y"
        End With
        '不良構造体を作成
        With Furyou
            .XTALC4 = BlkNow.CRYNUMC2   'ブロックID
            .INPOSC4 = BlkNow.INPOSC2   '結晶内開始位置
            .KCKNTC4 = BlkNow.KCNTC2    '工程連番
            .HINBC4 = "Z"               '品番
''            .REVNUMC4 =                '製品番号改訂番号
''            .FACTORYC4                  '工場
''            .OPEC4                      '操業条件
'            .WKKTC4 = PROCD_NUKISI_HENKOU
            .WKKTC4 = PROCD_KESSYOU_SIYOUJOUHOU_HENKOU    ' 2003/04/12 okazaki
            '「仮処理」。登録前にもう一度不良長さ・重量・枚数を求めなおす start -------------
            .PUCUTLC4 = CLng(BlkNow.GNLC2) - CLng(BlkOld.GNLC2) '不良長さ(前工程-現在工程（良品）)
            .PUCUTWC4 = CLng(BlkNow.GNWC2) - CLng(BlkOld.GNWC2)    'upd 2003/05/31 hitec)matsumoto 重量は再計算しない
            .PUCUTMC4 = 0 'upd 2003/05/31 hitec)matsumoto 枚数は再計算しない
        End With
    End If
    
    cmbc035_1_CreateNowProc = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    cmbc035_1_CreateNowProc = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

' ^^^^^ 2003.04.11 ADD BY HITEC)会田  END

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

'概要      :テーブル「TBCME042」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCME042 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村  (2002/07 s_cmzcTBCME042_SQL.basより移動)
Public Function DBDRV_GetTBCME042(records() As typ_TBCME042, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select CRYNUM, INGOTPOS, LENGTH, SXLID, KRPROCCD, NOWPROC, LPKRPROCCD, LASTPASS, DELCLS, LSTATCLS, HOLDCLS," & _
              " HINBAN, REVNUM, FACTORY, OPECOND, BDCAUS, COUNT, REGDATE, UPDDATE, SUMMITSENDFLAG, SENDFLAG, SENDDATE, " & _
              " PASSFLAG "   '02/04/16 Yam
    sqlBase = sqlBase & "From TBCME042"
    sql = sqlBase
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
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
            .INGOTPOS = rs("INGOTPOS")       ' 結晶内開始位置
            .LENGTH = rs("LENGTH")           ' 長さ
            .SXLID = rs("SXLID")             ' SXLID
            .KRPROCCD = rs("KRPROCCD")       ' 管理工程
            .NOWPROC = rs("NOWPROC")         ' 現在工程
            .LPKRPROCCD = rs("LPKRPROCCD")   ' 最終通過管理工程
            .LASTPASS = rs("LASTPASS")       ' 最終通過工程
            .DELCLS = rs("DELCLS")           ' 削除区分
            .LSTATCLS = rs("LSTATCLS")       ' 最終状態区分
            .HOLDCLS = rs("HOLDCLS")         ' ホールド区分
            .hinban = rs("HINBAN")           ' 品番
            .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
            .factory = rs("FACTORY")         ' 工場
            .opecond = rs("OPECOND")         ' 操業条件
            .BDCAUS = rs("BDCAUS")           ' 不良理由
            .Count = rs("COUNT")             ' 枚数
            .REGDATE = rs("REGDATE")         ' 登録日付
            .UPDDATE = rs("UPDDATE")         ' 更新日付
            .SUMMITSENDFLAG = rs("SUMMITSENDFLAG")
            .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
            .SENDDATE = rs("SENDDATE")       ' 送信日付
            .PASSFLAG = " "   ' 通過フラグのスペースクリア '02/04/16 Yam
            If rs("PASSFLAG") = "1" Then
                .PASSFLAG = rs("PASSFLAG")   ' 通過フラグ '02/04/05 Yam
            End If
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
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long
Dim tHIN As tFullHinban '03/05/24
Dim sOT1    As String   '03/05/24
Dim sOT2    As String
Dim sMAI1    As String   '04/06/25
Dim sMAI2    As String
Dim rtn     As FUNCTION_RETURN
    ''SQLを組み立てる   '03/05/24
'    sqlBase = "Select CRYNUM, INGOTPOS, SMPKBN, SMPLID, HINBAN, REVNUM, FACTORY, OPECOND, KTKBN, WFINDRS, WFINDOI, WFINDB1," & _
'              " WFINDB2, WFINDB3, WFINDL1, WFINDL2, WFINDL3, WFINDL4, WFINDDS, WFINDDZ, WFINDSP, WFINDDO1, WFINDDO2, WFINDDO3," & _
'              " NVL(WFINDOT1,'0') as DOT1, NVL(WFINDOT2,'0') as DOT2," & _
'              " WFRESRS, WFRESOI, WFRESB1, WFRESB2, WFRESB3, WFRESL1, WFRESL2, WFRESL3, WFRESL4, WFRESDS, WFRESDZ, WFRESSP," & _
'              " WFRESDO1, WFRESDO2, WFRESDO3,NVL(WFRESOT1,'0') as SOT1, NVL(WFRESOT2,'0') as SOT2, REGDATE, UPDDATE, SENDFLAG, SENDDATE "
'    sqlBase = sqlBase & "From TBCME044"

    'GD項目追加　05/01/25 ooba
    sqlBase = "Select SXLIDCW, SMPKBNCW, TBKBNCW, REVNUMCW, XTALCW, INPOSCW, REPSMPLIDCW, HINBCW, FACTORYCW, OPECW, KTKBNCW, " & _
              " SMCRYNUMCW, WFSMPLIDRSCW, WFSMPLIDRS1CW, WFSMPLIDRS2CW, WFINDRSCW, WFRESRS1CW, WFRESRS2CW, WFSMPLIDOICW, WFINDOICW, " & _
              " WFRESOICW, WFSMPLIDB1CW, WFINDB1CW, WFRESB1CW, WFSMPLIDB2CW, WFINDB2CW, WFRESB2CW, WFSMPLIDB3CW, WFINDB3CW, " & _
              " WFRESB3CW, WFSMPLIDL1CW, WFINDL1CW, WFRESL1CW, WFSMPLIDL2CW, WFINDL2CW, WFRESL2CW, WFSMPLIDL3CW, WFINDL3CW, WFRESL3CW, " & _
              " WFSMPLIDL4CW, WFINDL4CW, WFRESL4CW, WFSMPLIDDSCW, WFINDDSCW, WFRESDSCW, WFSMPLIDDZCW, WFINDDZCW, WFRESDZCW, " & _
              " WFSMPLIDSPCW, WFINDSPCW, WFRESSPCW, WFSMPLIDDO1CW, WFINDDO1CW, WFRESDO1CW, WFSMPLIDDO2CW, WFINDDO2CW, WFRESDO2CW, " & _
              " WFSMPLIDDO3CW, WFINDDO3CW, WFRESDO3CW, WFSMPLIDOT1CW, NVL(WFINDOT1CW,'0') as DOT1, NVL(WFRESOT1CW,'0') as SOT1, " & _
              " WFSMPLIDOT2CW, NVL(WFINDOT2CW,'0') as DOT2, NVL(WFRESOT2CW,'0') as SOT2, WFSMPLIDAOICW, WFINDAOICW, WFRESAOICW, SMPLNUMCW, " & _
              " WFSMPLIDGDCW, WFINDGDCW, WFRESGDCW, WFHSGDCW, " & _
              " SMPLPATCW, TSTAFFCW, TDAYCW, KSTAFFCW, KDAYCW, SNDKCW, SNDDAYCW "
    sqlBase = sqlBase & "From XSDCW"

    sql = sqlBase
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
        With records(i)
            .SXLIDCW = rs("SXLIDCW")
            .SMPKBNCW = rs("SMPKBNCW")           ' サンプル区分
            .TBKBNCW = rs("TBKBNCW")
            .REPSMPLIDCW = rs("REPSMPLIDCW")           ' サンプルID
            .XTALCW = rs("XTALCW")           ' 結晶番号
            .INPOSCW = rs("INPOSCW")       ' 結晶内位置
            .HINBCW = rs("HINBCW")           ' 品番
            .REVNUMCW = rs("REVNUMCW")           ' 製品番号改訂番号
            .FACTORYCW = rs("FACTORYCW")         ' 工場
            .OPECW = rs("OPECW")         ' 操業条件
            .KTKBNCW = rs("KTKBNCW")             ' 確定区分
            .SMCRYNUMCW = rs("SMCRYNUMCW")
            .WFSMPLIDRSCW = rs("WFSMPLIDRSCW")
            If Not IsNull(rs("WFSMPLIDRS1CW")) Then .WFSMPLIDRS1CW = rs("WFSMPLIDRS1CW")
            If Not IsNull(rs("WFSMPLIDRS2CW")) Then .WFSMPLIDRS2CW = rs("WFSMPLIDRS2CW")
            .WFINDRSCW = rs("WFINDRSCW")         ' WF検査指示（Rs)
            .WFRESRS1CW = rs("WFRESRS1CW")         ' WF検査実績（Rs)
            If Not IsNull(rs("WFRESRS2CW")) Then .WFRESRS2CW = rs("WFRESRS2CW")
            .WFSMPLIDOICW = rs("WFSMPLIDOICW")
            .WFINDOICW = rs("WFINDOICW")         ' WF検査指示（Oi)
            .WFRESOICW = rs("WFRESOICW")         ' WF検査実績（Oi)
            .WFSMPLIDB1CW = rs("WFSMPLIDB1CW")
            .WFINDB1CW = rs("WFINDB1CW")         ' WF検査指示（B1)
            .WFRESB1CW = rs("WFRESB1CW")
            .WFSMPLIDB2CW = rs("WFSMPLIDB2CW")
            .WFINDB2CW = rs("WFINDB2CW")         ' WF検査指示（B2）
            .WFRESB2CW = rs("WFRESB2CW")         ' WF検査実績（B2）
            .WFSMPLIDB3CW = rs("WFSMPLIDB3CW")
            .WFINDB3CW = rs("WFINDB3CW")         ' WF検査指示（B3)
            .WFRESB3CW = rs("WFRESB3CW")         ' WF検査実績（B3)
            .WFSMPLIDL1CW = rs("WFSMPLIDL1CW")
            .WFINDL1CW = rs("WFINDL1CW")         ' WF検査指示（L1)
            .WFRESL1CW = rs("WFRESL1CW")         ' WF検査実績（L1)
            .WFSMPLIDL2CW = rs("WFSMPLIDL2CW")
            .WFINDL2CW = rs("WFINDL2CW")         ' WF検査指示（L2)
            .WFRESL2CW = rs("WFRESL2CW")         ' WF検査実績（L2)
            .WFSMPLIDL3CW = rs("WFSMPLIDL3CW")
            .WFINDL3CW = rs("WFINDL3CW")         ' WF検査指示（L3)
            .WFRESL3CW = rs("WFRESL3CW")         ' WF検査実績（L3)
            .WFSMPLIDL4CW = rs("WFSMPLIDL4CW")
            .WFINDL4CW = rs("WFINDL4CW")         ' WF検査指示（L4)
            .WFRESL4CW = rs("WFRESL4CW")         ' WF検査実績（L4)
            .WFSMPLIDDSCW = rs("WFSMPLIDDSCW")
            .WFINDDSCW = rs("WFINDDSCW")         ' WF検査指示（DS)
            .WFRESDSCW = rs("WFRESDSCW")         ' WF検査実績（DS)
            .WFSMPLIDDZCW = rs("WFSMPLIDDZCW")
            .WFINDDZCW = rs("WFINDDZCW")         ' WF検査指示（DZ)
            .WFRESDZCW = rs("WFRESDZCW")         ' WF検査実績（DZ)
            .WFSMPLIDSPCW = rs("WFSMPLIDSPCW")
            .WFINDSPCW = rs("WFINDSPCW")         ' WF検査指示（SP)
            .WFRESSPCW = rs("WFRESSPCW")         ' WF検査実績（SP)
            .WFSMPLIDDO1CW = rs("WFSMPLIDDO1CW")
            .WFINDDO1CW = rs("WFINDDO1CW")       ' WF検査指示（DO1)
            .WFRESDO1CW = rs("WFRESDO1CW")       ' WF検査実績（DO1)
            .WFSMPLIDDO2CW = rs("WFSMPLIDDO2CW")
            .WFINDDO2CW = rs("WFINDDO2CW")       ' WF検査指示（DO2)
            .WFRESDO2CW = rs("WFRESDO2CW")       ' WF検査実績（DO2)
            .WFSMPLIDDO3CW = rs("WFSMPLIDDO3CW")
            .WFINDDO3CW = rs("WFINDDO3CW")       ' WF検査指示（DO3)
            .WFRESDO3CW = rs("WFRESDO3CW")       ' WF検査実績（DO3)
            tHIN.hinban = .HINBCW   ''03/05/24
            tHIN.factory = .FACTORYCW
            tHIN.mnorevno = .REVNUMCW
            tHIN.opecond = .OPECW
            rtn = scmzc_getE036(tHIN, sOT1, sOT2, sMAI1, sMAI2)
            If rtn = FUNCTION_RETURN_FAILURE Then
                rs.Close
                DBDRV_GetTBCME044 = FUNCTION_RETURN_FAILURE
                Exit Function
            End If
            If sOT1 = "1" Then
                .WFINDOT1CW = rs!DOT1 '03/05/23
            Else
                .WFINDOT1CW = 0 '03/05/23
            End If
            If sOT2 = "1" Then
                .WFINDOT2CW = rs!DOT2 '03/05/23
            Else
                .WFINDOT2CW = 0 '03/05/23
            End If
            '#####################################################03/05/23 後藤
            .WFRESOT1CW = rs("SOT1")       ' WF検査実績（OT1)
            .WFRESOT2CW = rs("SOT2")       ' WF検査実績（OT2)
            '#####################################################03/05/23 後藤
            If Not IsNull(rs("WFSMPLIDAOICW")) Then .WFSMPLIDAOICW = rs("WFSMPLIDAOICW")
            If Not IsNull(rs("WFINDAOICW")) Then .WFINDAOICW = rs("WFINDAOICW")
            If Not IsNull(rs("WFRESAOICW")) Then .WFRESAOICW = rs("WFRESAOICW")
            If Not IsNull(rs("SMPLNUMCW")) Then .SMPLNUMCW = rs("SMPLNUMCW")
            If Not IsNull(rs("SMPLPATCW")) Then .SMPLPATCW = rs("SMPLPATCW")
            If Not IsNull(rs("TSTAFFCW")) Then .TSTAFFCW = rs("TSTAFFCW")
            .TDAYCW = rs("TDAYCW")         ' 登録日付
            If Not IsNull(rs("KSTAFFCW")) Then .KSTAFFCW = rs("KSTAFFCW")
            .KDAYCW = rs("KDAYCW")         ' 更新日付
            If Not IsNull(rs("SNDKCW")) Then .SNDKCW = rs("SNDKCW")       ' 送信フラグ
            If Not IsNull(rs("SNDDAYCW")) Then .SNDDAYCW = rs("SNDDAYCW")       ' 送信日付
            '' GD項目取得　05/01/25 ooba START ==================================================>
            If Not IsNull(rs("WFSMPLIDGDCW")) Then .WFSMPLIDGDCW = rs("WFSMPLIDGDCW")
            If Not IsNull(rs("WFINDGDCW")) Then .WFINDGDCW = rs("WFINDGDCW")    ' WF検査指示 (GD)
            If Not IsNull(rs("WFRESGDCW")) Then .WFRESGDCW = rs("WFRESGDCW")    ' WF検査実績 (GD)
            If Not IsNull(rs("WFHSGDCW")) Then .WFHSGDCW = rs("WFHSGDCW")       ' WF検査保証 (GD)
            '' GD項目取得　05/01/25 ooba END ====================================================>
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCME044 = FUNCTION_RETURN_SUCCESS
End Function



