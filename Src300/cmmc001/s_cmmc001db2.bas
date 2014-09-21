Attribute VB_Name = "s_cmmc001db2"
Option Explicit

'Type type_HinbanSyutoku   ''s_cmmc001db_sql　のOUT用
'    CRYNUM      As String * 12  ''結晶番号
'    INGOTPOS    As Integer      ''トップサンプル位置
'    HINBAN      As String * 12  ''ボトムサンプル位置
'    BLOCKID     As String * 12  ''チャージ量
'    LENGTH      As Integer      ''トップ重量
'End Type

'概要      :
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型             ,説明
'          :sCryNum       ,I   ,String        ,入力用
'          :pTbcmj002()   ,O   ,typ_TBCMJ002  ,抵抗実績情報構造体
'説明      :
'履歴      :2001/06/28　小林　作成
'　　      :2001/08/08　S.Sano改造
'　　      :2006/04/27　窪田　改造      サンプル有無対応
'Public Function s_cmmc001db2_sql(all As typ_AllTypes) As FUNCTION_RETURN
    
Public Function s_cmmc001db2_sql(sCrynum As String, ADDDPPOS As Integer, FREELENG As Integer, INGOTPOS As Integer, typ_rsz() As typ_TBCMJ002) As FUNCTION_RETURN
    Dim temp() As typ_TBCMJ002
    Dim sql As String
    Dim recCnt As Long      'レコード数
    Dim c0 As Integer
    Dim c1 As Integer
'    Dim ret As Integer
'    Dim pos(2) As Long
    Dim rs As OraDynaset
'    Dim rsz() As typ_TBCMJ002
    Dim MaxMin As String
    
    s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
    
    ReDim typ_rsz(2)
    
'    sql = "select " & _
'        "INGOTPOS " & _
'        " from TBCME040 " & _
'        " where CRYNUM = '" & sCryNum & "' and " & _
'            "INGOTPOS = ANY (SELECT MIN(INGOTPOS) FROM TBCME040 " & _
'                        "WHERE CRYNUM = '" & sCryNum & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    pos(1) = rs("INGOTPOS")
'    rs.Close
'
'    sql = "select " & _
'        "INGOTPOS, LENGTH " & _
'        " from TBCME040 " & _
'        " where CRYNUM = '" & sCryNum & "' and " & _
'            "INGOTPOS = ANY (SELECT MAX(INGOTPOS) FROM TBCME040 " & _
'                        "WHERE CRYNUM = '" & sCryNum & "')"
'
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
'    pos(2) = CDbl(rs("INGOTPOS")) + CDbl(rs("LENGTH"))
'    rs.Close
'
'    For ii = 1 To 2
'        sql = " where CRYNUM = '" & sCryNum & "' and " & _
'            "POSITION = " & pos(ii) & " and " & _
'            "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 " & _
'                    "WHERE CRYNUM = '" & sCryNum & "' and " & _
'                        "POSITION = " & pos(ii) & ")"
'        ReDim rsz(0)
'        '両方取れなければだめ
'        If DBDRV_GetTBCMJ002(rsz(), sql, "") <> FUNCTION_RETURN_SUCCESS Then
'            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'        If UBound(rsz) = 0 Then
'            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
'            Exit Function
'        End If
'        all.typ_rsz(ii) = rsz(1)
'    Next ii
    


    For c0 = 1 To 2
        If c0 = 1 Then
            MaxMin = "MIN"
        Else
            MaxMin = "MAX"
        End If
        sql = "Select CRYNUM, POSITION, SMPKBN, TRANCOND, TRANCNT, SMPLNO, SMPLUMU, KRPROCCD, PROCCODE, HINBAN, REVNUM, FACTORY," & _
                  " OPECOND, GOUKI, TYPE, MEAS1, MEAS2, MEAS3, MEAS4, MEAS5, EFEHS, RRG, JUDGDATA, TSTAFFID, REGDATE, KSTAFFID, UPDDATE," & _
                  " SENDFLAG, SENDDATE "
        sql = sql & "From TBCMJ002 "
        
        If ADDDPPOS > 0 And ADDDPPOS < FREELENG Then
            If INGOTPOS < ADDDPPOS Then
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION < '" & ADDDPPOS & "') and "
                sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION < '" & ADDDPPOS & "'))"
            Else
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION > '" & ADDDPPOS & "') and "
                sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
                sql = sql & "where CRYNUM = '" & sCrynum & "' and "
                sql = sql & "POSITION > '" & ADDDPPOS & "'))"
            End If
        Else
            sql = sql & "where CRYNUM = '" & sCrynum & "' and "
            sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
            
'>>>>> サンプル有無対応 SETsw kubota(2006/04/27)
            'サンプル有の中でTop位置、Bot位置を取得
            'sql = sql & "where CRYNUM = '" & sCrynum & "') and "
            sql = sql & "where CRYNUM = '" & sCrynum & "'"
            sql = sql & "  and SMPLUMU = '0'"
            sql = sql & "  and (POSITION,TRANCNT) in"
            sql = sql & "      (SELECT POSITION,MAX(TRANCNT) FROM TBCMJ002 where CRYNUM = '" & sCrynum & "' group by POSITION)"
            sql = sql & ") and "
'<<<<< サンプル有無対応 SETsw kubota(2006/04/27)
            
            sql = sql & "TRANCNT = ANY (SELECT MAX(TRANCNT) FROM TBCMJ002 "
            sql = sql & "where CRYNUM = '" & sCrynum & "' and "
            sql = sql & "POSITION = ANY (SELECT " & MaxMin & "(POSITION) FROM TBCMJ002 "
            
'>>>>> サンプル有無対応 SETsw kubota(2006/04/27)
            'サンプル有の中でTop位置、Bot位置を取得
            'sql = sql & "where CRYNUM = '" & sCrynum & "'))"
            sql = sql & "where CRYNUM = '" & sCrynum & "'"
            sql = sql & "  and SMPLUMU = '0'"
            sql = sql & "  and (POSITION,TRANCNT) in"
            sql = sql & "      (SELECT POSITION,MAX(TRANCNT) FROM TBCMJ002 where CRYNUM = '" & sCrynum & "' group by POSITION)"
            sql = sql & "))"
'<<<<< サンプル有無対応 SETsw kubota(2006/04/27)
        
        End If
        
'Debug.Print sql
        ''データを抽出する
        Set rs = OraDB.DBCreateDynaset(sql, ORADYN_NO_BLANKSTRIP)
        If rs Is Nothing Then
            s_cmmc001db2_sql = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        ''抽出結果を格納する
        recCnt = rs.RecordCount
            
        If recCnt <> 0 Then

            ReDim temp(recCnt) As typ_TBCMJ002
            For c1 = 1 To recCnt
                With temp(c1)
                    .CRYNUM = rs("CRYNUM")           ' 結晶番号
                    .POSITION = rs("POSITION")       ' 位置
                    .SMPKBN = rs("SMPKBN")           ' サンプル区分
                    .TRANCOND = rs("TRANCOND")       ' 処理条件
                    .TRANCNT = rs("TRANCNT")         ' 処理回数
                    .SMPLNO = rs("SMPLNO")           ' サンプルＮｏ
                    .SMPLUMU = rs("SMPLUMU")         ' サンプル有無
                    .KRPROCCD = rs("KRPROCCD")       ' 管理工程コード
                    .PROCCODE = rs("PROCCODE")       ' 工程コード
                    .hinban = rs("HINBAN")           ' 品番
                    .REVNUM = rs("REVNUM")           ' 製品番号改訂番号
                    .factory = rs("FACTORY")         ' 工場
                    .opecond = rs("OPECOND")         ' 操業条件
                    .GOUKI = rs("GOUKI")             ' 号機
                    .TYPE = rs("TYPE")               ' タイプ
                    .MEAS1 = rs("MEAS1")             ' 測定値１
                    .MEAS2 = rs("MEAS2")             ' 測定値２
                    .MEAS3 = rs("MEAS3")             ' 測定値３
                    .MEAS4 = rs("MEAS4")             ' 測定値４
                    .MEAS5 = rs("MEAS5")             ' 測定値５
                    .EFEHS = rs("EFEHS")             ' 実効偏析
                    .RRG = rs("RRG")                 ' ＲＲＧ
                    .JudgData = rs("JUDGDATA")       ' 検索対象値
                    .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
                    .REGDATE = rs("REGDATE")         ' 登録日付
                    .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
                    .UPDDATE = rs("UPDDATE")         ' 更新日付
                    .SENDFLAG = rs("SENDFLAG")       ' 送信フラグ
                    .SENDDATE = rs("SENDDATE")       ' 送信日付
                End With
                rs.MoveNext
            Next
            
            c1 = 1
            If recCnt > 1 Then
                Select Case c0
                Case 1
                    If temp(2).SMPKBN = "B" Then
                        c1 = 2
                    End If
                Case 2
                    If temp(2).SMPKBN = "T" Then
                        c1 = 2
                    End If
                End Select
            End If
            
            With typ_rsz(c0)
                .CRYNUM = temp(c1).CRYNUM       ' 結晶番号
                .POSITION = temp(c1).POSITION   ' 位置
                .SMPKBN = temp(c1).SMPKBN       ' サンプル区分
                .TRANCOND = temp(c1).TRANCOND   ' 処理条件
                .TRANCNT = temp(c1).TRANCNT     ' 処理回数
                .SMPLNO = temp(c1).SMPLNO       ' サンプルＮｏ
                .SMPLUMU = temp(c1).SMPLUMU     ' サンプル有無
                .KRPROCCD = temp(c1).KRPROCCD   ' 管理工程コード
                .PROCCODE = temp(c1).PROCCODE   ' 工程コード
                .hinban = temp(c1).hinban       ' 品番
                .REVNUM = temp(c1).REVNUM       ' 製品番号改訂番号
                .factory = temp(c1).factory     ' 工場
                .opecond = temp(c1).opecond     ' 操業条件
                .GOUKI = temp(c1).GOUKI         ' 号機
                .TYPE = temp(c1).TYPE           ' タイプ
                .MEAS1 = temp(c1).MEAS1         ' 測定値１
                .MEAS2 = temp(c1).MEAS2         ' 測定値２
                .MEAS3 = temp(c1).MEAS3         ' 測定値３
                .MEAS4 = temp(c1).MEAS4         ' 測定値４
                .MEAS5 = temp(c1).MEAS5         ' 測定値５
                .EFEHS = temp(c1).EFEHS         ' 実効偏析
                .RRG = temp(c1).RRG             ' ＲＲＧ
                .JudgData = temp(c1).JudgData   ' 検索対象値
                .TSTAFFID = temp(c1).TSTAFFID   ' 登録社員ID
                .REGDATE = temp(c1).REGDATE     ' 登録日付
                .KSTAFFID = temp(c1).KSTAFFID   ' 更新社員ID
                .UPDDATE = temp(c1).UPDDATE     ' 更新日付
                .SENDFLAG = temp(c1).SENDFLAG   ' 送信フラグ
                .SENDDATE = temp(c1).SENDDATE   ' 送信日付
            End With
        Else
            With typ_rsz(c0)
                .CRYNUM = ""       ' 結晶番号
            End With
        End If
        rs.Close
    Next
    s_cmmc001db2_sql = FUNCTION_RETURN_SUCCESS
End Function
