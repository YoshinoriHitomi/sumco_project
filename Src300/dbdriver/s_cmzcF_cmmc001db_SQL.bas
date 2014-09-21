Attribute VB_Name = "s_cmzcF_cmmc001db_SQL"

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
            .CRYNUM = rs("CRYNUM")           ' 結晶番号
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



