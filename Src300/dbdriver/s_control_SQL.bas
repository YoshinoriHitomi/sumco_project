Attribute VB_Name = "s_control_SQL"

'------------------------------------------------
' DBアクセス関数
'------------------------------------------------

'概要      :テーブル「TBCMB005」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO ,型           ,説明
'          :records()     ,O  ,typ_TBCMB005 ,抽出レコード
'          :sqlWhere      ,I  ,String       ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I  ,String       ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O  ,FUNCTION_RETURN ,抽出の成否
'説明      :
'履歴      :2001/08/24作成　野村 (2002/07 s_cmzcTBCMB005_SQL.basより移動)
Public Function DBDRV_GetTBCMB005(records() As typ_TBCMB005, Optional sqlWhere$ = vbNullString, Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
Dim sql As String       'SQL全体
Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
Dim rs As OraDynaset    'RecordSet
Dim recCnt As Long      'レコード数
Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select SYSCLASS, CLASS, CODE, INFO1, INFO2, INFO3, INFO4, INFO5, INFO6, INFO7, INFO8, INFO9, NOTE, TSTAFFID," & _
              " REGDATE, KSTAFFID, UPDDATE "
    sqlBase = sqlBase & "From TBCMB005"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetTBCMB005 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            .SYSCLASS = rs("SYSCLASS")       ' システム区分
            .Class = rs("CLASS")             ' 区分
            .CODE = rs("CODE")               ' コード
            .INFO1 = rs("INFO1")             ' 情報１
            .INFO2 = rs("INFO2")             ' 情報２
            .INFO3 = rs("INFO3")             ' 情報３
            .INFO4 = rs("INFO4")             ' 情報４
            .INFO5 = rs("INFO5")             ' 情報５
            .INFO6 = rs("INFO6")             ' 情報６
            .INFO7 = rs("INFO7")             ' 情報７
            .INFO8 = rs("INFO8")             ' 情報８
            .INFO9 = rs("INFO9")             ' 情報９
            .NOTE = rs("NOTE")               ' 備考
            .TSTAFFID = rs("TSTAFFID")       ' 登録社員ID
            .REGDATE = rs("REGDATE")         ' 登録日付
            .KSTAFFID = rs("KSTAFFID")       ' 更新社員ID
            .UPDDATE = rs("UPDDATE")         ' 更新日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetTBCMB005 = FUNCTION_RETURN_SUCCESS
End Function


