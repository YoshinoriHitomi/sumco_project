Attribute VB_Name = "s_XSDC4_SQL"
'不良内訳 (XSDC4) ｱｸｾｽ関数


'***テーブル「XSDC4」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●不良内訳
Public Type typ_XSDC4
    XTALC4 As String * 12      ' ﾌﾞﾛｯｸID･結晶番号
    INPOSC4 As Integer         ' 結晶内開始位置
    KCKNTC4 As Integer         ' 工程連番
    HINBC4 As String * 8       ' 品番
    REVNUMC4 As Integer        ' 製品番号改訂番号
    FACTORYC4 As String * 1    ' 工場
    OPEC4 As String * 1        ' 操業条件
    KNKTC4 As String * 5       ' 管理工程
    WKKTC4 As String * 5       ' 工程
    WKKDC4 As String * 2       ' 作業区分
    MACOC4 As Integer          ' 処理回数
    SXLIDC4 As String * 13     ' SXLID
    FCODEC4 As String * 3      ' 判定コード
    PUCUTLC4 As Integer        ' 不良長さ
    PUCUTWC4 As Long           ' 不良重量
    PUCUTMC4 As Integer        ' 不良枚数
    FKUBC4 As String * 1       ' 不良区分
    TDAYC4 As Date             ' 登録日付
    KDAYC4 As Date             ' 更新日付
    SUMITBC3 As String * 1     ' SUMMIT送信フラグ
    SNDKC3 As String * 1       ' 送信フラグ
    SNDDAYC3 As Date           ' 送信日付
End Type

'更新用
Public Type typ_XSDC4_Update
    XTALC4 As String           ' ﾌﾞﾛｯｸID･結晶番号
    INPOSC4 As String          ' 結晶内開始位置
    KCKNTC4 As String          ' 工程連番
    HINBC4 As String           ' 品番
    REVNUMC4 As String         ' 製品番号改訂番号
    FACTORYC4 As String        ' 工場
    OPEC4 As String            ' 操業条件
    KNKTC4 As String           ' 管理工程
    WKKTC4 As String           ' 工程
    WKKDC4 As String           ' 作業区分
    MACOC4 As String           ' 処理回数
    SXLIDC4 As String          ' SXLID
    FCODEC4 As String          ' 判定コード
    PUCUTLC4 As String         ' 不良長さ
    PUCUTWC4 As String         ' 不良重量
    PUCUTMC4 As String         ' 不良枚数
    FKUBC4 As String           ' 不良区分
    TDAYC4 As String           ' 登録日付
    KDAYC4 As String           ' 更新日付
    SUMITBC3 As String         ' SUMMIT送信フラグ
    SNDKC3 As String           ' 送信フラグ
    SNDDAYC3 As String         ' 送信日付
End Type

'●SELECT●

'概要      :テーブル「XSDC4」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDC4     ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDC4(records() As typ_XSDC4, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select * From XSDC4"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC4 = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    For i = 1 To recCnt
        With records(i)
            If IsNull(rs.Fields("XTALC4")) = False Then .XTALC4 = rs.Fields("XTALC4")           ' ﾌﾞﾛｯｸID･結晶番号
            If IsNull(rs.Fields("INPOSC4")) = False Then .INPOSC4 = rs.Fields("INPOSC4")         ' 結晶内開始位置
            If IsNull(rs.Fields("KCKNTC4")) = False Then .KCKNTC4 = rs.Fields("KCKNTC4")         ' 工程連番
            If IsNull(rs.Fields("HINBC4")) = False Then .HINBC4 = rs.Fields("HINBC4")           ' 品番
            If IsNull(rs.Fields("REVNUMC4")) = False Then .REVNUMC4 = rs.Fields("REVNUMC4")       ' 製品番号改訂番号
            If IsNull(rs.Fields("FACTORYC4")) = False Then .FACTORYC4 = rs.Fields("FACTORYC4")     ' 工場
            If IsNull(rs.Fields("OPEC4")) = False Then .OPEC4 = rs.Fields("OPEC4")             ' 操業条件
            If IsNull(rs.Fields("KNKTC4")) = False Then .KNKTC4 = rs.Fields("KNKTC4")           ' 管理工程
            If IsNull(rs.Fields("WKKTC4")) = False Then .WKKTC4 = rs.Fields("WKKTC4")           ' 工程
            If IsNull(rs.Fields("WKKDC4")) = False Then .WKKDC4 = rs.Fields("WKKDC4")           ' 作業区分
            If IsNull(rs.Fields("MACOC4")) = False Then .MACOC4 = rs.Fields("MACOC4")           ' 処理回数
            If IsNull(rs.Fields("SXLIDC4")) = False Then .SXLIDC4 = rs.Fields("SXLIDC4")         ' SXLID
            If IsNull(rs.Fields("FCODEC4")) = False Then .FCODEC4 = rs.Fields("FCODEC4")         ' 判定コード
            If IsNull(rs.Fields("PUCUTLC4")) = False Then .PUCUTLC4 = rs.Fields("PUCUTLC4")       ' 不良長さ
            If IsNull(rs.Fields("PUCUTWC4")) = False Then .PUCUTWC4 = rs.Fields("PUCUTWC4")       ' 不良重量
            If IsNull(rs.Fields("PUCUTMC4")) = False Then .PUCUTMC4 = rs.Fields("PUCUTMC4")       ' 不良枚数
            If IsNull(rs.Fields("FKUBC4")) = False Then .FKUBC4 = rs.Fields("FKUBC4")           ' 不良区分
            If IsNull(rs.Fields("TDAYC4")) = False Then .TDAYC4 = rs.Fields("TDAYC4")           ' 登録日付
            If IsNull(rs.Fields("KDAYC4")) = False Then .KDAYC4 = rs.Fields("KDAYC4")           ' 更新日付
            If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")       ' SUMMIT送信フラグ
            If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")           ' 送信フラグ
            If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")       ' 送信日付
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC4 = FUNCTION_RETURN_SUCCESS
End Function


'●INSERT●

'概要      :テーブル「XSDC4」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDC4 　　  ,I  ,typ_XSDC4_Update   ,XSDC4更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDC4(pXSDC4 As typ_XSDC4_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      'レコード数
    Dim nowtime As Date
    Dim nowtime_sql     As String   'サーバ時間(SQL文)
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDC4_SQL.bas -- Function CreateXSDC4"
    sErrMsg = ""
    sDbName = "XSDC4"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku
   
'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC4
        
        sql = "INSERT INTO XSDC4 ("
        sql = sql & " XTALC4"           ' 1:ﾌﾞﾛｯｸID・結晶番号
        sql = sql & ",INPOSC4"          ' 2:結晶内開始位置
        sql = sql & ",KCKNTC4"          ' 3:工程連番
        sql = sql & ",HINBC4"           ' 4:品番
        sql = sql & ",REVNUMC4"         ' 5:製品番号改訂番号
        sql = sql & ",FACTORYC4"        ' 6:工場
        sql = sql & ",OPEC4"            ' 7:操業条件
        sql = sql & ",KNKTC4"           ' 8:管理工程
        sql = sql & ",WKKTC4"           ' 9:工程
        sql = sql & ",WKKDC4"           '10:作業区分
        sql = sql & ",MACOC4"           '11:処理回数
        sql = sql & ",SXLIDC4"          '12:SXLID
        sql = sql & ",FCODEC4"          '13:判定ｺｰﾄﾞ
        sql = sql & ",PUCUTLC4"         '14:不良長さ
        sql = sql & ",PUCUTWC4"         '15:不良重量
        sql = sql & ",PUCUTMC4"         '16:不良枚数
        sql = sql & ",FKUBC4"           '17:不良区分
        sql = sql & ",TDAYC4"           '18:登録日付
        sql = sql & ",KDAYC4"           '19:更新日付
        sql = sql & ",SUMITBC3"         '20:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",SNDKC3"           '21:送信ﾌﾗｸﾞ
        sql = sql & ",SNDDAYC3"         '22:送信日付
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:ﾌﾞﾛｯｸID・結晶番号
        If .XTALC4 <> "" Then
            sql = sql & " '" & .XTALC4 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:結晶内開始位置
        If .INPOSC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:工程連番
        If .KCKNTC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCKNTC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:品番
        If .HINBC4 <> "" Then
            sql = sql & ",'" & .HINBC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 5:製品番号改訂番号
        If .REVNUMC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.REVNUMC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:工場
        If .FACTORYC4 <> "" Then
            sql = sql & ",'" & .FACTORYC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:操業条件
        If .OPEC4 <> "" Then
            sql = sql & ",'" & .OPEC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 8:管理工程
        If .KNKTC4 <> "" Then
            sql = sql & ",'" & .KNKTC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 9:工程
        If .WKKTC4 <> "" Then
            sql = sql & ",'" & .WKKTC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '10:作業区分
        If .WKKDC4 <> "" Then
            sql = sql & ",'" & .WKKDC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '11:処理回数
        If .MACOC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.MACOC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '12:SXLID
        If .SXLIDC4 <> "" Then
            sql = sql & ",'" & .SXLIDC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        '13:判定ｺｰﾄﾞ
        If .FCODEC4 <> "" Then
            sql = sql & ",'" & .FCODEC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '14:不良長さ
        If .PUCUTLC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.PUCUTLC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:不良重量
        If .PUCUTWC4 <> "" Then
            sql = sql & ",'" & CStr(CLng(.PUCUTWC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:不良枚数
        If .PUCUTMC4 <> "" Then
            sql = sql & ",'" & CStr(CInt(.PUCUTMC4)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:不良区分
        If (.FKUBC4 <> "") Then
            sql = sql & ",'" & .FKUBC4 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '18:登録日付
        sql = sql & "," & nowtime_sql & vbLf

        '19:更新日付
        sql = sql & "," & nowtime_sql & vbLf

        '20:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '21:送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '22:送信日付
        sql = sql & ",NULL" & vbLf

        sql = sql & ")" & vbLf
    
        'SQLを実行
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If

    End With
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/29 SETsw kubota ------------------

    CreateXSDC4 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    CreateXSDC4 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function




