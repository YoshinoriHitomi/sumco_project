Attribute VB_Name = "s_XSDC3_SQL"
'工程実績 (XSDC3) ｱｸｾｽ関数


'***テーブル「XSDC3」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●工程実績
Public Type typ_XSDC3
    CRYNUMC3 As String * 12      ' ﾌﾞﾛｯｸID･結晶番号
    INPOSC3 As Integer           ' 結晶内開始位置
    KCNTC3 As Integer            ' 工程連番
    HINBC3 As String * 8         ' 品番
    REVNUMC3 As Integer          ' 製品番号改訂番号
    FACTORYC3 As String * 1       ' 工場
    OPEC3 As String * 1          ' 操業条件
    LENC3 As Integer             ' 長さ
    XTALC3 As String * 12        ' 結晶番号
    SXLIDC3 As String * 13        ' SXLID
    KNKTC3 As String * 5         ' 管理工程
    WKKTC3 As String * 5         ' 工程
    WKKBC3 As String * 2         ' 作業区分
    MACOC3 As Integer            ' 処理回数
    MODKBC3 As String * 1        ' 赤黒区分
    SUMKBC3 As String * 1        ' 集計区分
    FRKNKTC3 As String * 5       ' (受入)管理工程
    FRWKKTC3 As String * 5       ' (受入)工程
    FRWKKBC3 As String * 2       ' (受入)作業区分
    FRMACOC3 As Integer          ' (受入)処理回数
    TOWNKTC3 As String * 5       ' (払出)管理工程
    TOWKKTC3 As String * 5       ' (払出)工程
    TOMACOC3 As Integer          ' (払出)処理回数
    FRLC3 As Integer             ' 受入長さ
'    FRWC3 As Integer             ' 受入重量←データ型をLongに変更 2003/09/24 オーバーフローするため
    FRWC3 As Long                ' 受入重量
    FRMC3 As Integer             ' 受入枚数
    FULC3 As Integer             ' 不良長さ
    FUWC3 As Integer             ' 不良重量
    FUMC3 As Integer             ' 不良枚数
    LOSWC3 As Integer            ' ロス長さ
    LOSLC3 As Integer            ' ロス重量
    LOSMC3 As Integer            ' ロス枚数
    TOLC3 As Integer             ' 払出長さ
'    TOWC3 As Integer            ' 払出重量 ←データ型をLongに変更 2003/09/22 オーバーフローするため
    TOWC3 As Long                ' 払出重量
    TOMC3 As Integer             ' 払出枚数
    SUMITLC3 As Integer          ' SUMIT長さ
'    SUMITWC3 As Integer          ' SUMIT重量←データ型をLongに変更 2003/09/23 オーバーフローするため
    SUMITWC3 As Long             ' SUMIT重量
    SUMITMC3 As Integer          ' SUMIT枚数
    MOTHINC3 As String * 12      ' 振替品番(元)
    XTWORKC3 As String * 2       ' 製造工場
    WFWORKC3 As String * 2       ' ｳｪｰﾊ製造
    STATIMEC3 As Date            ' 処理時間開始
    STOTIMEC3 As Date            ' 処理時間終了
    ETIMEC3 As Date              ' 実績時間
    HOLDCC3 As String * 3        ' ホールドコード
    HOLDBC3 As String * 1        ' ホールド区分
    LDFRCC3 As String * 3        ' 格下コード
    LDFRBC3 As String * 1        ' 格下区分
    TSTAFFC3 As String * 8       ' 登録社員ID
    TDAYC3 As Date               ' 登録日付
    KSTAFFC3 As String * 8       ' 更新社員ID
    KDAYC3 As Date               ' 更新日付
    SUMITBC3 As String * 1       ' SUMIT送信フラグ
    SNDKC3 As String * 1         ' 送信フラグ
    SNDDAYC3 As Date             ' 登録日付
    'add start 2003/03/25 hitec)matsumoto ----
    SUMDAYC3 As Date             ' SUMCO時間
    PAYCLASSC3 As String * 1     ' 転送先工場フラグ
    'add end 2003/03/25 hitec)matsumoto ----
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTC3 As String * 1       ' 新規／再切区分 '1':再切
    HINBFLGC3 As String * 1      ' 代表品番フラグ '1'：代表品番
    '2005/11
    RPCRYNUMC3 As String
''>>>>> パワーON時間追加対応 SETsw H.Iwamoto 2005/11/28
    PROTMC3     As Integer       ' パワーON時間(時)
    PROMNC3     As Integer       ' パワーON時間(分)
    PROTM2C3    As Integer       ' (累計)パワーON時間(時)
    PROMN2C3    As Integer       ' (累計)パワーON時間(分)
''<<<<< パワーON時間追加対応 SETsw H.Iwamoto 2005/11/28
    PLANTCATC3 As String * 2     ' 向先　2007/08/15 SPK Tsutsumi
End Type

'更新用
Public Type typ_XSDC3_Update
    CRYNUMC3 As String           ' ﾌﾞﾛｯｸID･結晶番号
    INPOSC3 As String            ' 結晶内開始位置
    KCNTC3 As String             ' 工程連番
    HINBC3 As String             ' 品番
    REVNUMC3 As String           ' 製品番号改訂番号
    FACTORYC3 As String          ' 工場
    OPEC3 As String              ' 操業条件
    LENC3 As String              ' 長さ
    XTALC3 As String             ' 結晶番号
    SXLIDC3 As String             ' SXLID
    KNKTC3 As String             ' 管理工程
    WKKTC3 As String             ' 工程
    WKKBC3 As String             ' 作業区分
    MACOC3 As String             ' 処理回数
    MODKBC3 As String            ' 赤黒区分
    SUMKBC3 As String            ' 集計区分
    FRKNKTC3 As String           ' (受入)管理工程
    FRWKKTC3 As String           ' (受入)工程
    FRWKKBC3 As String           ' (受入)作業区分
    FRMACOC3 As String           ' (受入)処理回数
    TOWNKTC3 As String           ' (払出)管理工程
    TOWKKTC3 As String           ' (払出)工程
    TOMACOC3 As String           ' (払出)処理回数
    FRLC3 As String              ' 受入長さ
    FRWC3 As String              ' 受入重量
    FRMC3 As String              ' 受入枚数
    FULC3 As String              ' 不良長さ
    FUWC3 As String              ' 不良重量
    FUMC3 As String              ' 不良枚数
    LOSWC3 As String             ' ロス長さ
    LOSLC3 As String             ' ロス重量
    LOSMC3 As String             ' ロス枚数
    TOLC3 As String              ' 払出長さ
    TOWC3 As String              ' 払出重量
    TOMC3 As String              ' 払出枚数
    SUMITLC3 As String           ' SUMIT長さ
    SUMITWC3 As String           ' SUMIT重量
    SUMITMC3 As String           ' SUMIT枚数
    MOTHINC3 As String           ' 振替品番(元)
    XTWORKC3 As String           ' 製造工場
    WFWORKC3 As String           ' ｳｪｰﾊ製造
    STATIMEC3 As Date            ' 処理時間開始
    STOTIMEC3 As Date            ' 処理時間終了
    ETIMEC3 As Date              ' 実績時間
    HOLDCC3 As String            ' ホールドコード
    HOLDBC3 As String            ' ホールド区分
    LDFRCC3 As String            ' 格下コード
    LDFRBC3 As String            ' 格下区分
    TSTAFFC3 As String           ' 登録社員ID
    TDAYC3 As Date               ' 登録日付
    KSTAFFC3 As String           ' 更新社員ID
    KDAYC3 As Date               ' 更新日付
    SUMITBC3 As String           ' SUMIT送信フラグ
    SNDKC3 As String             ' 送信フラグ
    SNDDAYC3 As Date             ' 登録日付
    MODMACOC3 As String * 2       ' 赤黒処理回数
    KAKUCC3 As String * 5         ' 確定コード
    'add start 2003/03/25 hitec)matsumoto ----
    SUMDAYC3 As Date             ' SUMCO時間
    PAYCLASSC3 As String         ' 転送先工場フラグ
    'add end 2003/03/25 hitec)matsumoto ----
    '2003.06.11 (SPK)Y.Katabami tuika
    CUTCNTC3 As String * 1       ' 新規／再切区分 '1':再切
    HINBFLGC3 As String * 1      ' 代表品番フラグ '1'：代表品番
    '2005/11
    RPCRYNUMC3 As String
''>>>>> パワーON時間追加対応 SETsw H.Iwamoto 2005/11/28
    PROTMC3     As Integer       ' パワーON時間(時)
    PROMNC3     As Integer       ' パワーON時間(分)
    PROTM2C3    As Integer       ' (累計)パワーON時間(時)
    PROMN2C3    As Integer       ' (累計)パワーON時間(分)
''<<<<< パワーON時間追加対応 SETsw H.Iwamoto 2005/11/28
    PLANTCATC3 As String * 2     ' 向先　2007/08/15 SPK Tsutsumi
End Type

'●SELECT●

'概要      :テーブル「XSDC3」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDC3     ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDC3(records() As typ_XSDC3, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long


    ''SQLを組み立てる
    sqlBase = "Select * From XSDC3"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC3 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("CRYNUMC3")) = False Then .CRYNUMC3 = rs.Fields("CRYNUMC3")
            If IsNull(rs.Fields("INPOSC3")) = False Then .INPOSC3 = rs.Fields("INPOSC3")
            If IsNull(rs.Fields("KCNTC3")) = False Then .KCNTC3 = rs.Fields("KCNTC3")
            If IsNull(rs.Fields("HINBC3")) = False Then .HINBC3 = rs.Fields("HINBC3")
            If IsNull(rs.Fields("REVNUMC3")) = False Then .REVNUMC3 = rs.Fields("REVNUMC3")
            If IsNull(rs.Fields("FACTORYC3")) = False Then .FACTORYC3 = rs.Fields("FACTORYC3")
            If IsNull(rs.Fields("OPEC3")) = False Then .OPEC3 = rs.Fields("OPEC3")
            If IsNull(rs.Fields("LENC3")) = False Then .LENC3 = rs.Fields("LENC3")
            If IsNull(rs.Fields("XTALC3")) = False Then .XTALC3 = rs.Fields("XTALC3")
            If IsNull(rs.Fields("SXLIDC3")) = False Then .SXLIDC3 = rs.Fields("SXLIDC3")
            If IsNull(rs.Fields("KNKTC3")) = False Then .KNKTC3 = rs.Fields("KNKTC3")
            If IsNull(rs.Fields("WKKTC3")) = False Then .WKKTC3 = rs.Fields("WKKTC3")
            If IsNull(rs.Fields("WKKBC3")) = False Then .WKKBC3 = rs.Fields("WKKBC3")
            If IsNull(rs.Fields("MACOC3")) = False Then .MACOC3 = rs.Fields("MACOC3")
            If IsNull(rs.Fields("MODKBC3")) = False Then .MODKBC3 = rs.Fields("MODKBC3")
            If IsNull(rs.Fields("SUMKBC3")) = False Then .SUMKBC3 = rs.Fields("SUMKBC3")
            If IsNull(rs.Fields("FRKNKTC3")) = False Then .FRKNKTC3 = rs.Fields("FRKNKTC3")
            If IsNull(rs.Fields("FRWKKTC3")) = False Then .FRWKKTC3 = rs.Fields("FRWKKTC3")
            If IsNull(rs.Fields("FRWKKBC3")) = False Then .FRWKKBC3 = rs.Fields("FRWKKBC3")
            If IsNull(rs.Fields("FRMACOC3")) = False Then .FRMACOC3 = rs.Fields("FRMACOC3")
            If IsNull(rs.Fields("TOWNKTC3")) = False Then .TOWNKTC3 = rs.Fields("TOWNKTC3")
            If IsNull(rs.Fields("TOWKKTC3")) = False Then .TOWKKTC3 = rs.Fields("TOWKKTC3")
            If IsNull(rs.Fields("TOMACOC3")) = False Then .TOMACOC3 = rs.Fields("TOMACOC3")
            If IsNull(rs.Fields("FRLC3")) = False Then .FRLC3 = rs.Fields("FRLC3")
            If IsNull(rs.Fields("FRWC3")) = False Then .FRWC3 = rs.Fields("FRWC3")
            If IsNull(rs.Fields("FRMC3")) = False Then .FRMC3 = rs.Fields("FRMC3")
            If IsNull(rs.Fields("FULC3")) = False Then .FULC3 = rs.Fields("FULC3")
            If IsNull(rs.Fields("FUWC3")) = False Then .FUWC3 = rs.Fields("FUWC3")
            If IsNull(rs.Fields("FUMC3")) = False Then .FUMC3 = rs.Fields("FUMC3")
            If IsNull(rs.Fields("LOSWC3")) = False Then .LOSWC3 = rs.Fields("LOSWC3")
            If IsNull(rs.Fields("LOSLC3")) = False Then .LOSLC3 = rs.Fields("LOSLC3")
            If IsNull(rs.Fields("LOSMC3")) = False Then .LOSMC3 = rs.Fields("LOSMC3")
            If IsNull(rs.Fields("TOLC3")) = False Then .TOLC3 = rs.Fields("TOLC3")
            If IsNull(rs.Fields("TOWC3")) = False Then .TOWC3 = rs.Fields("TOWC3")
            If IsNull(rs.Fields("TOMC3")) = False Then .TOMC3 = rs.Fields("TOMC3")
            If IsNull(rs.Fields("SUMITLC3")) = False Then .SUMITLC3 = rs.Fields("SUMITLC3")
            If IsNull(rs.Fields("SUMITWC3")) = False Then .SUMITWC3 = rs.Fields("SUMITWC3")
            If IsNull(rs.Fields("SUMITMC3")) = False Then .SUMITMC3 = rs.Fields("SUMITMC3")
            If IsNull(rs.Fields("MOTHINC3")) = False Then .MOTHINC3 = rs.Fields("MOTHINC3")
            If IsNull(rs.Fields("XTWORKC3")) = False Then .XTWORKC3 = rs.Fields("XTWORKC3")
            If IsNull(rs.Fields("WFWORKC3")) = False Then .WFWORKC3 = rs.Fields("WFWORKC3")
            If IsNull(rs.Fields("STATIMEC3")) = False Then .STATIMEC3 = rs.Fields("STATIMEC3")
            If IsNull(rs.Fields("STOTIMEC3")) = False Then .STOTIMEC3 = rs.Fields("STOTIMEC3")
            If IsNull(rs.Fields("ETIMEC3")) = False Then .ETIMEC3 = rs.Fields("ETIMEC3")
            If IsNull(rs.Fields("HOLDCC3")) = False Then .HOLDCC3 = rs.Fields("HOLDCC3")
            If IsNull(rs.Fields("HOLDBC3")) = False Then .HOLDBC3 = rs.Fields("HOLDBC3")
            If IsNull(rs.Fields("LDFRCC3")) = False Then .LDFRCC3 = rs.Fields("LDFRCC3")
            If IsNull(rs.Fields("LDFRBC3")) = False Then .LDFRBC3 = rs.Fields("LDFRBC3")
            If IsNull(rs.Fields("TSTAFFC3")) = False Then .TSTAFFC3 = rs.Fields("TSTAFFC3")
            If IsNull(rs.Fields("TDAYC3")) = False Then .TDAYC3 = rs.Fields("TDAYC3")
            If IsNull(rs.Fields("KSTAFFC3")) = False Then .KSTAFFC3 = rs.Fields("KSTAFFC3")
            If IsNull(rs.Fields("KDAYC3")) = False Then .KDAYC3 = rs.Fields("KDAYC3")
            If IsNull(rs.Fields("SUMITBC3")) = False Then .SUMITBC3 = rs.Fields("SUMITBC3")
            If IsNull(rs.Fields("SNDKC3")) = False Then .SNDKC3 = rs.Fields("SNDKC3")
            If IsNull(rs.Fields("SNDDAYC3")) = False Then .SNDDAYC3 = rs.Fields("SNDDAYC3")
            'add start 2003/03/25 hitec)matsumoto ------
            If IsNull(rs.Fields("SUMDAYC3")) = False Then .SUMDAYC3 = rs.Fields("SUMDAYC3") 'SUMCO時間
            If IsNull(rs.Fields("PAYCLASSC3")) = False Then .PAYCLASSC3 = rs.Fields("PAYCLASSC3") '転送先フラグ
           'add end 2003/03/25 hitec)matsumoto ------
            '2003.06.11 (SPK)Y.katabami tuika
            If IsNull(rs.Fields("CUTCNTC3")) = False Then .CUTCNTC3 = rs.Fields("CUTCNTC3")
            If IsNull(rs.Fields("HINBFLGC3")) = False Then .HINBFLGC3 = rs.Fields("HINBFLGC3")
            '2005/11
            If IsNull(rs.Fields("RPCRYNUMC3")) = False Then .RPCRYNUMC3 = rs.Fields("RPCRYNUMC3")
            If IsNull(rs.Fields("PLANTCATC3")) = False Then .PLANTCATC3 = rs.Fields("PLANTCATC3")   ' 2007/09/04 SPK Tsutsumi Add
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC3 = FUNCTION_RETURN_SUCCESS
End Function


'●INSERT●  NULLの場合、charならスペース、NumberならNULLを入れる

'概要      :テーブル「XSDC3」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDC3 　　  ,I  ,typ_XSDC3_Update   ,XSDC3更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDC3(pXSDC3 As typ_XSDC3_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sql2 As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
    Dim rs2 As OraDynaset    'RecordSet
'    Dim recCnt As Long      'レコード数
    Dim nowtime         As Date
    Dim nowtime_sql     As String   'サーバ時間(SQL文)
    Dim justNowTime     As Date
    Dim justNowTime_sql As String   'サーバ時間(SQL文)

    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDC3_SQL.bas -- Function CreateXSDC3"
    sErrMsg = ""
    sDbName = "XSDC3"
    'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    'justNowTime = Format(Time, "hh:mm:ss")
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku
    justNowTime = Format(nowtime, "hh:mm:ss")
    
'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/29 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    justNowTime_sql = "TO_DATE('" & Format$(justNowTime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC3
        
        sql = "INSERT INTO XSDC3 ("
        sql = sql & " CRYNUMC3"         ' 1:ﾌﾞﾛｯｸID・結晶番号
        sql = sql & ",INPOSC3"          ' 2:結晶内開始位置
        sql = sql & ",KCNTC3"           ' 3:工程連番
        sql = sql & ",HINBC3"           ' 4:品番
        sql = sql & ",REVNUMC3"         ' 5:製品番号改訂番号
        sql = sql & ",FACTORYC3"        ' 6:工場
        sql = sql & ",OPEC3"            ' 7:操業条件
        sql = sql & ",LENC3"            ' 8:長さ
        sql = sql & ",XTALC3"           ' 9:結晶番号
        sql = sql & ",SXLIDC3"          '10:SXLID
        sql = sql & ",KNKTC3"           '11:管理工程
        sql = sql & ",WKKTC3"           '12:工程
        sql = sql & ",WKKBC3"           '13:作業区分
        sql = sql & ",MACOC3"           '14:処理回数
        sql = sql & ",MODKBC3"          '15:赤黒区分
        sql = sql & ",SUMKBC3"          '16:集計区分
        sql = sql & ",FRKNKTC3"         '17:(受入)管理工程
        sql = sql & ",FRWKKTC3"         '18:(受入)工程
        sql = sql & ",FRWKKBC3"         '19:(受入)作業区分
        sql = sql & ",FRMACOC3"         '20:(受入)処理回数
        sql = sql & ",TOWNKTC3"         '21:(払出)管理工程
        sql = sql & ",TOWKKTC3"         '22:(払出)工程
        sql = sql & ",TOMACOC3"         '23:(払出)処理回数
        sql = sql & ",FRLC3"            '24:受入長さ
        sql = sql & ",FRWC3"            '25:受入重量
        sql = sql & ",FRMC3"            '26:受入枚数
        sql = sql & ",FULC3"            '27:不良長さ
        sql = sql & ",FUWC3"            '28:不良重量
        sql = sql & ",FUMC3"            '29:不良枚数
        sql = sql & ",LOSWC3"           '30:ロス長さ
        sql = sql & ",LOSLC3"           '31:ロス重量
        sql = sql & ",LOSMC3"           '32:ロス枚数
        sql = sql & ",TOLC3"            '33:払出長さ
        sql = sql & ",TOWC3"            '34:払出重量
        sql = sql & ",TOMC3"            '35:払出枚数
        sql = sql & ",SUMITLC3"         '36:SUMMIT長さ
        sql = sql & ",SUMITWC3"         '37:SUMMIT重量
        sql = sql & ",SUMITMC3"         '38:SUMMIT枚数
        sql = sql & ",MOTHINC3"         '39:振替品番(元)
        sql = sql & ",XTWORKC3"         '40:製造工場
        sql = sql & ",WFWORKC3"         '41:ｳｪｰﾊ製造
        sql = sql & ",STATIMEC3"        '42:処理時間開始
        sql = sql & ",STOTIMEC3"        '43:処理時間終了
        sql = sql & ",ETIMEC3"          '44:実績時間
        sql = sql & ",HOLDCC3"          '45:ﾎｰﾙﾄﾞｺｰﾄﾞ
        sql = sql & ",HOLDBC3"          '46:ﾎｰﾙﾄﾞ区分
        sql = sql & ",LDFRCC3"          '47:格下ｺｰﾄﾞ
        sql = sql & ",LDFRBC3"          '48:格下区分
        sql = sql & ",TSTAFFC3"         '49:登録社員ID
        sql = sql & ",TDAYC3"           '50:登録日付
        sql = sql & ",KSTAFFC3"         '51:更新社員ID
        sql = sql & ",KDAYC3"           '52:更新日付
        sql = sql & ",SUMITBC3"         '53:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",SNDKC3"           '54:送信ﾌﾗｸﾞ
        sql = sql & ",SNDDAYC3"         '55:送信日付
        sql = sql & ",SUMDAYC3"         '56:SUMCO時間
        sql = sql & ",PAYCLASSC3"       '57:払出区分
        sql = sql & ",CUTCNTC3"         '58:新規／再切区分
        sql = sql & ",HINBFLGC3"        '59:代表品番フラグ
        sql = sql & ",RPCRYNUMC3"       '60:親ﾌﾞﾛｯｸID
        sql = sql & ",PROTMC3"          '61:パワーON時間(時)
        sql = sql & ",PROMNC3"          '62:パワーON時間(分)
        sql = sql & ",PROTM2C3"         '63:(累計)パワーON時間(時)
        sql = sql & ",PROMN2C3"         '64:(累計)パワーON時間(分)
        sql = sql & ",PLANTCATC3"       '65:向先
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf
        
        ' 1:ﾌﾞﾛｯｸID・結晶番号
        If .CRYNUMC3 <> "" Then
            sql = sql & " '" & .CRYNUMC3 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:結晶内開始位置
        If .INPOSC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 3:工程連番
        If .KCNTC3 <> "" Then
            sql = sql & ",'" & .KCNTC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 4:品番
        If .HINBC3 <> "" Then
            sql = sql & ",'" & .HINBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        ' 5:製品番号改訂番号
        If .REVNUMC3 <> "" Then
            sql = sql & ",'" & .REVNUMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 6:工場
        If .FACTORYC3 <> "" Then
            sql = sql & ",'" & .FACTORYC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 7:操業条件
        If .OPEC3 <> "" Then
            sql = sql & ",'" & .OPEC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        ' 8:長さ
        If .LENC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.LENC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 9:結晶番号
        If .XTALC3 <> "" Then
            sql = sql & ",'" & .XTALC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '10:SXLID
        If .SXLIDC3 <> "" And Left(.SXLIDC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .SXLIDC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(13) & "'" & vbLf
        End If

        '11:管理工程
        If .KNKTC3 <> "" Then
            sql = sql & ",'" & .KNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '12:工程
        If .WKKTC3 <> "" Then
            sql = sql & ",'" & .WKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '13:作業区分
        If .WKKBC3 <> "" Then
            sql = sql & ",'" & .WKKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '14:処理回数
        If .MACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.MACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:赤黒区分
        If .MODKBC3 <> "" Then
            sql = sql & ",'" & .MODKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '16:集計区分
        If .SUMKBC3 <> "" Then
            sql = sql & ",'" & .SUMKBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '17:(受入)管理工程
        If .FRKNKTC3 <> "" Then
            sql = sql & ",'" & .FRKNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '18:(受入)工程
        If .FRWKKTC3 <> "" Then
            sql = sql & ",'" & .FRWKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If
        
        '19:(受入)作業区分
        If .FRWKKBC3 <> "" Then
            sql = sql & ",'" & .FRWKKBC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '20:(受入)処理回数
        If .FRMACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRMACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:(払出)管理工程
        If .TOWNKTC3 <> "" Then
            sql = sql & ",'" & .TOWNKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '22:(払出)工程
        If .TOWKKTC3 <> "" Then
            sql = sql & ",'" & .TOWKKTC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '23:(払出)処理回数
        If .TOMACOC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.TOMACOC3)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        'ﾚｺｰﾄﾞの前回の値を取得
        sql2 = "SELECT TOLC3, TOWC3, TOMC3 FROM XSDC3 WHERE CRYNUMC3 = '" & .CRYNUMC3
        sql2 = sql2 & "' AND INPOSC3 = " & .INPOSC3
        sql2 = sql2 & " AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & .CRYNUMC3
        sql2 = sql2 & "' AND MODKBC3 != '1"   ' 赤処理レコード以外
        sql2 = sql2 & "' AND INPOSC3 = " & .INPOSC3 & ")"
        Set rs2 = OraDB.DBCreateDynaset(sql2, ORADYN_DEFAULT)

        '24:受入長さ
        If .FRLC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRLC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOLC3")) Then
                    .FRLC3 = rs2.Fields("TOLC3")
                    sql = sql & ",'" & CStr(CInt(.FRLC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '25:受入重量
        If .FRWC3 <> "" Then
            sql = sql & ",'" & CStr(CLng(.FRWC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOWC3")) Then
                    .FRWC3 = rs2.Fields("TOWC3")
                    sql = sql & ",'" & CStr(CLng(.FRWC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '26:受入枚数
        If .FRMC3 <> "" Then
            sql = sql & ",'" & CStr(CInt(.FRMC3)) & "'" & vbLf
        Else
            If rs2.RecordCount = 0 Or .WKKTC3 = "CC705" Then
                sql = sql & ",0" & vbLf
            Else
                If Not IsNull(rs2.Fields("TOMC3")) Then
                    .FRMC3 = rs2.Fields("TOMC3")
                    sql = sql & ",'" & CStr(CInt(.FRMC3)) & "'" & vbLf
                Else
                    sql = sql & ",0" & vbLf
                End If
            End If
        End If

        '27:不良長さ
        If .FULC3 <> "" Then
            sql = sql & ",'" & .FULC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '28:不良重量
        If .FUWC3 <> "" Then
            sql = sql & ",'" & .FUWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '29:不良枚数
        If .FUMC3 <> "" Then
            sql = sql & ",'" & .FUMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '30:ロス長さ
        If .LOSWC3 <> "" Then
            sql = sql & ",'" & .LOSWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '31:ロス長さ
        If .LOSLC3 <> "" Then
            sql = sql & ",'" & .LOSLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '32:ロス枚数
        If .LOSMC3 <> "" Then
            sql = sql & ",'" & .LOSMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '33:払出長さ
        If .TOLC3 <> "" Then
            sql = sql & ",'" & .TOLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '34:払出重量
        If .TOWC3 <> "" Then
            sql = sql & ",'" & .TOWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '35:払出枚数
        If .TOMC3 <> "" Then
            sql = sql & ",'" & .TOMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '36:SUMMIT長さ
        If .SUMITLC3 <> "" Then
            sql = sql & ",'" & .SUMITLC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '37:SUMMIT重量
        If .SUMITWC3 <> "" Then
            sql = sql & ",'" & .SUMITWC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '38:SUMMIT枚数
        If .SUMITMC3 <> "" Then
            sql = sql & ",'" & .SUMITMC3 & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '39:振替品番(元)
        If .MOTHINC3 <> "" Then
            sql = sql & ",'" & .MOTHINC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '40:製造工場
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '41:ｳｪｰﾊ製造
        If .WFWORKC3 <> "" Then
            sql = sql & ",'" & .WFWORKC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '42:処理時間開始
        sql = sql & "," & justNowTime_sql & vbLf

        '43:処理時間終了
        sql = sql & "," & justNowTime_sql & vbLf

        '44:実績時間
        sql = sql & ",0" & vbLf

        '45:ﾎｰﾙﾄﾞｺｰﾄﾞ
        If .HOLDCC3 <> "" Then
            sql = sql & ",'" & .HOLDCC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '46:ﾎｰﾙﾄﾞ区分
        If .HOLDBC3 <> "" Then
            sql = sql & ",'" & .HOLDBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '47:格下ｺｰﾄﾞ
        If .LDFRCC3 <> "" Then
            sql = sql & ",'" & .LDFRCC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '48:格下区分
        If .LDFRBC3 <> "" Then
            sql = sql & ",'" & .LDFRBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '49:登録社員ID
        If XSDC3_StaffID <> "" Then
            sql = sql & ",'" & XSDC3_StaffID & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '50:登録日付
        sql = sql & "," & nowtime_sql & vbLf

        '51:更新社員ID
        If .KSTAFFC3 <> "" Then
            sql = sql & ",'" & .KSTAFFC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(8) & "'" & vbLf
        End If

        '52:更新日付
        sql = sql & "," & nowtime_sql & vbLf

        '53:SUMMIT送信ﾌﾗｸﾞ
        If .SUMITBC3 <> "" Then
            sql = sql & ",'" & .SUMITBC3 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '54:送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '55:送信日付
        sql = sql & ",NULL" & vbLf

        '56:SUMCO時間
        sql = sql & ",TO_DATE('" & Format$(CalcSumcoTime(nowtime), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf

        '57:払出区分
        If .PAYCLASSC3 <> "" Then
            sql = sql & ",'" & .PAYCLASSC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If
        
        '58:新規／再切区分
        If .CUTCNTC3 <> "" And Left(.CUTCNTC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '59:代表品番フラグ
        If .HINBFLGC3 <> "" And Left(.HINBFLGC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .HINBFLGC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '60:親ﾌﾞﾛｯｸID
        If .RPCRYNUMC3 <> "" And Left(.RPCRYNUMC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMC3 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '61:パワーON時間(時)
        '62:パワーON時間(分)
        If .PROTMC3 = 0 And .PROMNC3 = 0 Then
            sql = sql & ",NULL" & vbLf
            sql = sql & ",NULL" & vbLf
        Else
            sql = sql & ",'" & CStr(.PROTMC3) & "'" & vbLf
            sql = sql & ",'" & CStr(.PROMNC3) & "'" & vbLf
        End If
        
        '63:(累計)パワーON時間(時)
        '64:(累計)パワーON時間(分)
        If .PROTM2C3 = 0 And .PROMN2C3 = 0 Then
            sql = sql & ",NULL" & vbLf
            sql = sql & ",NULL" & vbLf
        Else
            sql = sql & ",'" & CStr(.PROTM2C3) & "'" & vbLf
            sql = sql & ",'" & CStr(.PROMN2C3) & "'" & vbLf
        End If

        '65:向先
        If .PLANTCATC3 <> "" And Left(.PLANTCATC3, 1) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATC3 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If
        
        sql = sql & ")" & vbLf
    
        'SQLを実行
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
        
    End With
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/29 SETsw kubota ------------------

    CreateXSDC3 = FUNCTION_RETURN_SUCCESS

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
    CreateXSDC3 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function
