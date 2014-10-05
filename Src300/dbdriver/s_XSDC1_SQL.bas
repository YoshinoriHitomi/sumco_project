Attribute VB_Name = "s_XSDC1_SQL"
'結晶引上ﾃｰﾌﾞﾙ(XSDC1) ｱｸｾｽ関数

'●TEST用

'***テーブル「XSDC1」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●結晶引上
Public Type typ_XSDC1
    XTALC1 As String * 12        ' 結晶番号
    KNKTC1 As String * 5         ' 管理工程ｺｰﾄﾞ
    WKKTC1 As String * 5         ' 工程ｺｰﾄﾞ
    LENTOC1 As Integer           ' 長さ（TOP）
    LENTKC1 As Integer           ' 長さ（直胴）
    LENTAC1 As Integer           ' 長さ（TAIL）
    PUFRELC1 As Integer          ' ﾌﾘｰ長さ
    DIA1C1 As Integer            ' 直胴直径1
    DIA2C1 As Integer            ' 直胴直径2
    DIA3C1 As Integer            ' 直胴直径3
    WGHTTOC1 As Long             ' 重量（TOP）
    WGHTTKC1 As Long             ' 重量（直胴）
    WGHTTAC1 As Long             ' 重量（TAIL）
    WGHTFRC1 As Long             ' 重量（ﾌﾘｰ長さ）
    PUTCUTWC1 As Long            ' ﾄｯﾌﾟｶｯﾄ重量
    PUWC1 As Long                ' 引上げ重量
    PUHINBC1 As String * 8       ' 狙い品番
    PUCHAGC1 As Long             ' ﾁｬｰｼﾞ量
    KAKOUBC1 As String * 1       ' 加工区分
    KEIDAYC1 As Date             ' 計上日付
    SEEDC1 As String * 4         ' ｼｰﾄﾞ
    PUBTKBC1 As String * 3       ' BOT状況区分
    JDGECC1 As String * 3        ' 判定ｺｰﾄﾞ
    PWTIMEC1 As Double           ' ﾊﾟﾜｰ時間
    ADDOPPC1 As Integer          ' 追加ﾄﾞｰﾌﾟ位置
    ADDOPCC1 As String * 7       ' 追加ﾄﾞｰﾊﾟﾝﾄ種類
    ADDOPVC1 As Long             ' 追加ﾄﾞｰﾌﾟ量
    ADDOPNC1 As String * 20      ' 追加ﾄﾞｰﾌﾟ名
    TSTAFFC1 As String * 8       ' 登録社員ID
    TDAYC1 As Date               ' 登録日付
    KSTAFFC1 As String * 8       ' 更新社員ID
    KDAYC1 As Date               ' 更新日付
    SUMITBC1 As String * 1       ' SUMMIT送信ﾌﾗｸﾞ
    SNDKC1 As String * 1         ' 送信ﾌﾗｸﾞ
    SNDDAYC1 As Date             ' 送信日付
    SUIFLG As String * 1         ' 推定FLG     2003/10/27 tuku
    PUREVNUMC1 As Integer        ' 狙い品番製造番号改訂番号　04/09/28 ooba
    PUFACTORYC1 As String * 1    ' 狙い品番工場　04/09/28 ooba
    PUOPEC1 As String * 1        ' 狙い品番操業条件　04/09/28 ooba
    LENPUFRC1 As Integer         ' 引上ﾌﾘｰ長さ　05/07/19 ooba
    WGHTPUFRC1 As Long           ' 引上ﾌﾘｰ重量　05/07/19 ooba
    WGHTCUTRHC1 As Long          ' 切断後良品重量　05/07/19 ooba
''09/02/16 FAE)akiyama start
    SUICHARGE As Long                           '推定チャージ量
''09/02/16 FAE)akiyama end
End Type
    
    
'更新用
Public Type typ_XSDC1_Update
    XTALC1 As String            ' 結晶番号
    KNKTC1 As String            ' 管理工程ｺｰﾄﾞ
    WKKTC1 As String            ' 工程ｺｰﾄﾞ
    LENTOC1 As String           ' 長さ（TOP）
    LENTKC1 As String           ' 長さ（直胴）
    LENTAC1 As String           ' 長さ（TAIL）
    PUFRELC1 As String          ' ﾌﾘｰ長さ
    DIA1C1 As String            ' 直胴直径1
    DIA2C1 As String            ' 直胴直径2
    DIA3C1 As String            ' 直胴直径3
    WGHTTOC1 As String          ' 重量（TOP）
    WGHTTKC1 As String          ' 重量（直胴）
    WGHTTAC1 As String          ' 重量（TAIL）
    WGHTFRC1 As String          ' 重量（ﾌﾘｰ長さ）
    PUTCUTWC1 As String         ' ﾄｯﾌﾟｶｯﾄ重量
    PUWC1 As String             ' 引上げ重量
    PUHINBC1 As String          ' 狙い品番
    PUCHAGC1 As String          ' ﾁｬｰｼﾞ量
    KAKOUBC1 As String          ' 加工区分
    KEIDAYC1 As String          ' 計上日付
    SEEDC1 As String            ' ｼｰﾄﾞ
    PUBTKBC1 As String          ' BOT状況区分
    JDGECC1 As String           ' 判定ｺｰﾄﾞ
    PWTIMEC1 As String          ' ﾊﾟﾜｰ時間
    ADDOPPC1 As String          ' 追加ﾄﾞｰﾌﾟ位置
    ADDOPCC1 As String          ' 追加ﾄﾞｰﾊﾟﾝﾄ種類
    ADDOPVC1 As String          ' 追加ﾄﾞｰﾌﾟ量
    ADDOPNC1 As String          ' 追加ﾄﾞｰﾌﾟ名
    TSTAFFC1 As String          ' 登録社員ID
    TDAYC1 As String            ' 登録日付
    KSTAFFC1 As String          ' 更新社員ID
    KDAYC1 As String            ' 更新日付
    SUMITBC1 As String          ' SUMMIT送信ﾌﾗｸﾞ
    SNDKC1 As String            ' 送信ﾌﾗｸﾞ
    SNDDAYC1 As String          ' 送信日付
    SUIFLG As String            ' 推定FLG　2003/10/27 tuku
    PUREVNUMC1 As String        ' 狙い品番製造番号改訂番号　04/09/28 ooba
    PUFACTORYC1 As String       ' 狙い品番工場　04/09/28 ooba
    PUOPEC1 As String           ' 狙い品番操業条件　04/09/28 ooba
    LENPUFRC1 As String         ' 引上ﾌﾘｰ長さ　05/07/19 ooba
    WGHTPUFRC1 As String        ' 引上ﾌﾘｰ重量　05/07/19 ooba
    WGHTCUTRHC1 As String       ' 切断後良品重量　05/07/19 ooba
'C－OSF3判定機能追加 2007/05/11 M.Kaga START  ---
    JDGEIDC1    As String       'C-OSF3判定ID
'C－OSF3判定機能追加 2007/05/11 M.Kaga END    ---
End Type

'''　送信フラグ、SUMMIT送信フラグ
'Public Const SNDKC_NOTSEND = 0     '' 未送信
'Public Const SNDKC_SENDING = 1     '' 送信中
'Public Const SNDKC_SENDED = 2      '' 送信済み
'Public Const SNDKC_WAITING = 3     '' 送信待ち


'●SELECT●

'概要      :テーブル「XSDC1」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDC1        ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDC1(records() As typ_XSDC1, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long

    ''SQLを組み立てる
    sqlBase = "Select * From XSDC1"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC1 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("XTALC1")) = False Then .XTALC1 = rs.Fields("XTALC1")             ' 結晶番号
            If IsNull(rs.Fields("KNKTC1")) = False Then .KNKTC1 = rs.Fields("KNKTC1")             ' 管理工程ｺｰﾄﾞ
            If IsNull(rs.Fields("WKKTC1")) = False Then .WKKTC1 = rs.Fields("WKKTC1")             ' 工程ｺｰﾄﾞ
            If IsNull(rs.Fields("LENTOC1")) = False Then .LENTOC1 = rs.Fields("LENTOC1")           ' 長さ（TOP）
            If IsNull(rs.Fields("LENTKC1")) = False Then .LENTKC1 = rs.Fields("LENTKC1")           ' 長さ（直胴）
            If IsNull(rs.Fields("LENTAC1")) = False Then .LENTAC1 = rs.Fields("LENTAC1")           ' 長さ（TAIL）
            If IsNull(rs.Fields("PUFRELC1")) = False Then .PUFRELC1 = rs.Fields("PUFRELC1")         ' ﾌﾘｰ長さ
            If IsNull(rs.Fields("DIA1C1")) = False Then .DIA1C1 = rs.Fields("DIA1C1")             ' 直胴直径1
            If IsNull(rs.Fields("DIA2C1")) = False Then .DIA2C1 = rs.Fields("DIA2C1")             ' 直胴直径2
            If IsNull(rs.Fields("DIA3C1")) = False Then .DIA3C1 = rs.Fields("DIA3C1")             ' 直胴直径3
            If IsNull(rs.Fields("WGHTTOC1")) = False Then .WGHTTOC1 = rs.Fields("WGHTTOC1")         ' 重量（TOP）
            If IsNull(rs.Fields("WGHTTKC1")) = False Then .WGHTTKC1 = rs.Fields("WGHTTKC1")         ' 重量（直胴）
            If IsNull(rs.Fields("WGHTTAC1")) = False Then .WGHTTAC1 = rs.Fields("WGHTTAC1")         ' 重量（TAIL）
            If IsNull(rs.Fields("WGHTFRC1")) = False Then .WGHTFRC1 = rs.Fields("WGHTFRC1")         ' 重量（ﾌﾘｰ長さ）
            If IsNull(rs.Fields("PUTCUTWC1")) = False Then .PUTCUTWC1 = rs.Fields("PUTCUTWC1")       ' ﾄｯﾌﾟｶｯﾄ重量
            If IsNull(rs.Fields("PUWC1")) = False Then .PUWC1 = rs.Fields("PUWC1")               ' 引上げ重量
            If IsNull(rs.Fields("PUHINBC1")) = False Then .PUHINBC1 = rs.Fields("PUHINBC1")         ' 狙い品番
            If IsNull(rs.Fields("PUCHAGC1")) = False Then .PUCHAGC1 = rs.Fields("PUCHAGC1")         ' ﾁｬｰｼﾞ量
            If IsNull(rs.Fields("KAKOUBC1")) = False Then .KAKOUBC1 = rs.Fields("KAKOUBC1")         ' 加工区分
            If IsNull(rs.Fields("KEIDAYC1")) = False Then .KEIDAYC1 = rs.Fields("KEIDAYC1")         ' 計上日付
            If IsNull(rs.Fields("SEEDC1")) = False Then .SEEDC1 = rs.Fields("SEEDC1")             ' ｼｰﾄﾞ
            If IsNull(rs.Fields("PUBTKBC1")) = False Then .PUBTKBC1 = rs.Fields("PUBTKBC1")         ' BOT状況区分
            If IsNull(rs.Fields("JDGECC1")) = False Then .JDGECC1 = rs.Fields("JDGECC1")           ' 判定ｺｰﾄﾞ
            If IsNull(rs.Fields("PWTIMEC1")) = False Then .PWTIMEC1 = rs.Fields("PWTIMEC1")         ' ﾊﾟﾜｰ時間
            If IsNull(rs.Fields("ADDOPPC1")) = False Then .ADDOPPC1 = rs.Fields("ADDOPPC1")         ' 追加ﾄﾞｰﾌﾟ位置
            If IsNull(rs.Fields("ADDOPCC1")) = False Then .ADDOPCC1 = rs.Fields("ADDOPCC1")         ' 追加ﾄﾞｰﾊﾟﾝﾄ種類
            If IsNull(rs.Fields("ADDOPVC1")) = False Then .ADDOPVC1 = rs.Fields("ADDOPVC1")         ' 追加ﾄﾞｰﾌﾟ量
            If IsNull(rs.Fields("ADDOPNC1")) = False Then .ADDOPNC1 = rs.Fields("ADDOPNC1")         ' 追加ﾄﾞｰﾌﾟ名
            If IsNull(rs.Fields("TSTAFFC1")) = False Then .TSTAFFC1 = rs.Fields("TSTAFFC1")         ' 登録社員ID
            If IsNull(rs.Fields("TDAYC1")) = False Then .TDAYC1 = rs.Fields("TDAYC1")             ' 登録日付
            If IsNull(rs.Fields("KSTAFFC1")) = False Then .KSTAFFC1 = rs.Fields("KSTAFFC1")         ' 更新社員ID
            If IsNull(rs.Fields("KDAYC1")) = False Then .KDAYC1 = rs.Fields("KDAYC1")             ' 更新日付
            If IsNull(rs.Fields("SUMITBC1")) = False Then .SUMITBC1 = rs.Fields("SUMITBC1")         ' SUMMIT送信ﾌﾗｸﾞ
            If IsNull(rs.Fields("SNDKC1")) = False Then .SNDKC1 = rs.Fields("SNDKC1")             ' 送信ﾌﾗｸﾞ
            If IsNull(rs.Fields("SNDDAYC1")) = False Then .SNDDAYC1 = rs.Fields("SNDDAYC1")         ' 送信日付
            If IsNull(rs.Fields("SUIFLG")) = False Then .SUIFLG = rs.Fields("SUIFLG")         ' 推定FLG
            If IsNull(rs.Fields("PUREVNUMC1")) = False Then .PUREVNUMC1 = rs.Fields("PUREVNUMC1")       '狙い品番製造番号改訂番号　04/09/28 ooba
            If IsNull(rs.Fields("PUFACTORYC1")) = False Then .PUFACTORYC1 = rs.Fields("PUFACTORYC1")    '狙い品番工場　04/09/28 ooba
            If IsNull(rs.Fields("PUOPEC1")) = False Then .PUOPEC1 = rs.Fields("PUOPEC1")                '狙い品番操業条件　04/09/28 ooba
            If IsNull(rs.Fields("LENPUFRC1")) = False Then .LENPUFRC1 = rs.Fields("LENPUFRC1")      '引上ﾌﾘｰ長さ　05/07/19 ooba
            If IsNull(rs.Fields("WGHTPUFRC1")) = False Then .WGHTPUFRC1 = rs.Fields("WGHTPUFRC1")   '引上ﾌﾘｰ重量　05/07/19 ooba
            If IsNull(rs.Fields("WGHTCUTRHC1")) = False Then .WGHTCUTRHC1 = rs.Fields("WGHTCUTRHC1") '切断後良品重量　05/07/19 ooba
''09/02/16 FAE)akiyama start
            If IsNull(rs.Fields("SUICHARGE")) = False Then .SUICHARGE = rs.Fields("SUICHARGE") '推定チャージ量
''09/02/16 FAE)akiyama end
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC1 = FUNCTION_RETURN_SUCCESS
End Function


'●UPDATE●

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDC1」を更新する ptrn1
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :records       ,O   ,typ_XSDC1_Update ,更新レコード
'          :sqlWhere      ,I   ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I   ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O   ,FUNCTION_RETURN  ,抽出の成否
'説明      :

Public Function UpdateXSDC1(records As typ_XSDC1_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err

    Dim sql As String       'SQL全体
'    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)

    'エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDC1_SQL.bas -- Function UpdateXSDC1"
    
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> .EditをSQL(UPDATE)文に変更　2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQLを組み立てる
        sql = "UPDATE XSDC1 SET" & vbLf
        
        ''更新日付
        sql = sql & " KDAYC1 = " & nowtime_sql & vbLf
        
        ''結晶番号
        If .XTALC1 <> "" And Left(.XTALC1, 1) <> vbNullChar Then
            sql = sql & ",XTALC1 = '" & .XTALC1 & "'" & vbLf
        End If

        ''管理工程ｺｰﾄﾞ
        If .KNKTC1 <> "" And Left(.KNKTC1, 1) <> vbNullChar Then
            sql = sql & ",KNKTC1 = '" & .KNKTC1 & "'" & vbLf
        End If
        
        ''工程ｺｰﾄﾞ
        If .WKKTC1 <> "" And Left(.WKKTC1, 1) <> vbNullChar Then
            sql = sql & ",WKKTC1 = '" & .WKKTC1 & "'" & vbLf
        End If
        
        ''長さ（TOP）
        If .LENTOC1 <> "" Then
            sql = sql & ",LENTOC1 = '" & CStr(CInt(.LENTOC1)) & "'" & vbLf
        End If
        
        ''長さ（直胴）
        If .LENTKC1 <> "" Then
            sql = sql & ",LENTKC1 = '" & CStr(CInt(.LENTKC1)) & "'" & vbLf
        End If
        
        ''長さ（TAIL）
        If .LENTAC1 <> "" Then
            sql = sql & ",LENTAC1 = '" & CStr(CInt(.LENTAC1)) & "'" & vbLf
        End If
        
        ''ﾌﾘｰ長さ
        If .PUFRELC1 <> "" Then
            sql = sql & ",PUFRELC1 = '" & CStr(CInt(.PUFRELC1)) & "'" & vbLf
        End If
        
        ''直胴直径1
        If .DIA1C1 <> "" Then
            sql = sql & ",DIA1C1 = '" & .DIA1C1 & "'" & vbLf
        End If
        
        ''直胴直径2
        If .DIA2C1 <> "" Then
            sql = sql & ",DIA2C1 = '" & .DIA2C1 & "'" & vbLf
        End If
        
        ''直胴直径3
        If .DIA3C1 <> "" Then
            sql = sql & ",DIA3C1 = '" & .DIA3C1 & "'" & vbLf
        End If
        
        ''重量（TOP）
        If .WGHTTOC1 <> "" Then
            sql = sql & ",WGHTTOC1 = '" & CStr(CLng(.WGHTTOC1)) & "'" & vbLf
        End If
        
        ''重量（直胴）
        If .WGHTTKC1 <> "" Then
            sql = sql & ",WGHTTKC1 = '" & CStr(CLng(.WGHTTKC1)) & "'" & vbLf
        End If
        
        ''重量（TAIL）
        If .WGHTTAC1 <> "" Then
            sql = sql & ",WGHTTAC1 = '" & CStr(CLng(.WGHTTAC1)) & "'" & vbLf
        End If
        
        ''重量（ﾌﾘｰ長さ）
        If .WGHTFRC1 <> "" Then
            sql = sql & ",WGHTFRC1 = '" & CStr(CLng(.WGHTFRC1)) & "'" & vbLf
        End If
        
        ''ﾄｯﾌﾟｶｯﾄ重量
        If .PUTCUTWC1 <> "" Then
            sql = sql & ",PUTCUTWC1 = '" & CStr(CLng(.PUTCUTWC1)) & "'" & vbLf
        End If
        
        ''引上げ重量
        If .PUWC1 <> "" Then
            sql = sql & ",PUWC1 = '" & CStr(CLng(.PUWC1)) & "'" & vbLf
        End If
        
        ''狙い品番
        If .PUHINBC1 <> "" And Left(.PUHINBC1, 1) <> vbNullChar Then
            sql = sql & ",PUHINBC1 = '" & .PUHINBC1 & "'" & vbLf
        End If
        
        ''ﾁｬｰｼﾞ量
        If .PUCHAGC1 <> "" Then
            sql = sql & ",PUCHAGC1 = '" & CStr(CLng(.PUCHAGC1)) & "'" & vbLf
        End If
        
        ''加工区分
        If .KAKOUBC1 <> "" And Left(.KAKOUBC1, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBC1 = '" & .KAKOUBC1 & "'" & vbLf
        End If
        
        ''計上日付
        If .KEIDAYC1 <> "" Then
            sql = sql & ",KEIDAYC1 = TO_DATE('" & Format$(CDate(.KEIDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''ｼｰﾄﾞ
        If .SEEDC1 <> "" And Left(.SEEDC1, 1) <> vbNullChar Then
            sql = sql & ",SEEDC1 = '" & .SEEDC1 & "'" & vbLf
        End If
        
        ''BOT状況区分
        If .PUBTKBC1 <> "" And Left(.PUBTKBC1, 1) <> vbNullChar Then
            sql = sql & ",PUBTKBC1 = '" & .PUBTKBC1 & "'" & vbLf
        End If
        
        ''判定ｺｰﾄﾞ
        If .JDGECC1 <> "" And Left(.JDGECC1, 1) <> vbNullChar Then
            sql = sql & ",JDGECC1 = '" & .JDGECC1 & "'" & vbLf
        End If
        
        ''ﾊﾟﾜｰ時間
        If .PWTIMEC1 <> "" Then
            sql = sql & ",PWTIMEC1 = '" & .PWTIMEC1 & "'" & vbLf
        End If
        
        ''追加ﾄﾞｰﾌﾟ位置
        If .ADDOPPC1 <> "" Then
            sql = sql & ",ADDOPPC1 = '" & CStr(CInt(.ADDOPPC1)) & "'" & vbLf
        End If
        
        ''追加ﾄﾞｰﾊﾟﾝﾄ種類
        If .ADDOPCC1 <> "" And Left(.ADDOPCC1, 1) <> vbNullChar Then
            sql = sql & ",ADDOPCC1 = '" & .ADDOPCC1 & "'" & vbLf
        End If
        
        ''追加ﾄﾞｰﾌﾟ量
        If .ADDOPVC1 <> "" Then
            sql = sql & ",ADDOPVC1 = '" & CStr(CLng(.ADDOPVC1)) & "'" & vbLf
        End If
        
        ''追加ﾄﾞｰﾌﾟ名
        If .ADDOPNC1 <> "" And Left(.ADDOPNC1, 1) <> vbNullChar Then
            sql = sql & ",ADDOPNC1 = '" & .ADDOPNC1 & "'" & vbLf
        End If
        
        ''登録社員ID
        If .TSTAFFC1 <> "" And Left(.TSTAFFC1, 1) <> vbNullChar Then
            sql = sql & ",TSTAFFC1 = '" & .TSTAFFC1 & "'" & vbLf
        End If
        
        ''登録日付
        If .TDAYC1 <> "" Then
            sql = sql & ",TDAYC1 = TO_DATE('" & Format$(CDate(.TDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''更新社員ID
        If .KSTAFFC1 <> "" And Left(.KSTAFFC1, 1) <> vbNullChar Then
            sql = sql & ",KSTAFFC1 = '" & .KSTAFFC1 & "'" & vbLf
        End If
        
        ''SUMMIT送信ﾌﾗｸﾞ
        If .SUMITBC1 <> "" And Left(.SUMITBC1, 1) <> vbNullChar Then
            sql = sql & ",SUMITBC1 = '" & .SUMITBC1 & "'" & vbLf
        End If
        
        ''送信ﾌﾗｸﾞ
        If .SNDKC1 <> "" And Left(.SNDKC1, 1) <> vbNullChar Then
            sql = sql & ",SNDKC1 = '" & .SNDKC1 & "'" & vbLf
        End If
        
        ''送信日付
        If .SNDDAYC1 <> "" Then
            sql = sql & ",SNDDAYC1 = TO_DATE('" & Format$(CDate(.SNDDAYC1), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''推定FLG
        If .SUIFLG <> "" Then
            sql = sql & ",SUIFLG = '" & .SUIFLG & "'" & vbLf
        End If
        
        ''狙い品番製造番号改訂番号
        If .PUREVNUMC1 <> "" Then
            sql = sql & ",PUREVNUMC1 = '" & .PUREVNUMC1 & "'" & vbLf
        End If
        
        ''狙い品番工場
        If .PUFACTORYC1 <> "" Then
            sql = sql & ",PUFACTORYC1 = '" & .PUFACTORYC1 & "'" & vbLf
        End If
        
        ''狙い品番操業条件
        If .PUOPEC1 <> "" Then
            sql = sql & ",PUOPEC1 = '" & .PUOPEC1 & "'" & vbLf
        End If
        
        ''引上ﾌﾘｰ長さ
        If .LENPUFRC1 <> "" Then
            sql = sql & ",LENPUFRC1 = '" & CStr(CInt(.LENPUFRC1)) & "'" & vbLf
        End If
        
        ''引上ﾌﾘｰ重量
        If .WGHTPUFRC1 <> "" Then
            sql = sql & ",WGHTPUFRC1 = '" & CStr(CLng(.WGHTPUFRC1)) & "'" & vbLf
        End If
        
        ''切断後良品重量
        If .WGHTCUTRHC1 <> "" Then
            sql = sql & ",WGHTCUTRHC1 = '" & CStr(CLng(.WGHTCUTRHC1)) & "'" & vbLf
        End If
        
        ''C-OSF3判定ID
        If .JDGEIDC1 <> "" Then
            sql = sql & ",JDGEIDC1 = '" & .JDGEIDC1 & "'" & vbLf
        End If

        sql = sql & " " & sqlWhere & vbLf
    
        'SQLを実行
        recCnt = OraDB.ExecuteSQL(sql)
        
        '返り値が1以外はエラー
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDC1 = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDC1 = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
        
    End With
'<<<<< .EditをSQL(UPDATE)文に変更　2009/06/16 SETsw kubota ------------------

    UpdateXSDC1 = FUNCTION_RETURN_SUCCESS

proc_exit:
    '終了
    gErr.Pop
    Exit Function

proc_err:
    'エラーハンドラ
    UpdateXSDC1 = FUNCTION_RETURN_FAILURE
    Debug.Print "==== ERROR SQL ===="
    Debug.Print sql
    gErr.HandleError
    Resume proc_exit
End Function

'概要      :引上重量を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sXtal      ,I  ,String           ,結晶番号
'      　　:戻り値       ,O  ,Long           　,処理回数
'          :作成者　04/09/30 ooba
Public Function GetPutWeight(p_sXtal As String) As Long

    Dim sSQL As String
    Dim rs As OraDynaset
    
    If Left(p_sXtal, 1) = vbNullChar Then
        GetPutWeight = 0
        Exit Function
    End If
    
    sSQL = "SELECT PUWC1 FROM XSDC1 WHERE XTALC1 = '" & p_sXtal & "'"
    
    Set rs = OraDB.DBCreateDynaset(sSQL, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetPutWeight = 0
    Else
        If IsNull(rs.Fields("PUWC1")) = False Then GetPutWeight = CLng(rs.Fields("PUWC1")) Else GetPutWeight = 0
    End If
    
End Function
