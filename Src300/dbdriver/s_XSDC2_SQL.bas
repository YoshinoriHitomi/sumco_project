Attribute VB_Name = "s_XSDC2_SQL"
'分割結晶(ﾌﾞﾛｯｸ) (XSDC2) ｱｸｾｽ関数

'***テーブル「XSDC2」へのデータアクセス関数***
'＊注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

'●分割結晶(ﾌﾞﾛｯｸ)
Public Type typ_XSDC2
    CRYNUMC2 As String * 12      ' ﾌﾞﾛｯｸID･結晶番号
    KCNTC2 As Integer            ' 工程連番
    XTALC2 As String * 12        ' 結晶番号
    INPOSC2 As Integer           ' 結晶内開始位置
    NEKKNTC2 As String * 5       ' 最終通過管理工程
    NEWKNTC2 As String * 5       ' 最終通過工程
    NEWKKBC2 As String * 2       ' 最終通過作業区分
    NEMACOC2 As Integer          ' 最終通過処理回数
    GNKKNTC2 As String * 5       ' 現在管理工程
    GNWKNTC2 As String * 5       ' 現在工程
    GNWKKBC2 As String * 2       ' 現在作業区分
    GNMACOC2 As Integer          ' 現在処理回数
    GNDAYC2 As Date              ' 現在処理日付
    GNLC2 As Integer             ' 現在長さ
    GNWC2 As Long                ' 現在重量
    GNMC2 As Integer             ' 現在枚数
    SUMITLC2 As Integer          ' SUMMIT長さ
    SUMITWC2 As Long             ' SUMMIT重量
    SUMITMC2 As Integer          ' SUMMIT枚数
    CHGC2 As Long                ' ﾁｬｰｼﾞ量
    KAKOUBC2 As String * 1       ' 加工区分
    KEIDAYC2 As Date             ' 計上日付
    GNTKUBC2 As String * 3       ' 棚区分
    GNTNOC2 As String * 4        ' 棚番号
    XTWORKC2 As String * 2       ' 製造工場
    WFWORKC2 As String * 2       ' ｳｪｰﾊ製造
    LSTATBC2 As String * 1       ' 最終状態区分
    RSTATBC2 As String * 1       ' 流動状態区分
    LUFRCC2 As String * 3        ' 格上ｺｰﾄﾞ
    LUFRBC2 As String * 1        ' 格上区分
    LDFRCC2 As String * 3        ' 格下ｺｰﾄﾞ
    LDFRBC2 As String * 1        ' 格下区分
    HOLDCC2 As String * 3        ' ﾎｰﾙﾄﾞｺｰﾄﾞ
    HOLDBC2 As String * 1        ' ホールド区分
    EXKUBC2 As String * 1        ' 例外区分
    HENPKC2 As String * 1        ' 返品区分
    LIVKC2 As String * 1         ' 生死区分
    KANKC2 As String * 1         ' 完了区分
    NFC2 As String * 1           ' 入庫区分
    SAKJC2 As String * 1         ' 削除区分
    TDAYC2 As Date               ' 登録日付
    KDAYC2 As Date               ' 更新日付
    SUMITBC2 As String * 1       ' SUMMIT送信フラグ
    SNDKC2 As String * 1         ' 送信フラグ
    SNDDAYC2 As Date             ' 送信日付
' 2003.06.11 Y.KATABAMI tuika
    PRIORITYC2 As String * 1     ' 優先度
    CUTCNTC2 As String * 1       ' 新規／再切区分
    '2005/07
    HOLDKTC2 As String * 5
    RPCRYNUMC2 As String * 12    ' 親ﾌﾞﾛｯｸID　05/09/20 ooba
    BDCAUSC2 As String * 3       ' 不良理由　05/12/01 ooba
''↓追加 START SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
    REALLC2 As Integer           ' 実長さ
    REALWC2 As Long              ' 実重量
''↑追加 END   SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
    KBLKFLGC2 As String * 1      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　06/10/31 ooba
    KIKBNC2 As String            ' 期判別区分   2006/11/10 SETsw kubota
    PLANTCATC2 As String         ' 向先 2007/08/22 SPK Tsutsumi Add
    STCIDC2 As String            ' STCﾌﾞﾛｯｸID　08/06/16 ooba
End Type

'更新用
Public Type typ_XSDC2_Update
    CRYNUMC2 As String        ' ﾌﾞﾛｯｸID･結晶番号
    KCNTC2 As String          ' 工程連番
    XTALC2 As String          ' 結晶番号
    INPOSC2 As String         ' 結晶内開始位置
    NEKKNTC2 As String        ' 最終通過管理工程
    NEWKNTC2 As String        ' 最終通過工程
    NEWKKBC2 As String        ' 最終通過作業区分
    NEMACOC2 As String        ' 最終通過処理回数
    GNKKNTC2 As String        ' 現在管理工程
    GNWKNTC2 As String        ' 現在工程
    GNWKKBC2 As String        ' 現在作業区分
    GNMACOC2 As String        ' 現在処理回数
    GNDAYC2 As String         ' 現在処理日付
    GNLC2 As String           ' 現在長さ
    GNWC2 As String           ' 現在重量
    GNMC2 As String           ' 現在枚数
    SUMITLC2 As String        ' SUMMIT長さ
    SUMITWC2 As String        ' SUMMIT重量
    SUMITMC2 As String        ' SUMMIT枚数
    CHGC2 As String           ' ﾁｬｰｼﾞ量
    KAKOUBC2 As String        ' 加工区分
    KEIDAYC2 As String        ' 計上日付
    GNTKUBC2 As String        ' 棚区分
    GNTNOC2 As String         ' 棚番号
    XTWORKC2 As String        ' 製造工場
    WFWORKC2 As String        ' ｳｪｰﾊ製造
    LSTATBC2 As String        ' 最終状態区分
    RSTATBC2 As String        ' 流動状態区分
    LUFRCC2 As String         ' 格上ｺｰﾄﾞ
    LUFRBC2 As String         ' 格上区分
    LDFRCC2 As String         ' 格下ｺｰﾄﾞ
    LDFRBC2 As String         ' 格下区分
    HOLDCC2 As String         ' ﾎｰﾙﾄﾞｺｰﾄﾞ
    HOLDBC2 As String         ' ホールド区分
    EXKUBC2 As String         ' 例外区分
    HENPKC2 As String         ' 返品区分
    LIVKC2 As String          ' 生死区分
    KANKC2 As String          ' 完了区分
    NFC2 As String            ' 入庫区分
    SAKJC2 As String          ' 削除区分
    TDAYC2 As String          ' 登録日付
    KDAYC2 As String          ' 更新日付
    SUMITBC2 As String        ' SUMMIT送信フラグ
    SNDKC2 As String          ' 送信フラグ
    SNDDAYC2 As String        ' 送信日付
' 2003.06.11 Y.KATABAMI tuika
    PRIORITYC2 As String * 1     ' 優先度
    CUTCNTC2 As String * 1       ' 新規／再切区分
    '2005/07
    HOLDKTC2 As String * 5
    RPCRYNUMC2 As String * 12    ' 親ﾌﾞﾛｯｸID　05/09/20 ooba
    BDCAUSC2 As String * 3       ' 不良理由　05/12/01 ooba
''↓追加 START SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
    REALLC2 As String           ' 実長さ
    REALWC2 As String           ' 実重量
''↑追加 END   SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
    KBLKFLGC2 As String * 1      ' 関連ﾌﾞﾛｯｸﾌﾗｸﾞ　06/10/31 ooba
    KIKBNC2   As String          ' 期判別区分   2006/11/10 SETsw kubota
    PLANTCATC2 As String         ' 向先 2007/08/22 SPK Tsutsumi Add
    STCIDC2 As String            ' STCﾌﾞﾛｯｸID　08/06/16 ooba
End Type

'●SELECT●

'概要      :テーブル「XSDC2」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDC2     ,抽出レコード
'          :sqlWhere      ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDC2(records() As typ_XSDC2, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim sql As String       'SQL全体
    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
    Dim rs As OraDynaset    'RecordSet
    Dim recCnt As Long      'レコード数
    Dim i As Long


    ''SQLを組み立てる
    sqlBase = "Select * From XSDC2"
    sql = sqlBase
    If (sqlWhere <> vbNullString) Or (sqlOrder <> vbNullString) Then
        sql = sql & " " & sqlWhere & " " & sqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDC2 = FUNCTION_RETURN_FAILURE
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
            If IsNull(rs.Fields("CRYNUMC2")) = False Then .CRYNUMC2 = rs.Fields("CRYNUMC2")
            If IsNull(rs.Fields("KCNTC2")) = False Then .KCNTC2 = rs.Fields("KCNTC2")
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
            If IsNull(rs.Fields("GNLC2")) = False Then .GNLC2 = rs.Fields("GNLC2")
            If IsNull(rs.Fields("GNWC2")) = False Then .GNWC2 = rs.Fields("GNWC2")
            If IsNull(rs.Fields("GNMC2")) = False Then .GNMC2 = rs.Fields("GNMC2")
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
            '2003.06.11 Y.Katabami tuika
            If IsNull(rs.Fields("PRIORITYC2")) = False Then .PRIORITYC2 = rs.Fields("PRIORITYC2")
            If IsNull(rs.Fields("CUTCNTC2")) = False Then .CUTCNTC2 = rs.Fields("CUTCNTC2")
            '2005/07
            If IsNull(rs.Fields("HOLDKTC2")) = False Then .HOLDKTC2 = rs.Fields("HOLDKTC2")
            If IsNull(rs.Fields("RPCRYNUMC2")) = False Then .RPCRYNUMC2 = rs.Fields("RPCRYNUMC2")   '05/09/20 ooba
            If IsNull(rs.Fields("BDCAUSC2")) = False Then .BDCAUSC2 = rs.Fields("BDCAUSC2")         '05/12/01 ooba
            ''↓追加 START SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
            If IsNull(rs.Fields("REALLC2")) = False Then .REALLC2 = rs.Fields("REALLC2")        ''実長さ
            If IsNull(rs.Fields("REALWC2")) = False Then .REALWC2 = rs.Fields("REALWC2")        ''実重量
            ''↑追加 END   SPT用実績作成方法変更 2006/06/05 SMP-OKAMOTO
            If IsNull(rs.Fields("KBLKFLGC2")) = False Then .KBLKFLGC2 = rs.Fields("KBLKFLGC2")      '06/10/31 ooba
            If IsNull(rs.Fields("KIKBNC2")) = False Then .KIKBNC2 = rs.Fields("KIKBNC2")            '06/11/10 SETsw kubota
            If IsNull(rs.Fields("PLANTCATC2")) = False Then .PLANTCATC2 = rs.Fields("PLANTCATC2")            '07/08/22 SPK Tsutsumi Add
            If IsNull(rs.Fields("STCIDC2")) = False Then .STCIDC2 = rs.Fields("STCIDC2")            '08/06/16 ooba
        End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDC2 = FUNCTION_RETURN_SUCCESS
End Function

'●UPDATE●

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDC2」を更新する ptrn1
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :records()     ,O   ,typ_XSDC2     ,更新レコード
'          :sqlWhere      ,I   ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :sqlOrder      ,I   ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O   ,FUNCTION_RETURN  ,抽出の成否
'説明      :

Public Function UpdateXSDC2(records As typ_XSDC2_Update, _
                                  Optional sqlWhere$ = vbNullString, _
                                  Optional sqlOrder$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDC2_SQL.bas -- Function UpdateXSDC2"

    Dim sql As String       'SQL全体
'    Dim sqlBase As String   'SQL基本部(WHERE節の前まで)
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset
    Dim recCnt As Long      'レコード数
'    Dim i As Long
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)
    
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku
    
'>>>>> .EditをSQL(UPDATE)文に変更　2009/06/22 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"

    With records
        
        ''SQLを組み立てる
        sql = "UPDATE XSDC2 SET" & vbLf
        
        ''更新日付
        sql = sql & " KDAYC2 = " & nowtime_sql & vbLf
        
        ''ブロックID・結晶番号
        If .CRYNUMC2 <> "" And Left(.CRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",CRYNUMC2 = '" & .CRYNUMC2 & "'" & vbLf
        End If
        
        ''工程連番
        If .KCNTC2 <> "" Then
            sql = sql & ",KCNTC2 = '" & CStr(CInt(.KCNTC2)) & "'" & vbLf
        End If
        
        ''結晶番号
        If .XTALC2 <> "" And Left(.XTALC2, 1) <> vbNullChar Then
            sql = sql & ",XTALC2 = '" & .XTALC2 & "'" & vbLf
        End If
        
        ''結晶内開始位置
        If .INPOSC2 <> "" Then
            sql = sql & ",INPOSC2 = '" & CStr(CInt(.INPOSC2)) & "'" & vbLf
        End If
        
        ''最終通過管理工程
        If .NEKKNTC2 <> "" And Left(.NEKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",NEKKNTC2 = '" & .NEKKNTC2 & "'" & vbLf
        End If
        
        ''最終通過工程
        If .NEWKNTC2 <> "" And Left(.NEWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",NEWKNTC2 = '" & .NEWKNTC2 & "'" & vbLf
        End If
        
        ''最終通過作業区分
        If .NEWKKBC2 <> "" And Left(.NEWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",NEWKKBC2 = '" & .NEWKKBC2 & "'" & vbLf
        End If
        
        ''最終通過処理回数
        If .NEMACOC2 <> "" Then
            sql = sql & ",NEMACOC2 = '" & CStr(CInt(.NEMACOC2)) & "'" & vbLf
        End If
        
        ''現在管理工程
        If .GNKKNTC2 <> "" And Left(.GNKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",GNKKNTC2 = '" & .GNKKNTC2 & "'" & vbLf
        End If
        
        ''現在工程
        If .GNWKNTC2 <> "" And Left(.GNWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",GNWKNTC2 = '" & .GNWKNTC2 & "'" & vbLf
        End If
        
        ''現在作業区分
        If .GNWKKBC2 <> "" And Left(.GNWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",GNWKKBC2 = '" & .GNWKKBC2 & "'" & vbLf
        End If

        ''現在処理回数
        If .GNMACOC2 <> "" Then
            sql = sql & ",GNMACOC2 = '" & CStr(CInt(.GNMACOC2)) & "'" & vbLf
        End If

        ''現在処理日付
        If .GNDAYC2 <> "" Then
            sql = sql & ",GNDAYC2 = TO_DATE('" & Format$(CDate(.GNDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''現在長さ
        If .GNLC2 <> "" Then
            sql = sql & ",GNLC2 = '" & CStr(CInt(.GNLC2)) & "'" & vbLf
        End If
        
        ''現在重量
        If .GNWC2 <> "" Then
            sql = sql & ",GNWC2 = '" & CStr(CLng(.GNWC2)) & "'" & vbLf
        End If
        
        ''現在枚数
        If .GNMC2 <> "" Then
            sql = sql & ",GNMC2 = '" & CStr(CInt(.GNMC2)) & "'" & vbLf
        End If
        
        ''SUMIT長さ
        If .SUMITLC2 <> "" Then
            sql = sql & ",SUMITLC2 = '" & CStr(CInt(.SUMITLC2)) & "'" & vbLf
        End If
        
        ''SUMIT重量
        If .SUMITWC2 <> "" Then
            sql = sql & ",SUMITWC2 = '" & CStr(CLng(.SUMITWC2)) & "'" & vbLf
        End If
        
        ''SUMIT枚数
        If .SUMITMC2 <> "" Then
            sql = sql & ",SUMITMC2 = '" & CStr(CInt(.SUMITMC2)) & "'" & vbLf
        End If
        
        ''チャージ量
        If .CHGC2 <> "" Then
            sql = sql & ",CHGC2 = '" & CStr(CLng(.CHGC2)) & "'" & vbLf
        End If
        
        ''加工区分
        If .KAKOUBC2 <> "" And Left(.KAKOUBC2, 1) <> vbNullChar Then
            sql = sql & ",KAKOUBC2 = '" & .KAKOUBC2 & "'" & vbLf
        End If
        
        ''計上日付
        If .KEIDAYC2 <> "" Then
            sql = sql & ",KEIDAYC2 = TO_DATE('" & Format$(CDate(.KEIDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''棚区分
        If .GNTKUBC2 <> "" And Left(.GNTKUBC2, 1) <> vbNullChar Then
            sql = sql & ",GNTKUBC2 = '" & .GNTKUBC2 & "'" & vbLf
        End If
        
        ''棚番号
        If .GNTNOC2 <> "" And Left(.GNTNOC2, 1) <> vbNullChar Then
            sql = sql & ",GNTNOC2 = '" & .GNTNOC2 & "'" & vbLf
        End If
        
        ''製造工場
        If .XTWORKC2 <> "" And Left(.XTWORKC2, 1) <> vbNullChar Then
            sql = sql & ",XTWORKC2 = '" & .XTWORKC2 & "'" & vbLf
        End If
        
        ''ウェーハ製造
        If .WFWORKC2 <> "" And Left(.WFWORKC2, 1) <> vbNullChar Then
            sql = sql & ",WFWORKC2 = '" & .WFWORKC2 & "'" & vbLf
        End If
        
        ''最終状態区分
        If .LSTATBC2 <> "" And Left(.LSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",LSTATBC2 = '" & .LSTATBC2 & "'" & vbLf
        End If
        
        ''流動状態区分
        If .RSTATBC2 <> "" And Left(.RSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",RSTATBC2 = '" & .RSTATBC2 & "'" & vbLf
        End If
        
        ''格上コード
        If .LUFRCC2 <> "" And Left(.LUFRCC2, 1) <> vbNullChar Then
            sql = sql & ",LUFRCC2 = '" & .LUFRCC2 & "'" & vbLf
        End If
        
        ''格上区分
        If .LUFRBC2 <> "" And Left(.LUFRBC2, 1) <> vbNullChar Then
            sql = sql & ",LUFRBC2 = '" & .LUFRBC2 & "'" & vbLf
        End If
        
        ''格下コード
        If .LDFRCC2 <> "" And Left(.LDFRCC2, 1) <> vbNullChar Then
            sql = sql & ",LDFRCC2 = '" & .LDFRCC2 & "'" & vbLf
        End If
        
        ''格下区分
        If .LDFRBC2 <> "" And Left(.LDFRBC2, 1) <> vbNullChar Then
            sql = sql & ",LDFRBC2 = '" & .LDFRBC2 & "'" & vbLf
        End If
        
        ''ホールドコード
        If .HOLDCC2 <> "" And Left(.HOLDCC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDCC2 = '" & .HOLDCC2 & "'" & vbLf
        End If
        
        ''ホールド区分
        If .HOLDBC2 <> "" And Left(.HOLDBC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDBC2 = '" & .HOLDBC2 & "'" & vbLf
        End If
        
        ''例外区分
        If .EXKUBC2 <> "" And Left(.EXKUBC2, 1) <> vbNullChar Then
            sql = sql & ",EXKUBC2 = '" & .EXKUBC2 & "'" & vbLf
        End If
        
        ''返品区分
        If .HENPKC2 <> "" And Left(.HENPKC2, 1) <> vbNullChar Then
            sql = sql & ",HENPKC2 = '" & .HENPKC2 & "'" & vbLf
        End If
        
        ''生死区分
        If .LIVKC2 <> "" And Left(.LIVKC2, 1) <> vbNullChar Then
            sql = sql & ",LIVKC2 = '" & .LIVKC2 & "'" & vbLf
        End If
        
        ''完了区分
        If .KANKC2 <> "" And Left(.KANKC2, 1) <> vbNullChar Then
            sql = sql & ",KANKC2 = '" & .KANKC2 & "'" & vbLf
        End If
        
        ''入庫区分
        If .NFC2 <> "" And Left(.NFC2, 1) <> vbNullChar Then
            sql = sql & ",NFC2 = '" & .NFC2 & "'" & vbLf
        End If
        
        ''削除区分
        If .SAKJC2 <> "" And Left(.SAKJC2, 1) <> vbNullChar Then
            sql = sql & ",SAKJC2 = '" & .SAKJC2 & "'" & vbLf
        End If
        
        ''登録日付
        If .TDAYC2 <> "" Then
            sql = sql & ",TDAYC2 = TO_DATE('" & Format$(CDate(.TDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''SUMIT送信フラグ
        If .SUMITBC2 <> "" And Left(.SUMITBC2, 1) <> vbNullChar Then
            sql = sql & ",SUMITBC2 = '" & .SUMITBC2 & "'" & vbLf
        End If
        
        ''送信フラグ
        If .SNDKC2 <> "" And Left(.SNDKC2, 1) <> vbNullChar Then
            sql = sql & ",SNDKC2 = '" & .SNDKC2 & "'" & vbLf
        End If
        
        ''送信日付
        If .SNDDAYC2 <> "" Then
            sql = sql & ",SNDDAYC2 = TO_DATE('" & Format$(CDate(.SNDDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        
        ''優先度
        If .PRIORITYC2 <> "" And Left(.PRIORITYC2, 1) <> vbNullChar Then
            sql = sql & ",PRIORITYC2 = '" & .PRIORITYC2 & "'" & vbLf
        End If
        
        ''切断処理区分
        If .CUTCNTC2 <> "" And Left(.CUTCNTC2, 1) <> vbNullChar Then
            sql = sql & ",CUTCNTC2 = '" & .CUTCNTC2 & "'" & vbLf
        End If
        
        ''ﾎｰﾙﾄﾞ工程
        If .HOLDKTC2 <> "" And Left(.HOLDKTC2, 1) <> vbNullChar Then
            sql = sql & ",HOLDKTC2 = '" & .HOLDKTC2 & "'" & vbLf
        End If
        
        ''親ブロックID
        If .RPCRYNUMC2 <> "" And Left(.RPCRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",RPCRYNUMC2 = '" & .RPCRYNUMC2 & "'" & vbLf
        End If
        
        ''不良理由
        If .BDCAUSC2 <> "" And Left(.BDCAUSC2, 1) <> vbNullChar Then
            sql = sql & ",BDCAUSC2 = '" & .BDCAUSC2 & "'" & vbLf
        End If
        
        ''実長さ
        If .REALLC2 <> "" Then
            sql = sql & ",REALLC2 = '" & CStr(CInt(.REALLC2)) & "'" & vbLf
        End If
        
        ''実重量
        If .REALWC2 <> "" Then
            sql = sql & ",REALWC2 = '" & CStr(CLng(.REALWC2)) & "'" & vbLf
        End If
        
        ''関連ブロックフラグ
        If .KBLKFLGC2 <> "" And Left(.KBLKFLGC2, 1) <> vbNullChar Then
            sql = sql & ",KBLKFLGC2 = '" & .KBLKFLGC2 & "'" & vbLf
        End If
        
        ''事業所区分
        If .PLANTCATC2 <> "" And Left(.PLANTCATC2, 2) <> vbNullChar Then
            sql = sql & ",PLANTCATC2 = '" & .PLANTCATC2 & "'" & vbLf
        End If
        
        ''STCブロックID
        If .STCIDC2 <> "" And Left(.STCIDC2, 1) <> vbNullChar Then
            sql = sql & ",STCIDC2 = '" & .STCIDC2 & "'" & vbLf
        End If
    
        sql = sql & " " & sqlWhere & vbLf
    
        'SQLを実行
        recCnt = OraDB.ExecuteSQL(sql)
        
        '返り値が1以外はエラー
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDC2 = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDC2 = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .EditをSQL(UPDATE)文に変更　2009/06/22 SETsw kubota ------------------

    UpdateXSDC2 = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print sql
    gErr.HandleError
    UpdateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'●INSERT●  NULLの場合、charならスペース、NumberならNULLを入れる

'概要      :テーブル「XSDC2」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDC2 　　  ,I  ,typ_XSDC2_Update   ,XSDC2更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDC2(pXSDC2 As typ_XSDC2_Update, sErrMsg As String) As FUNCTION_RETURN


    Dim sql As String
    Dim sDbName As String
'    Dim rs As OraDynaset    'RecordSet
'    Dim rs2 As OraDynaset
'    Dim recCnt As Long      'レコード数
    Dim nowtime As Date
    Dim nowtime_sql As String   'サーバ時間(SQL文)
    
    '' エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDC2_SQL.bas -- Function CreateXSDC2"
    sErrMsg = ""
    sDbName = "XSDC2"
     'nowtime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    nowtime = getSvrTime()    'サーバーの時間を取得するように変更 2003/6/4 tuku

'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/22 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With pXSDC2
        
        sql = "INSERT INTO XSDC2 ("
        sql = sql & " CRYNUMC2"         ' 1:ﾌﾞﾛｯｸID・結晶番号
        sql = sql & ",KCNTC2"           ' 2:工程連番取得
        sql = sql & ",XTALC2"           ' 3:結晶番号
        sql = sql & ",INPOSC2"          ' 4:結晶内開始位置
        sql = sql & ",NEKKNTC2"         ' 5:最終通過管理工程
        sql = sql & ",NEWKNTC2"         ' 6:最終通過工程
        sql = sql & ",NEWKKBC2"         ' 7:最終通過作業区分
        sql = sql & ",NEMACOC2"         ' 8:最終通過処理回数
        sql = sql & ",GNKKNTC2"         ' 9:現在管理工程
        sql = sql & ",GNWKNTC2"         '10:現在工程
        sql = sql & ",GNWKKBC2"         '11:現在作業区分
        sql = sql & ",GNMACOC2"         '12:現在処理回数
        sql = sql & ",GNDAYC2"          '13:現在処理日付(登録日時)
        sql = sql & ",GNLC2"            '14:現在長さ
        sql = sql & ",GNWC2"            '15:現在重量
        sql = sql & ",GNMC2"            '16:現在枚数
        sql = sql & ",SUMITLC2"         '17:SUMMIT長さ
        sql = sql & ",SUMITWC2"         '18:SUMMIT重量
        sql = sql & ",SUMITMC2"         '19:SUMMIT枚数
        sql = sql & ",CHGC2"            '20:ﾁｬｰｼﾞ量
        sql = sql & ",KAKOUBC2"         '21:加工区分
        If .KEIDAYC2 <> "" Then
            sql = sql & ",KEIDAYC2"         '22:計上日付
        End If
        sql = sql & ",GNTKUBC2"         '23:棚区分
        sql = sql & ",GNTNOC2"          '24:棚番号
        sql = sql & ",XTWORKC2"         '25:製造工場
        sql = sql & ",WFWORKC2"         '26:ｳｪｰﾊ製造
        sql = sql & ",LSTATBC2"         '27:最終状態区分
        sql = sql & ",RSTATBC2"         '28:流動状態区分
        sql = sql & ",LUFRCC2"          '29:格上ｺｰﾄﾞ
        sql = sql & ",LUFRBC2"          '30:格上区分
        sql = sql & ",LDFRCC2"          '31:格下ｺｰﾄﾞ
        sql = sql & ",LDFRBC2"          '32:格下区分
        sql = sql & ",HOLDCC2"          '33:ﾎｰﾙﾄﾞｺｰﾄﾞ
        sql = sql & ",HOLDBC2"          '34:ﾎｰﾙﾄﾞ区分
        sql = sql & ",EXKUBC2"          '35:例外区分
        sql = sql & ",HENPKC2"          '36:返品区分
        sql = sql & ",LIVKC2"           '37:生死区分
        sql = sql & ",KANKC2"           '38:完了区分
        sql = sql & ",NFC2"             '39:入庫区分
        sql = sql & ",SAKJC2"           '40:削除区分
        sql = sql & ",TDAYC2"           '41:登録日付
        sql = sql & ",KDAYC2"           '42:更新日付
        sql = sql & ",SUMITBC2"         '43:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",SNDKC2"           '44:送信ﾌﾗｸﾞ
        sql = sql & ",SNDDAYC2"         '45:送信日付
        sql = sql & ",PRIORITYC2"       '46:優先度
        sql = sql & ",CUTCNTC2"         '47:新規／再切区分
        sql = sql & ",HOLDKTC2"         '48:ﾎｰﾙﾄﾞ工程
        sql = sql & ",RPCRYNUMC2"       '49:親ﾌﾞﾛｯｸID
        sql = sql & ",BDCAUSC2"         '50:不良理由
        sql = sql & ",REALLC2"          '51:実長さ
        sql = sql & ",REALWC2"          '52:実重量
        sql = sql & ",KBLKFLGC2"        '53:関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        sql = sql & ",PLANTCATC2"       '54:向先
        sql = sql & ",STCIDC2"          '55:STCﾌﾞﾛｯｸID
        sql = sql & ")"
        sql = sql & "VALUES (" & vbLf

        ' 1:ﾌﾞﾛｯｸID・結晶番号
        If .CRYNUMC2 <> "" And Left(.CRYNUMC2, 1) <> vbNullChar Then
            sql = sql & " '" & .CRYNUMC2 & "'" & vbLf
        Else
            sql = sql & " '" & Space(12) & "'" & vbLf
        End If

        ' 2:工程連番取得
        If .KCNTC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.KCNTC2)) & "'" & vbLf
        Else
            sql = sql & ",1" & vbLf
        End If

        ' 3:結晶番号
        If .XTALC2 <> "" And Left(.XTALC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .XTALC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        ' 4:結晶内開始位置
        If .INPOSC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.INPOSC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If
        
        ' 5:最終通過管理工程
        If .NEKKNTC2 <> "" And Left(.NEKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEKKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 6:最終通過工程
        If .NEWKNTC2 <> "" And Left(.NEWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        ' 7:最終通過作業区分
        If .NEWKKBC2 <> "" And Left(.NEWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NEWKKBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        ' 8:最終通過処理回数
        If .NEMACOC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.NEMACOC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        ' 9:現在管理工程
        If .GNKKNTC2 <> "" And Left(.GNKKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNKKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '10:現在工程
        If .GNWKNTC2 <> "" And Left(.GNWKNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKNTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '11:現在作業区分
        If .GNWKKBC2 <> "" And Left(.GNWKKBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNWKKBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '12:現在処理回数
        If .GNMACOC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMACOC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '13:現在処理日付(登録日時)
        sql = sql & "," & nowtime_sql & vbLf
        
        '14:現在長さ
        If .GNLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '15:現在重量
        If .GNWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.GNWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '16:現在枚数
        If .GNMC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.GNMC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '17:SUMMIT長さ
        If .SUMITLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '18:SUMMIT重量
        If .SUMITWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.SUMITWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '19:SUMMIT枚数
        If .SUMITMC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.SUMITMC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '20:ﾁｬｰｼﾞ量
        If .CHGC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.CHGC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '21:加工区分
        If .KAKOUBC2 <> "" And Left(.KAKOUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KAKOUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '22:計上日付
        If .KEIDAYC2 <> "" Then
            sql = sql & ",TO_DATE('" & Format$(CDate(.KEIDAYC2), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If

        '23:棚区分
        If .GNTKUBC2 <> "" And Left(.GNTKUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTKUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '24:棚番号
        If .GNTNOC2 <> "" And Left(.GNTNOC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .GNTNOC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(4) & "'" & vbLf
        End If

        '25:製造工場
        sql = sql & ",'" & FACTORYCD & "'" & vbLf

        '26:ｳｪｰﾊ製造
        If .WFWORKC2 <> "" And Left(.WFWORKC2, 1) <> vbNullChar Then
            'sql = sql & ",'" & .WFWORKC2 & "'"
            sql = sql & ",'" & .XTWORKC2 & "'" & vbLf   '既存通りに
        Else
            sql = sql & ",'" & Space(2) & "'" & vbLf
        End If

        '27:最終状態区分
        If .LSTATBC2 <> "" And Left(.LSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LSTATBC2 & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '28:流動状態区分
        If .RSTATBC2 <> "" And Left(.RSTATBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .RSTATBC2 & "'" & vbLf
        Else
            sql = sql & ",'T'" & vbLf
        End If

        '29:格上ｺｰﾄﾞ
        If .LUFRCC2 <> "" And Left(.LUFRCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '30:格上区分
        If .LUFRBC2 <> "" And Left(.LUFRBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LUFRBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '31:格下ｺｰﾄﾞ
        If .LDFRCC2 <> "" And Left(.LDFRCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '32:格下区分
        If .LDFRBC2 <> "" And Left(.LDFRBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LDFRBC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '33:ﾎｰﾙﾄﾞｺｰﾄﾞ
        If .HOLDCC2 <> "" And Left(.HOLDCC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDCC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '34:ﾎｰﾙﾄﾞ区分
        If .HOLDBC2 <> "" And Left(.HOLDBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDBC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '35:例外区分
        If .EXKUBC2 <> "" And Left(.EXKUBC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .EXKUBC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '36:返品区分
        If .HENPKC2 <> "" And Left(.HENPKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HENPKC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(1) & "'" & vbLf
        End If

        '37:生死区分
        If .LIVKC2 <> "" And Left(.LIVKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .LIVKC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '38:完了区分
        If .KANKC2 <> "" And Left(.KANKC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KANKC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '39:入庫区分
        If .NFC2 <> "" And Left(.NFC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .NFC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '40:削除区分
        If .SAKJC2 <> "" And Left(.SAKJC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .SAKJC2 & "'" & vbLf
        Else
            sql = sql & ",'0'" & vbLf
        End If

        '41:登録日付
        sql = sql & "," & nowtime_sql & vbLf
        
        '42:更新日付
        sql = sql & "," & nowtime_sql & vbLf

        '43:SUMMIT送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '44:送信ﾌﾗｸﾞ
        sql = sql & ",'0'" & vbLf

        '45:送信日付
        sql = sql & ",NULL" & vbLf

        '46:優先度
        If .PRIORITYC2 <> "" And Left(.PRIORITYC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .PRIORITYC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '47:新規／再切区分
        If .CUTCNTC2 <> "" And Left(.CUTCNTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .CUTCNTC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '48:ﾎｰﾙﾄﾞ工程
        If .HOLDKTC2 <> "" And Left(.HOLDKTC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .HOLDKTC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(5) & "'" & vbLf
        End If

        '49:親ﾌﾞﾛｯｸID
        If .RPCRYNUMC2 <> "" And Left(.RPCRYNUMC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .RPCRYNUMC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(12) & "'" & vbLf
        End If

        '50:不良理由
        If .BDCAUSC2 <> "" And Left(.BDCAUSC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .BDCAUSC2 & "'" & vbLf
        Else
            sql = sql & ",'" & Space(3) & "'" & vbLf
        End If

        '51:実長さ
        If .REALLC2 <> "" Then
            sql = sql & ",'" & CStr(CInt(.REALLC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '52:実重量
        If .REALWC2 <> "" Then
            sql = sql & ",'" & CStr(CLng(.REALWC2)) & "'" & vbLf
        Else
            sql = sql & ",0" & vbLf
        End If

        '53:関連ﾌﾞﾛｯｸﾌﾗｸﾞ
        If .KBLKFLGC2 <> "" And Left(.KBLKFLGC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .KBLKFLGC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '54:向先
        If .PLANTCATC2 <> "" And Left(.PLANTCATC2, 2) <> vbNullChar Then
            sql = sql & ",'" & .PLANTCATC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If

        '55:STCﾌﾞﾛｯｸID
        If .STCIDC2 <> "" And Left(.STCIDC2, 1) <> vbNullChar Then
            sql = sql & ",'" & .STCIDC2 & "'" & vbLf
        Else
            sql = sql & ",NULL" & vbLf
        End If
        
        sql = sql & ")" & vbLf
    
        'SQLを実行
        If OraDB.ExecuteSQL(sql) < 1 Then
            GoTo proc_err
        End If
        
    End With
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/22 SETsw kubota ------------------

    CreateXSDC2 = FUNCTION_RETURN_SUCCESS

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
    CreateXSDC2 = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function


'概要      :現在処理回数を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sGenKotei  ,I  ,String           ,現在工程
'      　　:戻り値       ,O  ,Integer        　,処理回数
Public Function GetGNMACOC(p_sCrynum As String, p_sGenKotei As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
''    sql = "SELECT COUNT(WKKTC3) FROM XSDC3 WHERE WKKTC3 = '" & p_sGenKotei & "'"
'    sql = "SELECT COUNT(DISTINCT(WKKTC3)) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
'    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "'"
''    sql = sql & "' AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
''    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "')"
'    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
'    GetGNMACOC = rs.Fields("COUNT(DISTINCT(WKKTC3))") + 1
    
    sql = "SELECT MACOC3 FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei
    sql = sql & "' AND KCNTC3 = (SELECT MAX(KCNTC3) FROM XSDC3 WHERE CRYNUMC3 = '" & p_sCrynum
    sql = sql & "' AND WKKTC3 = '" & p_sGenKotei & "')"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    If IsNull(rs.Fields("MACOC3")) Then
        GetGNMACOC = 1
    Else
        GetGNMACOC = CInt(rs.Fields("MACOC3")) + 1
    End If
    
End Function


'概要      :最終通過処理回数を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sCrynum    ,I  ,String           ,ブロックID
'      　　:p_iInpos     ,I  ,Integer          ,開始位置
'      　　:戻り値       ,O  ,Integer        　,処理回数
'          :作成者　　2002/11/21　tuku
Public Function GetNEMACOC2(p_sCrynum As String) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    If Left(p_sCrynum, 1) = vbNullChar Then
        GetNEMACOC2 = 1
        Exit Function
    End If
    
    sql = "SELECT GNMACOC2 FROM XSDC2 WHERE CRYNUMC2 = '" & p_sCrynum & "'"
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetNEMACOC2 = 1
    Else
        GetNEMACOC2 = CInt(rs.Fields("GNMACOC2"))
    End If

End Function

