Attribute VB_Name = "s_XSDCZ_SQL"
'@'s_XSDCZ_SQL.bas              '( 06/04/14 ) SMP-OKAMOTO 新規追加
''切断指示 (XSDCZ) ｱｸｾｽ関数

''***テーブル「XSDCZ」へのデータアクセス関数***
'※注意 ﾊﾟﾗﾒｰﾀに値をｾｯﾄする時、まず全て初期化すること

Option Explicit

''XSDCZ用構造体
''配列の場合、データ格納はインデックス=1から始めること。
Public Type typ_XSDCZ
    CRYNUMCZ    As String       ''ブロックID･結晶番号
    HINBCZ      As String       ''品番
    INPOSCZ     As String       ''結晶内開始位置
    REVNUMCZ    As String       ''製品番号改訂番号
    FACTORYCZ   As String       ''工場
    OPECZ       As String       ''操業条件
    KCKNTCZ     As String       ''工程連番
    SXLIDCZ     As String       ''SXLID
    XTALCZ      As String       ''結晶番号
    NEKKNTCZ    As String       ''最終通過管理工程
    NEWKNTCZ    As String       ''最終通過工程
    NEWKKBCZ    As String       ''最終通過作業区分
    NEMACOCZ    As String       ''最終通過処理回数
    GNKKNTCZ    As String       ''現在管理工程
    GNWKNTCZ    As String       ''現在工程
    GNWKKBCZ    As String       ''現在作業区分
    GNMACOCZ    As String       ''現在処理回数
    GNDAYCZ     As String       ''現在処理日付
    GNLCZ       As String       ''現在長さ
    GNWCZ       As String       ''現在重量
    GNMCZ       As String       ''現在枚数
    SUMITLCZ    As String       ''SUMMIT長さ
    SUMITWCZ    As String       ''SUMMIT重量
    SUMITMCZ    As String       ''SUMMIT枚数
    CHGCZ       As String       ''チャージ量
    KAKOUBCZ    As String       ''加工区分
    KEIDAYCZ    As String       ''計上日付
    GNTKUBCZ    As String       ''棚区分
    GNTNOCZ     As String       ''棚番号
    XTWORKCZ    As String       ''製造工場
    WFWORKCZ    As String       ''ウェーハ製造
    LSTATBCZ    As String       ''最終状態区分
    RSTATBCZ    As String       ''流動状態区分
    LUFRCCZ     As String       ''格上ｺｰﾄﾞ
    LUFRBCZ     As String       ''格上区分
    LDFRCCZ     As String       ''格下ｺｰﾄﾞ
    LDFRBCZ     As String       ''格下区分
    HOLDCCZ     As String       ''ﾎｰﾙﾄﾞｺｰﾄﾞ
    HOLDBCZ     As String       ''ホールド区分
    EXKUBCZ     As String       ''例外区分
    HENPKCZ     As String       ''返品区分
    LIVKCZ      As String       ''生死区分
    KANKCZ      As String       ''完了区分
    NFCZ        As String       ''入庫区分
    SAKJCZ      As String       ''削除区分
    TDAYCZ      As String       ''登録日付
    KDAYCZ      As String       ''更新日付
    SUMITBCZ    As String       ''SUMMIT送信フラグ
    SNDKCZ      As String       ''送信フラグ
    SNDDAYCZ    As String       ''送信日付
    LBLFLGCZ    As String       ''ラベル出力確認フラグ
    CUTCNTCZ    As String * 1   ''切断処理区分      '1:再切
    HINBFLGCZ   As String * 1   ''代表品番フラグ    '1：代表品番
    WFHOLDFLGCZ As String       ''ホールド区分(WF)
    HOLDKTCZ    As String * 5   ''ホールド工程
    RPCRYNUMCZ  As String * 12  ''親ブロックID
    FCODECZ     As String       ''不良コード
    SGNKCZ      As String       ''精製原料区分
    CUTKCZ      As String       ''切断区分
    STOPFLG     As String       ''長さ変更フラグ
    PLANTCATCZ  As String       ''向先　2007/08/15 SPK Tsutsumi
End Type

'●SELECT●

'概要      :テーブル「XSDCZ」から条件にあったレコードを抽出する
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO   ,型               ,説明
'          :records()     ,O    ,typ_XSDCZ        ,抽出レコード
'          :lsSqlWhere    ,I    ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :lsSqlOrder    ,I    ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :戻り値        ,O    ,FUNCTION_RETURN   ,抽出の成否
'説明      :

Public Function DBDRV_GetXSDCZ(records() As typ_XSDCZ, _
                                Optional lsSqlWhere$ = vbNullString, _
                                Optional lsSqlOrder$ = vbNullString) As FUNCTION_RETURN
    
    Dim lsSql       As String           ''SQL全体
    Dim lsSqlBase   As String           ''SQL基本部(WHERE節の前まで)
    Dim rs          As OraDynaset       ''RecordSet
    Dim recCnt      As Long             ''レコード数
    Dim i           As Long             ''カウンタ

    ''SQLを組み立てる
    lsSqlBase = "Select * From XSDCZ"
    lsSql = lsSqlBase
    If (lsSqlWhere <> vbNullString) Or (lsSqlOrder <> vbNullString) Then
        lsSql = lsSql & " " & lsSqlWhere & " " & lsSqlOrder
    End If

    ''データを抽出する
    Set rs = OraDB.DBCreateDynaset(lsSql, ORADYN_DEFAULT)
    If rs Is Nothing Then
        ReDim records(0)
        DBDRV_GetXSDCZ = FUNCTION_RETURN_FAILURE
        Exit Function
    End If

    ''抽出結果を格納する
    recCnt = rs.RecordCount
    ReDim records(recCnt)
    If recCnt = 0 Then
        Exit Function
    End If
    ReDim records(recCnt)
    For i = 1 To recCnt
        With records(i)
            ''ブロックID･結晶番号
            If IsNull(rs.Fields("CRYNUMCZ")) = False Then .CRYNUMCZ = CStr(rs.Fields("CRYNUMCZ"))
            ''品番
            If IsNull(rs.Fields("HINBCZ")) = False Then .HINBCZ = CStr(rs.Fields("HINBCZ"))
            ''結晶内開始位置
            If IsNull(rs.Fields("INPOSCZ")) = False Then .INPOSCZ = CStr(rs.Fields("INPOSCZ"))
            ''製品番号改訂番号
            If IsNull(rs.Fields("REVNUMCZ")) = False Then .REVNUMCZ = CStr(rs.Fields("REVNUMCZ"))
            ''工場
            If IsNull(rs.Fields("FACTORYCZ")) = False Then .FACTORYCZ = CStr(rs.Fields("FACTORYCZ"))
            ''操業条件
            If IsNull(rs.Fields("OPECZ")) = False Then .OPECZ = CStr(rs.Fields("OPECZ"))
            ''工程連番
            If IsNull(rs.Fields("KCKNTCZ")) = False Then .KCKNTCZ = CStr(rs.Fields("KCKNTCZ"))
            ''SXLID
            If IsNull(rs.Fields("SXLIDCZ")) = False Then .SXLIDCZ = CStr(rs.Fields("SXLIDCZ"))
            ''結晶番号
            If IsNull(rs.Fields("XTALCZ")) = False Then .XTALCZ = CStr(rs.Fields("XTALCZ"))
            ''最終通過管理工程
            If IsNull(rs.Fields("NEKKNTCZ")) = False Then .NEKKNTCZ = CStr(rs.Fields("NEKKNTCZ"))
            ''最終通過工程
            If IsNull(rs.Fields("NEWKNTCZ")) = False Then .NEWKNTCZ = CStr(rs.Fields("NEWKNTCZ"))
            ''最終通過作業区分
            If IsNull(rs.Fields("NEWKKBCZ")) = False Then .NEWKKBCZ = CStr(rs.Fields("NEWKKBCZ"))
            ''最終通過処理回数
            If IsNull(rs.Fields("NEMACOCZ")) = False Then .NEMACOCZ = CStr(rs.Fields("NEMACOCZ"))
            ''現在管理工程
            If IsNull(rs.Fields("GNKKNTCZ")) = False Then .GNKKNTCZ = CStr(rs.Fields("GNKKNTCZ"))
            ''現在工程
            If IsNull(rs.Fields("GNWKNTCZ")) = False Then .GNWKNTCZ = CStr(rs.Fields("GNWKNTCZ"))
            ''現在作業区分
            If IsNull(rs.Fields("GNWKKBCZ")) = False Then .GNWKKBCZ = CStr(rs.Fields("GNWKKBCZ"))
            ''現在処理回数
            If IsNull(rs.Fields("GNMACOCZ")) = False Then .GNMACOCZ = CStr(rs.Fields("GNMACOCZ"))
            ''現在処理日付
            If IsNull(rs.Fields("GNDAYCZ")) = False Then .GNDAYCZ = Format(CStr(rs.Fields("GNDAYCZ")), "yyyy/mm/dd hh:mm")
            ''現在長さ
            If IsNull(rs.Fields("GNLCZ")) = False Then .GNLCZ = CStr(rs.Fields("GNLCZ"))
            ''現在重量
            If IsNull(rs.Fields("GNWCZ")) = False Then .GNWCZ = CStr(rs.Fields("GNWCZ"))
            ''現在枚数
            If IsNull(rs.Fields("GNMCZ")) = False Then .GNMCZ = CStr(rs.Fields("GNMCZ"))
            ''SUMMIT長さ
            If IsNull(rs.Fields("SUMITLCZ")) = False Then .SUMITLCZ = CStr(rs.Fields("SUMITLCZ"))
            ''SUMMIT重量
            If IsNull(rs.Fields("SUMITWCZ")) = False Then .SUMITWCZ = CStr(rs.Fields("SUMITWCZ"))
            ''SUMMIT枚数
            If IsNull(rs.Fields("SUMITMCZ")) = False Then .SUMITMCZ = CStr(rs.Fields("SUMITMCZ"))
            ''チャージ量
            If IsNull(rs.Fields("CHGCZ")) = False Then .CHGCZ = CStr(rs.Fields("CHGCZ"))
            ''加工区分
            If IsNull(rs.Fields("KAKOUBCZ")) = False Then .KAKOUBCZ = CStr(rs.Fields("KAKOUBCZ"))
            ''計上日付
            If IsNull(rs.Fields("KEIDAYCZ")) = False Then .KEIDAYCZ = CStr(rs.Fields("KEIDAYCZ"))
            ''棚区分
            If IsNull(rs.Fields("GNTKUBCZ")) = False Then .GNTKUBCZ = CStr(rs.Fields("GNTKUBCZ"))
            ''棚番号
            If IsNull(rs.Fields("GNTNOCZ")) = False Then .GNTNOCZ = CStr(rs.Fields("GNTNOCZ"))
            ''製造工場
            If IsNull(rs.Fields("XTWORKCZ")) = False Then .XTWORKCZ = CStr(rs.Fields("XTWORKCZ"))
            ''ウェーハ製造
            If IsNull(rs.Fields("WFWORKCZ")) = False Then .WFWORKCZ = CStr(rs.Fields("WFWORKCZ"))
            ''最終状態区分
            If IsNull(rs.Fields("LSTATBCZ")) = False Then .LSTATBCZ = CStr(rs.Fields("LSTATBCZ"))
            ''流動状態区分
            If IsNull(rs.Fields("RSTATBCZ")) = False Then .RSTATBCZ = CStr(rs.Fields("RSTATBCZ"))
            ''格上ｺｰﾄﾞ
            If IsNull(rs.Fields("LUFRCCZ")) = False Then .LUFRCCZ = CStr(rs.Fields("LUFRCCZ"))
            ''格上区分
            If IsNull(rs.Fields("LUFRBCZ")) = False Then .LUFRBCZ = CStr(rs.Fields("LUFRBCZ"))
            ''格上ｺｰﾄﾞ
            If IsNull(rs.Fields("LDFRCCZ")) = False Then .LDFRCCZ = CStr(rs.Fields("LDFRCCZ"))
            ''格上区分
            If IsNull(rs.Fields("LDFRBCZ")) = False Then .LDFRBCZ = CStr(rs.Fields("LDFRBCZ"))
            ''ﾎｰﾙﾄﾞｺｰﾄﾞ
            If IsNull(rs.Fields("HOLDCCZ")) = False Then .HOLDCCZ = CStr(rs.Fields("HOLDCCZ"))
            ''ホールド区分
            If IsNull(rs.Fields("HOLDBCZ")) = False Then .HOLDBCZ = CStr(rs.Fields("HOLDBCZ"))
            ''例外区分
            If IsNull(rs.Fields("EXKUBCZ")) = False Then .EXKUBCZ = CStr(rs.Fields("EXKUBCZ"))
            ''返品区分
            If IsNull(rs.Fields("HENPKCZ")) = False Then .HENPKCZ = CStr(rs.Fields("HENPKCZ"))
            ''生死区分
            If IsNull(rs.Fields("LIVKCZ")) = False Then .LIVKCZ = CStr(rs.Fields("LIVKCZ"))
            ''完了区分
            If IsNull(rs.Fields("KANKCZ")) = False Then .KANKCZ = CStr(rs.Fields("KANKCZ"))
            ''入庫区分
            If IsNull(rs.Fields("NFCZ")) = False Then .NFCZ = CStr(rs.Fields("NFCZ"))
            ''削除区分
            If IsNull(rs.Fields("SAKJCZ")) = False Then .SAKJCZ = CStr(rs.Fields("SAKJCZ"))
            ''登録日付
            If IsNull(rs.Fields("TDAYCZ")) = False Then .TDAYCZ = Format(CStr(rs.Fields("TDAYCZ")), "yyyy/mm/dd hh:mm")
            ''更新日付
            If IsNull(rs.Fields("KDAYCZ")) = False Then .KDAYCZ = Format(CStr(rs.Fields("KDAYCZ")), "yyyy/mm/dd hh:mm")
            ''SUMMIT送信フラグ
            If IsNull(rs.Fields("SUMITBCZ")) = False Then .SUMITBCZ = CStr(rs.Fields("SUMITBCZ"))
            ''送信フラグ
            If IsNull(rs.Fields("SNDKCZ")) = False Then .SNDKCZ = CStr(rs.Fields("SNDKCZ"))
            ''送信日付
            If IsNull(rs.Fields("SNDDAYCZ")) = False Then .SNDDAYCZ = Format(CStr(rs.Fields("SNDDAYCZ")), "yyyy/mm/dd hh:mm")
            ''ラベル出力確認フラグ
            If IsNull(rs.Fields("LBLFLGCZ")) = False Then .LBLFLGCZ = CStr(rs.Fields("LBLFLGCZ"))
            ''切断処理区分
            If IsNull(rs.Fields("CUTCNTCZ")) = False Then .CUTCNTCZ = CStr(rs.Fields("CUTCNTCZ"))
            ''代表品番
            If IsNull(rs.Fields("HINBFLGCZ")) = False Then .HINBFLGCZ = CStr(rs.Fields("HINBFLGCZ"))
            ''ホールド区分(WF)
            If IsNull(rs.Fields("WFHOLDFLGCZ")) = False Then .WFHOLDFLGCZ = CStr(rs.Fields("WFHOLDFLGCZ"))
            ''ホールド工程
            If IsNull(rs.Fields("HOLDKTCZ")) = False Then .HOLDKTCZ = CStr(rs.Fields("HOLDKTCZ"))
            ''親ブロックID
            If IsNull(rs.Fields("RPCRYNUMCZ")) = False Then .RPCRYNUMCZ = CStr(rs.Fields("RPCRYNUMCZ"))
            ''不良コード
            If IsNull(rs.Fields("FCODECZ")) = False Then .FCODECZ = CStr(rs.Fields("FCODECZ"))
            ''精製原料区分
            If IsNull(rs.Fields("SGNKCZ")) = False Then .SGNKCZ = CStr(rs.Fields("SGNKCZ"))
            ''切断区分
            If IsNull(rs.Fields("CUTKCZ")) = False Then .CUTKCZ = CStr(rs.Fields("CUTKCZ"))
            ''向先 2007/08/17 SPK Tsutsumi Add Start
            If IsNull(rs.Fields("PLANTCATCZ")) = False Then .PLANTCATCZ = CStr(rs.Fields("PLANTCATCZ"))
            ''向先 2007/08/17 SPK Tsutsumi Add End
            End With
        rs.MoveNext
    Next
    rs.Close

    DBDRV_GetXSDCZ = FUNCTION_RETURN_SUCCESS
End Function

'●UPDATE●

'●更新項目を構造体にセットして引き渡す

'概要      :テーブル「XSDCZ」を更新する ptrn1
'ﾊﾟﾗﾒｰﾀ    :変数名        ,IO  ,型               ,説明
'          :records()     ,O   ,typ_XSDCZ        ,更新レコード
'          :lsSqlWhere    ,I   ,String           ,抽出条件(SQLのWhere節:省略可能)
'          :lsSqlOrder    ,I   ,String           ,抽出順序(SQLのOrder by節:省略可能)
'          :lsUpdate      ,I   ,String           ,更新箇所設定(省略可能)
'          :戻り値        ,O   ,FUNCTION_RETURN  ,抽出の成否
'説明      :

Public Function UpdateXSDCZ(records As typ_XSDCZ, _
                                Optional lsSqlWhere$ = vbNullString, _
                                Optional lsSqlOrder$ = vbNullString, _
                                Optional lsUpdate$ = vbNullString) As FUNCTION_RETURN
On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function UpdateXSDCZ"

    Dim lsSql       As String       ''SQL全体
'    Dim lsSqlBase   As String       ''SQL基本部(WHERE節の前まで)
'    Dim rs          As OraDynaset   ''RecordSet
    Dim recCnt      As Long         ''レコード数
    Dim nowtime     As Date         ''サーバ時間
    Dim nowtime_sql As String       ''サーバ時間(SQL文)
    
    ''サーバー時間取得
    nowtime = getSvrTime()

'>>>>> .EditをSQL(UPDATE)文に変更　2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    
    With records
        
        ''SQLを組み立てる
        lsSql = "UPDATE XSDCZ SET" & vbLf
        
        ''更新日付
        lsSql = lsSql & " KDAYCZ = " & nowtime_sql & vbLf
        
        ''ブロックID･結晶番号
        If .CRYNUMCZ <> "" And Left(.CRYNUMCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CRYNUMCZ = '" & .CRYNUMCZ & "'" & vbLf
        End If
        ''品番
        If .HINBCZ <> "" And Left(.HINBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HINBCZ = '" & .HINBCZ & "'" & vbLf
        End If
        ''結晶内開始位置
        If .INPOSCZ <> "" Then
            lsSql = lsSql & ",INPOSCZ = '" & CStr(CInt(.INPOSCZ)) & "'" & vbLf
        End If
        ''製品番号改訂番号
        If .REVNUMCZ <> "" Then
            lsSql = lsSql & ",REVNUMCZ = '" & CStr(CInt(.REVNUMCZ)) & "'" & vbLf
        End If
        ''工場
        If .FACTORYCZ <> "" And Left(.FACTORYCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",FACTORYCZ = '" & .FACTORYCZ & "'" & vbLf
        End If
        ''操業条件
        If .OPECZ <> "" And Left(.OPECZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",OPECZ = '" & .OPECZ & "'" & vbLf
        End If
        ''↓工程連番は登録しない
'        ''工程連番
'        If .KCKNTCZ <> "" Then
'            lsSql = lsSql & ",KCKNTCZ = '" & CStr(CInt(.KCKNTCZ)) & "'" & vbLf
'        End If
        ''↑工程連番は登録しない
        ''SXLID
        If .SXLIDCZ <> "" And Left(.SXLIDCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SXLIDCZ = '" & .SXLIDCZ & "'" & vbLf
        End If
        ''結晶番号
        If .XTALCZ <> "" And Left(.XTALCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",XTALCZ = '" & .XTALCZ & "'" & vbLf
        End If
        ''最終通過管理工程
        If .NEKKNTCZ <> "" And Left(.NEKKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEKKNTCZ = '" & .NEKKNTCZ & "'" & vbLf
        End If
        ''最終通過工程
        If .NEWKNTCZ <> "" And Left(.NEWKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEWKNTCZ = '" & .NEWKNTCZ & "'" & vbLf
        End If
        ''最終通過作業区分
        If .NEWKKBCZ <> "" And Left(.NEWKKBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NEWKKBCZ = '" & .NEWKKBCZ & "'" & vbLf
        End If
        ''最終通過処理回数
        If .NEMACOCZ <> "" Then
            lsSql = lsSql & ",NEMACOCZ = '" & CStr(CInt(.NEMACOCZ)) & "'" & vbLf
        End If
        ''現在管理工程
        If .GNKKNTCZ <> "" And Left(.GNKKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNKKNTCZ = '" & .GNKKNTCZ & "'" & vbLf
        End If
        ''現在工程
        If .GNWKNTCZ <> "" And Left(.GNWKNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNWKNTCZ = '" & .GNWKNTCZ & "'" & vbLf
        End If
        ''現在作業区分
        If .GNWKKBCZ <> "" And Left(.GNWKKBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNWKKBCZ = '" & .GNWKKBCZ & "'" & vbLf
        End If
        ''現在処理回数
        If .GNMACOCZ <> "" Then
            lsSql = lsSql & ",GNMACOCZ = '" & CStr(CInt(.GNMACOCZ)) & "'" & vbLf
        End If
        ''現在処理日付
        If lsUpdate = "NEW" Then    ''XSDCAは新規登録だった場合
            lsSql = lsSql & ",GNDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCAも更新だった場合
            If .GNDAYCZ <> "" And Left(.GNDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",GNDAYCZ = TO_DATE('" & Format$(CDate(.GNDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''現在長さ
        If .GNLCZ <> "" Then
            lsSql = lsSql & ",GNLCZ = '" & CStr(CInt(.GNLCZ)) & "'" & vbLf
        End If
        ''現在重量
        If .GNWCZ <> "" Then
            lsSql = lsSql & ",GNWCZ = '" & CStr(CLng(.GNWCZ)) & "'" & vbLf
        End If
        ''現在枚数
        If .GNMCZ <> "" Then
            lsSql = lsSql & ",GNMCZ = '" & CStr(CInt(.GNMCZ)) & "'" & vbLf
        End If
        ''SUMMIT長さ
        If .SUMITLCZ <> "" Then
            lsSql = lsSql & ",SUMITLCZ = '" & CStr(CInt(.SUMITLCZ)) & "'" & vbLf
        End If
        ''SUMMIT重量
        If .SUMITWCZ <> "" Then
            lsSql = lsSql & ",SUMITWCZ = '" & CStr(CLng(.SUMITWCZ)) & "'" & vbLf
        End If
        ''SUMMIT枚数
        If .SUMITMCZ <> "" Then
            lsSql = lsSql & ",SUMITMCZ = '" & CStr(CInt(.SUMITMCZ)) & "'" & vbLf
        End If
        ''チャージ量
        If .CHGCZ <> "" Then
            lsSql = lsSql & ",CHGCZ = '" & CStr(CLng(.CHGCZ)) & "'" & vbLf
        End If
        ''加工区分
        If .KAKOUBCZ <> "" And Left(.KAKOUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",KAKOUBCZ = '" & .KAKOUBCZ & "'" & vbLf
        End If
        ''計上日付
        If .KEIDAYCZ <> "" Then
            lsSql = lsSql & ",KEIDAYCZ = TO_DATE('" & Format$(CDate(.KEIDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
        End If
        ''棚区分
        If .GNTKUBCZ <> "" And Left(.GNTKUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNTKUBCZ = '" & .GNTKUBCZ & "'" & vbLf
        End If
        ''棚区分
        If .GNTNOCZ <> "" And Left(.GNTNOCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",GNTNOCZ = '" & .GNTNOCZ & "'" & vbLf
        End If
        ''製造工場
        If .XTWORKCZ <> "" And Left(.XTWORKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",XTWORKCZ = '" & .XTWORKCZ & "'" & vbLf
        End If
        ''ウェーハ製造
        If .WFWORKCZ <> "" And Left(.WFWORKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",WFWORKCZ = '" & .WFWORKCZ & "'" & vbLf
        End If
        ''最終状態区分
        If .LSTATBCZ <> "" And Left(.LSTATBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LSTATBCZ = '" & .LSTATBCZ & "'" & vbLf
        End If
        ''流動状態区分
        If .RSTATBCZ <> "" And Left(.RSTATBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",RSTATBCZ = '" & .RSTATBCZ & "'" & vbLf
        End If
        ''格上ｺｰﾄﾞ
        If .LUFRCCZ <> "" And Left(.LUFRCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LUFRCCZ = '" & .LUFRCCZ & "'" & vbLf
        End If
        ''格上区分
        If .LUFRBCZ <> "" And Left(.LUFRBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LUFRBCZ = '" & .LUFRBCZ & "'" & vbLf
        End If
        ''格下ｺｰﾄﾞ
        If .LDFRCCZ <> "" And Left(.LDFRCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LDFRCCZ = '" & .LDFRCCZ & "'" & vbLf
        End If
        ''格下区分
        If .LDFRBCZ <> "" And Left(.LDFRBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LDFRBCZ = '" & .LDFRBCZ & "'" & vbLf
        End If
        ''ﾎｰﾙﾄﾞｺｰﾄﾞ
        If .HOLDCCZ <> "" And Left(.HOLDCCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDCCZ = '" & .HOLDCCZ & "'" & vbLf
        End If
        ''ホールド区分
        If .HOLDBCZ <> "" And Left(.HOLDBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDBCZ = '" & .HOLDBCZ & "'" & vbLf
        End If
        ''例外区分
        If .EXKUBCZ <> "" And Left(.EXKUBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",EXKUBCZ = '" & .EXKUBCZ & "'" & vbLf
        End If
        ''返品区分
        If .HENPKCZ <> "" And Left(.HENPKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HENPKCZ = '" & .HENPKCZ & "'" & vbLf
        End If
        ''生死区分
        If .LIVKCZ <> "" And Left(.LIVKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",LIVKCZ = '" & .LIVKCZ & "'" & vbLf
        End If
        ''完了区分
        If .KANKCZ <> "" And Left(.KANKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",KANKCZ = '" & .KANKCZ & "'" & vbLf
        End If
        ''入庫区分
        If .NFCZ <> "" And Left(.NFCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",NFCZ = '" & .NFCZ & "'" & vbLf
        End If
        ''削除区分
        If .SAKJCZ <> "" And Left(.SAKJCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SAKJCZ = '" & .SAKJCZ & "'" & vbLf
        End If
        ''登録日付
        If lsUpdate = "NEW" Then    ''XSDCAは新規登録だった場合
            lsSql = lsSql & ",TDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCAも更新だった場合
            If .TDAYCZ <> "" And Left(.TDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",TDAYCZ = TO_DATE('" & Format$(CDate(.TDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''SUMMIT送信フラグ
        If .SUMITBCZ <> "" And Left(.SUMITBCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SUMITBCZ = '" & .SUMITBCZ & "'" & vbLf
        End If
        ''送信フラグ
        If .SNDKCZ <> "" And Left(.SNDKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SNDKCZ = '" & .SNDKCZ & "'" & vbLf
        End If
        ''送信日付
        If lsUpdate = "NEW" Then    ''XSDCAは新規登録だった場合
            lsSql = lsSql & ",SNDDAYCZ = " & nowtime_sql & vbLf
        Else                        ''XSDCAも更新だった場合
            If .SNDDAYCZ <> "" And Left(.SNDDAYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",SNDDAYCZ = TO_DATE('" & Format$(CDate(.SNDDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
        End If
        ''ラベル出力確認フラグ
        If .LBLFLGCZ <> "" Then
            lsSql = lsSql & ",LBLFLGCZ = '" & .LBLFLGCZ & "'" & vbLf
        End If
        ''切断処理区分
        If .CUTCNTCZ <> "" And Left(.CUTCNTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CUTCNTCZ = '" & .CUTCNTCZ & "'" & vbLf
        End If
        ''代表品番フラグ
        If .HINBFLGCZ <> "" And Left(.HINBFLGCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HINBFLGCZ = '" & .HINBFLGCZ & "'" & vbLf
        End If
        ''ホールド区分(WF)
        If .WFHOLDFLGCZ <> "" And Left(.WFHOLDFLGCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",WFHOLDFLGCZ = '" & .WFHOLDFLGCZ & "'" & vbLf
        End If
        ''ホールド工程
        If .HOLDKTCZ <> "" And Left(.HOLDKTCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",HOLDKTCZ = '" & .HOLDKTCZ & "'" & vbLf
        End If
        ''親ブロックID
        If .RPCRYNUMCZ <> "" And Left(.RPCRYNUMCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",RPCRYNUMCZ = '" & .RPCRYNUMCZ & "'" & vbLf
        End If
        ''不良コード
        If .FCODECZ <> "" And Left(.FCODECZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",FCODECZ = '" & .FCODECZ & "'" & vbLf
        End If
        ''精製原料区分
        If .SGNKCZ <> "" And Left(.SGNKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",SGNKCZ = '" & .SGNKCZ & "'" & vbLf
        End If
        ''切断区分
        If .CUTKCZ <> "" And Left(.CUTKCZ, 1) <> vbNullChar Then
            lsSql = lsSql & ",CUTKCZ = '" & .CUTKCZ & "'" & vbLf
        End If
        ''向先
        If .PLANTCATCZ <> "" And Left(.PLANTCATCZ, 2) <> vbNullChar And Trim(.HINBCZ) <> "Z" And Trim(.HINBCZ) <> "G" Then
            lsSql = lsSql & ",PLANTCATCZ = '" & .PLANTCATCZ & "'" & vbLf
        End If
    
        lsSql = lsSql & " " & lsSqlWhere & vbLf
    
        'SQLを実行
        recCnt = OraDB.ExecuteSQL(lsSql)
        
        '返り値が1以外はエラー
        If recCnt < 0 Then
            GoTo proc_err
        ElseIf recCnt = 0 Then
            '0件更新…エラー(既存通り)
            UpdateXSDCZ = FUNCTION_RETURN_FAILURE
            Exit Function
        ElseIf recCnt > 1 Then
            '複数件更新…エラー(既存は複数SELECTした最初の一件のみ更新)
            UpdateXSDCZ = FUNCTION_RETURN_FAILURE
            Exit Function
        End If
    
    End With
'<<<<< .EditをSQL(UPDATE)文に変更　2009/06/16 SETsw kubota ------------------

    UpdateXSDCZ = FUNCTION_RETURN_SUCCESS
    

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    UpdateXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'●INSERT●

'概要      :テーブル「XSDCZ」にレコードを挿入する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:pXSDCZ 　　  ,I  ,typ_XSDCZ        ,XSDCZ更新用ﾃﾞｰﾀ
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function CreateXSDCZ(pXSDCZ() As typ_XSDCZ, sErrMsg As String) As FUNCTION_RETURN

    Dim lsSql       As String       ''SQL全体
    Dim sDbName     As String       ''ﾃｰﾌﾞﾙ名
'    Dim rs          As OraDynaset   ''RecordSet
    Dim nowtime     As Date         ''サーバ時間
    Dim nowtime_sql As String       ''サーバ時間(SQL文)
    Dim i           As Long         ''カウンタ
    
    ''エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function CreateXSDCZ"
    sErrMsg = ""
    sDbName = "XSDCZ"
    
    ''配列にデータがない場合
    If UBound(pXSDCZ()) < 1 Then
        CreateXSDCZ = FUNCTION_RETURN_FAILURE
        Exit Function
    End If
    
    ''配列1番目の親ブロックIDをキーにXSDCZを削除する
    lsSql = ""
    lsSql = lsSql & " DELETE XSDCZ"
    lsSql = lsSql & " WHERE RPCRYNUMCZ = '" & pXSDCZ(LBound(pXSDCZ()) + 1).RPCRYNUMCZ & "'"
'    lsSql = lsSql & " WHERE RPCRYNUMCZ = '" & pXSDCZ(LBound(pXSDCZ()) + 1).CRYNUMCZ & "'"
    ''SQL実行
    OraDB.ExecuteSQL lsSql
    ''LOG出力
'    WriteDBLog lsSql        'ｺﾒﾝﾄ　07/06/20 ooba

    ''サーバー時間取得
    nowtime = getSvrTime()
    
'>>>>> .AddNewをSQL(INSERT)文に変更　2009/06/16 SETsw kubota ------------------
    nowtime_sql = "TO_DATE('" & Format$(nowtime, "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')"
    For i = LBound(pXSDCZ()) + 1 To UBound(pXSDCZ())
        With pXSDCZ(i)
            
            lsSql = "INSERT INTO XSDCZ ("
            lsSql = lsSql & " CRYNUMCZ"     ' 1:ﾌﾞﾛｯｸID・結晶番号
            lsSql = lsSql & ",HINBCZ"       ' 2:品番
            lsSql = lsSql & ",INPOSCZ"      ' 3:結晶内開始位置
            lsSql = lsSql & ",REVNUMCZ"     ' 4:製品番号改訂番号
            lsSql = lsSql & ",FACTORYCZ"    ' 5:工場
            lsSql = lsSql & ",OPECZ"        ' 6:操業条件
            'lssql = lssql & ",KCKNTCZ"     ' 7:工程連番   工程連番は登録しない(既存のまま)
            lsSql = lsSql & ",SXLIDCZ"      ' 8:SXLID
            lsSql = lsSql & ",XTALCZ"       ' 9:結晶番号
            lsSql = lsSql & ",NEKKNTCZ"     '10:最終通過管理工程
            lsSql = lsSql & ",NEWKNTCZ"     '11:最終通過工程
            lsSql = lsSql & ",NEWKKBCZ"     '12:最終通過作業区分
            lsSql = lsSql & ",NEMACOCZ"     '13:最終通過処理回数
            lsSql = lsSql & ",GNKKNTCZ"     '14:現在管理工程
            lsSql = lsSql & ",GNWKNTCZ"     '15:現在工程
            lsSql = lsSql & ",GNWKKBCZ"     '16:現在作業区分
            lsSql = lsSql & ",GNMACOCZ"     '17:現在処理回数
            lsSql = lsSql & ",GNDAYCZ"      '18:現在処理日付
            lsSql = lsSql & ",GNLCZ"        '19:現在長さ
            lsSql = lsSql & ",GNWCZ"        '20:現在重量
            lsSql = lsSql & ",GNMCZ"        '21:現在枚数
            lsSql = lsSql & ",SUMITLCZ"     '22:SUMMIT長さ
            lsSql = lsSql & ",SUMITWCZ"     '23:SUMMIT重量
            lsSql = lsSql & ",SUMITMCZ"     '24:SUMMIT枚数
            lsSql = lsSql & ",CHGCZ"        '25:ﾁｬｰｼﾞ量
            lsSql = lsSql & ",KAKOUBCZ"     '26:加工区分
            If .KEIDAYCZ <> "" Then
                lsSql = lsSql & ",KEIDAYCZ"     '27:計上日付
            End If
            lsSql = lsSql & ",GNTKUBCZ"     '28:棚区分
            lsSql = lsSql & ",GNTNOCZ"      '29:棚番号
            lsSql = lsSql & ",XTWORKCZ"     '30:製造工場
            lsSql = lsSql & ",WFWORKCZ"     '31:ｳｪｰﾊ製造
            lsSql = lsSql & ",LSTATBCZ"     '32:最終状態区分
            lsSql = lsSql & ",RSTATBCZ"     '33:流動状態区分
            lsSql = lsSql & ",LUFRCCZ"      '34:格上ｺｰﾄﾞ
            lsSql = lsSql & ",LUFRBCZ"      '35:格上区分
            lsSql = lsSql & ",LDFRCCZ"      '36:格下ｺｰﾄﾞ
            lsSql = lsSql & ",LDFRBCZ"      '37:格下区分
            lsSql = lsSql & ",HOLDCCZ"      '38:ﾎｰﾙﾄﾞｺｰﾄﾞ
            lsSql = lsSql & ",HOLDBCZ"      '39:ﾎｰﾙﾄﾞ区分
            lsSql = lsSql & ",EXKUBCZ"      '40:例外区分
            lsSql = lsSql & ",HENPKCZ"      '41:返品区分
            lsSql = lsSql & ",LIVKCZ"       '42:生死区分
            lsSql = lsSql & ",KANKCZ"       '43:完了区分
            lsSql = lsSql & ",NFCZ"         '44:入庫区分
            lsSql = lsSql & ",SAKJCZ"       '45:削除区分
            lsSql = lsSql & ",TDAYCZ"       '46:登録日付
            lsSql = lsSql & ",KDAYCZ"       '47:更新日付
            lsSql = lsSql & ",SUMITBCZ"     '48:SUMMIT送信ﾌﾗｸﾞ
            lsSql = lsSql & ",SNDKCZ"       '49:送信ﾌﾗｸﾞ
            lsSql = lsSql & ",SNDDAYCZ"     '50:送信日付
            lsSql = lsSql & ",LBLFLGCZ"     '51:ラベル出力確認フラグ
            lsSql = lsSql & ",CUTCNTCZ"     '52:新規／再切区分
            lsSql = lsSql & ",HINBFLGCZ"    '53:代表品番フラグ
            lsSql = lsSql & ",HOLDKTCZ"     '54:ﾎｰﾙﾄﾞ工程
            lsSql = lsSql & ",RPCRYNUMCZ"   '55:親ﾌﾞﾛｯｸID
            lsSql = lsSql & ",FCODECZ"      '56:不良ｺｰﾄﾞ
            lsSql = lsSql & ",SGNKCZ"       '57:精製原料区分
            lsSql = lsSql & ",CUTKCZ"       '58:切断区分
            lsSql = lsSql & ",PLANTCATCZ"   '59:向先
            lsSql = lsSql & ")"
            lsSql = lsSql & "VALUES (" & vbLf
            
            ' 1:ﾌﾞﾛｯｸID・結晶番号
            If .CRYNUMCZ <> "" And Left(.CRYNUMCZ, 1) <> vbNullChar Then
                lsSql = lsSql & " '" & .CRYNUMCZ & "'" & vbLf
            Else
                lsSql = lsSql & " '" & Space(12) & "'" & vbLf
            End If
            
            ' 2:品番
            If .HINBCZ <> "" And Left(.HINBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HINBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(8) & "'" & vbLf
            End If
            
            ' 3:結晶内開始位置
            If .INPOSCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.INPOSCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            ' 4:製品番号改訂番号
            If .REVNUMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.REVNUMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            ' 5:工場
            If .FACTORYCZ <> "" And Left(.FACTORYCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .FACTORYCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            ' 6:操業条件
            If .OPECZ <> "" And Left(.OPECZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .OPECZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            ' 7:工程連番   工程連番は登録しない(既存のまま)
            'If .KCKNTCZ <> "" Then
            '    lsSql = lsSql & ",'" & CStr(CInt(.KCKNTCZ)) & "'" & vbLf
            'Else
            '    lsSql = lsSql & ",0" & vbLf
            'End If
            
            ' 8:SXLID
            If .SXLIDCZ <> "" And Left(.SXLIDCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SXLIDCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(13) & "'" & vbLf
            End If
            
            ' 9:結晶番号
            If .XTALCZ <> "" And Left(.XTALCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .XTALCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(12) & "'" & vbLf
            End If
            
            '10:最終通過管理工程
            If .NEKKNTCZ <> "" And Left(.NEKKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEKKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '11:最終通過工程
            If .NEWKNTCZ <> "" And Left(.NEWKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEWKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '12:最終通過作業区分
            If .NEWKKBCZ <> "" And Left(.NEWKKBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NEWKKBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If
            
            '13:最終通過処理回数
            If .NEMACOCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.NEMACOCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '14:現在管理工程
            If .GNKKNTCZ <> "" And Left(.GNKKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNKKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '15:現在工程
            If .GNWKNTCZ <> "" And Left(.GNWKNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNWKNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If
            
            '16:現在作業区分
            If .GNWKKBCZ <> "" And Left(.GNWKKBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNWKKBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If
            
            '17:現在処理回数
            If .GNMACOCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNMACOCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '18:現在処理日付
            lsSql = lsSql & "," & nowtime_sql & vbLf
            
            '19:現在長さ
            If .GNLCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNLCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '20:現在重量
            If .GNWCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.GNWCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '21:現在枚数
            If .GNMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.GNMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '22:SUMMIT長さ
            If .SUMITLCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.SUMITLCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '23:SUMMIT重量
            If .SUMITWCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.SUMITWCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '24:SUMMIT枚数
            If .SUMITMCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CInt(.SUMITMCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '25:ﾁｬｰｼﾞ量
            If .CHGCZ <> "" Then
                lsSql = lsSql & ",'" & CStr(CLng(.CHGCZ)) & "'" & vbLf
            Else
                lsSql = lsSql & ",0" & vbLf
            End If
            
            '26:加工区分
            If .KAKOUBCZ <> "" And Left(.KAKOUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .KAKOUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If
            
            '27:計上日付
            If .KEIDAYCZ <> "" Then
                lsSql = lsSql & ",TO_DATE('" & Format$(CDate(.KEIDAYCZ), "yyyy/mm/dd hh:nn:ss") & "','YYYY/MM/DD HH24:MI:SS')" & vbLf
            End If
            
            '28:棚区分
            If .GNTKUBCZ <> "" And Left(.GNTKUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNTKUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If
            
            '29:棚番号
            If .GNTNOCZ <> "" And Left(.GNTNOCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .GNTNOCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(4) & "'" & vbLf
            End If
            
            '30:製造工場
            lsSql = lsSql & ",'" & FACTORYCD & "'" & vbLf
            
            '31:ｳｪｰﾊ製造
            If .WFWORKCZ <> "" And Left(.WFWORKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .WFWORKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(2) & "'" & vbLf
            End If

            '32:最終状態区分
            If .LSTATBCZ <> "" And Left(.LSTATBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LSTATBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'T'" & vbLf       '通常
            End If

            '33:流動状態区分
            If .RSTATBCZ <> "" And Left(.RSTATBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .RSTATBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'T'" & vbLf       '通常
            End If

            '34:格上ｺｰﾄﾞ
            If .LUFRCCZ <> "" And Left(.LUFRCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LUFRCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '35:格上区分
            If .LUFRBCZ <> "" And Left(.LUFRBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LUFRBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '36:格下ｺｰﾄﾞ
            If .LDFRCCZ <> "" And Left(.LDFRCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LDFRCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '37:格下区分
            If .LDFRBCZ <> "" And Left(.LDFRBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LDFRBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '38:ﾎｰﾙﾄﾞｺｰﾄﾞ
            If .HOLDCCZ <> "" And Left(.HOLDCCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDCCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '39:ﾎｰﾙﾄﾞ区分
            If .HOLDBCZ <> "" And Left(.HOLDBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '40:例外区分
            If .EXKUBCZ <> "" And Left(.EXKUBCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .EXKUBCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '41:返品区分
            If .HENPKCZ <> "" And Left(.HENPKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HENPKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '42:生死区分
            If .LIVKCZ <> "" And Left(.LIVKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LIVKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '43:完了区分
            If .KANKCZ <> "" And Left(.KANKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .KANKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '44:入庫区分
            If .NFCZ <> "" And Left(.NFCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .NFCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '45:削除区分
            If .SAKJCZ <> "" And Left(.SAKJCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SAKJCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'0'" & vbLf
            End If

            '46:登録日付
            lsSql = lsSql & "," & nowtime_sql & vbLf

            '47:更新日付
            lsSql = lsSql & "," & nowtime_sql & vbLf

            '48:SUMMIT送信ﾌﾗｸﾞ
            lsSql = lsSql & ",'0'" & vbLf

            '49:送信ﾌﾗｸﾞ
            lsSql = lsSql & ",'0'" & vbLf

            '50:送信日付
            lsSql = lsSql & ",NULL" & vbLf

            '51:ラベル出力確認フラグ
            If .LBLFLGCZ <> "" And Left(.LBLFLGCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .LBLFLGCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(1) & "'" & vbLf
            End If

            '52:新規／再切区分
            If .CUTCNTCZ <> "" And Left(.CUTCNTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .CUTCNTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '53:代表品番フラグ
            If .HINBFLGCZ <> "" And Left(.HINBFLGCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HINBFLGCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '54:ﾎｰﾙﾄﾞ工程
            If .HOLDKTCZ <> "" And Left(.HOLDKTCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .HOLDKTCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(5) & "'" & vbLf
            End If

            '55:親ﾌﾞﾛｯｸID
            If .RPCRYNUMCZ <> "" And Left(.RPCRYNUMCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .RPCRYNUMCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(12) & "'" & vbLf
            End If

            '56:不良ｺｰﾄﾞ
            If .FCODECZ <> "" And Left(.FCODECZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .FCODECZ & "'" & vbLf
            Else
                lsSql = lsSql & ",'" & Space(3) & "'" & vbLf
            End If

            '57:精製原料区分
            If .SGNKCZ <> "" And Left(.SGNKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .SGNKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '58:切断区分
            If .CUTKCZ <> "" And Left(.CUTKCZ, 1) <> vbNullChar Then
                lsSql = lsSql & ",'" & .CUTKCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If

            '59:向先
            If .PLANTCATCZ <> "" And Left(.PLANTCATCZ, 2) <> vbNullChar Then
                lsSql = lsSql & ",'" & .PLANTCATCZ & "'" & vbLf
            Else
                lsSql = lsSql & ",NULL" & vbLf
            End If
            
            lsSql = lsSql & ")" & vbLf
        
            'SQLを実行
            If OraDB.ExecuteSQL(lsSql) < 1 Then
                GoTo proc_err
            End If

            ' del SIRD対応 SETkimizuka Start 2010/02/15
            '''XODY3作成 流動監視機能追加に伴う修正  add SETkimizuka Start 09/08/03
            'If Y3Flg = True Then
            '    If pXSDCZ(i).SGNKCZ = "0" Then
            '        Call CreateOrUpdateXODY3(pXSDCZ(i).CRYNUMCZ, pXSDCZ(i).SXLIDCZ, pXSDCZ(i).LIVKCZ, "", "1")
            '    End If
            'End If
            ''XODY3作成 流動監視機能追加に伴う修正  add SETkimizuka End 09/08/03
            ' del SIRD対応 SETkimizuka End 2010/02/15

        End With
    Next i
'<<<<< .AddNewをSQL(INSERT)文に変更　2009/06/16 SETsw kubota ------------------

    CreateXSDCZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    CreateXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

'概要      :最終通過処理回数を取得する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_sCrynum    ,I  ,String           ,ブロックID
'      　　:p_iInpos     ,I  ,Integer          ,開始位置
'      　　:戻り値       ,O  ,Integer        　,処理回数
Public Function GetGNMACOCZ(p_sCrynum As String, p_iInpos As Integer) As Integer
    Dim sql As String
    Dim rs As OraDynaset
    
    
    sql = "SELECT GNMACOCZ FROM XSDCZ WHERE CRYNUMCZ = '" & p_sCrynum
    sql = sql & "' AND INPOSCZ = " & p_iInpos
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        GetGNMACOCZ = 1
    Else
        GetGNMACOCZ = CInt(rs.Fields("GNMACOCZ"))
    End If

End Function

'概要      :該当するﾚｺｰﾄﾞ有無をﾁｪｯｸ
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:p_BlockID    ,I  ,String           ,SXLID
'      　　:p_Hinban     ,O  ,String           ,長さ
'      　　:p_Inpos      ,O  ,Integer          ,長さ
'      　　:戻り値       ,O  ,Boolean        　,ﾚｺｰﾄﾞなし(TRUE)/あり(FALSE)
'説明　　　：品番振替、ｸﾘｽﾀﾙ格上など死ﾚｺｰﾄﾞと同品番への変更に対応
'履歴　　　：2002/08/29 ohno
Public Function CheckUniqueRecordXSDCZ(p_BlockID As String, p_Hinban As String, p_Inpos As Integer) As Boolean
    Dim sql As String
    Dim rs As OraDynaset
    
    sql = "SELECT * FROM XSDCZ WHERE CRYNUMCZ = '" & p_BlockID
    sql = sql & "' AND HINBCZ = '" & p_Hinban
    sql = sql & "' AND INPOSCZ = " & p_Inpos
    
    Set rs = OraDB.DBCreateDynaset(sql, ORADYN_DEFAULT)
    
    If rs.RecordCount = 0 Then
        CheckUniqueRecordXSDCZ = True
    Else
        CheckUniqueRecordXSDCZ = False
    End If
    
End Function

'●DELETE●

'概要      :テーブル「XSDCZ」のレコードを削除する
'ﾊﾟﾗﾒｰﾀ　　:変数名       ,IO ,型               ,説明
'      　　:sCRYNUMCZ 　 ,I  ,String           ,ブロックID
'      　　:sHINBCZ 　   ,I  ,String           ,品番
'      　　:sINPOSCZ 　  ,I  ,String           ,結晶内開始位置
'      　　:sErrMsg　　　,O  ,String         　,エラーメッセージ
'      　　:戻り値       ,O  ,FUNCTION_RETURN　,書き込みの成否
Public Function DeleteXSDCZ(sCRYNUMCZ As String, sHINBCZ As String, sINPOSCZ As String, _
                                sErrMsg As String) As FUNCTION_RETURN
    Dim lsSql       As String       ''SQL全体
    Dim sDbName     As String       ''ﾃｰﾌﾞﾙ名
    
    ''エラーハンドラの設定
    On Error GoTo proc_err
    gErr.Push "s_XSDCZ_SQL.bas -- Function DeleteXSDCZ"
    sErrMsg = ""
    sDbName = "XSDCZ"
    
    DeleteXSDCZ = FUNCTION_RETURN_FAILURE
    
    ''ブロックIDをキーにXSDCZを削除する
    lsSql = ""
    lsSql = lsSql & " DELETE XSDCZ"
    lsSql = lsSql & " WHERE CRYNUMCZ = '" & sCRYNUMCZ & "'"
    lsSql = lsSql & "   AND HINBCZ   = '" & sHINBCZ & "'"
    lsSql = lsSql & "   AND INPOSCZ  = '" & sINPOSCZ & "'"
    ''SQL実行
    OraDB.ExecuteSQL lsSql
    ''LOG出力
'    WriteDBLog lsSql        'ｺﾒﾝﾄ　07/06/20 ooba

    DeleteXSDCZ = FUNCTION_RETURN_SUCCESS

proc_exit:
    '' 終了
    gErr.Pop
    Exit Function

proc_err:
    '' エラーハンドラ
    Debug.Print "====== Error SQL ======"
    Debug.Print lsSql
    gErr.HandleError
    sErrMsg = GetMsgStr("ENG11", "DB", sDbName)
    DeleteXSDCZ = FUNCTION_RETURN_FAILURE
    Resume proc_exit

End Function

